'**Attribute VB_Name = "RekursivesUmbenennen"
'
'*******************************************************************************
' (c) 2007      Moldware GmbH
'*******************************************************************************
' Purpose     : Suchen/Ersetzen und Namenskonsistenz im Produktbaum.
' ActiveDoc   : CATProduct
' Author      : Stefan Urbanus stefan.urbanus@moldware.com)
' Version     : 5.12.1.1 / 10.12.2007
' Assumptions : Hauptprodukt ist geladen.
' Effects     : Instanz-, Produkt- und Dateinamen werden angepasst.
'               Die ganze Produktstruktur wird aktualisiert und gespeichert.
' Languages   : VBA, VBS
' Locales     : Alle
' Platforms   : Alle
' CATIA Level : V5R10
' Details     :
' History     : 2003-11-11  Alex Watzal  Erzeugt.
'*******************************************************************************
Option Explicit

'*******************************************************************************
' Funktionalitaets-Parameter und Makroname
'*******************************************************************************

'Const sstrTitle2 = "Umbenennen im Produktbaum" 'As String
Const sstrTitle2 = "Suchen/Ersetzen im Produktbaum" 'As String
'Const sstrTitle2 = "Namenskonsistenz im Produktbaum" 'As String
Const sblnUmbenennen = True 'As Boolean         'Suchen/Ersetzen wird aktiviert
Const sblnNamensKonsistenz = True 'As Boolean   'Namenskonsistenz wird aktiviert  (Dateiname = Productname)
Const sbln_xSET = True 'As Boolean              'bei Namenskonsistenz wird "_xSET" eliminiert
Const sbln_CATPart = True 'As Boolean           'bei Namenskonsistenz wird ".CATPart" eliminiert
Const sblnChangeFileName = True 'As Boolean     'Dateiname wird auch angepasst
Const sblnEnglish = False 'As Boolean           'Namensbaum auf Englisch
 
'*******************************************************************************
' Ende: Funktionalitaets-Parameter und Makroname
'*******************************************************************************

'*******************************************************************************
' Installations-Parameter
'*******************************************************************************

' Viewer
Const sstrViewerIntelHTML = "iexplore.exe" 'As String
'Const sstrViewerIntelHTML = "C:\Programme\INTERN~1\iexplore.exe" 'As String

Const sstrViewerUnixHTML = "mozilla" 'As String

Const sstrViewerIntelText = "notepad" 'As String
'Const sstrViewerIntelText = "write" 'As String
'Const sstrViewerIntelText = "C:\Programme\Window~1\Zubehoer\wordpad.exe" 'As String

Const sstrViewerUnixText = "dtpad -standAlone" 'As String
'Const sstrViewerUnixText = "xterm -e vi" 'As String
          
'*******************************************************************************
' Ende: Installations-Parameter
'*******************************************************************************
  
'*******************************************************************************
' Debug-Parameter
'*******************************************************************************

' Debug Modus
Const sblnDebug = False 'As Boolean
  
'*******************************************************************************
' Ende: Debug-Parameter
'*******************************************************************************
  
'*******************************************************************************
' Lizenz: Keine zeitliche Beschraenkung
'*******************************************************************************

'*******************************************************************************
'*******************************************************************************
'**Start Encode**
'Ab hier wird das Makro verschluesselt!
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
' Lizenz-Parameter
'*******************************************************************************

' Lizenz, zeitgesteuert
Const sdatExpires = #12/31/2999# 'As Date
  
'*******************************************************************************
' Ende: Lizenz-Parameter
'*******************************************************************************
  
'*******************************************************************************
' Titel fuer Dialogfenster, Makroversion
'*******************************************************************************

' Der Titel setzt sich aus sstrTitel1, sstrTitel2 und sstrTitel3 zusammen
Dim sstrTitle 'As String
Const sstrTitle1 = "Moldware GmbH: " 'As String
Const sstrTitle3 = " V1.12" 'As String

'*******************************************************************************
' Ende: Titel fuer Dialogfenster, Makroversion
'*******************************************************************************



Sub CATMain()
  
  ' Titel fuer Dialogfenster
  sstrTitle = sstrTitle1 & sstrTitle2 & sstrTitle3
  
  ' Debug Modus
  Dim blnDebug 'As Boolean
  blnDebug = sblnDebug
  If (blnDebug) Then CATIA.SystemService.Print "-----> CATMain"
  
  ' Return Code von MsgBox
  Dim intRC 'As Integer
    
  ' Mehrzeiliger String fuer MsgBox
  Dim strMsg 'As String

  If (Date > sdatExpires) Then
    strMsg = "Ihre Lizenz ist am " & FormatDateTime(sdatExpires, vbShortDate) & " abgelaufen!" & vbLf _
            & "Bitte beantragen Sie eine neue Lizenz."
    Call MsgBox(strMsg, vbCritical, sstrTitle)
    Exit Sub
  ElseIf (sdatExpires - Date <= 14) Then
    strMsg = "Ihre Lizenz laeuft am " & FormatDateTime(sdatExpires, vbShortDate) & " ab!" & vbLf _
            & "Bitte beantragen Sie eine neue Lizenz."
    Call MsgBox(strMsg, vbExclamation, sstrTitle)
  End If

  If (blnDebug) Then
    strMsg = "Die Lizenz laeuft am " & FormatDateTime(sdatExpires, vbShortDate) _
            & " also in " & sdatExpires - Date & " Tagen ab."
    CATIA.SystemService.Print strMsg
  End If
  
  ' Error Werte
  Dim lngErrNumber 'As Long
  Dim strErrDescr 'As String
  Dim strErrSource 'As String
  
  Dim objActiveDoc 'As Document
  Set objActiveDoc = CATIA.ActiveDocument
    
  Dim objActiveProdDoc 'As ProductDocument
  Dim objActivePartDoc 'As PartDocument
  Dim objRootProd 'As Product
    
  If (TypeName(objActiveDoc) = "ProductDocument") Then
    Set objActiveProdDoc = objActiveDoc
    Set objRootProd = objActiveProdDoc.Product
  ElseIf (TypeName(objActiveDoc) = "PartDocument") Then
    Set objActivePartDoc = objActiveDoc
    Set objRootProd = objActivePartDoc.Product
  Else
    Call Err.Raise(9999, "CATMain", "Nicht-unterstuetzter Dokumenttyp: " & TypeName(objActiveDoc))
  End If
        
  Dim aobjInst() 'As Product
  ReDim aobjInst(0)
  
  Dim aintStructure() 'As Integer
  ReDim aintStructure(1, 0)
  
  Dim aobjProd() 'As Product
  ReDim aobjProd(0)
  Set aobjProd(0) = objRootProd
  
  Dim aintNbInst() 'As Integer
  ReDim aintNbInst(0)
  aintNbInst(0) = 0

  ' Liste der Instanzen und Produkte rekursiv aufbauen
  Call GetAllInstances(objRootProd, 0, blnDebug, aobjInst, aintStructure, aobjProd, aintNbInst)
    
  Dim intNbInst 'As Integer
  intNbInst = UBound(aobjInst)
    
  Dim intNbProd 'As Integer
  intNbProd = UBound(aobjProd) + 1
    
  Dim strOldSNR 'As String
  Dim intLenOldSNR 'As Integer
  Dim strNewSNR 'As String
  strOldSNR = ""
  intLenOldSNR = 0
  strNewSNR = ""
  
  If (sblnUmbenennen) Then
    ' Alte Sachnummer, die ersetzt werden soll. Leer bedeutet keine Namensersetzung.
    strMsg = "Namensbestandteil, der ersetzt werden soll:" & vbLf & vbLf _
            & "Es kann ein ganzer Name oder auch nur ein Namensbestandteil ersetzt werden," & vbLf _
            & "z.B. eine Sachnummer." 
            
    strOldSNR = InputBox(strMsg, sstrTitle, objRootProd.Name)
    intLenOldSNR = Len(strOldSNR)
    
    ' Neue Sachnummer des Root-Produkts, falls leer wird die alte Sachnummer beibehalten
    If (intLenOldSNR > 0) Then
      strMsg = """" & strOldSNR & """ wird ersetzt durch:" & vbLf & vbLf _
              & "Geben Sie ein, wodurch der vorher definierte Namensbestandteil ersetzt werden soll," & vbLf _
              & "z.B. eine geaenderte Sachnummer." & vbLf _
              & "Ersetzt werden kann nur am Namensanfang."
      strNewSNR = InputBox(strMsg, sstrTitle, strOldSNR)
    End If
    
    If (strNewSNR = "" Or intLenOldSNR = 0 Or strNewSNR = strOldSNR) Then
      strNewSNR = ""
      strOldSNR = ""
      intLenOldSNR = 0
    End If
  End If
  
  If (Not sblnNamensKonsistenz) Then
    If (intLenOldSNR = 0) Then
      If (blnDebug) Then CATIA.SystemService.Print "<----- CATMain"
      Exit Sub
    End If
  End If
  
  ' Anzeige der Produktstruktur als HTML-Datei oder Text-Datei
  Dim blnHTML 'As Boolean
  strMsg = "Anzeigen der Produktstruktur im HTML-Format." & vbLf & vbLf _
          & "Druecken Sie <Ja> fuer das farbige HTML-Format." & vbLf _
          & "Druecken Sie <Nein> fuer das einfache Text-Format."
  intRC = MsgBox(strMsg, vbYesNo, sstrTitle)
  If (intRC = vbYes) Then blnHTML = True Else blnHTML = False
    
  ' Name der temporaeren Datei
  Dim strTmpFile 'As String
  If (blnHTML) Then
    strTmpFile = "ProductStructure.html"
  Else
    strTmpFile = "ProductStructure.txt"
  End If
    
  ' Temporaeres Verzeichnis
  Dim objTmpDir 'As Folder
  Set objTmpDir = CATIA.FileSystem.TemporaryDirectory
    
  ' Kompletter Pfad zur temporaeren Datei
  Dim strTmpPath 'As String
  strTmpPath = CATIA.FileSystem.ConcatenatePaths(objTmpDir.Path, strTmpFile)
  'Call MsgBox(strTmpPath)
    
  ' Temporaere Textdatei oeffnen
  Dim objTmpFile 'As File
  On Error Resume Next
  Set objTmpFile = CATIA.FileSystem.CreateFile(strTmpPath, True)
  lngErrNumber = Err.Number
  On Error GoTo 0
  If (lngErrNumber <> 0) Then
    'Call MsgBox("delete")
    On Error Resume Next
    Call CATIA.FileSystem.DeleteFile(strTmpPath)
    On Error GoTo 0
    
    Set objTmpFile = CATIA.FileSystem.CreateFile(strTmpPath, True)
  End If
    
  ' Flag, dass die Datei spaeter noch geloescht werden muss
  Dim blnDelete 'As Boolean
  blnDelete = True
    
  Dim objTmpStream 'As TextStream
  Set objTmpStream = objTmpFile.OpenAsTextStream("ForWriting")
    
  ' Neue Namen fuer die Products und deren Dateinamen ermitteln
  Dim astrNewProdName() 'As String
  ReDim astrNewProdName(1, intNbProd - 1)
    
  Dim ii 'As Integer
  For ii = 0 To intNbProd - 1
    ' Neuer Produkt-Name
    Dim strName 'As String
    strName = aobjProd(ii).Name
    
    ' Sachnummer ersetzen
'    If (intLenOldSNR > 0 And strOldSNR = Left(strName, intLenOldSNR)) Then
'      strName = strNewSNR & Mid(strName, intLenOldSNR + 1)
'    End If

    If (intLenOldSNR > 0) Then
	   strName = Replace(strName, strOldSNR,  strNewSNR)
    End If
    
    Dim intPos 'As Integer
    
    If (sblnNamensKonsistenz And sbln_xSET) Then
      'Dim intPos2 'As Integer
      
      ' "_xSETnnnn" entfernen
      intPos = InStr(strName, "_xSET")
      
      If (intPos > 0) Then
        'For intPos2 = intPos + 5 To Len(strName)
        '  If (Not IsNumeric(Mid(strName, intPos2, 1))) Then Exit For
        'Next
        'strName = Left(strName, intPos - 1) & Mid(strName, intPos2)
        
        ' aus "_xSETnnnn" den String "xSET" entfernen
        strName = Left(strName, intPos) & Mid(strName, intPos + 5)
      End If
      
      ' "_xxSETnnnn" entfernen
      intPos = InStr(strName, "_xxSET")
      
      If (intPos > 0) Then
        'For intPos2 = intPos + 6 To Len(strName)
        '  If (Not IsNumeric(Mid(strName, intPos2, 1))) Then Exit For
        'Next
        'strName = Left(strName, intPos - 1) & Mid(strName, intPos2)
        
        ' aus "_xxSETnnnn" den String "xSET" entfernen
        strName = Left(strName, intPos + 1) & Mid(strName, intPos + 6)
      End If
    End If
    
    If (strName = aobjProd(ii).Name) Then
      astrNewProdName(0, ii) = ""
    Else
      astrNewProdName(0, ii) = strName
    End If
    
    If (sblnChangeFileName) Then
    
      ' Neuer Datei Name
      Dim strFile 'As String
      
      ' Typ ermitteln
      If (aobjProd(ii).Name = aobjProd(ii).Parent.Product.Name) Then
        If (sblnNamensKonsistenz) Then
          ' Dateiname an Produktname anpassen
          
          ' Part oder Product
          If (TypeName(aobjProd(ii).Parent) = "PartDocument") Then
            strFile = strName & ".CATPart"
          Else
            strFile = strName & ".CATProduct"
          End If
        ElseIf (sblnUmbenennen) Then
          ' Sachnummer in Dateiname ersetzen
          
          strFile = aobjProd(ii).Parent.Name
          
          If (intLenOldSNR > 0 And strOldSNR = Left(strFile, intLenOldSNR)) Then
            strFile = strNewSNR & Mid(strFile, intLenOldSNR + 1)
          End If
        End If
        
        If (strFile <> aobjProd(ii).Parent.Name) Then
          astrNewProdName(1, ii) = strFile
        Else
          astrNewProdName(1, ii) = ""
        End If
      
      ElseIf (aobjProd(ii).HasAMasterShapeRepresentation = False) Then
        ' Component
        astrNewProdName(1, ii) = ""
      Else
        ' cgr oder model
        astrNewProdName(1, ii) = ""
      End If
    Else
      astrNewProdName(1, ii) = ""
    End If
    
  Next
    
  ' Root-Product Name und Dateiname ausgeben
  Call WriteHeader(blnHTML, objRootProd.Name, objTmpStream)
  If (sblnEnglish) Then
    If (TypeName(objActiveDoc) = "PartDocument") Then
      Call WriteLine(blnHTML, "Part:    " & objRootProd.Name, astrNewProdName(0, 0), objTmpStream)
    Else
      Call WriteLine(blnHTML, "Product: " & objRootProd.Name, astrNewProdName(0, 0), objTmpStream)
    End If
    Call WriteLine(blnHTML, "File:    " & objActiveDoc.Name, astrNewProdName(1, 0), objTmpStream)
  Else
    If (TypeName(objActiveDoc) = "PartDocument") Then
      Call WriteLine(blnHTML, "Part:    " & objRootProd.Name, astrNewProdName(0, 0), objTmpStream)
    Else
      Call WriteLine(blnHTML, "Produkt: " & objRootProd.Name, astrNewProdName(0, 0), objTmpStream)
    End If
    Call WriteLine(blnHTML, "Datei:   " & objActiveDoc.Name, astrNewProdName(1, 0), objTmpStream)
  End If
  
  ' Array fuer neue Instanz-Namen
  Dim astrNewInstName() 'As String
  ReDim astrNewInstName(intNbInst)
    
  If (intNbInst > 0) Then
    ' Variable fuer Indentation, enthaelt die grafischen Baumelemente und Blanks
    Dim astrIndent() 'As String
    
    ' Startwerte fuer Iteration durch Baumstruktur
    Dim intDepth 'As Integer
    intDepth = 0
    
    For ii = 0 To intNbInst - 1
      ' Tiefe vom Vorgaenger retten
      Dim intDepthP 'As Integer
      intDepthP = intDepth
      intDepth = aintStructure(0, ii)
                
      Dim blnExpand 'As Boolean
      blnExpand = False
      If (aintStructure(1, ii) >= 2) Then blnExpand = True
            
      Dim blnLast 'As Boolean
      blnLast = False
      If (aintStructure(1, ii) Mod 2 = 1) Then blnLast = True
            
      ' Index zum Product-Array suchen
      Dim jj 'As Integer
      For jj = 0 To intNbProd - 1
        If (aobjProd(jj) Is aobjInst(ii).ReferenceProduct) Then Exit For
      Next
            
      ' Alter Produkt-Name
      Dim strOldProdName 'As String
      strOldProdName = aobjInst(ii).PartNumber
      
      ' Neuer Produkt-Name
      Dim strProdName 'As String
      strProdName = astrNewProdName(0, jj)
      If (strProdName = "") Then strProdName = strOldProdName
      
      ' Alter Instanz-Name
      Dim strOldInstName 'As String
      strOldInstName = aobjInst(ii).Name
      
      ' Neuer Instanz-Name
      Dim strInstName 'As String
      
      If (sblnNamensKonsistenz) Then
        ' Instanzname an Produktname anpassen
        
        ' Suffix des Instanz-Namens
        Dim strSuffix 'As String
        strSuffix = ""
        
        If (Left(strOldInstName, Len(strOldProdName)) = strOldProdName) Then
          ' Produktname ist Bestandteil des Instanznamens
          strSuffix = Mid(strOldInstName, Len(strOldProdName) + 1)
        
          If (sbln_CATPart) Then
            ' ".CATPart" entfernen
            intPos = InStr(strSuffix, ".CATPart")
            If (intPos > 0) Then
              strSuffix = Left(strSuffix, intPos - 1) & Mid(strSuffix, intPos + 8)
            End If
          End If
        Else
          ' Instanzname hat ein Suffix der Form ".nnn"
          Dim intPosDot 'As Integer
          intPosDot = InStrRev(strOldInstName, ".")
          If (intPosDot > 0 And Len(strOldInstName) - intPosDot <= 3 _
              And IsNumeric(Mid(strOldInstName, intPosDot + 1))) Then
            strSuffix = Mid(strOldInstName, intPosDot)
          End If
        End If
        
        strInstName = strProdName & strSuffix
      ElseIf (sblnUmbenennen) Then
        ' Sachnummer in Instanzname ersetzen
      
        strInstName = strOldInstName
        If (intLenOldSNR > 0 And strOldSNR = Left(strInstName, intLenOldSNR)) Then
          strInstName = strNewSNR & Mid(strInstName, intLenOldSNR + 1)
        End If
      End If
            
      ' Neuen Instanz-Name speichern
      astrNewInstName(ii) = ""
      If (strInstName <> aobjInst(ii).Name) Then astrNewInstName(ii) = strInstName
            
      ' Dateiname
      Dim strFileName 'As String
      
      ' Typ ermitteln
      Dim strType 'As String
      
      If (aobjProd(jj).Name = aobjProd(jj).Parent.Product.Name) Then
        ' Part oder Product
        strFileName = aobjProd(jj).Parent.Name
        
        If (TypeName(aobjProd(jj).Parent) = "PartDocument") Then
          strType = "Part:       "
        Else
          If (sblnEnglish) Then
            strType = "Product:    "
          Else
            strType = "Produkt:    "
          End If
        End If
      ElseIf (aobjProd(jj).HasAMasterShapeRepresentation = False) Then
        ' Component
        If (sblnEnglish) Then
          strType = "Component:  "
        Else
          strType = "Komponente: "
        End If
      Else
        ' cgr oder model
        strFileName = aobjProd(jj).GetMasterShapeRepresentationPathName()
        strFileName = Mid(strFileName, InStrRev(strFileName, CATIA.FileSystem.FileSeparator) + 1)
        
        If (Right(strFileName, 4) = ".cgr") Then
          strType = "CGR:        "
        ElseIf (Right(strFileName, 6) = ".model") Then
          If (sblnEnglish) Then
            strType = "V4-Model:   "
          Else
            strType = "V4-Modell:  "
          End If
        Else
          If (sblnEnglish) Then
            strType = "Object:     "
          Else
            strType = "Objekt:     "
          End If
        End If
      End If
      
      ReDim Preserve astrIndent(intDepth - 1)
            
      If (intDepthP + 1 = intDepth) Then
        astrIndent(intDepthP) = "|"
      End If
        
      ' Einrueckung fuer aktuelle Zeile
      Dim strIndent 'As String
            
      ' Leerzeile schreiben
      strIndent = Join(astrIndent, "  ") & "  "
      Call WriteLine(blnHTML, strIndent, "", objTmpStream)
            
      ' Zeile mit Instanzname schreiben
      If (blnExpand Or strType = "Komponente: " Or strType = "Component: ") Then
        strIndent = Left(strIndent, Len(strIndent) - 3) & "+--"
      Else
        strIndent = Left(strIndent, Len(strIndent) - 3) & "+-+"
      End If
      If (sblnEnglish) Then
        Call WriteLine(blnHTML, strIndent & "Instance:   " & aobjInst(ii).Name, astrNewInstName(ii), objTmpStream)
      Else
        Call WriteLine(blnHTML, strIndent & "Instanz:    " & aobjInst(ii).Name, astrNewInstName(ii), objTmpStream)
      End If
      
      If (blnLast) Then astrIndent(UBound(astrIndent)) = " "
            
      ' Zeile mit Produktname/Partname schreiben
      strIndent = Join(astrIndent, "  ") & "  "
      Call WriteLine(blnHTML, strIndent & strType & aobjInst(ii).PartNumber, astrNewProdName(0, jj), objTmpStream)
            
      ' Zeile mit Dateiname schreiben
      If (blnExpand And strType <> "Komponente: " And strType <> "Component: ") Then
        If (sblnEnglish) Then
          Call WriteLine(blnHTML, strIndent & "File:       " & strFileName, astrNewProdName(1, jj), objTmpStream)
        Else
          Call WriteLine(blnHTML, strIndent & "Datei:      " & strFileName, astrNewProdName(1, jj), objTmpStream)
        End If
      End If
            
      'strMsg = "Instanz: " & aobjProdInst(ii).Name & vbLf & _
      '         "Tiefe: " & aintStructure(0, ii) & vbLf & _
      '         "Name: " & aobjProdInst(ii).PartNumber & vbLf & _
      '         "Type: " & TypeName(aobjProdInst(ii).ReferenceProduct.Parent) & vbLf & _
      '         "Datei: " & aobjProdInst(ii).ReferenceProduct.Parent.Name
      'Call MsgBox(strMsg)
    Next
  End If
    
  Call WriteFooter(blnHTML, objTmpStream)
    
  Call objTmpStream.Close
  Set objTmpStream = Nothing
  Set objTmpFile = Nothing
  'Call MsgBox(strTmpPath & ": " & CATIA.FileSystem.FileExists(strTmpPath))
  'Call InputBox("", "", strTmpPath)
    
  ' Viewer und Prozess-Art bestimmen
  Dim strViewer 'As String
  If (blnHTML) Then
    If (CATIA.SystemConfiguration.OperatingSystem = "intel_a") Then
    
      ' Internet Explorer selbst suchen, steht meistens nicht in Pfad-Variable.
      If (sstrViewerIntelHTML = "iexplore.exe") Then
        strViewer = FindIExplorer()
        If (strViewer = "") Then strViewer = sstrViewerIntelHTML
        If (blnDebug) Then CATIA.SystemService.Print "Viewer: " & strViewer
      Else
        strViewer = sstrViewerIntelHTML
      End If
      
      On Error Resume Next
      Call CATIA.SystemService.ExecuteBackgroundProcessus(strViewer & " " & strTmpPath)
      lngErrNumber = Err.Number
      strErrSource = Err.Source
      strErrDescr = Err.Description
      On Error GoTo 0
      If (lngErrNumber <> 0) Then
        If (InStr(1, strViewer, "\") > 0) Then
          ' Viewer mit Pfadangabe
          strMsg = "Das Programm """ & strViewer & """ konnte nicht ausgefuehrt werden!" & vbLf & vbLf _
                  & "Fehlerbeschreibung: " & strErrDescr
        Else
          ' Viewer ohne Pfadangabe
          strMsg = "Das Programm """ & strViewer & """ konnte nicht ausgefuehrt werden!" & vbLf _
                  & "Bitte erweitern Sie die Pfad-Variable von Windows, so dass" & vbLf _
                  & "sie auf das Programmverzeichnis verweist." & vbLf & vbLf _
                  & "Fehlerbeschreibung: " & strErrDescr
        End If
        Call Err.Raise(lngErrNumber, strErrSource, strMsg)
      End If
      Call CATIA.FileSystem.DeleteFile(strTmpPath)
      blnDelete = False
    End If
  Else
    If (CATIA.SystemConfiguration.OperatingSystem = "intel_a") Then
      strViewer = sstrViewerIntelText
    Else
      strViewer = sstrViewerUnixText
    End If
            
    On Error Resume Next
    Call CATIA.SystemService.ExecuteProcessus(strViewer & " " & strTmpPath)
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescr = Err.Description
    On Error GoTo 0
    If (lngErrNumber <> 0) Then
      strMsg = "Das Programm """ & strViewer & """ konnte nicht ausgefuehrt werden!" & vbLf _
              & "Bitte erweitern Sie die Pfad-Variable, so dass" & vbLf _
              & "sie auf das Programmverzeichnis verweist." & vbLf & vbLf _
              & "Fehlerbeschreibung: " & strErrDescr
      
      Call Err.Raise(lngErrNumber, strErrSource, strMsg)
    End If
    Call CATIA.FileSystem.DeleteFile(strTmpPath)
    blnDelete = False
  End If
    
  ' Stehen Aenderungen an?
  Dim blnChanges 'As Boolean
  blnChanges = False
  For ii = 0 To intNbProd - 1
    If (astrNewProdName(0, ii) <> "") Then blnChanges = True
    If (astrNewProdName(1, ii) <> "") Then blnChanges = True
  Next
  For ii = 0 To intNbInst - 1
    If (astrNewInstName(ii) <> "") Then blnChanges = True
  Next
        
  Dim blnSaveIt 'As Boolean
  blnSaveIt = False
    
  Dim blnPrompt 'As Boolean
  blnPrompt = False
  
  ' Warnhinweise vom System beim Speichern ausschalten
  CATIA.DisplayFileAlerts = False
  
  If (blnChanges) Then
    strMsg = "Die Aenderungen werden automatisch durchgefuehrt." & vbLf & vbLf _
            & "Druecken Sie <Ja> zum automatischen Speichern ohne weitere Zwischenfragen." & vbLf _
            & "Druecken Sie <Nein> um Sicherheitsabfragen beim Speichern zu konfigurieren." & vbLf _
            & "Druecken Sie <Abbrechen> um die Aenderungen zu verwerfen."
    intRC = MsgBox(strMsg, vbYesNoCancel + vbQuestion, sstrTitle)
    
    If (intRC = vbYes) Then
      blnSaveIt = True
    ElseIf (intRC = vbNo) Then
      blnSaveIt = True
      
      strMsg = "Moechten Sie jeden Speichervorgang einzeln bestaetigen?"
      intRC = MsgBox(strMsg, vbYesNo + vbQuestion, sstrTitle)
    
      If (intRC = vbYes) Then blnPrompt = True
      
      strMsg = "Moechten Sie Warnhinweise vom System angezeigt bekommen?"
      intRC = MsgBox(strMsg, vbYesNo + vbQuestion, sstrTitle)
      
      If (intRC = vbYes) Then
        CATIA.DisplayFileAlerts = True
      End If
    End If
  Else
    Call MsgBox("Die Produkt-Struktur ist konsistent.", vbInformation, sstrTitle)
  End If
    
  ' Aenderungen durchfuehren
  If (blnSaveIt) Then
    
    ' Instanzen umbenennen
    On Error Resume Next
    For ii = 0 To intNbInst - 1
      If (astrNewInstName(ii) <> "") Then
        aobjInst(ii).Name = astrNewInstName(ii)
        If (Err.Number <> 0) Then
          Dim kk 'As Integer
          For kk = 1 To 99
            Err.Clear
            aobjInst(ii).Name = astrNewInstName(ii) & "." & CInt(kk)
            If (Err.Number = 0) Then Exit For
          Next
          
          If (Err.Number <> 0) Then
            lngErrNumber = Err.Number
            strErrSource = Err.Source
            strErrDescr = Err.Description
            On Error GoTo 0
            strMsg = "Fehler beim Umbenennen des Instanznamens von" & vbLf _
                    & aobjInst(ii).Name & " zu " & astrNewInstName(ii) & ".99" & vbLf & vbLf _
                    & "Fehlerbeschreibung: " & strErrDescr
            Call Err.Raise(lngErrNumber, strErrSource, strMsg)
          End If
        End If
      End If
    Next
    On Error GoTo 0
        
    ' Products umbenennen
    For ii = 0 To intNbProd - 1
      If (astrNewProdName(0, ii) <> "") Then
        aobjProd(ii).PartNumber = astrNewProdName(0, ii)
        If (aobjProd(ii).PartNumber <> astrNewProdName(0, ii)) Then
          strMsg = "Fehler beim Umbenennen der Teilenummer von" & vbLf _
                  & aobjProd(ii).PartNumber & " zu " & astrNewProdName(0, ii) & vbLf _
                  & "Die Teilenummer muss im gesamten Strukturbaum eindeutig sein." & vbLf _
                  & "Bitte pruefen Sie, ob die neue Teilenummer schon anderweitig im Gebrauch ist," & vbLf _
                  & "und korrigieren Sie die Teilenummer gegebenenfalls manuell."
          Call Err.Raise(9999, "CATMain", strMsg)
        End If
      End If
    Next
        
    ' rekursives Speichern
    Dim ablnProcessed() 'As Boolean
    ReDim ablnProcessed(intNbProd - 1)
    For ii = 0 To intNbProd - 1
      ablnProcessed(ii) = False
    Next
    Dim blnCancel 'As Boolean
    blnCancel = False
    Dim blnSaveParent 'As Boolean
    Call SaveChanges(0, aobjProd, astrNewProdName, aobjInst, _
                    astrNewInstName, blnPrompt, blnDebug, ablnProcessed, blnCancel, blnSaveParent)
    If (blnCancel) Then
      If (blnDebug) Then CATIA.SystemService.Print "<----- CATMain"
      Exit Sub
    End If
    
    On Error Resume Next
    Call objRootProd.Update
    On Error GoTo 0
    
    ' Bei Bedarf Root-Produkt speichern
    If (objActiveDoc.Saved = False) Then
      If (objActiveDoc.ReadOnly <> False) Then
        Call MsgBox("Keine Schreibrecht: " & objActiveDoc.Name, vbExclamation)
      Else
        If (blnPrompt) Then
          intRC = MsgBox("Haupt-Produkt speichern: " & objActiveDoc.Name, vbYesNoCancel)
          If (intRC = vbYes) Then
            On Error Resume Next
            Call objActiveDoc.Save
            lngErrNumber = Err.Number
            strErrSource = Err.Source
            strErrDescr = Err.Description
            On Error GoTo 0
            If (lngErrNumber <> 0) Then
              strMsg = "Fehler beim Speichern von " & objActiveDoc.Name & vbLf & vbLf _
                      & "Fehlerbeschreibung: " & strErrDescr
              Call Err.Raise(lngErrNumber, strErrSource, strMsg)
            End If
          ElseIf (intRC = vbCancel) Then
            If (blnDebug) Then CATIA.SystemService.Print "<----- CATMain"
            Exit Sub
          End If
        Else
          On Error Resume Next
          Call objActiveDoc.Save
          lngErrNumber = Err.Number
          strErrSource = Err.Source
          strErrDescr = Err.Description
          On Error GoTo 0
          If (lngErrNumber <> 0) Then
            strMsg = "Fehler beim Speichern von " & objActiveDoc.Name & vbLf & vbLf _
                    & "Fehlerbeschreibung: " & strErrDescr
            Call Err.Raise(lngErrNumber, strErrSource, strMsg)
          End If
        End If
      End If
    End If
    
    ' Eigentlich sollte hier schon alles gespeichert sein, aber bei Querverlinkungen muessen manche
    ' Dokumente nochmals gespeichert werden.
    Dim intNbUnsaved 'As Integer
    intNbUnsaved = 0
    Dim intNbWritable, intNbWritableP 'As Integer
    intNbWritable = 0
        
    ' Anzahl der "echten" Produkte mit Save-Flag ermitteln
    Dim intNbProd2 'As Integer
    intNbProd2 = 0
    
    For ii = 0 To intNbProd - 1
      If (aobjProd(ii).Name = aobjProd(ii).Parent.Product.Name) Then
        ' Part oder Product
        intNbProd2 = intNbProd2 + 1
        
        If (aobjProd(ii).Parent.Saved = False) Then
          intNbUnsaved = intNbUnsaved + 1
          If (aobjProd(ii).Parent.ReadOnly = False) Then intNbWritable = intNbWritable + 1
        End If
      End If
    Next
        
    intNbWritableP = intNbWritable + 1
    While (intNbWritable > 0 And intNbWritable < intNbWritableP)
      
      If (blnDebug) Then
        strMsg = intNbUnsaved & " Dokumente sind noch als ungespeichert markiert." _
                  & vbLf & intNbWritable & " Dokumente davon haben Schreibrecht."
        Call MsgBox(strMsg)
      End If
      
      ' Alten Wert retten
      intNbWritableP = intNbWritable
          
      ' Produkte speichern
      For ii = 0 To intNbProd - 1
        If (aobjProd(ii).Name = aobjProd(ii).Parent.Product.Name) Then
        
          Dim objDoc 'As Document
          Set objDoc = aobjProd(ii).Parent
                
          'Call MsgBox(objDoc.Saved & (Not objDoc.Saved) & objDoc.Saved)
          ' Achtung!!! If (Not objDoc.Saved) Then .. funktioniert nicht.
          If (objDoc.Saved = False And objDoc.ReadOnly = False) Then
            On Error Resume Next
            Call aobjProd(ii).Update
            On Error GoTo 0
            If (blnPrompt) Then
              intRC = MsgBox("Speichern: " & objDoc.Name, vbYesNoCancel)
              If (intRC = vbYes) Then
                On Error Resume Next
                Call objDoc.Save
                lngErrNumber = Err.Number
                strErrSource = Err.Source
                strErrDescr = Err.Description
                On Error GoTo 0
                If (lngErrNumber <> 0) Then
                  strMsg = "Fehler beim Speichern von " & objDoc.Name & vbLf & vbLf _
                          & "Fehlerbeschreibung: " & strErrDescr
                  Call Err.Raise(lngErrNumber, strErrSource, strMsg)
                End If
              ElseIf (intRC = vbCancel) Then
                If (blnDebug) Then CATIA.SystemService.Print "<----- CATMain"
                Exit Sub
              End If
            Else
              On Error Resume Next
              Call objDoc.Save
              lngErrNumber = Err.Number
              strErrSource = Err.Source
              strErrDescr = Err.Description
              On Error GoTo 0
              If (lngErrNumber <> 0) Then
                strMsg = "Fehler beim Speichern von " & objDoc.Name & vbLf & vbLf _
                        & "Fehlerbeschreibung: " & strErrDescr
                Call Err.Raise(lngErrNumber, strErrSource, strMsg)
              End If
            End If
          End If
        End If
      Next
          
      ' Anzahl der Produkte mit Save-Flag ermitteln
      intNbWritable = 0
      intNbUnsaved = 0
      For ii = 0 To intNbProd - 1
        If (aobjProd(ii).Name = aobjProd(ii).Parent.Product.Name) Then
          If (aobjProd(ii).Parent.Saved = False) Then
            intNbUnsaved = intNbUnsaved + 1
            If (aobjProd(ii).Parent.ReadOnly = False) Then intNbWritable = intNbWritable + 1
          End If
        End If
      Next
    Wend
        
    If (intNbUnsaved > 0) Then
      strMsg = intNbUnsaved & " Dokumente sind als ungespeichert markiert!" _
                & vbLf & CStr(intNbUnsaved - intNbWritable) & " Dokumente davon " _
                & "haben kein Schreibrecht." & vbLf _
                & "Bitte speichern Sie die Dokumente von Hand."
      Call MsgBox(strMsg, vbExclamation, sstrTitle)
    End If
  
    Call MsgBox("Die Aenderungen wurden erfolgreich durchgefuehrt.", vbInformation, sstrTitle)
        
  End If
    
  If (blnDelete) Then
    Call CATIA.FileSystem.DeleteFile(strTmpPath)
  End If
  
  If (blnDebug) Then CATIA.SystemService.Print "<----- CATMain"
End Sub


'VBA: Sub GetAllInstances(iobjProd As Product, iintDepth As Integer, iblnDebug As Boolean, _
'VBA:                     ioaobjInst() As Product, ioaintStructure() As Integer, _
'VBA:                     ioaobjProd() As Product, ioaintNbInst() As Integer)
 Sub GetAllInstances(iobjProd, iintDepth, iblnDebug, _
                     ioaobjInst(), ioaintStructure(), _
                     ioaobjProd(), ioaintNbInst())

  Dim blnDebug 'As Boolean
  blnDebug = iblnDebug
  If (blnDebug) Then CATIA.SystemService.Print "-----> GetAllInstances"
  
  Dim strMsg 'As String
    
  ' Error Werte
  Dim lngErrNumber 'As Long
  Dim strErrDescr 'As String
  Dim strErrSource 'As String
  
  Dim intNbChildren 'As Integer
  intNbChildren = iobjProd.Products.Count
    
  Dim jj 'As Integer
  For jj = 1 To intNbChildren
    ' Ist das Attribut ReferenceProduct vorhanden?
    If (jj = 1) Then
      Dim objRefProd 'As Product
      On Error Resume Next
      Set objRefProd = iobjProd.Products.Item(jj).ReferenceProduct
      lngErrNumber = Err.Number
      strErrSource = Err.Source
      strErrDescr = Err.Description
      On Error GoTo 0
      If (lngErrNumber <> 0) Then
        strMsg = "Die Catia Produktstruktur ist nicht vollstaendig geladen." & vbLf _
                & "Bitte aktualisieren Sie das Hauptprodukt (Update)." & vbLf & vbLf _
                & "Fehlerbeschreibung: " & strErrDescr
        Call Err.Raise(lngErrNumber, strErrSource, strMsg)
      End If
    End If
    
    ' unmittelbare Instanz speichern
    Dim intNbInst 'As Integer
    intNbInst = UBound(ioaobjInst)
    Set ioaobjInst(intNbInst) = iobjProd.Products.Item(jj)
    ioaintStructure(0, intNbInst) = iintDepth + 1
    ioaintStructure(1, intNbInst) = 0
    If (jj = intNbChildren) Then
      ' Flag setzen fuer letzte Kind-Instanz
      ioaintStructure(1, intNbInst) = ioaintStructure(1, intNbInst) + 1
    End If
    intNbInst = intNbInst + 1
    ReDim Preserve ioaobjInst(intNbInst)
    ReDim Preserve ioaintStructure(1, intNbInst)
    'Call MsgBox(objInst.Name & ", " & (iintDepth + 1))
        
    ' Zaehler zum referenzierten Produkt erhoehen
    Dim intNbProd 'As Integer
    intNbProd = UBound(ioaobjProd) + 1
    Dim ii 'As Integer
    For ii = 0 To intNbProd - 1
      If (ioaobjProd(ii) Is iobjProd.Products.Item(jj).ReferenceProduct) Then
        ioaintNbInst(ii) = ioaintNbInst(ii) + 1
        Exit For
      End If
    Next
        
    If (ii = intNbProd) Then
      ' referenziertes Produkt zum Array hinzufuegen
      ReDim Preserve ioaobjProd(intNbProd)
      ReDim Preserve ioaintNbInst(intNbProd)
      Set ioaobjProd(intNbProd) = iobjProd.Products.Item(jj).ReferenceProduct
      ioaintNbInst(intNbProd) = 1
      intNbProd = intNbProd + 1
            
      ' Flag setzen fuer erste Referenz auf Produkt
      ioaintStructure(1, intNbInst - 1) = ioaintStructure(1, intNbInst - 1) + 2
            
      ' alle tieferen Instanzen speichern
      Call GetAllInstances(iobjProd.Products.Item(jj).ReferenceProduct, iintDepth + 1, _
                          blnDebug, ioaobjInst, ioaintStructure, _
                          ioaobjProd, ioaintNbInst)
    End If
  Next
  
  If (blnDebug) Then CATIA.SystemService.Print "<----- GetAllInstances"
End Sub

'VBA: Sub SaveChanges(iintIdxProd As Integer, iaobjProd() As Product, iastrNewProdName() As String, _
'VBA:                 iaobjInst() As Product, iastrNewInstName() As String, iblnPrompt As Boolean, _
'VBA:                 iblnDebug As Boolean, ioablnProcessed() As Boolean, oblnCancel As Boolean, _
'VBA:                 oblnSaveParent As Boolean)
 Sub SaveChanges(iintIdxProd, iaobjProd(), iastrNewProdName(), _
                 iaobjInst(), iastrNewInstName(), iblnPrompt, _
                 iblnDebug, ioablnProcessed(), oblnCancel, _
                 oblnSaveParent)
                
  Dim blnDebug 'As Boolean
  blnDebug = iblnDebug
  If (blnDebug) Then CATIA.SystemService.Print "-----> SaveChanges"
  
  oblnSaveParent = False
  Dim blnSaveParent 'As Boolean
  
  Dim intRC 'As Integer
  Dim strMsg 'As String
    
  ' Error Werte
  Dim lngErrNumber 'As Long
  Dim strErrDescr 'As String
  Dim strErrSource 'As String
  
  ' Flag fuer Speichern
  Dim blnStore 'As Boolean
  blnStore = False
    
  ' Indexgrenzen fuer Schleifen
  Dim intNbProd 'As Integer
  intNbProd = UBound(iaobjProd) + 1
  Dim intNbInst 'As Integer
  intNbInst = UBound(iaobjInst)
    
  ' Loop ueber alle Kinder (Instanzen)
  Dim objInst 'As Product
  For Each objInst In iaobjProd(iintIdxProd).Products
    ' Index zum Produkt-Array herausfinden
    Dim ii 'As Integer
    For ii = 0 To intNbProd - 1
      If (objInst.ReferenceProduct Is iaobjProd(ii)) Then Exit For
    Next
        
    ' Ist Produkt schon verarbeitet worden?
    If (Not ioablnProcessed(ii)) Then
      Call SaveChanges(ii, iaobjProd, iastrNewProdName, iaobjInst, iastrNewInstName, _
                      iblnPrompt, blnDebug, ioablnProcessed, oblnCancel, blnSaveParent)
      If (oblnCancel) Then
        If (blnDebug) Then CATIA.SystemService.Print "<----- SaveChanges"
        Exit Sub
      End If
      ioablnProcessed(ii) = True
      If (blnSaveParent) Then blnStore = True
    End If
                
    ' Hat sich der externe Name geaendert?
    If (iastrNewProdName(1, ii) <> "") Then blnStore = True
        
    ' Index zum Instanzen-Array herausfinden
    For ii = 0 To intNbInst - 1
      If (objInst Is iaobjInst(ii)) Then Exit For
    Next
        
    ' Hat sich der Instanz-Name geaendert?
    If (iastrNewInstName(ii) <> "") Then blnStore = True
  Next
    
  ' Hat sich der Produkt-Name geaendert?
  If (iastrNewProdName(0, iintIdxProd) <> "") Then blnStore = True
    
  ' Hat sich der Datei-Name geaendert?
  If (iastrNewProdName(1, iintIdxProd) <> "") Then blnStore = True
    
  ' Produkt speichern
  If (blnStore) Then
    Dim objDoc 'As Document
    Set objDoc = iaobjProd(iintIdxProd).Parent
        
    If (iastrNewProdName(1, iintIdxProd) = "") Then
      ' Save
      
      ' Speichern nur falls CATPart oder CATProduct
      If (iaobjProd(iintIdxProd).Name = iaobjProd(iintIdxProd).Parent.Product.Name) Then
        On Error Resume Next
        Call iaobjProd(iintIdxProd).Update
        On Error GoTo 0
        
        Dim blnEnforceSave 'As Boolean
        blnEnforceSave = False
        
        Dim strFullPath 'As String   ' Pfad fuer EnforceSave
        
        If (objDoc.Saved <> False) Then
          If (blnDebug) Then
            strMsg = "Fehler in Catia! Dokument " & objDoc.Name & " ist als bereits " _
                     & "gespeichert markiert." & vbLf & "Speichern wird erzwungen!"
            Call MsgBox(strMsg, vbExclamation)
          End If
          
          ' Speichern durch Namensaenderung erzwingen
          Dim strName 'As String
          strName = iaobjProd(iintIdxProd).PartNumber
          iaobjProd(iintIdxProd).PartNumber = strName & "###"
          'On Error Resume Next
          'Call iaobjProd(iintIdxProd).Update
          'On Error Goto 0
          iaobjProd(iintIdxProd).PartNumber = strName
          On Error Resume Next
          Call iaobjProd(iintIdxProd).Update
          On Error GoTo 0
          
          If (objDoc.Saved <> False) Then
            'strMsg = "Fehler in Catia! Speichern laesst sich nicht erzwingen!" & vbLf _
            '        & "Bitte speichern Sie das Produkt " & iaobjProd(iintIdxProd).PartNumber & vbLf _
            '        & "nach Beendigung des Makros von Hand."
            'Call MsgBox(strMsg, vbCritical)
            
            If (blnDebug) Then
              strMsg = "Speichern laesst sich nicht erzwingen!" & vbLf _
                      & "Die Methode ""SaveAs"" wird angewandt."
              Call MsgBox(strMsg, vbExclamation)
            End If
            
            blnEnforceSave = True
          End If
        End If
                  
        If (objDoc.ReadOnly <> False) Then
          intRC = MsgBox("Kein Schreibrecht: " & objDoc.Name, vbOKCancel + vbExclamation)
          If (intRC = vbCancel) Then
            oblnCancel = True
            If (blnDebug) Then CATIA.SystemService.Print "<----- SaveChanges"
            Exit Sub
          End If
        Else
          If (iblnPrompt) Then
            intRC = MsgBox("Speichern: " & objDoc.Name, vbYesNoCancel)
            If (intRC = vbYes) Then
              On Error Resume Next
              If (blnEnforceSave) Then
                strFullPath = objDoc.FullName
                Call objDoc.SaveAs(strFullPath)
              Else
                Call objDoc.Save
              End If
              lngErrNumber = Err.Number
              strErrSource = Err.Source
              strErrDescr = Err.Description
              On Error GoTo 0
              If (lngErrNumber <> 0) Then
                strMsg = "Fehler beim Speichern von " & objDoc.Name & vbLf & vbLf _
                        & "Fehlerbeschreibung: " & strErrDescr
                Call Err.Raise(lngErrNumber, strErrSource, strMsg)
              End If
            ElseIf (intRC = vbCancel) Then
              oblnCancel = True
              If (blnDebug) Then CATIA.SystemService.Print "<----- SaveChanges"
              Exit Sub
            End If
          Else
            On Error Resume Next
            If (blnEnforceSave) Then
              strFullPath = objDoc.FullName
              Call objDoc.SaveAs(strFullPath)
            Else
              Call objDoc.Save
            End If
            lngErrNumber = Err.Number
            strErrSource = Err.Source
            strErrDescr = Err.Description
            On Error GoTo 0
            If (lngErrNumber <> 0) Then
              strMsg = "Fehler beim Speichern von " & objDoc.Name & vbLf & vbLf _
                      & "Fehlerbeschreibung: " & strErrDescr
              Call Err.Raise(lngErrNumber, strErrSource, strMsg)
            End If
          End If
        End If
      Else
        ' Vater-Dokument muss gespeichert werden
        oblnSaveParent = True
      End If
    Else
      ' Save As
      Dim strPath 'As String
      strPath = CATIA.FileSystem.ConcatenatePaths(objDoc.Path, iastrNewProdName(1, iintIdxProd))
          
      ' Alten Pfad retten
      Dim strOldPath 'As String
      strOldPath = objDoc.FullName
      
      ' gleiche Datei unter Windows aber unterschiedliche Gross-/Klein-Schreibung
      Dim blnSameDoc 'As Boolean
      blnSameDoc = False
      If (CATIA.SystemConfiguration.OperatingSystem = "intel_a") Then
        If (UCase(strPath) = UCase(strOldPath)) Then blnSameDoc = True
      End If
      
      On Error Resume Next
      Call iaobjProd(iintIdxProd).Update
      On Error GoTo 0
      
      If (iblnPrompt) Then
        intRC = MsgBox("Speichern unter: " & strPath, vbYesNoCancel)
        If (intRC = vbYes) Then
          If (CATIA.FileSystem.FileExists(strPath) <> False _
              And CATIA.DisplayFileAlerts = False _
              And Not blnSameDoc) Then
            strMsg = "Speichern unter: Die Datei " & strPath & " existiert bereits!" & vbLf _
                    & "Soll die Datei ueberschrieben werden?"
            intRC = MsgBox(strMsg, vbYesNo + vbExclamation)
          End If
          
          If (intRC = vbYes) Then
            On Error Resume Next
            Call objDoc.SaveAs(strPath)
            lngErrNumber = Err.Number
            strErrSource = Err.Source
            strErrDescr = Err.Description
            On Error GoTo 0
            If (lngErrNumber <> 0) Then
              strMsg = "Fehler beim Speichern-Unter von " & objDoc.Name & vbLf _
                      & "Neuer Pfad: " & strPath & vbLf & vbLf _
                      & "Fehlerbeschreibung: " & strErrDescr
              Call Err.Raise(lngErrNumber, strErrSource, strMsg)
            End If
          End If
            
          ' Altes Dokument loeschen
          If (objDoc.FullName <> strOldPath And Not blnSameDoc) Then
            On Error Resume Next
            Call CATIA.FileSystem.DeleteFile(strOldPath)
            If (Err.Number <> 0) Then
              strMsg = "Fehler beim Loeschen von " & strOldPath & vbLf _
                      & "Fehlernummer: " & Err.Number & vbLf _
                      & Err.Description
              Call MsgBox(strMsg, vbCritical)
            End If
            On Error GoTo 0
          End If
        ElseIf (intRC = vbCancel) Then
          oblnCancel = True
          If (blnDebug) Then CATIA.SystemService.Print "<----- SaveChanges"
          Exit Sub
        End If
      Else
        intRC = vbYes
        If (CATIA.FileSystem.FileExists(strPath) <> False _
            And CATIA.DisplayFileAlerts = False _
            And Not blnSameDoc) Then
          strMsg = "Speichern unter: Die Datei " & strPath & " existiert bereits!" & vbLf _
                  & "Soll die Datei ueberschrieben werden?"
          intRC = MsgBox(strMsg, vbYesNo + vbExclamation)
        End If
        
        If (intRC = vbYes) Then
          On Error Resume Next
          Call objDoc.SaveAs(strPath)
          lngErrNumber = Err.Number
          strErrSource = Err.Source
          strErrDescr = Err.Description
          On Error GoTo 0
          If (lngErrNumber <> 0) Then
            strMsg = "Fehler beim Speichern-Unter von " & objDoc.Name & vbLf _
                    & "Neuer Pfad: " & strPath & vbLf & vbLf _
                    & "Fehlerbeschreibung: " & strErrDescr
            Call Err.Raise(lngErrNumber, strErrSource, strMsg)
          End If
        
        End If
          
        ' Altes Dokument loeschen
        If (objDoc.FullName <> strOldPath And Not blnSameDoc) Then
          On Error Resume Next
          Call CATIA.FileSystem.DeleteFile(strOldPath)
          If (Err.Number <> 0) Then
            strMsg = "Fehler beim Loeschen von " & strOldPath & vbLf _
                    & "Fehlernummer: " & Err.Number & vbLf _
                    & Err.Description
            Call MsgBox(strMsg, vbCritical)
          End If
          On Error GoTo 0
        End If
      End If
    End If
  End If
  
  If (blnDebug) Then CATIA.SystemService.Print "<----- SaveChanges"
End Sub


'VBA: Sub WriteHeader(iblnHTML As Boolean, istrName As String, oobjStream As TextStream)
 Sub WriteHeader(iblnHTML, istrName, oobjStream)

  If (iblnHTML) Then
    Call oobjStream.Write("<html>" & vbLf)
    Call oobjStream.Write("<head>" & vbLf)
    Call oobjStream.Write("<meta http-equiv=""content-type"" content=""text/html; charset=iso-8859-1"">" & vbLf)
    Call oobjStream.Write("<meta http-equiv=""expires"" content=""0"">" & vbLf)
    Call oobjStream.Write("<meta http-equiv=""pragma"" content=""no-cache"">" & vbLf)
    Call oobjStream.Write("<meta http-equiv=""cache-control"" content=""no-cache"">" & vbLf)
    Call oobjStream.Write("<meta name=""author"" content=""INCAT GmbH, Germany"">" & vbLf)
    Call oobjStream.Write("<meta name=""copyright"" content=""2003 INCAT GmbH"">" & vbLf)
    Call oobjStream.Write("<title>Product Structure " & istrName & "</title>" & vbLf)
    Call oobjStream.Write("<style type=""text/css"">" & vbLf)
    Call oobjStream.Write("<!--" & vbLf)
    Call oobjStream.Write("body {font-family: ""Courier New"", Courier, monospace; font-size: 12pt;}" & vbLf)
    Call oobjStream.Write("pre {font-family: ""Courier New"", Courier, monospace; font-size: 12pt;}" & vbLf)
    Call oobjStream.Write("-->" & vbLf)
    Call oobjStream.Write("</style>" & vbLf)
    Call oobjStream.Write("</head>" & vbLf)
    Call oobjStream.Write("<body>" & vbLf)
    Call oobjStream.Write("<pre>" & vbLf)
  End If
End Sub

'VBA: Sub WriteFooter(iblnHTML As Boolean, oobjStream As TextStream)
 Sub WriteFooter(iblnHTML, oobjStream)

  If (iblnHTML) Then
    Call oobjStream.Write("</pre>" & vbLf)
    Call oobjStream.Write("</body>" & vbLf)
    Call oobjStream.Write("</html>" & vbLf)
  End If
End Sub

'VBA: Sub WriteLine(iblnHTML As Boolean, istr1 As String, istr2 As String, oobjStream As TextStream)
 Sub WriteLine(iblnHTML, istr1, istr2, oobjStream)

  If (iblnHTML) Then
    If (istr2 = "") Then
      Call oobjStream.Write(istr1 & vbLf)
    Else
      Call oobjStream.Write(istr1 & ", " & "<font color=""#ff0000"">" & istr2 & "</font>" & vbLf)
    End If
  Else
    If (istr2 = "") Then
      Call oobjStream.Write(istr1 & vbLf)
    Else
      Call oobjStream.Write(istr1 & ", " & istr2 & vbLf)
    End If
  End If
End Sub

'VBA: Function FindIExplorer() As String
 Function FindIExplorer()
  
  FindIExplorer = ""
  
  Dim objFSO 'As FileSystemObject
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
  Dim objFile
  Dim strLongPath 'As String
  
  strLongPath = "C:\Programme\Internet Explorer\iexplore.exe"
  If (objFSO.FileExists(strLongPath) <> False) Then
    Set objFile = objFSO.GetFile(strLongPath)
    FindIExplorer = objFile.ShortPath
    Set objFile = Nothing
  Else
    strLongPath = "C:\Program Files\Internet Explorer\iexplore.exe"
    If (objFSO.FileExists(strLongPath) <> False) Then
      Set objFile = objFSO.GetFile(strLongPath)
      FindIExplorer = objFile.ShortPath
      Set objFile = Nothing
    End If
  End If
  
  Set objFSO = Nothing
End Function


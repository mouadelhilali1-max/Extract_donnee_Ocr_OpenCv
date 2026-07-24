Sub CATMain()
  On Error Resume Next

  ' --- Paramètres ---
  outFolder = "C:\Users\pc\Desktop\Projet_Stage\Projet_analyse\CATIA_AI_Comparator\output"
  outFileName = "catia_spec_inspect_vbs.json"
  outPath = outFolder & "\" & outFileName

  ' --- Obtenir CATIA et le document actif ---
  Dim CATIA, doc, sel
  Set CATIA = Nothing
  Set doc = Nothing
  Set sel = Nothing

  Set CATIA = GetObject(, "CATIA.Application")
  If Err.Number <> 0 Then
    MsgBox "Impossible d'obtenir CATIA.Application"
    Exit Sub
  End If

  Set doc = CATIA.ActiveDocument
  If doc Is Nothing Then
    MsgBox "Aucun document actif"
    Exit Sub
  End If

  ' --- Choix de l'extraction ---
  targetName = InputBox("Nom du noeud ou ensemble d'annotations à extraire (laisser vide pour tout) :", "Extraction CATIA", "")
  targetName = Trim(targetName)
  targetNameNormalized = LCase(targetName)

  ' --- Préparer la sélection et la recherche ---
  Set sel = doc.Selection
  sel.Clear
  sel.Search "Name=*,all"

  ' Récupérer le nombre d'éléments (Count ou Count2)
  Dim count
  count = 0
  On Error Resume Next
  count = sel.Count
  If Err.Number <> 0 Then
    Err.Clear
    count = sel.Count2
  End If
  On Error GoTo 0

  If count = 0 Then
    MsgBox "Aucun élément trouvé par Selection.Search."
    ' continuer quand même pour produire un JSON vide
  End If

  ' --- Créer dossier de sortie si nécessaire ---
  CreateFolderIfNotExists outFolder

  ' --- Construire JSON et CSV ---
  Dim json, csv
  json = "{""items"":[]}" 
  json = "{""items"":[]}" ' placeholder will be replaced below
  json = "{""items"":[]}" ' ensure variable exists
  json = "{""items"":["
  csv = "index,type,name,Caption,Label,GetCaption,GetName,ParentName,ParentType,Container,Folder,Path" & vbCrLf

  Dim firstItem
  firstItem = True

  Dim i
  For i = 1 To count
    On Error Resume Next
    Dim selItem
    Set selItem = Nothing

    ' Essayer Item, puis Item2
    Set selItem = Nothing
    Err.Clear
    On Error Resume Next
    Set selItem = sel.Item(i).Value
    If Err.Number <> 0 Then
      Err.Clear
      Set selItem = Nothing
      On Error Resume Next
      Set selItem = sel.Item2(i).Value
      If Err.Number <> 0 Then
        Err.Clear
        Set selItem = Nothing
      End If
    End If
    On Error GoTo 0

    If selItem Is Nothing Then
      ' Saute l'itération si élément non accessible
    Else
      ' Rassembler les informations
      Dim tname, itemName, valCaption, valLabel, valGetCaption, valGetName
      Dim parentName, parentType

      tname = SafeTypeName(selItem)

      ' Name
      itemName = SafePropertyGet(selItem, "Name")

      ' Caption, Label
      valCaption = SafePropertyGet(selItem, "Caption")
      valLabel = SafePropertyGet(selItem, "Label")

      ' GetCaption() et GetName() (méthodes)
      valGetCaption = SafeMethodCall(selItem, "GetCaption")
      valGetName = SafeMethodCall(selItem, "GetName")

      ' Parent (nom et type)
      parentName = ""
      parentType = ""
      On Error Resume Next
      Dim parentObj
      Set parentObj = Nothing
      Err.Clear
      On Error Resume Next
      Set parentObj = selItem.Parent
      If Err.Number <> 0 Then
        Err.Clear
        Set parentObj = Nothing
      End If
      On Error GoTo 0
      If Not parentObj Is Nothing Then
        parentName = SafePropertyGet(parentObj, "Name")
        parentType = SafeTypeName(parentObj)
      End If

      itemPath = BuildItemPathString(selItem)
      typeLower = LCase(SafeTypeName(selItem))
      pathLower = LCase(itemPath)

      ' Déterminer le container (nom de l'ensemble d'annotations s'il existe)
      containerName = GetAnnotationSetNameFromPath(itemPath)
      containerLower = LCase(containerName)

      includeItem = True
      ' Est-ce lié aux annotations ? heuristiques
      If Not IsAnnotationRelated(pathLower, typeLower) Then
        includeItem = False
      End If

      ' Filtre par nom demandé : si demandé, vérifier que le container ou le chemin contient la chaine
      If Len(targetNameNormalized) > 0 Then
        If InStr(containerLower, targetNameNormalized) = 0 And InStr(pathLower, targetNameNormalized) = 0 Then
          includeItem = False
        End If
      End If

      If includeItem Then
        ' Déterminer le libellé de dossier visible (Captures, Vues, ...)
        folderLabel = GetFolderLabelFromPath(pathLower, typeLower)

        ' Ajouter virgule si besoin
        If Not firstItem Then
          json = json & ","
        Else
          firstItem = False
        End If

        ' Concaténer l'entrée JSON (en échappant les chaînes)
        json = json & "{"
        json = json & """index"":" & i & ","
        json = json & """type"":""" & JsonEscape(tname) & """" & ","
        json = json & """name"":""" & JsonEscape(itemName) & """" & ","
        json = json & """Caption"":""" & JsonEscape(valCaption) & """" & ","
        json = json & """Label"":""" & JsonEscape(valLabel) & """" & ","
        json = json & """GetCaption"":""" & JsonEscape(valGetCaption) & """" & ","
        json = json & """GetName"":""" & JsonEscape(valGetName) & """" & ","
        json = json & """ParentName"":""" & JsonEscape(parentName) & """" & ","
        json = json & """ParentType"":""" & JsonEscape(parentType) & """" & ","
        json = json & """Container"":""" & JsonEscape(containerName) & """" & ","
        json = json & """Folder"":""" & JsonEscape(folderLabel) & """" & ","
        json = json & """Path"":""" & JsonEscape(itemPath) & """"
        json = json & "}"

        csv = csv & CsvEscape(CStr(i)) & "," & CsvEscape(tname) & "," & CsvEscape(itemName) & "," & CsvEscape(valCaption) & "," & CsvEscape(valLabel) & "," & CsvEscape(valGetCaption) & "," & CsvEscape(valGetName) & "," & CsvEscape(parentName) & "," & CsvEscape(parentType) & "," & CsvEscape(containerName) & "," & CsvEscape(folderLabel) & "," & CsvEscape(itemPath) & vbCrLf
      End If
    End If
  Next

  json = json & "]}"

  ' --- Écrire fichier JSON ---
  If Not WriteTextFile(outPath, json) Then
    MsgBox "Impossible d'écrire le fichier JSON : " & outPath
    Exit Sub
  End If

  csvPath = outFolder & "\" & "catia_spec_inspect_vbs.csv"
  If Not WriteTextFile(csvPath, csv) Then
    MsgBox "Impossible d'écrire le fichier CSV : " & csvPath
    Exit Sub
  End If

  ' Tentative de conversion en .xlsx via Excel COM
  xlsxPath = outFolder & "\" & "catia_spec_inspect_vbs.xlsx"
  On Error Resume Next
  Dim xl, wb
  Set xl = CreateObject("Excel.Application")
  If Err.Number <> 0 Then
    Err.Clear
    Set xl = Nothing
    MsgBox "Export CSV créé. Excel non disponible pour convertir en XLSX." & vbCrLf & "CSV: " & csvPath
  Else
    On Error GoTo 0
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(csvPath)
    On Error Resume Next
    wb.SaveAs xlsxPath, 51 ' xlOpenXMLWorkbook (xlsx)
    If Err.Number <> 0 Then
      Err.Clear
      MsgBox "CSV exporté mais conversion en XLSX a échoué. CSV disponible: " & csvPath
    Else
      MsgBox "Inspect exporté: " & outPath & vbCrLf & "CSV exporté: " & csvPath & vbCrLf & "XLSX exporté: " & xlsxPath
    End If
    On Error Resume Next
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing
  End If
End Sub

Function WriteTextFile(filePath, text)
  On Error Resume Next
  Dim stream, fso, file
  Set stream = CreateObject("ADODB.Stream")
  If Err.Number = 0 Then
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    If Err.Number = 0 Then
      WriteTextFile = True
      Exit Function
    End If
  End If

  Err.Clear
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set file = fso.CreateTextFile(filePath, True, False)
  file.Write text
  file.Close
  If Err.Number = 0 Then
    WriteTextFile = True
  Else
    WriteTextFile = False
  End If
  Err.Clear
  On Error GoTo 0
End Function

' -----------------------
' Helpers
' -----------------------

Sub CreateFolderIfNotExists(folderPath)
  On Error Resume Next
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(folderPath) Then
    On Error Resume Next
    fso.CreateFolder folderPath
    If Err.Number <> 0 Then
      Err.Clear
    End If
  End If
  On Error GoTo 0
End Sub

Function SafePropertyGet(obj, propName)
  On Error Resume Next
  Dim v
  v = ""
  If IsNull(obj) Then
    SafePropertyGet = ""
    Exit Function
  End If
  Err.Clear
  Select Case LCase(propName)
    Case "name"
      On Error Resume Next
      v = obj.Name
    Case "caption"
      On Error Resume Next
      v = obj.Caption
    Case "label"
      On Error Resume Next
      v = obj.Label
    Case "title"
      On Error Resume Next
      v = obj.Title
    Case "username"
      On Error Resume Next
      v = obj.UserName
    Case "displayedname"
      On Error Resume Next
      v = obj.DisplayedName
    Case "text"
      On Error Resume Next
      v = obj.Text
    Case Else
      v = ""
  End Select
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
  If IsNull(v) Then v = ""
  SafePropertyGet = CStr(v)
End Function

Function SafeMethodCall(obj, methodName)
  On Error Resume Next
  Dim res
  res = ""
  If IsNull(obj) Then
    SafeMethodCall = ""
    Exit Function
  End If
  Err.Clear
  Select Case LCase(methodName)
    Case "getcaption"
      On Error Resume Next
      res = obj.GetCaption()
    Case "getname"
      On Error Resume Next
      res = obj.GetName()
    Case Else
      res = ""
  End Select
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
  If IsNull(res) Then res = ""
  SafeMethodCall = CStr(res)
End Function

Function SafeTypeName(obj)
  On Error Resume Next
  Dim tn
  tn = ""
  tn = TypeName(obj)
  If IsNull(tn) Then tn = ""
  SafeTypeName = tn
End Function

Function BuildItemPathString(obj)
  Dim parts(), count, current, label
  ReDim parts(0)
  count = 0
  Set current = obj
  Do While Not current Is Nothing And count < 200
    label = SafePropertyGet(current, "Name")
    If Len(label) = 0 Then label = SafePropertyGet(current, "Caption")
    If Len(label) = 0 Then label = SafePropertyGet(current, "Label")
    If Len(label) = 0 Then label = SafeTypeName(current)
    If Len(label) > 0 Then
      parts(count) = label
      count = count + 1
      If count < 200 Then ReDim Preserve parts(count)
    End If
    On Error Resume Next
    Set current = current.Parent
    If Err.Number <> 0 Then Err.Clear : Exit Do
    On Error GoTo 0
  Loop

  If count = 0 Then
    BuildItemPathString = ""
    Exit Function
  End If

  Dim reversed(), i
  ReDim reversed(count - 1)
  For i = 0 To count - 1
    reversed(count - 1 - i) = parts(i)
  Next
  BuildItemPathString = Join(reversed, " / ")
End Function

Function PathContains(pathText, targetText)
  If Len(Trim(targetText)) = 0 Then
    PathContains = True
    Exit Function
  End If
  PathContains = (InStr(LCase(pathText), LCase(targetText)) > 0)
End Function

Function JsonEscape(s)
  If IsNull(s) Then JsonEscape = "" : Exit Function
  If Len(s) = 0 Then JsonEscape = "" : Exit Function
  s = Replace(s, "\", "\\")
  s = Replace(s, Chr(34), "\""" & Chr(34))
  s = Replace(s, vbCrLf, "\n")
  s = Replace(s, vbCr, "\n")
  s = Replace(s, vbLf, "\n")
  JsonEscape = s
End Function

Function CsvEscape(s)
  If IsNull(s) Then CsvEscape = "" : Exit Function
  s = CStr(s)
  If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
    s = Replace(s, """", """""")
    s = """" & s & """"
  End If
  CsvEscape = s
End Function

' ----- Helpers spécifiques aux annotations ----
Function GetAnnotationSetNameFromPath(pathText)
  On Error Resume Next
  If Len(Trim(pathText)) = 0 Then
    GetAnnotationSetNameFromPath = "Résultat d'un ensemble d'annotations"
    Exit Function
  End If
  parts = Split(pathText, " / ")
  For Each p In parts
    pl = LCase(Trim(p))
    If InStr(pl, "ensemble") > 0 And InStr(pl, "annotation") > 0 Then
      GetAnnotationSetNameFromPath = p
      Exit Function
    End If
    If InStr(pl, "annotationset") > 0 Then
      GetAnnotationSetNameFromPath = p
      Exit Function
    End If
    If InStr(pl, "résultat") > 0 Or InStr(pl, "resultat") > 0 Then
      GetAnnotationSetNameFromPath = p
      Exit Function
    End If
  Next
  GetAnnotationSetNameFromPath = "Résultat d'un ensemble d'annotations"
End Function

Function IsAnnotationRelated(pathLower, typeLower)
  ' heuristiques sur chemin et type
  If InStr(pathLower, "annotation") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(pathLower, "capture") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(pathLower, "vue") > 0 Or InStr(pathLower, "view") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(pathLower, "référence") > 0 Or InStr(pathLower, "reference") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(pathLower, "cadre") > 0 Or InStr(pathLower, "tolerance") > 0 Or InStr(pathLower, "tolérance") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(pathLower, "note") > 0 Or InStr(pathLower, "noa") > 0 Or InStr(pathLower, "flagnote") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(pathLower, "publication") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(typeLower, "annotation") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(typeLower, "tolerance") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(typeLower, "capture") > 0 Then IsAnnotationRelated = True : Exit Function
  If InStr(typeLower, "view") > 0 Then IsAnnotationRelated = True : Exit Function
  IsAnnotationRelated = False
End Function

Function GetFolderLabelFromPath(pathLower, typeLower)
  If InStr(pathLower, "capture") > 0 Then GetFolderLabelFromPath = "Captures" : Exit Function
  If InStr(pathLower, "vue") > 0 Or InStr(pathLower, "view") > 0 Then GetFolderLabelFromPath = "Vues" : Exit Function
  If InStr(pathLower, "référence") > 0 Or InStr(pathLower, "reference") > 0 Then GetFolderLabelFromPath = "Références" : Exit Function
  If InStr(pathLower, "cadre") > 0 Or InStr(pathLower, "tolerance") > 0 Or InStr(pathLower, "tolérance") > 0 Then GetFolderLabelFromPath = "Cadres de tolérances" : Exit Function
  If InStr(pathLower, "note") > 0 Or InStr(pathLower, "noa") > 0 Or InStr(pathLower, "flagnote") > 0 Then GetFolderLabelFromPath = "Notes" : Exit Function
  If InStr(pathLower, "publication") > 0 Then GetFolderLabelFromPath = "Publication" : Exit Function
  ' fallback sur type
  If InStr(typeLower, "capture") > 0 Then GetFolderLabelFromPath = "Captures" : Exit Function
  If InStr(typeLower, "view") > 0 Then GetFolderLabelFromPath = "Vues" : Exit Function
  If InStr(typeLower, "tolerance") > 0 Then GetFolderLabelFromPath = "Cadres de tolérances" : Exit Function
  If InStr(typeLower, "note") > 0 Then GetFolderLabelFromPath = "Notes" : Exit Function
  If InStr(typeLower, "publication") > 0 Then GetFolderLabelFromPath = "Publication" : Exit Function
  GetFolderLabelFromPath = "Annotations"
End Function

Sub deleteFilesByName()

Dim objFSO, objFolder, FF As Object
Dim delete As String
Dim arr(0 To 0) As String
    
folderPath = InputBox("Путь к основной папке")
fileNameToDelete = InputBox("Название файла для удаления без расширения. Например 02")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(folderPath)

For Each folder In objFolder.subfolders
    Set FF = objFSO.GetFolder(folder.Path)
    filesCountNow = objFSO.GetFolder(folder.Path).Files.Count
    filesCount = filesCount + filesCountNow
    For Each file In FF.Files
        delete = "no"
        If file.Name = fileNameToDelete & ".jpg" Then
            delete = "yes"
        ElseIf file.Name = fileNameToDelete & ".png" Then
            delete = "yes"
        ElseIf file.Name = fileNameToDelete & ".webp" Then
            delete = "yes"
        ElseIf file.Name = fileNameToDelete & ".jpeg" Then
            delete = "yes"
        End If
        If delete = "yes" Then Kill (file)
    Next file
Next folder

MsgBox "Готово"
End Sub


Sub renameFolders()

Dim objFSO, objFolder, FF As Object
Dim delete As String
    
folderPath = InputBox("Путь к основной папке")
addToFileName = InputBox("Текст, который нужно добавить к названию папок")


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(folderPath)

For Each folder In objFolder.subfolders
    Set FF = objFSO.GetFolder(folder.Path)
    folder.Name = folder.Name & addToFileName
Next folder

MsgBox "Готово"
End Sub


Sub createFolders()
Dim sPath
Dim x
Dim i, lastRow As Integer
Dim k, answer As String
i = 1
answer = MsgBox("Названия папок должны находиться в столбце A и начинаться с 1 строки. Все заполнено верно?", vbQuestion + vbYesNo)
If answer = vbYes Then
    x = InputBox("Введите путь")
    lastX = Right(x, 1)
    If x = "" Then Exit Sub
    If Not lastX = "\" Then x = x & "\"
    lastRow = Cells(1, 1).CurrentRegion.Rows.Count
    For i = 1 To lastRow
        k = Cells(i, 1)
        sPath = CreatePath(x & k)
    Next i
    If sPath Then MsgBox "Папки созданы", vbInformation
End If
End Sub
 
Function CreatePath(PathName) As Boolean
  ' Формирование папки с указанным именем.
  ' Имя папки должно быть полным и начинаться с имени драйвера (диска или сетевого сервера).
  ' Функция создает указанную папку с любым уровнем вложения папок (в отличие от стандраного средтва объекта FileSystemObject, который создает только концевую папку).
  ' Возвращает True если папка уже существует или если она успешно создана.
Dim FSO, cDrive$, cFolder$, aFolders, nFolder, pSp
   Set FSO = CreateObject("Scripting.FileSystemObject")
   pSp = Application.PathSeparator
  If FSO.FolderExists(PathName) Then
    CreatePath = True
    Exit Function
  End If
  cDrive = FSO.GetDriveName(PathName)
  cFolder = Mid$(PathName, Len(cDrive) + 2)
  If Right$(cFolder, 1) = pSp Then cFolder = Left$(cFolder, Len(cFolder) - 1)
  If Left$(cFolder, 1) = pSp Then cFolder = Mid$(cFolder, 2)
  aFolders = Split(cFolder, pSp, -1, 0)
  cFolder = cDrive & pSp
  If Not FSO.FolderExists(cFolder) Then
    On Error GoTo Break
    FSO.CreateFolder cFolder
  End If
  If Not IsEmpty(aFolders) Then
    For nFolder = 0 To UBound(aFolders)
      cFolder = cFolder & aFolders(nFolder) & pSp
      If Not FSO.FolderExists(cFolder) Then
        On Error GoTo Break
        FSO.CreateFolder cFolder
      End If
    Next
  End If
  If Not FSO.FolderExists(cFolder) Then GoTo Break
  CreatePath = True
  Set FSO = Nothing
  Exit Function
 
Break:
  If MsgBox(Err.Description, vbExclamation + vbOKCancel, "clsFSO.CreatePath") = vbCancel Then Stop
End Function

Sub mixFiles()

Dim objFSO, objFolder, FF As Object
Dim delete As String
    
folderPath = InputBox("Путь к основной папке")


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(folderPath)

k = "yes"
first = True
For Each file In objFolder.Files
    If first = False Then
        If k = "yes" Then
            sFileName = CStr(Left(file.Name, InStrRev(file.Name, ".") - 1))
            ext = Right(file.Name, Len(file.Name) - InStrRev(file.Name, "."))
            newFileName = sFileName * 21
            file.Name = newFileName & "." & ext
            k = "no"
        ElseIf k = "no" Then
            sFileName = CStr(Left(file.Name, InStrRev(file.Name, ".") - 1))
            ext = Right(file.Name, Len(file.Name) - InStrRev(file.Name, "."))
            newFileName = sFileName * -21
            file.Name = newFileName & "." & ext
            k = "yes"
        End If
    End If
    first = False
Next file
    
n = 2
For Each file In objFolder.Files
    ext = Right(file.Name, Len(file.Name) - InStrRev(file.Name, "."))
    If file.Name = "01" & "." & ext Then
        GoTo Continue
    End If
        ext = Right(file.Name, Len(file.Name) - InStrRev(file.Name, "."))
        If n < 10 Then
            file.Name = "0" & n & "." & ext
            n = n + 1
        Else
            file.Name = n & "." & ext
            n = n + 1
        End If
Continue:
Next file
    
MsgBox "Готово"
End Sub



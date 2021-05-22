  'Option Explicit == force all variables to be declared
  
  Function GetFullName() // <!-- [7] -->
    GetFullName = Replace(cmdoSE.commandLine,"""","") //WARNING! -- This depends on the HTA definition above!!!
  End Function

  Function GetDirectory()
    Dim strPath
    Dim objFSO, objFile, objShell


    //Set objShell = CreateObject("Wscript.Shell")
    //strPath = Wscript.ScriptFullName
    strPath = GetFullName()

    //<!-- [6] -->
    //Set fso = CreateObject("Scripting.FileSystemObject")

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    //Set objFile = objFSO.GetFile("./05_Search-Engine.hta")
    Set objFile = objFSO.GetFile(strPath)
    GetDirectory = objFSO.GetParentFolderName(objFile)
  End Function
  
  Sub ListFilesWriteToCSV(name)

    Dim fso, folder, files, OutputFile
    Dim strPath

    Dim outFileName
    outFileName = "ScriptOutput.csv"

    ' Create a FileSystemObject  
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Define folder we want to list files from
    strPath = GetDirectory()

    If  name <> "" Then
      outFileName = name
    End If

    Set folder = fso.GetFolder(strPath)
    Set files = folder.Files

    Set OutputFile = fso.CreateTextFile(strPath & "/" & outFileName, True)


    ' CSV Header
      OutputFile.Write("Name")
      OutputFile.Write(",")
      OutputFile.Write("DateCreated")
      OutputFile.Write(",")
      OutputFile.Write("DateLastAccessed")
      OutputFile.Write(",")
      OutputFile.Write("DateLastModified")
      OutputFile.Write(",")
      OutputFile.WriteLine("Size")
      OutputFile.Write(",")
      OutputFile.WriteLine("FileType")


    ' Loop through each file 
    Dim singleFile
    //As Variant //https://stackoverflow.com/questions/46510416/vba-illegal-assignment-in-for-each
    For Each singleFile In files

      ' Output file properties to a text file
      OutputFile.Write(singleFile.Name)
      OutputFile.Write(",")
      OutputFile.Write(singleFile.DateCreated)
      OutputFile.Write(",")
      OutputFile.Write(singleFile.DateLastAccessed)
      OutputFile.Write(",")
      OutputFile.Write(singleFile.DateLastModified)
      OutputFile.Write(",")
      OutputFile.WriteLine(singleFile.Size)
      OutputFile.Write(",")
      OutputFile.WriteLine(fso.GetExtensionName(singleFile.Name)) 'https://stackoverflow.com/questions/18920310/vbscript-to-search-for-all-files-with-an-extension-and-save-them-to-a-csv
      
    Next

    ' Close text file
    OutputFile.Close
  End Sub


' https://developer.rhino3d.com/guides/rhinoscript/testing-for-empty-arrays/
Function IsArrayDimmed(arr) 
  IsArrayDimmed = False
  If IsArray(arr) Then
    On Error Resume Next
    Dim ub : ub = UBound(arr)
    If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
  End If  
End Function

  'https://stackoverflow.com/questions/4605270/add-item-to-array-in-vbscript
  Function ArrayAddItem(arr, val)
      ReDim Preserve arr(UBound(arr) + 1)
      arr(UBound(arr)) = val
      ArrayAddItem = arr
  End Function

' Sub ArrayAddItem(arr, val)
'     ReDim Preserve arr(UBound(arr) + 1)
'     arr(UBound(arr)) = val
' End Sub

' objFSO.GetExtensionName(objFile.Name))
  Function GetDirectorFilesByType(folderDirectory, fileType)

    Dim fso, folder, files, singleFile
    Dim FileList
    FileList = Array()
    ' ReDim FileList(1)
    ' FileList(0) = ""

    ' FileList = Array()

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set folder = fso.GetFolder(folderDirectory)
    Set files = folder.Files

    ' Dim isFirstAdded
    ' isFirstAdded = false

    For Each singleFile In files
      If LCase(fso.GetExtensionName(singleFile.Name)) = fileType Then
        ' IF NOT isFirstAdded Then
        FileList = ArrayAddItem(FileList, singleFile.Name)
        ' Else
        '   FileList(0) = singleFile.Name
        '   isFirstAdded = true
        ' End If
      End If
      
    Next

    ' SubFolder Recursion!
    '---------------------
    ' For Each objSubFolder In objFolder.SubFolders
    '     Recurse objSubFolder
    ' Next
    GetDirectorFilesByType = FileList
  End Function


  Sub RunProgram()   //<!-- [ii] -->
      Dim objShell
    Set objShell = CreateObject("Wscript.Shell")  
    objShell.Run "notepad.exe " & GetFullName() //Self-hosted file opening?
  End Sub

  Sub OpenFile(name)   //<!-- [ii] -->
      Dim objShell
    Set objShell = CreateObject("Wscript.Shell")  
    objShell.Run "notepad.exe " & name //Self-hosted file opening?
  End Sub


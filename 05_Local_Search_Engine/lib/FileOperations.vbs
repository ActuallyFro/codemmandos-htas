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


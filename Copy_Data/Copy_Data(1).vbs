Set objFSO = CreateObject("Scripting.FileSystemObject")

' Defina o diretório de origem e destino para a cópia dos dados
strSourceFolder = "C:\"
strDestinationFolder = "D:\Backup"

' Função recursiva para copiar todos os arquivos e subpastas
Sub CopyFolder(strSource, strDestination)
    If Not objFSO.FolderExists(strDestination) Then
        objFSO.CreateFolder(strDestination)
    End If

    Set objFolder = objFSO.GetFolder(strSource)
    Set colFiles = objFolder.Files

    ' Copiar todos os arquivos
    For Each objFile in colFiles
        strFile = objFSO.GetFileName(objFile)
        objFSO.CopyFile objFile.Path, strDestination & "\" & strFile
    Next

    ' Copiar todas as subpastas
    For Each objSubfolder in objFolder.Subfolders
        strSubfolder = objFSO.GetFolderName(objSubfolder)
        CopyFolder objSubfolder.Path, strDestination & "\" & strSubfolder
    Next
End Sub

' Chamar a função para iniciar a cópia
CopyFolder strSourceFolder, strDestinationFolder

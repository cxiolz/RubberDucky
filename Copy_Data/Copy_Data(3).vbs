Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Obter o caminho da pasta de usuários
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

' Definir o caminho da pasta de origem e destino
strSourceFolder = strUserProfile & "\Caminho\para\pasta\origem"
strDestinationFolder = "D:\Caminho\para\pasta\destino"

' Verificar se a pasta de origem existe
If objFSO.FolderExists(strSourceFolder) Then
    ' Verificar se a pasta de destino existe, caso contrário, criar
    If Not objFSO.FolderExists(strDestinationFolder) Then
        objFSO.CreateFolder(strDestinationFolder)
    End If

    ' Obter os arquivos da pasta de origem
    Set objFolder = objFSO.GetFolder(strSourceFolder)
    Set colFiles = objFolder.Files

    ' Copiar cada arquivo para a pasta de destino
    For Each objFile in colFiles
        strFile = objFSO.GetFileName(objFile)
        objFSO.CopyFile objFile.Path, strDestinationFolder & "\" & strFile, True
    Next

    WScript.Echo "A cópia de dados foi concluída com sucesso."
Else
    WScript.Echo "A pasta de origem não existe."
End If

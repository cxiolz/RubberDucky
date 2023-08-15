Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Obter o caminho da pasta de usuários
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

' Definir o caminho da pasta de origem e destino
strSourceFolder = strUserProfile & "\Caminho\para\pasta\origem"
strDestinationFolder = "D:\Caminho\para\pasta\destino"

' Função para copiar arquivos com menos de 1000 KB
Sub CopyFiles()
    ' Verificar se a pasta de origem existe
    If objFSO.FolderExists(strSourceFolder) Then
        ' Verificar se a pasta de destino existe, caso contrário, criar
        If Not objFSO.FolderExists(strDestinationFolder) Then
            objFSO.CreateFolder(strDestinationFolder)
        End If

        ' Obter os arquivos da pasta de origem
        Set objFolder = objFSO.GetFolder(strSourceFolder)
        Set colFiles = objFolder.Files

        ' Copiar arquivos com menos de 1000 KB para a pasta de destino
        For Each objFile in colFiles
            If objFile.Size < 1000000 Then ' 1000 KB = 1000 * 1024 bytes
                strFile = objFSO.GetFileName(objFile)
                objFSO.CopyFile objFile.Path, strDestinationFolder & "\" & strFile, True
            End If
        Next

        WScript.Echo "A cópia de dados foi concluída com sucesso."
    Else
        WScript.Echo "A pasta de origem não existe."
    End If
End Sub

' Chamada da função para copiar arquivos
CopyFiles()

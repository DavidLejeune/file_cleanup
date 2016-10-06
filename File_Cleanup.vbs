Set objFSO = CreateObject("Scripting.FileSystemObject")
call ClearScreen

'This is the folder where all your shit will be moved to in an orderly fashion
objTargetFolder = "C:\Users\David\Desktop\DaLe\File Hub\"


'For each folder (subfolders not included) change the folder name and call the main function
'Can be repeated for as many folders as you want
objStartFolder = "C:\Users\David\Downloads\"
call Main




objStartFolder = "C:\Users\David\Desktop\DaLe\File Hub\"
call Main



objStartFolder = "C:\Users\David\Desktop\DaLe\"
call Main



Wscript.Echo ""
Wscript.Echo " #####################################################"
Wscript.Echo " #               PROGRAM CONCLUDED                   #"
Wscript.Echo " #                    HOORAAAH                       #"
Wscript.Echo " #####################################################"
Wscript.Echo ""
Wscript.Echo ""

















Sub Main()

Wscript.Echo "      ____              __        "  
Wscript.Echo "     / __ \   ____ _   / /      ___ "
Wscript.Echo "    / / / /  / __ `/  / /      / _ \"
Wscript.Echo "   / /_/ /  / /_/ /  / /___   /  __/"
Wscript.Echo "  /_____/   \__,_/  /_____/   \___/ "
Wscript.Echo ""
Wscript.Echo "           +-+-+-+-+-+-+-+-+"
Wscript.Echo "           |t|e|r|m|i|n|a|l|"
Wscript.Echo "           +-+-+-+-+-+-+-+-+"
Wscript.Echo ""
Wscript.Echo " ^> Name : CleanUp Files.vbs"
Wscript.Echo " ^> Description : trying to create some order in"
Wscript.Echo " ^>               chaos"
Wscript.Echo " ^> Author : David Lejeune"
Wscript.Echo " ^> Created : 10-07-15 13:50"
Wscript.Echo ""
Wscript.Echo " #####################################################"
Wscript.Echo " #               RUNNING VBS SCRIPT                  #"
Wscript.Echo " #                                                   #"
Wscript.Echo " #####################################################"
Wscript.Echo " Target : " & objStartFolder
Wscript.Echo ""
Wscript.Sleep 2000
iTotal = 0

iNew = 0
iCount = 0
Set objFolder = objFSO.GetFolder(objStartFolder)

Set colFiles = objFolder.Files
sDelete = 0

For Each objFile in colFiles

    strinput = objFile.Name
    If instr(strinput,".") >0 Then
        sType = Ucase( Mid(strinput,instrRev(strinput,".")+1))
        If instr(sType,".") >0 Then
            sType = Ucase( Mid(sType,instrRev(sType,".")+1))
        End If
    End If

    If Len(sType) > 4 Then
    If instr(sType,".") >0 Then
            sType = Ucase( Mid(sType,instrRev(sType,".")+1))
        End If
    End If
   

    sDestinationFolder = objTargetFolder + sType + "\"
    'Wscript.echo sDestinationFolder
    If NOT (objFSO.FolderExists(sDestinationFolder)) Then
        objFSO.CreateFolder(sDestinationFolder)

        Wscript.Echo ""
        Wscript.Echo "#########################"
        Wscript.Echo "# Subfolder " & sType & " created #"
        Wscript.Echo "#########################"
        Wscript.Echo ""
        Wscript.Sleep 300
        iNew = iNew + 1
    End If
    sName = objFile.Name
    If objFSO.FileExists(sDestinationFolder & sName) Then

    objFSO.DeleteFile sDestinationFolder & sName
    sDelete = sDelete + 1
    else
    iCount = iCount + 1
    End If

    objFSO.MoveFile objFile.Path , sDestinationFolder & sName
    Wscript.Echo "[" & sType & "]" & vbTab & sName 

    
    
Next


Wscript.Echo vbCrlf & "Folder : " & objStartFolder & vbCrlf & "Created " & iNew & " new Type Folder(s)" & vbCrlf & "Moved " & sDelete & " duplicate file(s)" & vbCrlf & "Moved " & iCount & " new file(s)"
Wscript.Sleep 3500


End Sub






Sub ClearScreen()
For i = 1 to 45
    Wscript.Echo ""
Next

End Sub

Option Explicit
' Run as Admin
If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
    WScript.Quit
End If
' To let executing just one insctance of this script
If AppPrevInstance() Then 
    MsgBox "There is an existing proceeding !" & VbCrLF &_
    CommandLineLike(WScript.ScriptName),VbExclamation,"There is an existing proceeding !"    
    WScript.Quit   
Else
    Do
        Call AutoSave_USB_SDCARD()
        Pause(30)
    Loop
End If
'**************************************************************************
Function AppPrevInstance()   
    With GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")   
        With .ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE "_
            & CommandLineLike(WScript.ScriptFullName) & _
            " AND CommandLine LIKE '%WScript%' OR CommandLine LIKE '%cscript%'")   
            AppPrevInstance = (.Count > 1)   
        End With   
    End With   
End Function   
'**************************************************************************
Function CommandLineLike(ProcessPath)   
    ProcessPath = Replace(ProcessPath, "\", "\\")   
    CommandLineLike = "'%" & ProcessPath & "%'"   
End Function
'*************************AutoSave_USB_SDCARD()****************************
Sub AutoSave_USB_SDCARD()
    Dim Ws,WshNetwork,NomMachine,MyDoc,strComputer,objWMIService,objDisk,colDisks
    Dim fso,Drive,NumSerie,volume,cible,Amovible,Dossier,chemin,Command,Result
    Set Ws = CreateObject("WScript.Shell")
    Set WshNetwork = CreateObject("WScript.Network")
    NomMachine = WshNetwork.ComputerName
    MyDoc = Ws.SpecialFolders("Mydocuments")
    cible = MyDoc & "\"
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colDisks = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_LogicalDisk")

    For Each objDisk in colDisks
        If objDisk.DriveType = 2 Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            For Each Drive In fso.Drives
                If Drive.IsReady Then
                    If Drive.DriveType = 1 Then
                        NumSerie=fso.Drives(Drive + "\").SerialNumber
                        Amovible=fso.Drives(Drive + "\")
                        Numserie=ABS(INT(Numserie))
                        volume=fso.Drives(Drive + "\").VolumeName
                        Dossier=NomMachine & "_" & volume &"_"& NumSerie
                        chemin=cible & Dossier
                        Command = "cmd /c Xcopy.exe " & Amovible &" "& chemin &" /I /D /Y /S /J /C"
                        Result = Ws.Run(Command,0,True)
                    end if
                End If   
            Next
        End If   
    Next
End Sub
'**************************End of AutoSave_USB_SDCARD()*******************
Sub Pause(Sec)
    Wscript.Sleep(Sec*1000)
End Sub 
'************************************************************************
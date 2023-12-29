#Requires AutoHotkey v2.0
mygui := Gui("+Resize")
phase := mygui.AddText("w700", "Getting apps information...")
msg := mygui.AddText("wp y+10 h500", "This may take minutes")
mygui.Show()
appsinfo := pwcmd(
    msg,
    "Get-StartApps",
    "|",
    "ForEach { 'appid=' + $_.AppID + ',prettyname=' + $_.Name }",
    ";",
    "Get-AppxPackage",
    "|",
    "ForEach",
    "{",
    " 'name=' + $_.Name + ',' +"
    " 'packagefamilyname=' + $_.PackageFamilyName + ',' +",
    " 'installlocation=' + $_.InstallLocation + ',' +",
    " 'ids=' +",
    "  (Get-AppxPackageManifest $_).package.applications.application.foreach{",
    "   '' + $_.id + ' '",
    "  }",
    "}"
)
if (appsinfo.ExitCode != 0) {
    install := MsgBox("Error getting apps information`r`nInstall PowerShell?",, "YesNo")
    if (install == "Yes") {
        Run("winget install --id Microsoft.Powershell --source winget")
    }
    ExitApp()
}

phase.Text := "Parsing apps information..."
name := ""
apps := Map()
appx := Map()
for (i in StrSplit(appsinfo.Output, "`n", "`r ")) {
    kva := StrSplit(i, ",", " ")
    if (kva.Length == 2) {
        aia := StrSplit(kva[1], "=", " ")
        pna := StrSplit(kva[2], "=", " ")
        if (
            aia.Length != 2 or
            pna.Length != 2 or
            aia[1] != "appid" or
            pna[1] != "prettyname"
        ) {
            continue
        }
        appid := aia[2]
        prettyname := pna[2]

        apps[appid] := Map()
        apps[appid]["name"] := prettyname
        msg.Text := Format("[{}]=[{}]`r`n", appid, prettyname)
    }
    else if (kva.Length == 4) {
        for (kv in kva) {
            a := StrSplit(kv, "=", " ")
            if (a.Length != 2) {
                continue
            }
            if (a[1] == "name") {
                name := a[2]
                appx[name] := Map()
                appx[name]["ids"] := Array()
            }
            else if (a[1] == "ids") {
                if (a[2] == "") {
                    continue
                }
                for (id in StrSplit(a[2], " ", " ")) {
                    if (id != "") {
                        appx[name]["ids"].Push(id)
                    }
                }
                msg.Text := Format(
                    "[{}]={}{}{}`r`n",
                    name,
                    appx[name]["packagefamilyname"],
                    appx[name]["ids"].Length ? "!" : "",
                    appx[name]["ids"].Length ? appx[name]["ids"][1] : ""
                )
            }
            else {
                appx[name][a[1]] := a[2]
            }
        }
    }
}

for (i in appx) {
    a := appx[i]["ids"]
    if (a.Length == 0) {
        appid := appx[i]["packagefamilyname"]
        if (apps.Has(appid)) {
            apps[appid]["installlocation"] := appx[i]["installlocation"]
        }
        continue
    }
    for (id in a) {
        if (appx[i].Has("packagefamilyname")) {
            appid := appx[i]["packagefamilyname"] . (id ? "!" . id : "")
            if (apps.Has(appid)) {
                apps[appid]["installlocation"] := appx[i]["installlocation"]
            }
        }
    }
}
mygui.Destroy()

mygui := Gui("+Resize")
lv := mygui.Add(
    "ListView",
    "r20 w700 -ReadOnly",
    ["Name", "AppID", "InstallLocation"]
)
for (i in apps) {
    ai := apps[i]
    if (ai.Has("installlocation")) {
        lv.Add(, ai["name"], i, ai["installlocation"])
    }
}
lv.ModifyCol(1)
lv.ModifyCol(2, 100)
mygui.OnEvent("Size", resizeList)
resizeList(g, minmax, w, h) {
    global lv
    lv.Move(,, w-20, h-15)
}

lv.OnEvent("ContextMenu", showContext)
showContext(v, item, isRightClick, x, y) {
    if (item == 0) {
        return
    }
    shellapp := Format('explorer.exe "shell:Appsfolder\{}"', v.GetText(item, 2))
    folder := Format('explorer.exe "{}"', v.GetText(item, 3))

    cm := Menu()
    cm.Add("Copy AppID", substClip.Bind(shellapp))
    cm.Add("Open Folder", runsimply.Bind(folder))
    cm.Add("Run", runsimply.Bind(shellapp))
    cm.Show(x, y)

    substClip(text, *) {
        A_Clipboard := text
    }

    runsimply(text, *) {
        Run(text)
    }
}

mygui.Show()


; powershell commands
pwcmd(tc, cmds*) {
    EnvSet("TERM", "dumb")
    str := ""
    for (cmd in cmds) {
        str .= cmd . " "
    }
    return StdoutToVar(
        "pwsh -NoProfile -Command " .
        "[console]::InputEncoding = [console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); "
        str,
        ,
        "CP65001",
        tc
    )
}

; ----------------------------------------------------------------------------------------------------------------------
; Function .....: StdoutToVar
; Description ..: Runs a command line program and returns its output.
; Parameters ...: sCmd - Commandline to be executed.
; ..............: sDir - Working directory.
; ..............: sEnc - Encoding used by the target process. Look at StrGet() for possible values.
; Return .......: Command output as a string on success, empty string on error.
; AHK Version ..: AHK v2 x32/64 Unicode
; Author .......: Sean (http://goo.gl/o3VCO8), modified by nfl and by Cyruz
; License ......: WTFPL - http://www.wtfpl.net/txt/copying/
; Changelog ....: Feb. 20, 2007 - Sean version.
; ..............: Sep. 21, 2011 - nfl version.
; ..............: Nov. 27, 2013 - Cyruz version (code refactored and exit code).
; ..............: Mar. 09, 2014 - Removed input, doesn't seem reliable. Some code improvements.
; ..............: Mar. 16, 2014 - Added encoding parameter as pointed out by lexikos.
; ..............: Jun. 02, 2014 - Corrected exit code error.
; ..............: Nov. 02, 2016 - Fixed blocking behavior due to ReadFile thanks to PeekNamedPipe.
; ..............: Apr. 13, 2021 - Code restyling. Fixed deprecated DllCall types.
; ..............: Oct. 06, 2022 - AHK v2 version. Throw exceptions on failure.
; ..............: Oct. 08, 2022 - Exceptions management and handles closure fix. Thanks to lexikos and iseahound.
; ----------------------------------------------------------------------------------------------------------------------
StdoutToVar(sCmd, sDir:="", sEnc:="CP0", gMsg:=False) {
    ; Create 2 buffer-like objects to wrap the handles to take advantage of the __Delete meta-function.
    oHndStdoutRd := { Ptr: 0, __Delete: delete(this) => DllCall("CloseHandle", "Ptr", this) }
    oHndStdoutWr := { Base: oHndStdoutRd }
    
    If !DllCall( "CreatePipe"
               , "PtrP" , oHndStdoutRd
               , "PtrP" , oHndStdoutWr
               , "Ptr"  , 0
               , "UInt" , 0 )
        Throw OSError(,, "Error creating pipe.")
    If !DllCall( "SetHandleInformation"
               , "Ptr"  , oHndStdoutWr
               , "UInt" , 1
               , "UInt" , 1 )
        Throw OSError(,, "Error setting handle information.")

    PI := Buffer(A_PtrSize == 4 ? 16 : 24,  0)
    SI := Buffer(A_PtrSize == 4 ? 68 : 104, 0)
    NumPut( "UInt", SI.Size,          SI,  0 )
    NumPut( "UInt", 0x100,            SI, A_PtrSize == 4 ? 44 : 60 )
    NumPut( "Ptr",  oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 60 : 88 )
    NumPut( "Ptr",  oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 64 : 96 )

    If !DllCall( "CreateProcess"
               , "Ptr"  , 0
               , "Str"  , sCmd
               , "Ptr"  , 0
               , "Ptr"  , 0
               , "Int"  , True
               , "UInt" , 0x08000000
               , "Ptr"  , 0
               , "Ptr"  , sDir ? StrPtr(sDir) : 0
               , "Ptr"  , SI
               , "Ptr"  , PI )
        Throw OSError(,, "Error creating process.")

    ; The write pipe must be closed before reading the stdout so we release the object.
    ; The reading pipe will be released automatically on function return.
    oHndStdOutWr := ""

    ; Before reading, we check if the pipe has been written to, so we avoid freezings.
    nAvail := 0, nLen := 0
    While DllCall( "PeekNamedPipe"
                 , "Ptr"   , oHndStdoutRd
                 , "Ptr"   , 0
                 , "UInt"  , 0
                 , "Ptr"   , 0
                 , "UIntP" , &nAvail
                 , "Ptr"   , 0 ) != 0
    {
        ; If the pipe buffer is empty, sleep and continue checking.
        If !nAvail && Sleep(100)
            Continue
        cBuf := Buffer(nAvail+1)
        DllCall( "ReadFile"
               , "Ptr"  , oHndStdoutRd
               , "Ptr"  , cBuf
               , "UInt" , nAvail
               , "PtrP" , &nLen
               , "Ptr"  , 0 )
        newStr := StrGet(cBuf, nLen, sEnc)
        sOutput .= newStr
        If gMsg
            gMsg.Text := newStr
    }
    
    ; Get the exit code, close all process handles and return the output object.
    DllCall( "GetExitCodeProcess"
           , "Ptr"   , NumGet(PI, 0, "Ptr")
           , "UIntP" , &nExitCode:=0 )
    DllCall( "CloseHandle", "Ptr", NumGet(PI, 0, "Ptr") )
    DllCall( "CloseHandle", "Ptr", NumGet(PI, A_PtrSize, "Ptr") )
    Return { Output: sOutput, ExitCode: nExitCode } 
}

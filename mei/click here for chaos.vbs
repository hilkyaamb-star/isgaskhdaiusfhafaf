Dim i, maxNotifications, delayTime
maxNotifications = 40 ' Set the maximum number of notifications you want
delayTime = 1000 ' Set the delay time in milliseconds (5000 ms = 5 seconds)

' Wait for the specified delay time before starting the notifications
WScript.Sleep delayTime

' Start displaying notifications
For i = 1 To maxNotifications
    CreateObject("WScript.Shell").Run "mshta.exe ""javascript:alert('HAPPY BIRTHDAYYY ME :3');close();"""
    WScript.Sleep 0000 ' Delay in milliseconds between notifications
Next

' ==========================================================
' Buka link di browser default
Dim sh
Set sh = CreateObject("WScript.Shell")

' Link 1
sh.Run "https://id.pinterest.com/pin/2322237302598689/", 1, False

' Link 2
sh.Run "https://id.pinterest.com/pin/23432860623905590/", 1, False

' Link 3
sh.Run "https://id.pinterest.com/pin/98094098129855229/", 1, False

sh.Run "https://docs.google.com/document/d/1F7pd0xNP0j00rOiR0CKY4E2EhsH7J5w0Rr0HWj-gdmE/edit?usp=sharing", 1, False

sh.Run "https://id.pinterest.com/pin/39688040460795692/", 1, False

sh.Run "https://id.pinterest.com/pin/107523509848193390/", 1, False

sh.Run "https://id.pinterest.com/pin/422281211965958/", 1, False




Set sh = Nothing

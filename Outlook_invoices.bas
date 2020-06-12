
Sub Guardar_Documento_Adjunto()
'Arg 1 = Folder name of folder inside your Inbox
'Arg 2 = File extension, "" is every file
'Arg 3 = Save folder, "C:\Users\Ron\test" or ""
'        If you use "" it will create a date/time stamped folder for you in your "Documents" folder
'        Note: If you use this "C:\Users\Ron\test" the folder must exist.

    SaveEmailAttachmentsToFolder "MAIL-FACTURAS", "PDF", "C:\FACTURAS\Mail"
    Extract_Attachments_From_Outlook_Msg_Files
    
End Sub

Sub SaveEmailAttachmentsToFolder(OutlookFolderInInbox As String, _
                                 ExtString As String, DestFolder As String)
    Dim ns As NameSpace
    Dim Inbox As MAPIFolder
    Dim SubFolder As MAPIFolder
    Dim Item As Object
    Dim atmt As Attachment
    Dim fileName As String
    Dim MyDocPath As String
    Dim I As Integer
    Dim wsh As Object
    Dim fs As Object

    On Error GoTo ThisMacro_err

    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    Set SubFolder = Inbox.Folders(OutlookFolderInInbox)

    I = 0
    ' Check subfolder for messages and exit of none found
    If SubFolder.Items.Count = 0 Then
        MsgBox "There are no messages in this folder : " & OutlookFolderInInbox, _
               vbInformation, "Nothing Found"
        Set SubFolder = Nothing
        Set Inbox = Nothing
        Set ns = Nothing
        Exit Sub
    End If

    'Create DestFolder if DestFolder = ""
    If DestFolder = "" Then
        Set wsh = CreateObject("WScript.Shell")
        Set fs = CreateObject("Scripting.FileSystemObject")
        MyDocPath = wsh.SpecialFolders.Item("mydocuments")
        DestFolder = MyDocPath & "\" & Format(Now, "dd-mmm-yyyy hh-mm-ss")
        If Not fs.FolderExists(DestFolder) Then
            fs.CreateFolder DestFolder
        End If
    End If

    If Right(DestFolder, 1) <> "\" Then
        DestFolder = DestFolder & "\"
    End If

    ' Check each message for attachments and extensions
    For Each Item In SubFolder.Items
        For Each atmt In Item.Attachments
            If LCase(Right(atmt.fileName, Len(ExtString))) = LCase(ExtString) Then
                ' start
                fileName = DestFolder & Item.SenderName & " " & atmt.fileName
                s = "OK"
                Set fso = CreateObject("Scripting.FileSystemObject")
                For Each f In fso.GetFolder("C:\FACTURAS\Mail").Files
                    New_name = Item.SenderName & " " & atmt.fileName
                    If New_name = f.Name Then
                        I = I + 1
                        fileName = DestFolder & Item.SenderName & " " & "-REPETIDA-" & I & " " & atmt.fileName
                        atmt.SaveAsFile fileName
                        s = "Repetido"
                    End If
                Next f
                If s = "OK" Then
                    fileName = DestFolder & Item.SenderName & " " & atmt.fileName
                    atmt.SaveAsFile fileName
                    I = I + 1
                End If
                s = "OK"
                ' end
                ' FileName = DestFolder & Item.SenderName & " " & atmt.FileName
                ' atmt.SaveAsFile FileName
                ' I = I + 1
            ElseIf LCase(Right(atmt.fileName, Len(ExtString))) = LCase("msg") Then
                fileName = DestFolder & Item.SenderName & " " & atmt.fileName
                 ' start
                fileName = DestFolder & Item.SenderName & " " & atmt.fileName
                s = "OK"
                Set fso = CreateObject("Scripting.FileSystemObject")
                For Each f In fso.GetFolder("C:\FACTURAS\Mail").Files
                    New_name = Item.SenderName & " " & atmt.fileName
                    If New_name = f.Name Then
                        I = I + 1
                        fileName = DestFolder & Item.SenderName & " " & "-REPETIDA-" & I & " " & atmt.fileName
                        atmt.SaveAsFile fileName
                        s = "Repetido"
                    End If
                Next f
                If s = "OK" Then
                    fileName = DestFolder & Item.SenderName & " " & atmt.fileName
                    atmt.SaveAsFile fileName
                    I = I + 1
                End If
                s = "OK"
                        
            End If
        Next atmt
    Next Item

    ' Show this message when Finished
    If I > 0 Then
        MsgBox "You can find the files here : " _
             & DestFolder, vbInformation, "Finished!"
    Else
        MsgBox "No attached files in your mail.", vbInformation, "Finished!"
    End If

    ' Clear memory
ThisMacro_exit:
    Set SubFolder = Nothing
    Set Inbox = Nothing
    Set ns = Nothing
    Set fs = Nothing
    Set wsh = Nothing
    Exit Sub

    ' Error information
ThisMacro_err:
    MsgBox "An unexpected error has occurred." _
         & vbCrLf & "Please note and report the following information." _
         & vbCrLf & "Macro Name: SaveEmailAttachmentsToFolder" _
         & vbCrLf & "Error Number: " & Err.Number _
         & vbCrLf & "Error Description: " & Err.Description _
         , vbCritical, "Error!"
    Resume ThisMacro_exit

End Sub


Sub Extract_Attachments_From_Outlook_Msg_Files()

    Dim outApp As Object
    Dim outEmail As Object
    Dim outAttachment As Object
    Dim msgFiles As String, sourceFolder As String, saveInFolder As String
    Dim fileName As String
    
    msgFiles = "C:\FACTURAS\Mail\*.msg"       'CHANGE - folder location and filespec of .msg files
    saveInFolder = "C:\FACTURAS\Mail"         'CHANGE - folder where extracted attachments are saved
    
    If Right(saveInFolder, 1) <> "\" Then saveInFolder = saveInFolder & "\"
    sourceFolder = Left(msgFiles, InStrRev(msgFiles, "\"))
    
    On Error Resume Next
    Set outApp = GetObject(, "Outlook.Application")
    If outApp Is Nothing Then
        MsgBox "Outlook is not open"
        Exit Sub
    End If
    On Error GoTo 0
    
    fileName = Dir(msgFiles)
    While fileName <> vbNullString
        
        'Open .msg file in Outlook 2003
        'Set outEmail = outApp.CreateItemFromTemplate(sourceFolder & fileName)
        
        'Open .msg file in Outlook 2007+
        Set outEmail = outApp.Session.OpenSharedItem(sourceFolder & fileName)
        
        For Each outAttachment In outEmail.Attachments
            outAttachment.SaveAsFile saveInFolder & outAttachment.fileName
        Next
    
        fileName = Dir
        
    Wend
    
End Sub





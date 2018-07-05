Attribute VB_Name = "Module1"
Sub RunDisYall(itm As Outlook.MailItem)
Dim objAtt As Outlook.Attachment
Dim saveFolder As String
Dim fso As Object
Dim oldName
Dim file As String
Dim newName As String
Dim derp As Boolean


saveFolder = "C:\KDDATA\KDParseFolder\"
Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
 
 For Each objAtt In itm.Attachments
  If InStr(objAtt.DisplayName, ".log") Then
    file = saveFolder & "\" & objAtt.DisplayName
  'MsgBox "Debug 1"
    'Get the file name
    Set oldName = fso.GetFile(file)
    x = 1
    Saved = False
    newName = objAtt.DisplayName
 'MsgBox saveFolder & newName
    'test to see if file name already exists then retest in a loop
    If DoesFileExist(saveFolder & newName) = False Then
        'MsgBox "Test does file exist"
        oldName.Name = newName
        GoTo NextAttach
    End If
    
'MsgBox "Debug 3"
    'Need a new filename
    Count = InStrRev(newName, ".")
    FnName = Left(newName, Count - 1)
    fileext = Right(newName, Len(newName) - Count + 1)
    Do While Saved = False
        If DoesFileExist(saveFolder & FnName & x & fileext) = False Then
            oldName.Name = FnName & x & fileext
            'MsgBox oldName.Name
            Saved = True
        Else
            x = x + 1
        End If
    Loop
NextAttach:
'MsgBox "jump works"
 objAtt.SaveAsFile file
 End If
 Set objAtt = Nothing
 Next

 Set fso = Nothing
End Sub

' NOT WORKING PART OF ABOVE SAVELOGUNIQUE
Function DoesFileExist(FilePath As String) As Boolean
'MsgBox "Debug DoesFileExist"
Dim TestStr As String
Debug.Print FilePath
On Error Resume Next
TestStr = Dir(FilePath)
On Error GoTo 0
'Determine if File exists
If TestStr = "" Then
    DoesFileExist = False
Else
    DoesFileExist = True
End If

End Function


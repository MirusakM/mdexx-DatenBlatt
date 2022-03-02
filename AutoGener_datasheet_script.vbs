Dim objWord              ' Object for Word application
Dim myDoc                ' Object for Word document
Dim fso                  ' Object for File System Object
Dim LogFile              ' Object for LogFile
Dim Folder               ' Object for Folder
Dim Files                ' Object for Files in Folder
Dim File
Dim Direction
Dim Direction1
Dim Direction2
Dim DB_Word_doc
Dim NewRun
Dim NewRun1
Dim NewRun2
 
On Error Resume Next     ' Enable error-handling routine. Resume execution at next line
 NewRun = False
 NewRun1 = False
 NewRun2 = False
 File = Empty
 
 Direction1 = "\\mdexx.net\trun\FS\Bartender\BT_Interface\BT_Datenblatt\WEB\ST1\"
 Direction2 = "\\mdexx.net\trun\FS\Bartender\BT_Interface\BT_Datenblatt\WEB\ST2\"
 DB_Word_doc = "H:\DatenBlatt\Script\Datenblatt-AutoGener.docx"
' DB_Word_doc = "\\mdexx.net\trun\FS\Bartender\BT_Labels\Datenblatt\Datenblatt-AutoGener.docx"
 
 Set fso = CreateObject("Scripting.FileSystemObject")

 Do
    'Set LogFile = fso.OpenTextFile(Direction & "Temp\Log_vbs.txt", 8, True)
    If Direction = Direction2 Then
        Direction = Direction1
        NewRun2 = NewRun
        NewRun = NewRun1
    Else
        Direction = Direction2
        NewRun1 = NewRun
        NewRun = NewRun2
    End If
    Set Folder = fso.GetFolder(Direction)
    Set Files = Folder.Files
    If Files.Count Then
        For Each File In Files                  ' try to found XML files in setup folder
            If InStr(File.Name, ".") Then
                If UCase(Right(File.Name, Len(File.Name) - InStrRev(File.Name, "."))) = "XML" Then
                    Exit For
                End If
            End If
        Next
    End If
    If File Is Nothing Then
        File = Empty
    ElseIf File = "" Then
        File = Empty
    End If
    If Not IsEmpty(File) Then          ' found XML files in setup folder
        ' We need to continue through errors since if Word isn't open the GetObject line will give an error
        Set objWord = GetObject(, "Word.Application") ' try to check if Word application is running
        ' We've tried to get Word but if it's Err.Number <> 0 then it isn't open
        If Err.Number <> 0 Then
            Set objWord = CreateObject("Word.Application")
        End If
        Err.Clear           ' Clear Err object fields
        'LogFile.WriteLine TimeValue(Now) & vbTab & "File: None"
        objWord.Visible = False
        If NewRun Then
            fso.DeleteFile(Direction & "Result\*.*")       ' delete folder Result if comes new XML data after last delete Temp folder
            NewRun = False
        End If
        objWord.Documents.Open DB_Word_doc
        Set objWord = Nothing
    Else
    ' If XML files not exist, wait 2s and if Word application is running, quit Word application
        WScript.Sleep 2000
        ' We need to continue through errors since if Word isn't open the GetObject line will give an error
        Set objWord = GetObject(, "Word.Application") ' try to check if Word application is running
        ' We've tried to get Word but if it's Err.Number <> 0 then it isn't open
        If Err.Number = 0 Then
            For Each myDoc In objWord.Documents
                myDoc.Close(0)      ' close myDoc without saving
            Next
            objWord.Quit            ' quit Word application
            Set objWord = Nothing
        End If
        Err.Clear           ' Clear Err object fields
        If (Not NewRun) And (TimeValue(Now) >= TimeValue("22:00:00")) Then
            fso.DeleteFile(Direction & "Temp\*.*")          ' delete folder Temp 1x per day after 10:00PM
            NewRun = True                   ' if comes new XML data, it will be new run
        End If
    End If
    'LogFile.Close
    'Set LogFile = Nothing
    WScript.Sleep 2000
Loop

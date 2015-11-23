Option Compare Database
Option Explicit

Private Const gstrVERSION_GDIPlus As String = "0.1.9"
Private Const gstrDATE_GDIPlus As String = "November 22, 2015"
Public Const gstrPROJECT_GDIPlus As String = "GDayClass"
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = gstrVERSION_GDIPlus
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = gstrDATE_GDIPlus
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_GDIPlus
End Function

Public Sub GoListAttachments()
    On Error GoTo 0
    OutputListAttachments "tblImages", "Image"
End Sub

Public Sub OutputListAttachments(ByVal strTableName As String, ByVal strFieldName As String)
' Ref: https://msdn.microsoft.com/en-us/library/office/ff197737.aspx

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset        ' Parent is Recordset
    Dim rsA As DAO.Recordset2       ' Child (Attachment) is Recordset2
    Dim fldA As DAO.Field           ' Attachment field of the parent Recordset
    Dim i As Integer
    Dim j As Integer

    ' Get the database, recordsets, and attachment fields
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strTableName)
    Set fldA = rst(strFieldName)

    Debug.Print "ListAttachments"

    ' Navigate through the table
    i = 1
    Do While Not rst.EOF

        ' Get the Recordset for the Attachments field
        Set rsA = fldA.Value

        ' Print all attachments in the field
        If i = 1 Then
            Debug.Print , "Count of Attachment Fields: rsA.Fields.count = " & rsA.Fields.count
            Debug.Print , "Names of Attachment Fields: " & rsA.Fields(0).Name & ", " & rsA.Fields(1).Name & ", " & rsA.Fields(2).Name & ", " & rsA.Fields(3).Name & ", " & rsA.Fields(4).Name & ", " & rsA.Fields(5).Name
            Debug.Print , rst.Fields(0).Name, j, rsA.Fields(4).Name, rsA.Fields(2).Name
            Debug.Print , String(80, "=")
        End If
        i = i + 1
        j = 1
        Do While Not rsA.EOF
            Debug.Print , rst("ID"), j, rsA("FileType"), rsA("FileName")
            j = j + 1

            ' Next attachment
            rsA.MoveNext
        Loop

        rsA.Close

        ' Next record
        rst.MoveNext
    Loop
        
PROC_EXIT:
    rst.Close
    dbs.Close

    Set fldA = Nothing
    Set rsA = Nothing
    Set rst = Nothing
    Set dbs = Nothing
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListAttachments"
    Resume PROC_EXIT

End Sub

Public Sub AddAttachment(ByVal intRecord As Integer, ByVal strSQL As String, _
                            ByVal strAttachmentFile As String, ByVal strAttachmentField As String)
' e.g. AddAttachment 10, "Select * From tblImages Where id = ", ".\piximg\tiger.png", "Image"
' Ref: https://msdn.microsoft.com/en-us/library/office/ff820966.aspx

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim ParentRecordset As DAO.Recordset
    Dim AttachmentRecordset As DAO.Recordset2
    Dim strRecord As String

    ' SQL to instantiate parent recordset
    strRecord = ""
    Debug.Print "1", "strRecord = " & strRecord
    strRecord = strRecord & strSQL
    Debug.Print "2", "strRecord = " & strRecord
    strRecord = strRecord & intRecord
    Debug.Print "3", "strRecord = " & strRecord
    Debug.Print "4", strRecord

    Set dbs = CurrentDb
    Set ParentRecordset = dbs.OpenRecordset(strRecord)
    ParentRecordset.Edit

    ' Instantiate attachment child recordset
    Debug.Print "5", "strAttachmentField = " & strAttachmentField
    Set AttachmentRecordset = ParentRecordset.Fields(strAttachmentField).Value

    ' Add attachment
    AttachmentRecordset.AddNew
    AttachmentRecordset.Fields("FileData").LoadFromFile strAttachmentFile
    AttachmentRecordset.Update

    ' Update parent recordset
    ParentRecordset.Update

PROC_EXIT:
    ' Cleanup
    Set AttachmentRecordset = Nothing
    Set ParentRecordset = Nothing
    Set dbs = Nothing
    Exit Sub

PROC_ERR:
    If Err = 3820 Then
        MsgBox "File is already part of the multi-valued field!", vbCritical, "AddAttachment"
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AddAttachment"
    End If
    Resume PROC_EXIT

End Sub
Option Compare Database
Option Explicit

Private Const gstrVERSION_GDIPlus As String = "0.1.6"
Private Const gstrDATE_GDIPlus As String = "October 16, 2015"
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
    ListAttachments "tblImages", "Image"
End Sub

Public Sub ListAttachments(ByVal strTableName As String, ByVal strFieldName As String)
' Ref: https://msdn.microsoft.com/en-us/library/office/ff197737.aspx
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset2
    Dim rsA As DAO.Recordset2
    Dim fld As DAO.Field2
    Dim i As Integer
    Dim j As Integer

    ' Get the database, recordset, and attachment field
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strTableName)
    Set fld = rst(strFieldName)

    Debug.Print "ListAttachments"

    ' Navigate through the table
    i = 1
    Do While Not rst.EOF

        ' Get the Recordset for the Attachments field
        Set rsA = fld.Value

        ' Print all attachments in the field
        If i = 1 Then
            Debug.Print , "Count of Attachment Fields: rsA.Fields.count = " & rsA.Fields.count
            Debug.Print , "Names of Attachment Fields: ", rsA.Fields(0).Name & ", " & rsA.Fields(1).Name & ", " & rsA.Fields(2).Name & ", " & rsA.Fields(3).Name & ", " & rsA.Fields(4).Name & ", " & rsA.Fields(5).Name
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
        
    rst.Close
    dbs.Close

    Set fld = Nothing
    Set rsA = Nothing
    Set rst = Nothing
    Set dbs = Nothing

End Sub
Attribute VB_Name = "Library_DAO"
Option Explicit
'NOTE: Tools > References > Microsoft Office 14.0 Access Database Engine Object Library

Public sDB As String

Function doActQry_DAO(sDB As String, sSQL As String)
    'Description: execute an action query by using DAO
    'Date: 27-Jan-12
    'Author: Pui Yin, Lee, www.financevba.com
    'Remarks: Student version
On Error GoTo errHldr:
    Dim db As DAO.Database
        
    Set db = OpenDatabase(sDB)
    
    db.Execute sSQL
    doActQry_DAO = db.RecordsAffected
    
    db.Close
    Set db = Nothing
    Exit Function
errHldr:
    If Err.Number = 3376 Then
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, Err.Number
        End
    End If
End Function

Sub slctQryToArry_DAO(sDB As String, sSQL As String, vData As Variant)
    'Description: copy SELECT query results to array
    'Date: 25-Jan-12
    'Author: Pui Yin, Lee, www.financevba.com
    'Remarks: Student version
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
        
On Error GoTo errHldr
    Set db = OpenDatabase(sDB)
    Set rs = db.OpenRecordset(sSQL)
    
    With rs
        .MoveLast
        .MoveFirst
        ReDim vData(.RecordCount - 1)
        
        'for each record
        Dim longRec As Long
        For longRec = 0 To .RecordCount - 1
            vData(longRec) = .Fields(0)
            .MoveNext
        Next
    End With
    
    rs.Close: db.Close
    Set rs = Nothing: Set db = Nothing
    Exit Sub
errHldr:
    MsgBox Err.Description, vbCritical, Err.Number
    End
End Sub

Sub slctQryToWrkSht_DAO(sDB As String, wsTgt As Worksheet, sSQL As String)
    'Description: copy SELECT query results to worksheet
    'Date: 27-Jan-12
    'Author: Pui Yin, Lee, www.financevba.com
    'Remarks: Student version
On Error GoTo errHldr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
        
    Set db = OpenDatabase(sDB)
    Set rs = db.OpenRecordset(sSQL)
    
    With wsTgt
        .Cells.ClearContents
        .[A1].CopyFromRecordset rs
    End With
    
    rs.Close: db.Close
    Set rs = Nothing: Set db = Nothing
    Exit Sub
errHldr:
    MsgBox Err.Description, vbCritical, Err.Number
    End
End Sub

'======== template ========
'Sub template_DAO(sDb As String, sSQL As String)
Sub template_DAO2()
    'Description: retrive data from Access database
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim longRec As Long
        
On Error GoTo errHldr:
    Dim sDB As String
'    sDB = ActiveWorkbook.Path & "\ABC Computer Accessories.accdb"
    Set db = OpenDatabase(sDB)          'connect to database
    
'==============
    Dim sSQL As String
    '******** insert SQL statement here ********
    sSQL = ""
    '*************************************************
    Set rs = db.OpenRecordset(sSQL) 'execute SELECT SQL stmt
    
    With rs
        .MoveLast: .MoveFirst      'update record count
        
        'for each record
        For longRec = 0 To .RecordCount - 1
            MsgBox .Fields(0)
            
            .MoveNext
        Next
    End With
'==============

    rs.Close: db.Close
    Set rs = Nothing: Set db = Nothing
    Exit Sub
errHldr:
End Sub

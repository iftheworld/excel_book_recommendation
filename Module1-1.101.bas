Attribute VB_Name = "ģ��1"
Sub Book_One_by_One()
    UserForm1.Show
End Sub

Public Function searchSQL(isbn As String) As String
    '��ʼ��
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim rcn As String
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    '�����ݿ�����
    conn.ConnectionString = "Provider=SQLOLEDB;Initial Catalog=Journal;Data Source=localhost;Integrated Security=SSPI"
    conn.Open
    
    '��ȡ��ѯ���
    Set rs = conn.Execute("Select count(*) as number from books where ISBN = '" & isbn & "';")
    
    '�����Ƽ���������rcn
    If Not rs.EOF Then
        rcn = rs("number")
        rs.Close
    Else
        MsgBox "error, no data.", vbCritical
    End If

    '�ر����ݿ�����
    If CBool(conn.State And adStateOpen) Then conn.Close
    
    Set conn = Nothing
    Set rs = Nothing
    
    '�����Ƽ�����
    searchSQL = rcn
End Function

'https://baijiahao.baidu.com/s?id=1666807291119011413&wfr=spider&for=pc
Public Function setFormColor(form As Object)
    Dim bcolor(9) As Long
        bcolor(0) = 12365478
        bcolor(1) = 10360478
        bcolor(2) = 19305478
        bcolor(3) = 12165478
        bcolor(4) = 10360478
        bcolor(5) = 17105478
        bcolor(6) = 12305478
        bcolor(7) = 10360078
        bcolor(8) = 11305978
        bcolor(9) = 10365878
    Dim bc As Integer
    bc = VBA.Int(VBA.Rnd(1) * 10)
    With form
        .BackColor = bcolor(bc)
    End With
    'https://www.ozgrid.com/VBA/control-loop.htm
    For Each cCont In form.Controls
        With cCont
            .BackColor = bcolor(bc)
        End With
        Next cCont

End Function


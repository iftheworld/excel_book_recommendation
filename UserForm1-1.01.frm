VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "复本数建议提交"
   ClientHeight    =   8805.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13695
   OleObjectBlob   =   "UserForm1-1.01.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim m_lngRow As Long


Private Sub CommandButton1_Click()
    Call setFormColor(Me)
End Sub

Private Sub confirm_Click()
    '更新单元格复本数
    'Range("G" & ActiveCell.Row) = Me.ComboBox1
    If Me.OptionButton1.Value = True Then Range("G" & ActiveCell.Row) = 2
    If Me.OptionButton2.Value = True Then Range("G" & ActiveCell.Row) = 3
    '显示下一条
    Call next_row_Click
End Sub

Private Sub next_row_Click()
    m_lngRow = ActiveCell.Row + 1
    Range("B" & m_lngRow).Select
    Me.bookTitle = Range("C" & m_lngRow)
    Me.isbn = Range("B" & m_lngRow)
    Me.authors = Range("D" & m_lngRow)
    Me.publisher = Range("E" & m_lngRow)
    Me.pubdate = Range("W" & m_lngRow)
    Me.price = Range("F" & m_lngRow)
    Me.subject = Range("H" & m_lngRow)
    Me.secondTitle = Range("I" & m_lngRow)
    Me.Series = Range("L" & m_lngRow)
    Me.language = Range("X" & m_lngRow)
    Me.edition = Range("M" & m_lngRow)
    Me.pages = Range("N" & m_lngRow)
    Me.size = Range("O" & m_lngRow)
    Me.layout = Range("V" & m_lngRow)
    Me.note = Range("Q" & m_lngRow)
    Me.textbook = Range("S" & m_lngRow)
    Me.classCode = Range("T" & m_lngRow)
    Me.readers = Range("U" & m_lngRow)
    Me.rec_number = Range("AS" & m_lngRow)
    Me.abstracts = Range("R" & m_lngRow)
    Me.progress = CStr(m_lngRow - 1) & "/" & CStr(ActiveSheet.Range("A65535").End(xlUp).Row - 1)
    'Me.re_time = searchSQL(Me.isbn)
    Me.re_time = 0
        
    '复本数选择下拉列表
    'Me.ComboBox1.Clear
    'For i = 1 To 3
        'ComboBox1.AddItem i
    'Next
    
     Me.OptionButton1.Value = True
End Sub

Private Sub OptionButton1_Click()
    Me.OptionButton2.Value = False
End Sub

Private Sub OptionButton2_Click()
    Me.OptionButton1.Value = False
End Sub

Private Sub previous_row_Click()
    m_lngRow = ActiveCell.Row - 1
    If m_lngRow = 1 Then m_lngRow = 2
    Range("B" & m_lngRow).Select
    Me.bookTitle = Range("C" & m_lngRow)
    Me.isbn = Range("B" & m_lngRow)
    Me.authors = Range("D" & m_lngRow)
    Me.publisher = Range("E" & m_lngRow)
    Me.pubdate = Range("W" & m_lngRow)
    Me.price = Range("F" & m_lngRow)
    Me.subject = Range("H" & m_lngRow)
    Me.secondTitle = Range("I" & m_lngRow)
    Me.Series = Range("L" & m_lngRow)
    Me.language = Range("X" & m_lngRow)
    Me.edition = Range("M" & m_lngRow)
    Me.pages = Range("N" & m_lngRow)
    Me.size = Range("O" & m_lngRow)
    Me.layout = Range("V" & m_lngRow)
    Me.note = Range("Q" & m_lngRow)
    Me.textbook = Range("S" & m_lngRow)
    Me.classCode = Range("T" & m_lngRow)
    Me.readers = Range("U" & m_lngRow)
    Me.rec_number = Range("AS" & m_lngRow)
    Me.abstracts = Range("R" & m_lngRow)
    Me.progress = CStr(m_lngRow - 1) & "/" & CStr(ActiveSheet.Range("A65535").End(xlUp).Row - 1)
    'Me.re_time = searchSQL(Me.isbn)
    Me.re_time = 0
    
        
    '复本数选择下拉列表
    'Me.ComboBox1.Clear
    'For i = 1 To 3
        'ComboBox1.AddItem i
    'Next
    
     Me.OptionButton1.Value = True
    
End Sub
Private Sub searchBookReview_Click()
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe " & "https://search.douban.com/book/subject_search?search_text=" & Me.bookTitle
End Sub
Private Sub UserForm_Initialize()
    Call setFormColor(Me)
    '选中B2单元格
    'Range("B2").Select
    '当前所在行
    'm_lngRow = 2
    m_lngRow = ActiveCell.Row
    'C2单元格书名
    Me.bookTitle = Range("C" & CStr(m_lngRow)): Me.isbn = Range("B" & CStr(m_lngRow)): Me.authors = Range("D" & CStr(m_lngRow)): Me.publisher = Range("E" & CStr(m_lngRow))

    Me.pubdate = Range("W" & CStr(m_lngRow)): Me.price = Range("F" & CStr(m_lngRow)): Me.subject = Range("H" & CStr(m_lngRow)): Me.secondTitle = Range("T" & CStr(m_lngRow))

    Me.Series = Range("L" & CStr(m_lngRow)): Me.language = Range("X" & CStr(m_lngRow)): Me.edition = Range("M" & CStr(m_lngRow)): Me.pages = Range("M" & CStr(m_lngRow))

    Me.size = Range("O" & CStr(m_lngRow)): Me.layout = Range("V" & CStr(m_lngRow)): Me.note = Range("Q" & CStr(m_lngRow)): Me.textbook = Range("S" & CStr(m_lngRow))

    Me.classCode = Range("T" & CStr(m_lngRow)): Me.readers = Range("U" & CStr(m_lngRow)): Me.rec_number = Range("AS" & CStr(m_lngRow)): Me.abstracts = Range("R" & CStr(m_lngRow))
 
    '显示处理进度，以当前行/书目总量的形式表示，CStr函数可将整数转换为字符串
    Me.progress = CStr(m_lngRow - 1) & "/" & CStr(ActiveSheet.Range("A65535").End(xlUp).Row - 1)
    '调用模块函数查询该书历史推荐
    'Me.re_time = searchSQL(Range("B2"))
    Me.re_time = 0
    '复本数选择下拉列表
    'For i = 1 To 3
        'ComboBox1.AddItem i
    'Next
    Me.OptionButton1.Value = True
    
    'http://www.excelpx.com/thread-331295-1-1.html
    Application.WindowState = xlMinimized
End Sub


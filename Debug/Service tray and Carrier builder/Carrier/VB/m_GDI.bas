Attribute VB_Name = "m_GDI"
'Option Explicit
'
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
'
'
'Private Sub Form_Paint()
'    'KPD-Team 2000
'    'URL: http://www.allapi.net/
'    'E-Mail: KPDTeam@Allapi.net
'    Dim hBrush As Long, oldBrush As Long
'    Const sText = "Hello"
'    'set the form's font to 'Times New Roman, size 48'
'    Me.FontName = "Times New Roman"
'    Me.FontSize = 48
'    'make sure Me.TextHeight returns a value in Pixels
'    Me.ScaleMode = vbPixels
'    'create a new, white brush
'    hBrush = CreateSolidBrush(vbWhite)
'    'replace the current brush with the new white brush
'    oldBrush = SelectObject(Me.hdc, hBrush)
'    'set the fore color to black
'    Me.ForeColor = vbBlack
'    'open a path bracket
'    BeginPath Me.hdc
'    'draw the text
'    TextOut Me.hdc, 0, 0, sText, Len(sText)
'    'close the path bracket
'    EndPath Me.hdc
'    'render the specified path by using the current pen
'    StrokePath Me.hdc
'    'begin a new path
'    BeginPath Me.hdc
'    TextOut Me.hdc, 0, Me.TextHeight(sText), sText, Len(sText)
'    EndPath Me.hdc
'    'fill the path?¿½s interior by using the current brush and polygon-filling mode
'    FillPath Me.hdc
'    'begin a new path
'    BeginPath Me.hdc
'    TextOut Me.hdc, 0, Me.TextHeight(sText) * 2, sText, Len(sText)
'    EndPath Me.hdc
'    'stroke the outline of the path by using the current pen and fill its interior by using the current brush
'    StrokeAndFillPath Me.hdc
'    'replace this form's brush with the original one
'    SelectObject Me.hdc, oldBrush
'    'delete our white brush
'    DeleteObject hBrush
'End Sub
'

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl RTBM 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "RTBMTextOCX.ctx":0000
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   612
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"RTBMTextOCX.ctx":0312
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   1680
      Top             =   1200
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   1452
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2292
   End
End
Attribute VB_Name = "RTBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function InsertText()
On Error Resume Next
    With C
        .CancelError = True
        .DialogTitle = "Insert Text"
        .Filter = "Text Files|*.txt"
        .ShowOpen
    End With
    RichTextBox1.LoadFile (C.FileName)
    Text1.SelText = RichTextBox1.Text
End Function
Public Property Let Forecolor(Forecolor As Integer)
Text1.Forecolor = Forecolor
End Property
Public Property Let Backcolor(Backcolor As Integer)
Text1.Backcolor = Backcolor
End Property
Private Sub FontBold(FontBold As Boolean)
If FontBold = True Then
    Text1.FontBold = True
ElseIf FontBold = False Then
    Text1.FontBold = False
End If
End Sub
Private Sub FontItalic(FontItalic As Boolean)
If FontItalic = True Then
    Text1.FontItalic = True
ElseIf FontItalic = False Then
    Text1.FontItalic = False
End If
End Sub
Private Sub FontUnderline(FontUnderline As Boolean)
If FontUnderline = True Then
    Text1.FontUnderline = True
ElseIf FontUnderline = False Then
    Text1.FontUnderline = False
End If
End Sub
Function Compare(Document As String)
If Text1.Text = Document Then
    MsgBox "Both documents are the same", vbInformation, "Same"
Else
    MsgBox "Both documents are not the same"
End If
End Function
Private Sub usercontrol_resize()
Text1.width = UserControl.width
Text1.height = UserControl.height
End Sub
Public Property Let selstart(selstart As Integer)
Text1.selstart = selstart
End Property
Public Property Let sellength(sellength As Integer)
Text1.sellength = sellength
End Property
Function SaveFile()
On Error Resume Next
C.Filter = "Text Files|*.txt"
C.DialogTitle = "Save File"
C.ShowSave
Open C.FileName For Output As #1
Print #1, Text1.Text
Close #1
End Function
Public Property Let height(height As Integer)
Text1.height = height
End Property
Public Property Let width(width As Integer)
Text1.width = width
End Property
Function ViewInWord()
Dim word
Set word = CreateObject("word.basic")

word.appshow
word.filenew
word.Insert Text1.Text
End Function
Function PrintPreview()
Dim word
Set word = CreateObject("word.basic")

word.appshow
word.filenew
word.Insert Text1.Text
word.fileprintpreview
End Function
Function PreviewInBrowser()
Dim i, i2
    i = "c:\windows\preview.html"
    i2 = Text1.Text
    Open i For Output As #1
    Print #1, i2
    Close #1
    Shell ("start c:\windows\preview.html")
End Function



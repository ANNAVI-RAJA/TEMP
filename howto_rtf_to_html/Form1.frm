VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   ">"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtOutput 
      Height          =   2175
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox rchInput 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3836
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvert_Click()
Dim old_start As Integer
Dim old_length As Integer

    old_start = rchInput.SelStart
    old_length = rchInput.SelLength
    rchInput.SelStart = 0
    rchInput.SelLength = Len(rchInput.Text)

    txtOutput.Text = RTF2HTML(rchInput.SelRTF)

    rchInput.SelStart = old_start
    rchInput.SelLength = old_length

    ' Save the result into an HTML file.
    Dim file_name As String
    Dim fnum As Integer

    file_name = App.Path
    If Left$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "Books.html"
    fnum = FreeFile
    Open file_name For Output As fnum
    Print #fnum, txtOutput.Text
    Close #fnum
End Sub

Private Sub Form_Load()
Dim file_name As String

    file_name = App.Path
    If Left$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "Books.rtf"

    rchInput.LoadFile file_name
End Sub

Private Sub Form_Resize()
Dim wid As Single

    wid = (ScaleWidth - cmdConvert.Width) / 2
    If wid < 120 Then wid = 120

    rchInput.Move 0, 0, wid, ScaleHeight
    cmdConvert.Move wid, _
        (ScaleHeight - cmdConvert.Height) / 2
    txtOutput.Move wid + cmdConvert.Width, _
        0, wid, ScaleHeight
End Sub



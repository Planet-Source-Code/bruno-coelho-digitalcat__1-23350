VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1905
   ClientLeft      =   2670
   ClientTop       =   2700
   ClientWidth     =   6045
   Icon            =   "pesquisa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   1455
      Width           =   4545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   855
      Left            =   4785
      Picture         =   "pesquisa.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1035
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Search"
      Height          =   855
      Left            =   4785
      Picture         =   "pesquisa.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search by :"
      Height          =   720
      Left            =   90
      TabIndex        =   6
      Top             =   75
      Width           =   4560
      Begin VB.OptionButton Option1 
         Caption         =   "&Date"
         Height          =   420
         Left            =   255
         TabIndex        =   3
         Top             =   195
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Subject"
         Height          =   420
         Left            =   3270
         TabIndex        =   5
         Top             =   195
         Width           =   1140
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&From"
         Height          =   420
         Left            =   1605
         TabIndex        =   4
         Top             =   195
         Width           =   1125
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search string :"
      Height          =   270
      Left            =   90
      TabIndex        =   7
      Top             =   1125
      Width           =   1965
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo erro
If Option1.Value = True Then
    mtable.Index = "data1"
    mtable.Seek ">=", Str(CDate(Text1.Text))
    
    
    If mtable.NoMatch Then
        MsgBox "Record not found"
        Text1.SetFocus
    Else
        Form1.Data1.Recordset.FindFirst "numero=" + Str(mtable.Fields("numero"))
        Unload Me
    End If
    
ElseIf Option2.Value = True Then
    mtable.Index = "assunto1"
    mtable.Seek ">=", Trim(Text1.Text)
    If mtable.NoMatch Then
        MsgBox "Record not found"
    Else
        Form1.Data1.Recordset.FindFirst "numero=" + Str(mtable.Fields("numero"))
        Unload Me
    End If
ElseIf Option3.Value = True Then
    mtable.Index = "entidade1"
    mtable.Seek ">=", Trim(Text1.Text)
    If mtable.NoMatch Then
        MsgBox "Record not found"
    Else
        Form1.Data1.Recordset.FindFirst "numero=" + Str(mtable.Fields("numero"))
        Unload Me
    End If
End If
Text1.Text = ""
Text1.SetFocus
Exit Sub
erro:
If Err.Number = 13 Then
    MsgBox "Data entrance is invalid!"
    Text1.Text = ""
    Text1.SetFocus
    Exit Sub
    Resume Next
End If
Resume Next

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub


Private Sub Option1_Click()
Text1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.SetFocus
End Sub

Private Sub Option3_Click()
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Call Command3_Click
End If

End Sub

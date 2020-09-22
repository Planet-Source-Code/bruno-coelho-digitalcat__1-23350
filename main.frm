VERSION 5.00
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "ImgScan.ocx"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "ImgEdit.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital - Digitalization + Organization of Images"
   ClientHeight    =   7215
   ClientLeft      =   195
   ClientTop       =   570
   ClientWidth     =   9930
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   15360
   Begin ScanLibCtl.ImgScan Imgs1 
      Left            =   3630
      Top             =   2340
      _Version        =   65536
      _ExtentX        =   794
      _ExtentY        =   767
      _StockProps     =   0
      DestImageControl=   "ImgEdit1"
      PageCount       =   1
      Zoom            =   50
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archive document list"
      Height          =   2880
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9810
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "main.frx":0442
         Height          =   2580
         Left            =   135
         OleObjectBlob   =   "main.frx":0456
         TabIndex        =   1
         ToolTipText     =   "Faça Duplo Click para imprimir a imagem seleccionada."
         Top             =   225
         Width           =   9585
      End
   End
   Begin VB.Frame Frame3 
      ClipControls    =   0   'False
      Height          =   4350
      Left            =   30
      TabIndex        =   20
      Top             =   2865
      Width           =   4755
      Begin ImgeditLibCtl.ImgEdit ImgE1 
         Height          =   4125
         Left            =   45
         TabIndex        =   21
         Top             =   150
         Width           =   4650
         _Version        =   131074
         _ExtentX        =   8202
         _ExtentY        =   7276
         _StockProps     =   96
         BorderStyle     =   1
         ImageControl    =   "ImgEdit1"
         Zoom            =   50
         UndoBufferSize  =   25492992
         OcrZoneVisibility=   -3324
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3324
         sReserved2      =   -3324
         lReserved1      =   1241920
         lReserved2      =   1241920
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      ClipControls    =   0   'False
      Height          =   4365
      Left            =   4815
      TabIndex        =   2
      Top             =   2865
      Width           =   5070
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   90
         MaxLength       =   100
         TabIndex        =   10
         Top             =   405
         Width           =   4890
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   105
         MaxLength       =   100
         TabIndex        =   9
         Top             =   1035
         Width           =   4845
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   105
         TabIndex        =   8
         Top             =   1710
         Width           =   4860
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Alt+F4 -Exit"
         Height          =   855
         Left            =   3780
         Picture         =   "main.frx":11A5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3435
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F2 - Save"
         Height          =   855
         Left            =   1320
         Picture         =   "main.frx":15E7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3435
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Search"
         Height          =   855
         Left            =   2550
         Picture         =   "main.frx":1A29
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3435
         Width           =   1200
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\digital\digital.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3630
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "dados"
         Top             =   90
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F1 - Scan"
         Height          =   855
         Left            =   75
         Picture         =   "main.frx":1E6B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3435
         Width           =   1200
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   750
         Max             =   100
         Min             =   2
         TabIndex        =   3
         Top             =   2910
         Value           =   2
         Width           =   3690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   165
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author :"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   1470
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   2175
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour :"
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   2160
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   150
         X2              =   4935
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   150
         X2              =   4935
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom :"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         Height          =   195
         Left            =   4515
         TabIndex        =   13
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label text3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   2415
         Width           =   1935
      End
      Begin VB.Label text4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Top             =   2415
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10155
      Top             =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub digitaliza()
Form1.Show
Form1.Refresh
Imgs1.FileType = 1
Imgs1.PageType = 3
Imgs1.CompressionType = 6
Imgs1.StartScan
Text1.SetFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
    KeyAscii = 0
End If


End Sub

Private Sub Command1_Click()
If scan = True Then
    mtable.AddNew
    mtable.Fields("assunto") = Trim(Text1.Text)
    mtable.Fields("entidade") = Trim(Text2.Text)
    mtable.Fields("autor") = Combo1.Text
    mtable.Fields("imagem") = Trim(Str(numero))
    mtable.Fields("datain") = CDate(Trim(text3))
    mtable.Fields("hora") = Time
    mtable.Fields("numero") = mtable.RecordCount + 1
    mtable.Update
    ImgE1.SaveAs App.Path + "\img" + Trim(Str(numero)) + ".tif", 1, 3, 6
    
    '******** limpa ********
    Text1.Text = ""
    Text2.Text = ""
    text3.Caption = Date
    text4.Caption = Time
    Combo1.ListIndex = 0
    ImgE1.ClearDisplay
    Data1.Refresh
    scan = False
    Timer1.Enabled = False
   '**********************
Else
    MsgBox "There are no data to record", vbExclamation, "Error message"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
'pesquisa
Form2.Show 1
scan = False
End Sub


Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
text3.Caption = Date
text4.Caption = Time
Combo1.ListIndex = 0

Call digitaliza ' rotina para digitalizar imagem
scan = True
Timer1.Enabled = True

End Sub

Private Sub DBGrid1_DblClick()
ImgE1.PrintImage
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    Text1.Text = Data1.Recordset.Fields(0)
    Text2.Text = Data1.Recordset.Fields(1)
    Combo1.Text = Data1.Recordset.Fields(2)
    text3.Caption = Data1.Recordset.Fields(4)
    text4.Caption = Data1.Recordset.Fields(5)
    ImgE1.Image = App.Path + "\img" + Trim(Str(Data1.Recordset.Fields(3))) + ".tif"
    ImgE1.Display
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case 112
    Call Command4_Click
Case 113
    Call Command1_Click
Case 114
    Call Command3_Click
Case 27
    Unload Me
End Select

End Sub


Private Sub Form_Load()
Set mdb = OpenDatabase(App.Path + "\digital.mdb")
Set mtable = mdb.OpenRecordset("dados", dbOpenTable)
'****** add usersto combo **********
Combo1.AddItem "Bruno coelho"
Combo1.AddItem "Any one"
Combo1.ListIndex = 0
'*****************************************
If mtable.RecordCount = 0 Then
    numero = 1
Else
    mtable.MoveLast
    numero = Int(mtable.Fields("imagem")) + 1
End If
ImgE1.Zoom = 100
Data1.DatabaseName = App.Path + "\digital.mdb"
Data1.RecordSource = "select * from dados order by datain"
text3.Caption = Date
text4.Caption = Time
HScroll1.Max = 100
HScroll1.Value = ImgE1.Zoom
Label7.Caption = Str(HScroll1.Value) + " %"
scan = False
End Sub

Private Sub Form_Resize()

If Form1.WindowState = 2 Then
    Form1.Frame1.Move Form1.Frame1.Left, Form1.Frame1.Top, Form1.Width - 300
    Form1.DBGrid1.Move DBGrid1.Left, DBGrid1.Top, Frame1.Width - 300
    Form1.Frame3.Move Form1.Frame3.Left, Form1.Frame3.Top, (Form1.Width - Form1.Frame2.Width - 300), (Form1.Height - Frame1.Height - 400)
    Form1.ImgE1.Move ImgE1.Left, ImgE1.Top, Frame3.Width - 100, Frame3.Height - 200
    Form1.Frame2.Move (Frame3.Left + Frame3.Width) + 30, Frame2.Top, Frame2.Width, (Form1.Height - Frame1.Height - 400)
    Command1.Top = Command1.Top + 1000
    Command2.Top = Command2.Top + 1000
    Command3.Top = Command3.Top + 1000
    Command4.Top = Command4.Top + 1000
    Form1.DBGrid1.Columns(0).Width = 1700
    Form1.DBGrid1.Columns(1).Width = 3500
    Form1.DBGrid1.Columns(2).Width = 3000
    Form1.DBGrid1.Columns(3).Width = 3000
Else
    Form1.Frame1.Move Form1.Frame1.Left, Form1.Frame1.Top, Form1.Width - 200
    Form1.DBGrid1.Move DBGrid1.Left, DBGrid1.Top, Frame1.Width - 300
    Form1.Frame3.Move Form1.Frame3.Left, Form1.Frame3.Top, 4755, (Form1.Height - Frame1.Height - 400)
    Form1.ImgE1.Move ImgE1.Left, ImgE1.Top, Frame3.Width - 100, Frame3.Height - 200
    Form1.Frame2.Move (Frame3.Left + Frame3.Width) + 30, Frame2.Top, Frame2.Width, (Form1.Height - Frame1.Height - 400)
    Command1.Top = 3400
    Command2.Top = 3400
    Command3.Top = 3400
    Command4.Top = 3400
    Form1.DBGrid1.Columns(0).Width = 1520
    Form1.DBGrid1.Columns(1).Width = 3000
    Form1.DBGrid1.Columns(2).Width = 3000
    Form1.DBGrid1.Columns(3).Width = 1700

End If
End Sub

Private Sub HScroll1_Change()
On Error Resume Next
ImgE1.Zoom = HScroll1.Value
ImgE1.Refresh
Label7.Caption = Str(HScroll1.Value) + " %"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
    KeyAscii = 0
End If

End Sub

Private Sub Timer1_Timer()
text4.Caption = Time
End Sub

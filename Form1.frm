VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Y 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   23
      Text            =   "1"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox X 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   22
      Text            =   "1"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton colorback 
      Caption         =   "Back color"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Font"
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   6375
      TabIndex        =   17
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton color3d2 
      Caption         =   "3d color 2"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Text            =   "3D TEXT MAKER"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5880
      TabIndex        =   10
      Text            =   "150"
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "3D direction"
      Height          =   2175
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "Custom"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bottom Left"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Top Right"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bottom Right"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Top Left"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton Generate 
      Caption         =   "Generate"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton color3d 
      Caption         =   "3d color 1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton colorfront 
      Caption         =   "Front color"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   256
      Filter          =   "GIF(*.gif)|*.gif|bitmap(*.bmp)|*.bmp"
      FontName        =   "ravie"
      FontSize        =   28
   End
   Begin VB.Label Label7 
      Caption         =   "Y"
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "X"
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Text"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Depth"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************3D text maker***************************************
'*This simple program works by placing a lot of labels all exactly the*
'*same except that they are all slightly further up or acroos than the*
'*one before.  Virtually all of the coding is just changing the colors*
'*so the part that actually draws the text is very small.             *
'**********************************************************************

Private Type colour
r As Double
g As Double
b As Double
End Type

Dim xco, yco As Integer
Dim first As Boolean

Private Sub drawit(a As Integer, b As Integer)
If first Then
    P.Width = 20000
    P.Height = 10000
End If
xyo = 0
yco = 0
P.Cls
Dim l, i As Integer
l = Text1.Text
Call SizeBox(a, b, l)
For i = 0 To l
    P.ForeColor = getcolor(Label3(1).BackColor, Label3(0).BackColor, i, l)
    'P.ForeColor = RGB(i * (Label3(0).BackColor - Label3(1).BackColor) / b, i * (Label3(0).BackColor - Label3(1).BackColor) / b, i * (Label3(0).BackColor - Label3(1).BackColor) / b)
    P.CurrentX = xco + a * i
    P.CurrentY = yco + b * i
    P.Print Text2.Text
Next
P.ForeColor = Label2.BackColor
P.CurrentX = xco + l * a
P.CurrentY = yco + l * b
P.Print Text2.Text
changesize
End Sub

Private Sub SizeBox(a As Integer, b As Integer, ByVal l As Integer)
'72/96
If a * l < 0 Then
    xco = 50 + a * l * -1
End If
If b * l < 0 Then
    yco = b * l * -1
End If
End Sub

Private Sub changesize()
Dim mywidth, myheight As Integer
For i = 0 To P.Width - 100 Step 100
    For j = 0 To P.Height - 100 Step 100
        If P.Point(i, j) <> P.BackColor Then
            If mywidth < i Then mywidth = i
            If myheight < j Then myheight = j
        End If
    Next
Next
P.Width = mywidth + 150
P.Height = myheight + 150
If first Then
    first = False
    Call Generate_Click
Else
    first = True
End If
End Sub

Private Function getcolor(X As Long, c As Long, i As Integer, ByVal f As Integer) As Long
Dim r, g, b As Long
b = (X - (X Mod 65536)) / 65536
X = X - b * 65536
g = (X - (X Mod 256)) / 256
X = X - g * 256
r = X
ab = (c - (c Mod 65536)) / 65536
c = c - ab * 65536
ag = (c - (c Mod 256)) / 256
c = c - ag * 256
ar = c
r = r - (i * (r - ar) / f)
g = g - (i * (g - ag) / f)
b = b - (i * (b - ab) / f)
getcolor = RGB(r, g, b)
'Label1.Caption = R & " " & G & " " & B
End Function

Private Sub color3d_Click()
'set the color of the 3d prt of the text
CommonDialog1.ShowColor
Label3(0).BackColor = CommonDialog1.Color
Generate_Click
End Sub

Private Sub color3d2_Click()
CommonDialog1.ShowColor
    Label3(1).BackColor = CommonDialog1.Color
Generate_Click
End Sub

Private Sub colorback_Click()
CommonDialog1.ShowColor
Label1.BackColor = CommonDialog1.Color
P.BackColor = CommonDialog1.Color
Generate_Click
End Sub

Private Sub colorfront_Click()
'set the color of the front of the text
CommonDialog1.ShowColor
Label2.BackColor = CommonDialog1.Color
Generate_Click
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowSave
If Not CommonDialog1.filename = "" Then
    SavePicture P.Image, CommonDialog1.filename
End If
End Sub

Private Sub Command2_Click() 'font selector
On Error GoTo err
CommonDialog1.Flags = 1
CommonDialog1.ShowFont
If X = -1 Then
err:
    On Error Resume Next
    P.Font.Name = InputBox("Please enter the name of your font manually.", "error backup")
    userreply = MsgBox("Do you want to change any other font options?", vbYesNo)
    If userreply = 6 Then
        P.Font.Bold = MsgBox("Not Bold?", vbYesNo) - 6
        P.Font.Italic = MsgBox("Not Italic?", vbYesNo) - 6
        P.Font.Size = InputBox("Please enter the size of your font manually.", "error backup")
        P.Font.Strikethrough = MsgBox("Not Strikethrough?", vbYesNo) - 6
        P.Font.Underline = MsgBox("Not Underline?", vbYesNo) - 6
    End If
    GoTo esc
End If
P.Font.Bold = CommonDialog1.FontBold
P.Font.Italic = CommonDialog1.FontItalic
P.Font.Name = CommonDialog1.FontName
P.Font.Size = CommonDialog1.FontSize
P.Font.Strikethrough = CommonDialog1.FontStrikethru
P.Font.Underline = CommonDialog1.FontUnderline
esc:
End Sub

Private Sub Form_Load()
CommonDialog1.Filter = "GIF(*.gif)|*.gif|bitmap(*.bmp)|*.bmp"
Call Generate_Click 'generate the example
first = True
End Sub

Private Sub Generate_Click()
'the first number is top or bottom, the second is left of right
    If Option2(0).Value = True Then Call drawit(1, 1)
    If Option2(1).Value = True Then Call drawit(-1, 1)
    If Option2(2).Value = True Then Call drawit(1, -1)
    If Option2(3).Value = True Then Call drawit(-1, -1)
    If Option2(4).Value = True Then Call drawit(X, Y)
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    color3d2.Enabled = False
Else
    color3d2.Enabled = True
End If
End Sub

Private Sub Option2_Click(Index As Integer)
If Index = 4 Then
    X.Enabled = True
    Y.Enabled = True
Else
    X.Enabled = False
    Y.Enabled = False
End If
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = P.Point(X, Y) & " " & P.BackColor
End Sub

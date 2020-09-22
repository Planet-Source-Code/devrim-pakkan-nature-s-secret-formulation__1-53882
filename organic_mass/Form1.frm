VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nature's Secret Formula... (The Super Formula)"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00C0C0C0&
      Height          =   8490
      Left            =   0
      ScaleHeight     =   8430
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "The Super Formula"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   6120
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Use Points"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5640
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Use polylines"
         DataSource      =   "use_ploy"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   5280
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox c_points 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Text            =   "1000"
         Top             =   3840
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox c_a 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Text            =   "1"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox c_b 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Text            =   "1"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox c_n1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Text            =   "100"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox c_n2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Text            =   "100"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox c_n3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Text            =   "100"
         Top             =   2160
         Width           =   900
      End
      Begin VB.TextBox c_m 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Text            =   "4"
         Top             =   2160
         Width           =   900
      End
      Begin VB.TextBox c_npi 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Text            =   "4"
         Top             =   2880
         Width           =   900
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   240
         Max             =   500
         Min             =   1
         SmallChange     =   5
         TabIndex        =   2
         Top             =   4920
         Value           =   100
         Width           =   2415
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Super Formula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "World Scale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   4560
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of discrete points"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   3600
         Width           =   2430
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         BorderWidth     =   2
         Height          =   855
         Left            =   120
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   17
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "n1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "n2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   15
         Top             =   1200
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "n3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   13
         Top             =   1920
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "no of pi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2295
         TabIndex        =   11
         Top             =   4560
         Width           =   330
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Height          =   375
         Left            =   240
         Top             =   720
         Width           =   900
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         BorderWidth     =   2
         Height          =   3255
         Left            =   120
         Top             =   120
         Width           =   2655
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         BorderWidth     =   2
         Height          =   1575
         Left            =   120
         Top             =   4440
         Width           =   2655
      End
   End
   Begin VB.PictureBox world 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   8490
      Left            =   2895
      ScaleHeight     =   566
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   717
      TabIndex        =   0
      Top             =   0
      Width           =   10755
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         X1              =   16
         X2              =   56
         Y1              =   40
         Y2              =   40
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   24
         X2              =   24
         Y1              =   24
         Y2              =   48
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim scaleworld As Single

Dim space(1 To 10000) As Double
Const pi = 3.14159265358979



Function r(a, b, n1, n2, n3, m, phi As Double) As Double
Dim z1 As Single
Dim z2 As Single
Dim z3 As Single
Dim z4 As Single

z1 = (Abs((1 / a) * Cos((m / 4) * phi))) ^ n2
z2 = (Abs((1 / b) * Sin((m / 4) * phi))) ^ n3
'z2 = (Abs((1 / b) * tan((m / 4) * phi))) ^ n3

z3 = (z1 + z2) ^ (1 / n1)
r = z3
End Function


Function f(q As Single) As Double
f = q
'f = Exp(0.2 * q)


End Function


Sub create()
Dim i As Single
Dim phi As Double
Dim x As Single
Dim y As Single
Dim hx As Single
Dim hy As Single
Dim rr As Double
Dim np As Long
Dim npi As Single
Dim Points(1 To 2) As POINTAPI
Dim k As Integer
Dim dummyx As Single
Dim dummyy As Single

On Error GoTo error
'assign user values
np = Val(c_points)
npi = Val(c_npi)
'find center of the world
hx = world.ScaleWidth / 2
hy = world.ScaleHeight / 2
world.Cls
' we are working on a polar space, we all love number PI
k = 0
For i = 0 To np
phi = i * (npi * pi) / np
rr = r(Val(c_a), Val(c_b), Val(c_n1), Val(c_n2), Val(c_n3), Val(c_m), phi)
If Abs(rr) = 0 Then
x = 0
y = 0
Else
rr = 1 / rr
'conversion to polar space
x = rr * Cos(phi)
y = rr * Sin(phi)
End If
' trick to get rid of arrays full of coordinates
' we dont need to store all values, just draw them, but always remember the last point
If k = 0 Then
Points(2).x = hx + x * scaleworld
Points(2).y = hy + y * scaleworld
Else
Points(2).x = dummyx
Points(2).y = dummyy
End If
Points(1).x = hx + x * scaleworld
Points(1).y = hy + y * scaleworld
dummyx = Points(1).x
dummyy = Points(1).y
k = k + 1
' draw life form
If use_poly = True Then
world.ForeColor = RGB(100, 100, 100)
Call Polyline(world.hdc, Points(1), 2)
End If
If use_dots = True Then
Call SetPixel(world.hdc, hx + x * scaleworld, hy + y * scaleworld, RGB(50, 250, 10))
End If
Next i
Exit Sub
error:
Call MsgBox("Calculation is beyond known mathematical formulations.", vbCritical, "Math Error")

End Sub

Private Sub c_a_Click()
Shape1.Left = c_a.Left
Shape1.Top = c_a.Top


End Sub


Private Sub c_b_Click()
Shape1.Left = c_b.Left
Shape1.Top = c_b.Top


End Sub

Private Sub c_m_Click()
Shape1.Left = c_m.Left
Shape1.Top = c_m.Top

End Sub

Private Sub c_n1_Click()
Shape1.Left = c_n1.Left
Shape1.Top = c_n1.Top

End Sub

Private Sub c_n2_Click()
Shape1.Left = c_n2.Left
Shape1.Top = c_n2.Top

End Sub

Private Sub c_n3_Click()
Shape1.Left = c_n3.Left
Shape1.Top = c_n3.Top

End Sub

Private Sub c_npi_Click()
Shape1.Left = c_npi.Left
Shape1.Top = c_npi.Top

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
use_poly = True

Else
use_poly = False
End If


End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
use_dots = True
Else
use_dots = False
End If

End Sub

Private Sub Command1_Click()
Call create
End Sub

Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Form_Load()
use_poly = True
use_dots = True
scaleworld = 100
Shape1.Width = c_a.Width + 5
Shape1.Height = c_a.Height + 5

End Sub

Private Sub Form_Resize()
world.Width = Abs(Me.Width - Picture1.Width)
Line1.Y1 = 0
Line1.Y2 = world.Height
Line2.X1 = 0
Line2.X2 = world.Width
Call create

End Sub

Private Sub HScroll1_Change()
scaleworld = HScroll1.Value
Label8.Caption = scaleworld
Call create

End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub world_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Line1.X1 = x
Line1.X2 = x
Line2.Y1 = y
Line2.Y2 = y
End Sub

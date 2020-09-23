VERSION 5.00
Begin VB.Form TstJZXPLb 
   Caption         =   "JZ XP Label - Demo/Tutor/Test"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin JZXPLb.JzXPLab Xlb10 
      Height          =   1065
      Left            =   135
      Top             =   5370
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1879
      Caption         =   "M E N U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JZXPLb.JzXPLab XLb 
      Height          =   360
      Index           =   0
      Left            =   525
      Top             =   5370
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   635
      Caption         =   "Green"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JZXPLb.JzXPLab XLb9 
      Height          =   450
      Left            =   3030
      ToolTipText     =   "Click/DblClick test (as a button)"
      Top             =   90
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   794
      Caption         =   "Click Me"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin JZXPLb.JzXPLab XLb8 
      Height          =   705
      Left            =   1545
      Top             =   5520
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1244
      Caption         =   "HoverTime property is delay to init  events ...  Test it"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox HTim 
         Height          =   285
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "400"
         ToolTipText     =   "Move mouse again to reflect it"
         Top             =   330
         Width           =   510
      End
   End
   Begin JZXPLb.JzXPLab Xlb7 
      Height          =   2100
      Left            =   1530
      Top             =   3330
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   3704
      Caption         =   $"TstJZXPLb.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4210752
      Alignment       =   2
      BorderColor     =   4210752
      BackColor       =   12648447
   End
   Begin VB.TextBox Txb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "TstJZXPLb.frx":00C8
      Top             =   2370
      Width           =   2865
   End
   Begin JZXPLb.JzXPLab Xlb4 
      Height          =   3195
      Left            =   210
      Top             =   2010
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   5636
      Caption         =   "All of this form  was made by using the Label control and design ideas ..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "TstJZXPLb.frx":013A
   End
   Begin JZXPLb.JzXPLab XLb3 
      Height          =   600
      Left            =   150
      Top             =   1215
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   1058
      Caption         =   "Can be a BANNER"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin JZXPLb.JzXPLab XLb2 
      Height          =   435
      Left            =   150
      Top             =   615
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   767
      Caption         =   "Border/Back/Text color (+Alignments)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   0
      Alignment       =   1
      BorderColor     =   0
      BackColor       =   14737632
      TextColor       =   255
      MousePointer    =   99
      MouseIcon       =   "TstJZXPLb.frx":0454
   End
   Begin JZXPLb.JzXPLab XLb1 
      Height          =   390
      Left            =   120
      Top             =   135
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   688
      Caption         =   "Simulate a Text Box Locked"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JZXPLb.JzXPLab Xlb5 
      Height          =   375
      Left            =   1560
      Top             =   2025
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   661
      Caption         =   "As a Header of Text Box"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   0
      Alignment       =   2
      BorderColor     =   0
      BackColor       =   0
      TextColor       =   16777215
   End
   Begin JZXPLb.JzXPLab XLb 
      Height          =   360
      Index           =   1
      Left            =   525
      Top             =   5715
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   635
      Caption         =   "Yellow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JZXPLb.JzXPLab XLb 
      Height          =   360
      Index           =   2
      Left            =   525
      Top             =   6060
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   635
      Caption         =   "Blue"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "TstJZXPLb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'.---------------------------------------------------------------------------------------
' Modulo    : TstJZXPLb - JZ XP Label Demo/Tutor/Test
' Edição    : 19/05/2005
' Autor     : JOZE Walter de Moura - RIO DE JANEIRO, Brazil.
' Aplicação :
' Objetivos : Some ideas about how use the Control on Applications design
'           :
'`=======================================================================================
'
Option Explicit
'.--------------------------------------------------------------------------------------
' Rotina     : Generalities
'`--------------------------------------------------------------------------------------
'showing defaut delay time for Xlb4
Private Sub Form_Load()
  HTim.Text = CStr(Xlb4.HoverTime)
End Sub

'.--------------------------------------------------------------------------------------
' Rotina     : ChangeHoverTime, HTim Text Box receiving new values
' Finalidade : Retard/Accelerate MouseHover trigger
'`--------------------------------------------------------------------------------------
'general routine limits min=100 max=2000 milliseconds
Private Sub ChangeHoverTime()
  Dim t As Long
  t = Val(Trim(HTim.Text))
  If t < 100 Then t = 100
  If t > 2000 Then t = 2000
  Xlb4.HoverTime = t
  HTim.Text = Format(t, "#000")
End Sub

'a kind of lost_focus
Private Sub HTim_Click()
  ChangeHoverTime
End Sub

Private Sub HTim_GotFocus()
  HTim.Text = CStr(Xlb4.HoverTime)
End Sub

Private Sub HTim_LostFocus()
  ChangeHoverTime
End Sub

'a kind of lost_focus
Private Sub Xlb4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ChangeHoverTime
End Sub

'.--------------------------------------------------------------------------------------
' Rotina     : Xlb10_MouseMove, XLb_MouseMove
' Finalidade : Explore Menu possibilities
'            : by using Control
'`--------------------------------------------------------------------------------------
'Resets Menu Background color
Private Sub Xlb10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   DoEvents ' mode waiting Windows works
   XLb(0).BackColor = vbWhite
   XLb(1).BackColor = vbWhite
   XLb(2).BackColor = vbWhite
End Sub

'Sets pointed Label Background color
Private Sub XLb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DoEvents ' mode waiting Windows works
   Select Case Index
     Case 0
       XLb(0).BackColor = vbGreen
       XLb(1).BackColor = vbWhite
       XLb(2).BackColor = vbWhite
     Case 1
       XLb(0).BackColor = vbWhite
       XLb(1).BackColor = vbYellow
       XLb(2).BackColor = vbWhite
     Case 2
       XLb(0).BackColor = vbWhite
       XLb(1).BackColor = vbWhite
       XLb(2).BackColor = vbBlue
   End Select
End Sub

'.--------------------------------------------------------------------------------------
' Rotina     : Xlb4_MouseHover, Xlb4_MouseLeave
' Finalidade : Shows/Hides other control (Xlb7)
'`--------------------------------------------------------------------------------------
'If user persists whith the mouse over 400 milliseconds ...
Private Sub Xlb4_MouseHover()
   Xlb7.Visible = True
End Sub


'When the mouse is got out, we Hide the previous showed control
Private Sub Xlb4_MouseLeave()
   Xlb7.Visible = False
End Sub

'.--------------------------------------------------------------------------------------
' Rotina     : Xlb9_Click, Xlb9_DblClick
' Finalidade : For study purposes only Click treatments
'`--------------------------------------------------------------------------------------
'If one click was apprehended
Private Sub XLb9_Click()
  With XLb9
    If Not .Caption = "Clicked" Then
       .Caption = "Clicked"
    Else
       .Caption = "Click Me"
    End If
  End With
End Sub

'If double click was apprehended
Private Sub XLb9_DblClick()
  With XLb9
    If Not .Caption = "DblClicked" Then
       .Caption = "DblClicked"
    Else
       .Caption = "Click Me"
    End If
  End With
End Sub

'.======================================================================================
'  THAT'S ALL, FOLKS!
'`--------------------------------------------------------------------------------------


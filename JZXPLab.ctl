VERSION 5.00
Begin VB.UserControl JzXPLab 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   LockControls    =   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   1770
   ToolboxBitmap   =   "JZXPLab.ctx":0000
   Begin VB.Label JzLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JzLabel"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   1275
   End
   Begin VB.Shape JzBorder 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      Height          =   375
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   1365
   End
End
Attribute VB_Name = "JzXPLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'.------------------------------------------------------------------
' Control   : JZ XP Label 1.0
' Edition   : 19-May-2005
' Author    : JOZE Walter de Moura - RIO DE JANEIRO, BRASIL.
'           : me: www.joze.kit.net   or   qualyum@globo.com
'           :
'           : Well, I wrote this control by using some few line
'           : codes from several authors at Internet whose
'           : credits I acknowledge e appreciate.
'           :
'           : But the Module.Bas and the SubClass on separate
'           : mode use MouseHover/MouseLeave events and HoverTime
'           : property was based on an Article from Source Guru, by
'           : HSPC (nickname). I credit to those persons all credits
'           : for these codes.
'           :
'           : Thanks and part of credits to Territop (Paul) who had email
'           : me the link to see the HSPC Article: "Adding MouseLeave and
'           : MouseHover events to Vb6". We were looking how to do a timed
'           : cursor query without includes a strong Timer Control in code.
'           :
'           : The Demo/Tutor/Test form has few line codes - I recommend
'           : people examine Object Properties where will find effects and
'           : easy design.
'           :
'           : I use this Label to receive and show data, variables, etc, as
'           : it were a Text Box with a Locked option, to inhibit any typing
'           : and focus.
'           :
'Application: Another all-Windows XP-style Label:
'           : - Shaped rectangle with rounded borders;
'           : - BorderColor extra property;
'           : - Proper Alignment adjust.
'           : - Click, DblClick, MouseDown, MouseMove, adjustable MouseHover,
'           :   MouseLeave events to complement and enhance your designed
'           :   project.
'           :
' License   : Freeware - you may distribute, alter, sold, anything
'           : as you want. This code is for you, don't it?
'           : I'm sure you will apply maximum of honesty and ethics
'           : concerning it.
'           :
'           : Joze.
'           :

Public Event Change()
Attribute Change.VB_Description = "Occurs when the Caption contents is modified."
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object.\r\n"
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object.\r\n"
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse to Control ambit.\r\n"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when user picks down the mouse over control."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when Cursor get out the control ambit.\r\n"
Public Event MouseHover()
Attribute MouseHover.VB_Description = "Occurs when User pauses the Cursor over Control for defined time (v.HoverTime property).\r\n"

'Alignment options
Public Enum AlignmentOpts
   [Left] = 0
   [Right] = 1
   [Centered] = 2
End Enum

'MousePointer options
Public Enum MousePointerOpts
   [Default] = vbDefault
   [Arrow] = vbArrow
   [Cross Hair] = vbCrosshair
   [I-Beam] = vbIbeam
   [Icon Pointer] = vbIconPointer
   [Size Pointer] = vbSizePointer
   [Size NESW] = vbSizeNESW
   [Size NS] = vbSizeNS
   [Size NWSE] = vbSizeNWSE
   [Size WE] = vbSizeWE
   [Up Arrow] = vbUpArrow
   [Hour Glass] = vbHourglass
   [No Drop] = vbNoDrop
   [Arrow Hour Glass] = vbArrowHourglass
   [Arrow Question] = vbArrowQuestion
   [Size All] = vbSizeAll
   [Custom] = vbCustom
End Enum

Const m_def_Border = &HB99D7F
Const m_def_BackColor = vbWhite
Const m_def_TextColor = vbBlack
Const m_def_HoverTime = 400

Dim WithEvents MyTrak As clsTrackInfo
Attribute MyTrak.VB_VarHelpID = -1

'Control Properties Variables
Private m_Font As Font
Private m_BorderColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_TextColor As OLE_COLOR
Private m_Mousepointer As MousePointerOpts

Private Sub JzLabel_Change()
   RaiseEvent Change
End Sub

Private Sub JzLabel_Click()
   RaiseEvent Click
End Sub

Private Sub JzLabel_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub JzLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub JzLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'code to subclass
Private Sub MyTrak_MouseHover()
   RaiseEvent MouseHover
End Sub

Private Sub MyTrak_MouseLeave()
   RaiseEvent MouseLeave
End Sub

Private Sub RePos()
   If Width < 405 Then Width = 405
   If Height < 315 Then Height = 315
   JzBorder.Width = Width
   JzBorder.Left = 0
   JzBorder.Height = Height
   JzBorder.Top = 0
   JzLabel.Width = Width - 180
   JzLabel.Height = Height - 165
   JzLabel.Top = 90 '60
   JzLabel.Left = 105 '75
End Sub

'
Private Sub UserControl_Change()
   RaiseEvent Change
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
   RePos
End Sub

'subclass
Private Sub UserControl_Terminate()
   EndTrack MyTrak
   Set MyTrak = Nothing
End Sub

Private Sub UserControl_InitProperties()

   Set MyTrak = New clsTrackInfo

   m_TextColor = m_def_TextColor
   m_BackColor = m_def_BackColor
   m_BorderColor = m_def_Border
   m_Mousepointer = [Default]
   BackColor = m_BackColor
   BorderColor = m_BorderColor
   TextColor = m_TextColor
   Set m_Font = Ambient.Font
   Alignment = [Left]
   JzLabel.Caption = Ambient.DisplayName

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    JzLabel.Caption = .ReadProperty("Caption", "JzLabel1")
    Set JzLabel.Font = .ReadProperty("Font", Ambient.Font)
    JzLabel.Alignment = .ReadProperty("Alignment", [Left])
    m_BorderColor = .ReadProperty("BorderColor", m_def_Border)
    m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
    m_TextColor = .ReadProperty("TextColor", m_def_TextColor)
    m_Mousepointer = .ReadProperty("MousePointer", [Default])
    Set JzLabel.MouseIcon = .ReadProperty("MouseIcon", Nothing)
    
  End With
    
    RePos
    JzLabel.MousePointer = m_Mousepointer
    JzBorder.BorderColor = m_BorderColor
    JzBorder.BackColor = m_BackColor
    JzLabel.BackColor = m_BackColor
    JzLabel.ForeColor = m_TextColor

 'subclass
    Set MyTrak = New clsTrackInfo
    MyTrak.hwnd = UserControl.hwnd

    MyTrak.HoverTime = PropBag.ReadProperty("HoverTime", m_def_HoverTime)

    If Ambient.UserMode Then 'only if is running
       StartTrack MyTrak
    End If
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("Caption", JzLabel.Caption, "Label1")
    Call .WriteProperty("Font", JzLabel.Font, Ambient.Font)
    Call .WriteProperty("BorderColor", m_BorderColor, &HB99D7F)
    Call .WriteProperty("Alignment", JzLabel.Alignment, [Left])
    Call .WriteProperty("BorderColor", m_BorderColor, m_def_Border)
    Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call .WriteProperty("TextColor", m_TextColor, m_def_TextColor)
    Call .WriteProperty("MousePointer", JzLabel.MousePointer, [Default])
    Call .WriteProperty("MouseIcon", JzLabel.MouseIcon, Nothing)
'subclass
    Call .WriteProperty("HoverTime", MyTrak.HoverTime, m_def_HoverTime)

  End With
End Sub

Public Property Get Alignment() As AlignmentOpts
Attribute Alignment.VB_Description = "Indicates disposition of Text in Box."
   Alignment = JzLabel.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentOpts)
   JzLabel.Alignment = New_Alignment
   PropertyChanged "Alignment"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the Background color."
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   m_BackColor = New_BackColor
   JzLabel.BackColor = m_BackColor
   JzBorder.BackColor = m_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the Border color."
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   JzBorder.BorderColor = m_BorderColor
   PropertyChanged "BorderColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text contained in the control."
   Caption = JzLabel.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   JzLabel.Caption = New_Caption
   PropertyChanged "Caption"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
   Set Font = JzLabel.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set JzLabel.Font = New_Font
   PropertyChanged "Font"
   RePos
End Property

Public Property Get HoverTime() As Long
Attribute HoverTime.VB_Description = "Time to have elapsed before the MouseHover event action (user staying cursor positioned into control ambit.\r\n"
   HoverTime = MyTrak.HoverTime
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hwnd = UserControl.hwnd
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Icon or Cursor picture will be the Mouse Pointer when MousePointer Property is setted to 99 - Custom."
   Set MouseIcon = JzLabel.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_Icon As IPictureDisp)
   Set JzLabel.MouseIcon = New_Icon
   PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerOpts
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object\r\n"
   m_Mousepointer = JzLabel.MousePointer
   MousePointer = m_Mousepointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerOpts)
   m_Mousepointer = New_MousePointer
   JzLabel.MousePointer = m_Mousepointer
   PropertyChanged "MousePointer"
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Returns/sets the Text color."
   TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
   m_TextColor = New_TextColor
   JzLabel.ForeColor = m_TextColor
   PropertyChanged "TextColor"
End Property

Public Property Let HoverTime(newHoverTime As Long)
   MyTrak.HoverTime = newHoverTime
   PropertyChanged "HoverTime"
End Property


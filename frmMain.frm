VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PNG Viewer"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvProps 
      Height          =   2895
      Left            =   60
      TabIndex        =   5
      Top             =   1860
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbProgress 
      Height          =   135
      Left            =   2820
      TabIndex        =   3
      Top             =   4650
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   4
      Scrolling       =   1
   End
   Begin MSComctlLib.TabStrip tabChoose 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      MultiRow        =   -1  'True
      Style           =   2
      ShowTips        =   0   'False
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "RGB"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "RGBA"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Grayscale"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Transparent Palette"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Choose..."
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   2820
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2730
      Begin VB.Timer tmrScrollBack 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   60
         Top             =   60
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2760
      Left            =   2820
      ScaleHeight     =   2700
      ScaleWidth      =   2700
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.PictureBox picBanner 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   60
      ScaleHeight     =   960
      ScaleWidth      =   4020
      TabIndex        =   4
      ToolTipText     =   "http://www.vbfrood.de"
      Top             =   660
      Width           =   4020
   End
   Begin VB.Line linLine 
      BorderColor     =   &H80000011&
      X1              =   60
      X2              =   5580
      Y1              =   450
      Y2              =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private WithEvents PNG As cPNG
Attribute PNG.VB_VarHelpID = -1
Private FileDlg As New cFileDialog
Private Function AddProp(ValName As String, Value As String)
    Dim NewItm As ListItem
    
    Set NewItm = lvProps.ListItems.Add
    
    With NewItm
        .Text = ValName
        .SubItems(1) = Value
    End With
End Function

Private Function GetProps()
    lvProps.ListItems.Clear
    With PNG
        AddProp "Width", .Width & " Pixels"
        AddProp "Height", .Height & " Pixels"
        AddProp "Image Type", Choose(.ColorType + 1, _
            "Grayscale", , "RGB", "Palette", "Grayscale+Alpha", , _
            "RGB+Alpha")
        AddProp "Bit Depth", .BitDepth
    End With
End Function
Private Sub Form_Load()
    Set PNG = New cPNG
    PNG.LoadPNGFile App.Path & "\banner.png"
    picBanner.Left = Me.ScaleWidth - picBanner.Width
    PNG.DrawToDC picBanner.hDC, 0, 0
    PNG.LoadPNGFile App.Path & "\back.png"
    PNG.DrawToDC picBack.hDC, 0, 0
    picBanner.Refresh
    Me.Visible = True
    Me.Refresh
    tabChoose_Click
End Sub


Private Sub PNG_LoadProgress(Max As Long, Value As Long)
    prbProgress.Max = Max
    prbProgress.Value = Value
End Sub


Private Sub tabChoose_Click()
    Dim CurCap As String
    
    CurCap = tabChoose.SelectedItem.Caption
    
    tmrScrollBack.Enabled = False
    If CurCap <> "Choose..." Then
        PNG.LoadPNGFile App.Path & "\" & LCase(CurCap) & ".png"
    Else
        FileDlg.Owner = Me
        FileDlg.Flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES _
            Or OFN_PATHMUSTEXIST
        FileDlg.Filter = "Portable Network Graphics (PNG)|*.png|All Files|*.*"
        FileDlg.ShowOpen
        If FileDlg.FileName <> "" Then
            Select Case PNG.LoadPNGFile(FileDlg.FileName)
                Case pngeFileNotFound
                    MsgBox "The specified file could not be found.", vbCritical
                Case pngeOpenError
                    MsgBox "Error opening the file.", vbCritical
                Case pngeInvalidFile
                    MsgBox "The specified file is no valid PNG file.", vbCritical
            End Select
        End If
    End If
    GetProps
    tmrScrollBack_Timer
    tmrScrollBack.Enabled = True
End Sub


Private Sub tmrScrollBack_Timer()
    Static MyY As Long
    
    MyY = MyY + 1
    If MyY > 19 Then MyY = 0
    
    BitBlt picDraw.hDC, 0, 0, picDraw.ScaleWidth, _
        picDraw.ScaleHeight, picBack.hDC, 0, MyY, vbSrcCopy
    BitBlt picDraw.hDC, 0, picDraw.ScaleHeight - MyY, picDraw.ScaleWidth, _
        MyY, picBack.hDC, 0, 0, vbSrcCopy
        
    PNG.DrawToDC picDraw.hDC, picDraw.ScaleWidth \ 2 - _
        PNG.Width \ 2, picDraw.ScaleHeight \ 2 - PNG.Height \ 2
    picDraw.Refresh
End Sub



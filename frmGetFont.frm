VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetFont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get Fonts"
   ClientHeight    =   3960
   ClientLeft      =   36
   ClientTop       =   288
   ClientWidth     =   5052
   ClipControls    =   0   'False
   Icon            =   "frmGetFont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5052
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkStrikethru 
      Caption         =   "&Strikethru"
      Height          =   288
      Left            =   1152
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3252
      Width           =   996
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "&Underline"
      Height          =   288
      Left            =   132
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   996
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Bol&d Italic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Index           =   3
      Left            =   3732
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1746
      Width           =   1020
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "&Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1456
      Width           =   1020
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "It&alic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Index           =   1
      Left            =   3732
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1166
      Width           =   1020
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "&Regular"
      Height          =   290
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   876
      Width           =   1020
   End
   Begin MSComctlLib.Slider sldSize 
      Height          =   1200
      Left            =   3096
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   828
      Width           =   408
      _ExtentX        =   720
      _ExtentY        =   2117
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   1
      Min             =   6
      Max             =   120
      SelStart        =   26
      TickStyle       =   2
      TickFrequency   =   19
      Value           =   26
      TextPosition    =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2244
      Top             =   3516
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGetFont.frx":030A
            Key             =   "True"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGetFont.frx":0466
            Key             =   "Fixed"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   108
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3672
      Visible         =   0   'False
      Width           =   1524
   End
   Begin MSComctlLib.ListView lvFonts 
      Height          =   1224
      Left            =   144
      TabIndex        =   0
      Top             =   830
      Width           =   2748
      _ExtentX        =   4847
      _ExtentY        =   2159
      View            =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   3756
      TabIndex        =   4
      Top             =   3288
      Width           =   996
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      Left            =   2748
      TabIndex        =   3
      Top             =   3288
      Width           =   996
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   852
      Left            =   156
      ScaleHeight     =   852
      ScaleWidth      =   4548
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
      Width           =   4548
   End
   Begin VB.Label lblSelSize 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Height          =   204
      Left            =   3072
      TabIndex        =   17
      Top             =   408
      Width           =   420
   End
   Begin VB.Label lblSty 
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      Height          =   204
      Left            =   3708
      TabIndex        =   14
      Top             =   144
      Width           =   492
   End
   Begin VB.Label lblStyle 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   204
      Left            =   3696
      TabIndex        =   13
      Top             =   408
      Width           =   888
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   252
      Left            =   3036
      TabIndex        =   12
      Top             =   144
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Font:"
      Height          =   192
      Left            =   144
      TabIndex        =   11
      Top             =   144
      Width           =   1032
   End
   Begin VB.Label lblSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   204
      Left            =   144
      TabIndex        =   10
      Top             =   408
      Width           =   2748
   End
End
Attribute VB_Name = "frmGetFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Dim oStyle
Dim tpx, tpy 'ScreenTwipsPerPixel

Public Sub DrawFrameOn(TopLeftControl As Control, LowestRightControl As Control, Style As String, Framewidth)
Dim dw, fs, sm
Dim st$
Dim Lft, Toplft, Hite
Dim Rite, Ritebotm
Dim lt As Long
Dim rb As Long

    'Routine to draw frame around controls.  Variations can be achieved by changing "DrawWidth", shadow colors
    ' and the width of the frame elements in the form Paint Event
    
    'Save the current settings
    dw = DrawWidth
    fs = FillStyle
    sm = ScaleMode
    
    DrawWidth = 1
    FillStyle = 1
    ScaleMode = 3
    
    st = LCase(Left$(Style, 1))
    Lft = TopLeftControl.Left
    Toplft = TopLeftControl.Top
    Hite = TopLeftControl.Height
    
    Rite = LowestRightControl.Left + LowestRightControl.Width
    Ritebotm = LowestRightControl.Top + LowestRightControl.Height
    
    If Ritebotm > Hite Then Hite = Ritebotm
       
    lt = vb3DHighlight
    rb = vbButtonShadow
    
    'Swap colors if "inward"
    If st = "i" Then
        lt = vb3DDKShadow
        rb = vb3DHighlight
    End If
    
    'Draw the frame
    Line (Lft - Framewidth, Toplft - Framewidth)-(Rite + Framewidth, Toplft - Framewidth), lt
    Line (Lft - Framewidth, Toplft - Framewidth)-(Lft - Framewidth, Hite + Framewidth), lt
    Line (Rite + Framewidth, Toplft - Framewidth)-(Rite + Framewidth, Ritebotm + Framewidth), rb
    Line (Rite + Framewidth, Ritebotm + Framewidth)-(Lft - Framewidth, Hite + Framewidth), rb
    
    'Restore original settings
    DrawWidth = dw
    FillStyle = fs
    ScaleMode = sm
     
End Sub

Private Sub chkStrikethru_Click()

    UpdateSample "Strikethru"

End Sub

Private Sub chkStrikethru_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picSample.SetFocus

End Sub

Private Sub chkUnderline_Click()
    
    UpdateSample "Underline"
    
End Sub

Private Sub chkUnderline_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picSample.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
  
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    'Outputs
    SelectedFont = lblSelected
    SelectedSize = CInt(lblSelSize)
    SelectedStyle = lblStyle
    fUnderline = chkUnderline
    fStrikethru = chkStrikethru
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
 
    lvFonts.SetFocus
    
End Sub

Private Sub Form_Load()
Dim i
Dim fnt As String
Dim itmX As ListItem
Dim hDC As Long

    Screen.MousePointer = vbHourglass
    
    hDC = GetDC(List1.hWnd)
    
    'Put all the TrueType's into List1
    ShowFontType = 4 'True Type
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamTypeProc, List1
       
    'Add the TrueTypes w/icon to the Listview
    With lvFonts
        .Icons = ImageList1
        .SmallIcons = ImageList1
        For i = 0 To List1.ListCount - 1
            fnt = List1.List(i)
            Set itmX = .ListItems.Add(, , fnt, , 1)
    
        Next i
 
        List1.Clear
        
        'Put the Fixed fonts into List1
        ShowFontType = 1 'Fixed Width
        EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamTypeProc, List1
            
        'Add these w/icon to the Listview
        For i = 0 To List1.ListCount - 1
            fnt = List1.List(i)
            Set itmX = .ListItems.Add(, , fnt)
            itmX.SmallIcon = 2
        Next i
         
       Caption = "Get Fonts  - " & .ListItems.Count & " Fonts found"
        
        lblSelected = .ListItems(1)
        .ListItems(1).Selected = True
        
    
    End With

    ReleaseDC List1.hWnd, hDC
           
    lblSelSize = sldSize
    optStyle(0) = True
     
    
    With picSample
        .Font = lblSelected
        .FontSize = lblSelSize
        .Font.Bold = False
        .Font.Italic = False
    End With
      
    UpdateSample ("Name")
      
    'Using ScreenTwips to make it look right in small or large font setting
    tpx = Screen.TwipsPerPixelX
    tpy = Screen.TwipsPerPixelY
    
    Width = 420 * tpx
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Paint()
    
    'Call the DrawFrameOn routine to draw frame around controls (in pixels).
    
    DrawFrameOn lvFonts, lvFonts, "outward", 4
    DrawFrameOn lvFonts, lvFonts, "inward", 1
    
    DrawFrameOn picSample, picSample, "outward", 4
    DrawFrameOn picSample, picSample, "inward", 1

    DrawFrameOn lblSelected, lblSelected, "outward", 4
    DrawFrameOn lblSelected, lblSelected, "inward", 1
    
    DrawFrameOn sldSize, sldSize, "outward", 4
  
    DrawFrameOn lblSelSize, lblSelSize, "outward", 4
    DrawFrameOn lblSelSize, lblSelSize, "inward", 1

    DrawFrameOn lblStyle, lblStyle, "outward", 4
    DrawFrameOn lblStyle, lblStyle, "inward", 1

    DrawFrameOn optStyle(0), optStyle(3), "outward", 4
   
    DrawFrameOn chkUnderline, chkStrikethru, "outward", 4
    DrawFrameOn chkUnderline, chkStrikethru, "inward", 1
   
    DrawFrameOn cmdOK, cmdCancel, "outward", 4
    DrawFrameOn cmdOK, cmdCancel, "inward", 1


End Sub

Private Sub Form_Resize()
Dim i, sp
Dim l, t, w, h
    
    sp = 16 * tpx 'Spacing
   
    'Layout the controls
    
    With lblName
        .Move 12 * tpx, 12 * tpy, 90 * tpx, sp
        lblSelected.Move sp, .Top + .Height + 6 * tpy, 230 * tpx, sp
        lblSize.Move 257 * tpx, .Top, 42 * tpx, .Height
        lblSty.Move 308 * tpx, .Top, 42 * tpx, .Height
    End With
    
    With lblSelected
        lvFonts.Move sp, .Top + .Height + sp, .Width, 100 * tpy
        lblSelSize.Move .Left + .Width + sp, .Top, 34 * tpy, .Height
    End With
    
    With lblSelSize
        lblStyle.Move .Left + .Width + sp, .Top, 85 * tpx, .Height
    
    End With
    
    With lvFonts
        sldSize.Move 2 * sp + .Width, .Top, lblSelSize.Width, .Height
        picSample.Move sp, .Top + .Height + sp, ScaleWidth - 2 * sp, 70 * tpy
        
    End With

    With sldSize
        optStyle(0).Move .Left + .Width + sp, .Top, 85 * tpx, 25 * tpy
    End With
    
    With optStyle(0)
        For i = 1 To 3
            optStyle(i).Move .Left, optStyle(i - 1).Top + .Height, .Width, .Height
        Next
    End With
       
    w = 82 * tpx
    h = 24 * tpy
    l = ScaleWidth - sp - w
    t = picSample.Top + picSample.Height + sp
    
    cmdCancel.Move l, t, w, h
    cmdOK.Move l - w, t, w, h
    chkUnderline.Move sp, t, w, h
    chkStrikethru.Move sp + w, t, w, h
        
    Height = t + h + sp + (Height - ScaleHeight)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Set frmGetFont = Nothing
    
End Sub

Private Sub lvFonts_ItemClick(ByVal Item As MSComctlLib.ListItem)

    lblSelected = lvFonts.ListItems(Item.Index)
    
    UpdateSample "Name"
    
End Sub

Private Sub optStyle_Click(Index As Integer)
    
    Select Case Index
        Case 0
            lblStyle = "Regular"
        Case 1
            lblStyle = "Italic"
        Case 2
            lblStyle = "Bold"
        Case 3
            lblStyle = "Bold Italic"
    End Select
        
    oStyle = Index
    UpdateSample "Style"
    
End Sub

Private Sub optStyle_GotFocus(Index As Integer)
    
    'Get rid of focus rectangle
    lvFonts.SetFocus

End Sub

Public Sub UpdateSample(Item As String)
Dim i, j
Dim Msg$
    
    'Update the sample
    
    Msg = "Sample"
    
    With picSample
        .Cls
        
        Select Case Item
            Case "Name"
                .Font = lblSelected
            
            Case "Size"
                .FontSize = lblSelSize
                lblSelSize.Refresh
            
            Case "Style"
                
                Select Case oStyle
                    
                    Case 0 'Regular
                        .Font.Bold = False
                        .Font.Italic = False
                    
                    Case 1 'Italic
                        .Font.Bold = False
                        .Font.Italic = True
          
                    Case 2 'Bold
                        .Font.Bold = True
                        .Font.Italic = False
            
                    Case 3 'Bold Italic
                        .Font.Bold = True
                        .Font.Italic = True
          
                End Select
       
            Case "Underline"
                .FontUnderline = chkUnderline
                
            Case "Strikethru"
                .FontStrikethru = chkStrikethru
        
        End Select
         
        'Center "Msg" in picSample
        i = .TextWidth(Msg) \ 2
        j = .TextHeight(Msg) \ 2
        .CurrentX = (.ScaleWidth \ 2) - i
        .CurrentY = (.ScaleHeight \ 2) - j
    
    End With
  
    picSample.Print Msg
  
End Sub

Private Sub sldSize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picSample.SetFocus

End Sub

Private Sub sldSize_Scroll()

    lblSelSize = sldSize
    UpdateSample "Size"
 
End Sub


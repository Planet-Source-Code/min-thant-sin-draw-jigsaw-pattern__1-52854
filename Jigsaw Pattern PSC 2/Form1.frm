VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Jigsaw Pattern"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBox 
      Height          =   6840
      Left            =   75
      ScaleHeight     =   452
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   607
      TabIndex        =   0
      Top             =   75
      Width           =   9165
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         FillColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   6315
         Left            =   225
         ScaleHeight     =   421
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   566
         TabIndex        =   1
         Top             =   300
         Width           =   8490
      End
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   150
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenPicture 
         Caption         =   "&Open Picture..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDrawPattern 
         Caption         =   "&Draw Jigsaw Pattern"
         Shortcut        =   ^D
      End
      Begin VB.Menu sepSavePictureAs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePictureAs 
         Caption         =   "Save Picture &As..."
      End
      Begin VB.Menu sepExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuCurvesColor 
         Caption         =   "Curves Color..."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPieceSize 
         Caption         =   "Piece &Size..."
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// This program is written by Min Thant Sin on Sunday, April 4, 2004
'/// Questions and comments are welcome
'/// Modify and improve the program as you wish.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// NOTE : Curve's width and height will always be the same in this program.
'///           If you know how to scale, feel free to change the curve's width and height
'///           so that they are of different dimensions.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Const DEFAULT_SIZE As Integer = 100     'Default piece size

Private MouseX As Integer, MouseY As Integer
Private NumCurvesAcross As Integer   'Number of curves horizontally
Private NumCurvesDown As Integer    'Number of curves vertically

Private CurveSize As Integer        'Curve's size (width = height)
Private PictureFilePath As String

Private Bezier As BEZIER_CURVE                'Base bezier
Private Bezier180 As BEZIER_CURVE           '180 degrees rotated (direction doesn't matter in this case)
Private Bezier90CW As BEZIER_CURVE        '90 degrees clock-wise rotated
Private Bezier90ACW As BEZIER_CURVE      '90 degress anti-clockwise rotated

Private Sub InitializeBezierCurves()
      '180 degrees rotated bezier
      InitializeBezierData Bezier180, CurveSize, CurveSize
      Rotate Bezier180, 180
      
      '-90 degrees rotated bezier
      InitializeBezierData Bezier90CW, CurveSize, CurveSize
      Rotate Bezier90CW, -90
      
      '90 degrees rotated bezier
      InitializeBezierData Bezier90ACW, CurveSize, CurveSize
      Rotate Bezier90ACW, 90
      
      'Normal bezier (base bezier)
      InitializeBezierData Bezier, CurveSize, CurveSize
End Sub

Private Sub DrawJigsawPattern()
      Dim Row As Integer, Col As Integer
      Dim OffsetX As Integer, OffsetY As Integer
      Dim boolFlip As Boolean
      
      'Draw horizontal curves
      For Row = 1 To NumCurvesDown - 1
            boolFlip = CBool(Row Mod 2 = 0)
            
            For Col = 0 To NumCurvesAcross - 1
                  InitializeBezierCurves
                  
                  'Positions of the curve
                  OffsetX = Col * CurveSize
                  OffsetY = Row * CurveSize
                  
                  If boolFlip Then
                        Translate Bezier180, OffsetX + CurveSize, OffsetY
                        PolyBezier picDraw.hdc, Bezier180.Points(1), NUM_BEZIER_POINTS
                  Else
                        Translate Bezier, OffsetX, OffsetY
                        PolyBezier picDraw.hdc, Bezier.Points(1), NUM_BEZIER_POINTS
                  End If
                  
                  boolFlip = Not boolFlip
            Next Col
      Next Row
      
      
      'Draw vertical curves
      For Row = 0 To NumCurvesDown - 1
            boolFlip = CBool(Row Mod 2 = 0)
            
            For Col = 1 To NumCurvesAcross - 1
                  InitializeBezierCurves
                  
                  'Positions of the curve
                  OffsetX = Col * CurveSize
                  OffsetY = Row * CurveSize
                  
                  If boolFlip Then
                        Translate Bezier90CW, OffsetX, OffsetY + CurveSize
                        PolyBezier picDraw.hdc, Bezier90CW.Points(1), NUM_BEZIER_POINTS
                  Else
                        Translate Bezier90ACW, OffsetX, OffsetY
                        PolyBezier picDraw.hdc, Bezier90ACW.Points(1), NUM_BEZIER_POINTS
                  End If
                  
                  boolFlip = Not boolFlip
            Next Col
      Next Row
End Sub

Private Sub Form_Load()
      ReDim Bezier.Points(1 To 7)
      ReDim Bezier90CW.Points(1 To 7)
      ReDim Bezier90ACW.Points(1 To 7)
      ReDim Bezier180.Points(1 To 7)
      
      CurveSize = 150
End Sub

Private Sub Form_Resize()
      picBox.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuCurvesColor_Click()
      With cdl
            .Color = 0
            .Flags = cdlCCFullOpen
            .ShowColor
      End With
      
      picDraw.ForeColor = cdl.Color
End Sub

Private Sub mnuDrawPattern_Click()
      'For clearing up picDraw.
      picDraw.Picture = LoadPicture("")
      
      If PictureFilePath <> "" Then
            picDraw.Picture = LoadPicture(PictureFilePath)
      End If
      
      'How many pieces can fit horizontally and vertically?
      NumCurvesAcross = (picDraw.ScaleWidth - 1) \ CurveSize
      NumCurvesDown = (picDraw.ScaleHeight - 1) \ CurveSize
      
      'Draw the jigsaw pattern
      Call DrawJigsawPattern
      
      'Draw rectangle around picDraw.
      Rectangle picDraw.hdc, 0, 0, (NumCurvesAcross * CurveSize) + 1, (NumCurvesDown * CurveSize) + 1
      
      picDraw.Picture = picDraw.Image
End Sub

Private Sub mnuExit_Click()
      picDraw.Picture = LoadPicture("")
      Unload Me
End Sub

Private Sub mnuPieceSize_Click()
      Dim strInput As String
      
      strInput = InputBox("Enter the piece size between 10 and 500 (whole number).", "I need some information...", "50")
      
      CurveSize = Int(Val(strInput))
      
      'Make sure curve's size lies within valid, reasonable values
      'NOTE : Please enter only interger values. If you enter a value greater than
      '           an interger can hold, an error will occur. No error-handling included here.
      If CurveSize < 10 Then CurveSize = DEFAULT_SIZE
      If CurveSize > (Screen.Height \ Screen.TwipsPerPixelY) Then CurveSize = (Screen.Height \ Screen.TwipsPerPixelY)
End Sub

Private Sub mnuOpenPicture_Click()
      With cdl
            .FileName = ""
            .Filter = "JPE Files (*.jpe)|*.jpe|JPG Files (*.jpg)|*.jpg|BMP Files (*.bmp)|*.bmp|All Files (*.*)|*.*"
            .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
            .ShowOpen
      End With
      
      If Trim(cdl.FileName) = "" Then Exit Sub
      
      'Save picture file path
      PictureFilePath = cdl.FileName
      
      picDraw.Picture = LoadPicture(PictureFilePath)
      picDraw.Move 0, 0
End Sub

Private Sub mnuSavePictureAs_Click()
      With cdl
            .FileName = ""
            .Filter = "BMP Files (*.bmp)|*.bmp|All Files (*.*)|*.*"
            .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
            .ShowSave
      End With
      
      If cdl.FileName = "" Then Exit Sub
      
      SavePicture picDraw.Image, cdl.FileName
End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Just a slapdash code for moving picDraw inside picBox
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Button Then
            MouseX = x
            MouseY = y
      End If
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Button Then
            picDraw.Left = picDraw.Left + (x - MouseX)
            picDraw.Top = picDraw.Top + (y - MouseY)
      End If
End Sub

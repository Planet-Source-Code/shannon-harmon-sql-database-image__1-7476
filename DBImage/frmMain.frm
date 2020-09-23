VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL DB Image Processor"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5475
   FillColor       =   &H00404040&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   90
      TabIndex        =   0
      Top             =   3105
      Width           =   3975
      Begin VB.TextBox txtConnectionString 
         Height          =   285
         Left            =   165
         TabIndex        =   2
         ToolTipText     =   "Enter valid database connection string"
         Top             =   375
         Width           =   3645
      End
      Begin VB.TextBox txtUID 
         Height          =   285
         Left            =   1305
         TabIndex        =   4
         ToolTipText     =   "Enter sql login id"
         Top             =   735
         Width           =   990
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2835
         PasswordChar    =   "*"
         TabIndex        =   6
         ToolTipText     =   "Enter your sql login password"
         Top             =   735
         Width           =   975
      End
      Begin VB.TextBox txtDBTable 
         Height          =   285
         Left            =   1305
         TabIndex        =   8
         ToolTipText     =   "Enter table to query"
         Top             =   1080
         Width           =   2505
      End
      Begin VB.TextBox txtDBColumn 
         Height          =   285
         Left            =   1305
         TabIndex        =   10
         ToolTipText     =   "Enter column to query"
         Top             =   1440
         Width           =   2505
      End
      Begin VB.TextBox txtWhere 
         Height          =   285
         Left            =   1305
         TabIndex        =   12
         ToolTipText     =   "Example:  CategoryID=1"
         Top             =   1800
         Width           =   2505
      End
      Begin VB.Label lblConnectionString 
         AutoSize        =   -1  'True
         Caption         =   "Connection String:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label lblUserID 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Left            =   690
         TabIndex        =   3
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblPW 
         AutoSize        =   -1  'True
         Caption         =   "PW:"
         Height          =   195
         Left            =   2475
         TabIndex        =   5
         Top             =   780
         Width           =   315
      End
      Begin VB.Label lblDBTable 
         AutoSize        =   -1  'True
         Caption         =   "DB Table:"
         Height          =   195
         Left            =   555
         TabIndex        =   7
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label lblDBColumn 
         AutoSize        =   -1  'True
         Caption         =   "DB Column:"
         Height          =   195
         Left            =   435
         TabIndex        =   9
         Top             =   1485
         Width           =   840
      End
      Begin VB.Label lblWhere 
         AutoSize        =   -1  'True
         Caption         =   "Where Clause:"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   1845
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdSchema 
      Caption         =   "&DB Schema"
      Height          =   360
      Left            =   4200
      TabIndex        =   18
      ToolTipText     =   "Shows schema for database tables"
      Top             =   4980
      Width           =   1170
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show &Table"
      Height          =   360
      Left            =   4200
      TabIndex        =   17
      ToolTipText     =   "Show all columns in table"
      Top             =   4620
      Width           =   1170
   End
   Begin VB.CommandButton cmdMakeASP 
      Caption         =   "&Make ASP"
      Height          =   360
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Makes an asp page for image proxy"
      Top             =   4260
      Width           =   1170
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Set &Null"
      Height          =   360
      Left            =   4200
      TabIndex        =   15
      ToolTipText     =   "Set database field to null"
      Top             =   3900
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   4785
      Top             =   345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "&Save Image"
      Height          =   360
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Save image to database"
      Top             =   3540
      Width           =   1170
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get Image"
      Height          =   345
      Left            =   4200
      TabIndex        =   13
      ToolTipText     =   "Get image from database"
      Top             =   3195
      Width           =   1170
   End
   Begin VB.PictureBox picHolder 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   2835
      Left            =   90
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Double click to view full screen"
      Top             =   105
      Width           =   5355
      Begin VB.Image imgGet 
         Height          =   915
         Left            =   2070
         ToolTipText     =   "Double click to view full screen"
         Top             =   990
         Width           =   975
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "By Shannon Harmon (c)2000 - All rights reserved!"
      Height          =   195
      Left            =   105
      TabIndex        =   19
      Top             =   5550
      Width           =   3480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Picture"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSavePixel 
         Caption         =   "&Save Single Pixel"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSaveOnGet 
         Caption         =   "&Save Image to Disk on Get"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuUseImage 
         Caption         =   "&Use viewing image on DB Save"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuAutosize 
         Caption         =   "&Autosize Picture"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCenterPic 
         Caption         =   "Center &Picture"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'*  By Shannon Harmon                        *
'*  Copyright 2000 - All rights reserved...  *
'*********************************************

Option Explicit

Dim xImage As New cImage
Dim strTable As String
Dim strColumn As String
Dim strWhere As String

Private Sub Form_Load()
On Error GoTo ProcErr

  Me.Line (0, 0)-(Me.ScaleWidth, 0), &H404040
  Me.Line (0, 1)-(Me.ScaleWidth, 1), &HFFFFFF
  Me.Line (0, 202)-(Me.ScaleWidth, 202), &H404040
  Me.Line (0, 203)-(Me.ScaleWidth, 203), &HFFFFFF
  Me.Line (0, 365)-(Me.ScaleWidth, 365), &H404040
  Me.Line (0, 366)-(Me.ScaleWidth, 366), &HFFFFFF

  Set xImage = New cImage

  '//Setting default values for testing...
  txtConnectionString = "DSN=Northwind;"
  txtUID = "sa"
  txtPassword = ""
  txtDBTable = "Categories"
  txtDBColumn = "Picture"
  txtWhere = "CategoryID=8"

ProcExit:
  Exit Sub
  
ProcErr:
  Debug.Print Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set xImage = Nothing
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
  
  Dim msgResult As VbMsgBoxResult
  msgResult = MsgBox("Are you sure?", vbYesNo, "Set field to null?")
  If msgResult = vbYes Then
  
    imgGet.Picture = LoadPicture()
    imgGet.Visible = False
    MousePointer = vbHourglass
    xImage.ClearImage strTable, strColumn, strWhere
    MousePointer = vbNormal
  
    If xImage.ErrNumber <> 0 Then
      lblInfo = xImage.ErrDesc
    Else
      lblInfo = "Field set to null successful!"
    End If
  Else
    lblInfo = "Process cancelled!"
  End If
End Sub

Private Sub cmdMakeASP_Click()
On Error Resume Next
  
  Dim strSession As String
  strSession = InputBox("Enter session variable name:", "Session Variable")
  If Trim(strSession) = "" Then
    lblInfo = "Process cancelled!"
    Exit Sub
  End If
  strSession = Trim(strSession)
  
  cDialog.CancelError = True
  cDialog.DialogTitle = "Save ASP Picture Proxy..."
  cDialog.Filename = ""
  cDialog.Filter = "Active Server Page(*.asp)|*.asp"
  cDialog.FilterIndex = 1
  cDialog.Flags = cdlOFNOverwritePrompt

  cDialog.ShowSave
  If Err = cdlCancel Or Trim(cDialog.Filename) = "" Then
    lblInfo = "Process cancelled!"
    Exit Sub
  End If
  
  MousePointer = vbHourglass
  xImage.MakeASPProxy Trim(cDialog.Filename), strSession
  MousePointer = vbNormal
  
  If xImage.ErrNumber <> 0 Then
    lblInfo = xImage.ErrDesc
  Else
    lblInfo = "ASP file created successful!"
  End If
End Sub

Private Sub cmdSaveImage_Click()
On Error Resume Next

  Dim blnUseImage As Boolean
  Dim strPath As String
  
  If mnuUseImage.Checked = True And imgGet.Picture <> 0 Then
    blnUseImage = True
    strPath = App.Path & "\dbtmpfl1.bmp"
    SavePicture imgGet.Picture, strPath
    If Err.Number <> 0 Then
      blnUseImage = False
      Err.Clear
    End If
  End If
      
  If Not blnUseImage Then
    cDialog.CancelError = True
    cDialog.DialogTitle = "Select Image..."
    cDialog.Filename = ""
    cDialog.Filter = "Pictures(*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp"
    cDialog.FilterIndex = 1
    cDialog.Flags = cdlOFNOverwritePrompt

    cDialog.ShowOpen
    If Err = cdlCancel Then
      strPath = ""
    Else
      strPath = Trim(cDialog.Filename)
    End If
  End If
  
  If strPath = "" Then
    lblInfo = "Operation cancelled!"
    Exit Sub
  Else
    If Not blnUseImage Then
      imgGet.Visible = False
      imgGet.Picture = LoadPicture()
    End If
    MousePointer = vbHourglass
    xImage.SaveImage strTable, strColumn, strWhere, Trim(strPath)
    If blnUseImage = True Then Kill strPath
    
    If xImage.ErrNumber <> 0 Then
      lblInfo = xImage.ErrDesc
    Else
      lblInfo = "Picture save to DB successful!"
      If Not blnUseImage Then
        imgGet.Picture = LoadPicture(Trim(strPath))
        CreateThumb picHolder, imgGet, mnuCenterPic.Checked, mnuAutosize.Checked
        imgGet.Visible = True
        imgGet.Refresh
      End If
    End If
    MousePointer = vbNormal
  End If
End Sub

Private Sub cmdShow_Click()
On Error Resume Next
  frmAllRows.Show vbModal, Me
End Sub

Private Sub cmdSchema_Click()
On Error Resume Next
  frmTables.Show vbModal, Me
End Sub

Private Sub cmdGet_Click()
On Error Resume Next
  
  Dim Filename As String
  If mnuSaveOnGet.Checked = True Then
    cDialog.CancelError = True
    cDialog.DialogTitle = "Save DB Image..."
    cDialog.Filename = ""
    cDialog.Filter = "Pictures(*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp"
    cDialog.FilterIndex = 1
    cDialog.Flags = cdlOFNOverwritePrompt
  
    cDialog.ShowSave
    If Err = cdlCancel Then
      Filename = ""
    Else
      Filename = Trim(cDialog.Filename)
    End If
  End If
  
  imgGet.Visible = False
  imgGet.Picture = LoadPicture()
  lblInfo = "Attempting to retrieve picture..."
  MousePointer = vbHourglass
  Set imgGet.Picture = xImage.GetImage(strTable, strColumn, strWhere, Filename)
  
  If xImage.ErrNumber <> 0 Then
    lblInfo = xImage.ErrDesc
  Else
    lblInfo = "Picture received successfully!"
    CreateThumb picHolder, imgGet, mnuCenterPic.Checked, mnuAutosize.Checked
    imgGet.Visible = True
    imgGet.Refresh
  End If
  MousePointer = vbNormal
End Sub

Public Sub CreateThumb(picTarget As PictureBox, imgActual As Image, Center As Boolean, Autosize As Boolean)
On Error Resume Next

  imgActual.Stretch = False
  If Autosize = True Then
    If picTarget.ScaleHeight < imgActual.Height _
      Or picTarget.ScaleWidth < imgActual.Width Then
        
        Dim intHeight As Integer, intWidth As Integer, dblMultiplyer As Double
        intHeight = imgActual.Height - picTarget.ScaleHeight
        intWidth = imgActual.Width - picTarget.ScaleWidth
        
        If intHeight >= intWidth Then
          dblMultiplyer = (imgActual.Height - intHeight) / imgActual.Height
        Else
          dblMultiplyer = picTarget.ScaleWidth / imgActual.Width
        End If
        
        imgActual.Height = imgActual.Height * dblMultiplyer
        imgActual.Width = imgActual.Width * dblMultiplyer
        imgActual.Stretch = True
        imgActual.Refresh
    Else
      imgActual.Stretch = False
    End If
  Else
    imgActual.Stretch = False
  End If
  
  If Center Then
    imgActual.Left = picTarget.ScaleWidth / 2 - imgActual.Width / 2
    imgActual.Top = picTarget.ScaleHeight / 2 - imgActual.Height / 2
  Else
    imgActual.Top = 0
    imgActual.Left = 0
  End If
End Sub

Private Sub imgGet_DblClick()
  frmView.Show vbModal, Me
End Sub

Private Sub mnuAbout_Click()
  MsgBox "SQL DBImage Tools by Shannon Harmon" & vbCrLf & "Copyright 2000 - All rights reserved!", vbInformation, "About"
End Sub

Private Sub mnuAutosize_Click()
  mnuAutosize.Checked = Not (mnuAutosize.Checked)
  CreateThumb picHolder, imgGet, mnuCenterPic.Checked, mnuAutosize.Checked
End Sub

Private Sub mnuCenterPic_Click()
  mnuCenterPic.Checked = Not (mnuCenterPic.Checked)
  CreateThumb picHolder, imgGet, mnuCenterPic.Checked, mnuAutosize.Checked
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuOpen_Click()
On Error Resume Next
  
  cDialog.CancelError = True
  cDialog.DialogTitle = "Open Image..."
  cDialog.Filename = ""
  cDialog.Filter = "Pictures(*.jpg;*.gif;*.bmp;*.ico;*.wmf;*.cur)|*.jpg;*.gif;*.bmp;*.ico;*.wmf;*.cur"
  cDialog.FilterIndex = 1
  
  cDialog.ShowOpen
  If Err = cdlCancel Or Trim(cDialog.Filename) = "" Then Exit Sub
  
  imgGet.Visible = False
  imgGet.Stretch = False
  imgGet.Picture = LoadPicture(Trim(cDialog.Filename))
  lblInfo = imgGet.Width & "x" & imgGet.Height
  CreateThumb picHolder, imgGet, mnuCenterPic.Checked, mnuAutosize.Checked
    
  If Err.Number <> 0 Then
    lblInfo = "Error opening picture file..."
  Else
    imgGet.Visible = True
    imgGet.Refresh
  End If
End Sub

Private Sub mnuSaveOnGet_Click()
  mnuSaveOnGet.Checked = Not (mnuSaveOnGet.Checked)
End Sub

Private Sub mnuSavePixel_Click()
On Error Resume Next
  
  Dim strColor As String, blnTransparent As Boolean, msgReturn As VbMsgBoxResult
  strColor = InputBox("Enter the web color to create (ie: #003366)...", "Save Pixel")
  strColor = Trim(strColor)
  strColor = Replace(strColor, "#", "")
  
  If strColor = "" Then
    lblInfo = "Operation cancelled!"
    Exit Sub
  ElseIf Len(strColor) <> 6 Then
    lblInfo = "Invalid web color!"
    Exit Sub
  End If
  
  msgReturn = MsgBox("Would you like this pixel to be transparent?", vbYesNo, "Save Pixel")
  If msgReturn = vbYes Then
    blnTransparent = True
  Else
    blnTransparent = False
  End If
  
  cDialog.CancelError = True
  cDialog.DialogTitle = "Save Gif 1x1 Pixel..."
  cDialog.Filename = ""
  cDialog.Filter = "Gif 89a(*.gif)|*.gif"
  cDialog.FilterIndex = 1
  cDialog.Flags = cdlOFNOverwritePrompt

  cDialog.ShowSave
  If Err = cdlCancel Then
    lblInfo = "Operation cancelled!"
    Exit Sub
  End If
  
  MousePointer = vbHourglass
  xImage.MakePixel Trim(cDialog.Filename), strColor, blnTransparent
  MousePointer = vbNormal
      
  If xImage.ErrNumber <> 0 Then
    lblInfo = xImage.ErrDesc
  Else
    lblInfo = "Pixel save successful!"
  End If
End Sub

Private Sub mnuUseImage_Click()
  mnuUseImage.Checked = Not (mnuUseImage.Checked)
End Sub

Private Sub picHolder_DblClick()
  frmView.Show vbModal, Me
End Sub

Private Sub txtConnectionString_Change()
  xImage.ConnectionString = Trim(txtConnectionString)
End Sub

Private Sub txtConnectionString_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtDBColumn_Change()
  strColumn = Trim(txtDBColumn)
End Sub

Private Sub txtDBColumn_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtDBTable_Change()
  strTable = Trim(txtDBTable)
End Sub

Private Sub txtDBTable_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtPassword_Change()
  xImage.Password = Trim(txtPassword)
End Sub

Private Sub txtPassword_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtUID_Change()
  xImage.UID = Trim(txtUID)
End Sub

Private Sub txtUID_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtWhere_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtWhere_Change()
  strWhere = Trim(txtWhere)
End Sub

Private Sub mnuEdit_Click()
On Error Resume Next
  If Clipboard.GetFormat(2) Then
    mnuPaste.Enabled = True
  Else
    mnuPaste.Enabled = False
  End If
  
  If imgGet.Picture <> 0 Then
    mnuCopy.Enabled = True
  Else
    mnuCopy.Enabled = False
  End If
End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
  Clipboard.Clear
  Clipboard.SetData imgGet.Picture
End Sub

Private Sub mnuPaste_Click()
On Error Resume Next
  imgGet.Visible = False
  imgGet.Stretch = False
  imgGet.Picture = Clipboard.GetData()
  lblInfo = imgGet.Width & "x" & imgGet.Height
  CreateThumb picHolder, imgGet, mnuCenterPic.Checked, mnuAutosize.Checked
  If Err.Number <> 0 Then
    lblInfo = Err.Description
  Else
    imgGet.Visible = True
    imgGet.Refresh
  End If
End Sub

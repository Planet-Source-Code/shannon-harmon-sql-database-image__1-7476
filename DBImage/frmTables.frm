VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTables 
   Caption         =   "Table Information"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   Icon            =   "frmTables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2550
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4498
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
'*  By Shannon Harmon                        *
'*  Copyright 2000 - All rights reserved...  *
'*********************************************

Option Explicit

Dim strSQL As String
Dim objRS As ADODB.Recordset
Dim objConn As ADODB.Connection

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ProcErr

  Screen.MousePointer = vbHourglass
  Set objConn = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  objConn.CursorLocation = adUseClient
  objConn.ConnectionTimeout = 15
  objConn.CommandTimeout = 30
  objConn.Open frmMain.txtConnectionString, frmMain.txtUID, frmMain.txtPassword
  
  objRS.CursorLocation = adUseClient
  
  '//Use this if you want to show only user tables
  'Set objRS = objConn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
  '//This shows entire schema
  Set objRS = objConn.OpenSchema(adSchemaTables)
  
  Set DataGrid1.DataSource = objRS
  DataGrid1.Refresh

  On Error Resume Next
  '//You will get an error if you have a valid database but no user
  '  tables so turn error checking off if necessary...
  '  (Only applies if you set the recordset to show tables only)
  Caption = "Schema for Database: " & objRS.Fields(0)
  If Err.Number <> 0 Then Caption = "Schema..."
  Err.Clear
  
ProcExit:
  Screen.MousePointer = vbNormal
  Exit Sub
  
ProcErr:
  Screen.MousePointer = vbNormal
  MsgBox Err.Description
  Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
  DataGrid1.Move 0, 0, Width - 90, Height - 90
  ColumnSize DataGrid1
  DataGrid1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Me.Hide
  Set DataGrid1.DataSource = Nothing
  If objRS.State = adStateOpen Then objRS.Close
  If objConn.State = adStateOpen Then objConn.Close
  Set objRS = Nothing
  Set objConn = Nothing
End Sub

Private Sub ColumnSize(dg As DataGrid)
On Error Resume Next

  Dim i As Integer
  Dim totalSize As Long, eachSize As Long
  
  totalSize = dg.Width
  eachSize = dg.Width / dg.Columns.Count
  
  For i = 0 To dg.Columns.Count - 1
    dg.Columns(i).Width = eachSize
  Next i
End Sub


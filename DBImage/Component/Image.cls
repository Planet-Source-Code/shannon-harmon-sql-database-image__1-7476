VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Allows easy image manipulation in a database field, insert, retrieve, web and applications...."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*********************************************
'*  By Shannon Harmon                        *
'*  Copyright 2000 - All rights reserved...  *
'*********************************************

Option Explicit

Private mvarConnectionString As String
Private mvarUID As String
Private mvarPassword As String

Public ErrDesc As String
Attribute ErrDesc.VB_VarDescription = "Returns last error description."
Public ErrNumber As Long
Attribute ErrNumber.VB_VarDescription = "Returns last error code."

Public Sub ErrClear()
Attribute ErrClear.VB_Description = "Clears the error value/description."
On Error Resume Next
  ErrDesc = ""
  ErrNumber = 0
End Sub

Public Sub DeleteFile(Filename As String)
On Error GoTo ProcErr

  ErrClear
  Kill Filename

ProcExit:
  Exit Sub

ProcErr:
  ErrNumber = Err.Number
  ErrDesc = Err.Description
End Sub

Public Sub MakePixel(Filename As String, Webcolor As String, Transparent As Boolean)
On Error GoTo ProcErr
Dim intRed As Integer, intGreen As Integer, intBlue As Integer

  ErrClear
  Filename = Trim(Filename)
  Webcolor = Replace(Webcolor, "#", "")
  intRed = Hex2Dec(Mid(Webcolor, 1, 2))
  intBlue = Hex2Dec(Mid(Webcolor, 3, 2))
  intGreen = Hex2Dec(Mid(Webcolor, 5, 2))

  Dim strPixel As String
  strPixel = "GIF89a" & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(128) & Chr(0) & Chr(0)
  strPixel = strPixel & Chr(intRed) & Chr(intGreen) & Chr(intBlue)
  strPixel = strPixel & Chr(0) & Chr(0) & Chr(0) & Chr(33) & Chr(249) & Chr(4)
  If Transparent Then
    strPixel = strPixel & Chr(1)
  Else
    strPixel = strPixel & Chr(0)
  End If
  strPixel = strPixel & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(44) & Chr(0) & Chr(0) & Chr(0)
  strPixel = strPixel & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(2) & Chr(2)
  strPixel = strPixel & Chr(68) & Chr(1) & Chr(0) & Chr(59)
  
  Dim Datafile As Integer
  Datafile = FreeFile
  Open Filename For Output As Datafile
    Print #Datafile, strPixel
  Close Datafile

ProcExit:
  Exit Sub

ProcErr:
  ErrNumber = Err.Number
  ErrDesc = Err.Description
End Sub

Public Sub MakeASPProxy(Filename As String, SessionVariable As String)
On Error GoTo ProcErr
  
  If Trim(Filename) = "" Then
    ErrNumber = 1002
    ErrDesc = "Invalid filename!"
    GoTo ProcExit
  Else
    If LCase(Right(Filename, 3)) <> "asp" Then Filename = Filename & ".asp"
  End If
  
  SessionVariable = "Session(""" & SessionVariable & """)"
  
  Dim Datafile As Integer
  Datafile = FreeFile
  Open Filename For Output As Datafile
    Print #Datafile, "<%"
    Print #Datafile, "Response.Expires=0"
    Print #Datafile, "Response.Buffer=True"
    Print #Datafile, "Response.Clear"
    Print #Datafile, "Response.contentType=""image/jpeg"""
    Print #Datafile, "Response.BinaryWrite " & SessionVariable
    Print #Datafile, SessionVariable & "="""""
    Print #Datafile, "Response.End"
    Print #Datafile, "%>"
  Close Datafile
  
ProcExit:
  Exit Sub

ProcErr:
  ErrNumber = Err.Number
  ErrDesc = Err.Description
End Sub

Public Function GetASPImage(Table As String, Column As String, Where As String) As Variant
On Error GoTo ProcErr

  Dim objConn As ADODB.Connection
  Dim objRS As ADODB.Recordset
  Dim strSQL As String
  Dim Chunk() As Byte
  Set objConn = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  
  ErrClear
  objConn.CursorLocation = adUseClient
  objConn.ConnectionTimeout = 15
  objConn.CommandTimeout = 30
  objConn.Open mvarConnectionString, mvarUID, mvarPassword
  
  strSQL = "SELECT " & Column & " FROM " & Table & " WHERE " & Where
  objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
  
  If objRS.BOF And objRS.EOF Then
    ErrNumber = 1000
    ErrDesc = "Record not found!"
    GetASPImage = 0
    GoTo ProcExit
  ElseIf IsNull(objRS.Fields(0)) Then
    ErrNumber = 1001
    ErrDesc = "Field is null"
    GetASPImage = 0
    GoTo ProcExit
  End If

  GetASPImage = objRS.Fields(0).Value

ProcExit:
  On Error Resume Next
  objRS.Close
  objConn.Close
  Set objRS = Nothing
  Set objConn = Nothing
  Exit Function

ProcErr:
  GetASPImage = 0
  ErrNumber = Err.Number
  ErrDesc = Err.Description
  Resume ProcExit
End Function

Public Function GetImage(Table As String, Column As String, Where As String, Optional Filename As String) As Variant
On Error GoTo ProcErr

  Dim objConn As ADODB.Connection
  Dim objRS As ADODB.Recordset
  Dim strSQL As String
  Dim Chunk() As Byte
  Set objConn = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  
  ErrClear
  objConn.CursorLocation = adUseClient
  objConn.ConnectionTimeout = 15
  objConn.CommandTimeout = 30
  objConn.Open mvarConnectionString, mvarUID, mvarPassword
  
  strSQL = "SELECT " & Column & " FROM " & Table & " WHERE " & Where
  objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly
  
  If objRS.BOF And objRS.EOF Then
    ErrNumber = 1000
    ErrDesc = "Record not found!"
    GetImage = 0
    GoTo ProcExit
  ElseIf IsNull(objRS.Fields(0)) Then
    ErrNumber = 1001
    ErrDesc = "Field is null"
    GoTo ProcExit
  End If
  
  Chunk() = objRS.Fields(0).GetChunk(objRS.Fields(0).ActualSize)
  Set GetImage = Chunk2Image(Chunk(), Filename)

ProcExit:
  On Error Resume Next
  objRS.Close
  objConn.Close
  Set objRS = Nothing
  Set objConn = Nothing
  Exit Function

ProcErr:
  GetImage = 0
  ErrNumber = Err.Number
  ErrDesc = Err.Description
  Resume ProcExit
End Function

Public Sub ClearImage(Table As String, Column As String, Where As String)
On Error GoTo ProcErr
  
  Dim objConn As ADODB.Connection
  Dim strSQL As String

  ErrClear
  Set objConn = New ADODB.Connection
  
  objConn.CursorLocation = adUseClient
  objConn.ConnectionTimeout = 15
  objConn.CommandTimeout = 30
  objConn.Open mvarConnectionString, mvarUID, mvarPassword
  
  strSQL = "UPDATE " & Table & " SET " & Column & " = NULL WHERE " & Where
  objConn.Execute strSQL
  
ProcExit:
  On Error Resume Next
  objConn.Close
  Set objConn = Nothing
  Exit Sub

ProcErr:
  ErrNumber = Err.Number
  ErrDesc = Err.Description
  Resume ProcExit
End Sub

Public Sub SaveImage(Table As String, Column As String, Where As String, Filename As String)
Attribute SaveImage.VB_Description = "Saves an image to a database from a file."
On Error GoTo ProcErr

  Dim objConn As ADODB.Connection
  Dim objRS As ADODB.Recordset
  Dim strSQL As String
  Dim Chunk() As Byte
  Set objConn = New ADODB.Connection
  Set objRS = New ADODB.Recordset
  
  ErrClear
  objConn.CursorLocation = adUseClient
  objConn.ConnectionTimeout = 15
  objConn.CommandTimeout = 30
  objConn.Open mvarConnectionString, mvarUID, mvarPassword
  
  strSQL = "SELECT " & Column & " FROM " & Table & " WHERE " & Where
  objRS.Open strSQL, objConn, adOpenStatic, adLockOptimistic
  
  If objRS.BOF And objRS.EOF Then
    ErrNumber = 1000
    ErrDesc = "Record not found!"
    GoTo ProcExit
  End If
  
  Chunk() = Image2Chunk(Filename)
    
  If Not objRS.EOF Then
    objRS.Fields(0).AppendChunk Chunk()
    objRS.Update
    objRS.Close
  End If
    
ProcExit:
  On Error Resume Next
  objRS.Close
  objConn.Close
  Set objRS = Nothing
  Set objConn = Nothing
  Exit Sub

ProcErr:
  ErrNumber = Err.Number
  ErrDesc = Err.Description
  Resume ProcExit
End Sub

Public Property Let Password(ByVal vData As String)
Attribute Password.VB_Description = "Gets/Sets the Password for the database connection."
  mvarPassword = vData
End Property

Public Property Get Password() As String
  Password = mvarPassword
End Property

Public Property Let UID(ByVal vData As String)
Attribute UID.VB_Description = "Gets/Sets the User ID for the database connection."
  mvarUID = vData
End Property

Public Property Get UID() As String
  UID = mvarUID
End Property

Public Property Let ConnectionString(ByVal vData As String)
Attribute ConnectionString.VB_Description = "Gets/Sets the database connection string"
  mvarConnectionString = vData
End Property

Public Property Get ConnectionString() As String
  ConnectionString = mvarConnectionString
End Property

'//Local Functions
'---------------------------------------------------------
Private Function Image2Chunk(Filename As String) As Variant
On Error GoTo ProcErr
Dim Datafile As Integer
Dim FileLength As Long
Dim Chunk() As Byte
    
  Datafile = FreeFile
  Open Filename For Binary Access Read As Datafile
    FileLength = LOF(Datafile)
    If FileLength = 0 Then GoTo ProcErr
    ReDim Chunk(FileLength)
    Get Datafile, , Chunk()
  Close Datafile
    
ProcExit:
  Image2Chunk = Chunk()
  Exit Function

ProcErr:
  Image2Chunk = 0
End Function

Private Function Chunk2Image(Chunk() As Byte, Optional Filename As String) As Variant
On Error GoTo ProcErr
Dim KeepFile As Boolean
Dim Datafile As Integer

    KeepFile = True
    If Trim(Filename) = "" Then
      Filename = "c:\tmpxxdb.fil"
      KeepFile = False
    End If

    Datafile = FreeFile
    Open Filename For Binary Access Write As Datafile
      Put Datafile, , Chunk()
    Close Datafile

ProcExit:
  Set Chunk2Image = LoadPicture(Filename)
  On Error Resume Next
  If Not KeepFile Then Kill Filename
  Exit Function

ProcErr:
  On Error Resume Next
  Kill Filename
  Chunk2Image = 0
End Function

Private Function Hex2Dec(valHex As String) As Integer
On Error Resume Next
Dim X As Integer, Y As Integer, chrInput As String

  For X = 0 To Len(valHex) - 1
    Y = Len(valHex) - X
    chrInput = Mid(valHex, Y, 1)

    If Asc(chrInput) >= 48 And Asc(chrInput) <= 57 Then
      chrInput = chrInput
    ElseIf Asc(chrInput) >= 65 And Asc(chrInput) <= 70 Then
      chrInput = Asc(chrInput) - 55
    ElseIf Asc(chrInput) >= 97 And Asc(chrInput) <= 102 Then
      chrInput = Asc(chrInput) - 87
    End If
    Hex2Dec = Hex2Dec + 16 ^ X * chrInput
  Next X
End Function

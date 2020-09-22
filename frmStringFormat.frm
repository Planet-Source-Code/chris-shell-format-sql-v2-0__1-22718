VERSION 5.00
Begin VB.Form frmStringFormat 
   Caption         =   "Format SQL v2.0"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   Icon            =   "frmStringFormat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClearAll 
      Caption         =   "Clear All"
      Height          =   390
      Left            =   2550
      TabIndex        =   19
      Top             =   7350
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   7425
      TabIndex        =   14
      Top             =   7350
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   6375
      TabIndex        =   13
      Top             =   7350
      Width           =   990
   End
   Begin VB.Frame fOutput 
      Caption         =   "Output"
      Height          =   615
      Left            =   150
      TabIndex        =   10
      Top             =   7200
      Width           =   2265
      Begin VB.OptionButton optFile 
         Caption         =   "Notepad"
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   225
         Width           =   915
      End
      Begin VB.OptionButton optClipBoard 
         Caption         =   "Clipboard"
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   225
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "Format String"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4035
      Width           =   3735
   End
   Begin VB.Frame fOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2040
      TabIndex        =   6
      Top             =   150
      Width           =   6375
      Begin VB.CheckBox ckContinue 
         Caption         =   "Line Continuation"
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   525
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.TextBox txtLineLen 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4350
         TabIndex        =   17
         Text            =   "30"
         Top             =   450
         Width           =   390
      End
      Begin VB.CheckBox ckQuotes 
         Caption         =   "Double Quotes to Single Quotes"
         Height          =   255
         Left            =   3300
         TabIndex        =   8
         Top             =   225
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox ckVarible 
         Caption         =   "Make Variable"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label lblLineLen 
         Caption         =   "Line Length:"
         Height          =   240
         Left            =   3300
         TabIndex        =   16
         Top             =   525
         Width           =   990
      End
   End
   Begin VB.TextBox txtNewString 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4440
      Width           =   8295
   End
   Begin VB.TextBox txtOldString 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1320
      Width           =   8295
   End
   Begin VB.TextBox txtVar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "strSQL"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "http://www.cshellvb.com"
      Height          =   240
      Left            =   3600
      TabIndex        =   21
      Top             =   7500
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "Written By: Chris Shell"
      Height          =   240
      Left            =   3600
      TabIndex        =   20
      Top             =   7200
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   3300
      TabIndex        =   15
      Top             =   375
      Width           =   1665
   End
   Begin VB.Label lblNewString 
      Caption         =   "New String:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lOldText 
      Caption         =   "String to be Formatted:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1950
   End
   Begin VB.Label lVar 
      Caption         =   "Varible Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmStringFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *************************************************************
'  Format String
'  Chris Shell
'  http://www.cshellvb.com
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************

Const DIM_STR1 As String = "Dim "
Const DIM_STR2 As String = " as String"
Const CONT_STR As String = " & _"
Const CONNECT_STR As String = " & "

Const SELECT_STR As String = "SELECT "
Const FROM_STR As String = " FROM "
Const WHERE_STR As String = " WHERE "
Const GROUPBY_STR As String = " GROUP BY "
Const UPDATE_STR As String = "UPDATE "
Const INSERT_STR As String = "INSERT INTO "
Const DELETE_STR As String = "DELETE "

Dim aSQLVar() As Integer

'**************************************
'Windows API/Global Declarations for :
'Create links from labels!
'**************************************

Public Enum OpType
    Startup = 1
    Click = 2
    FormMove = 3
    LinkMove = 4
End Enum

Dim Clicked As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Function FormatString(sOld As String, bSQLSmart As Boolean, bContinue As Boolean, _
    iLineCnt As Integer, sVariable As String, bFixQuotes As Boolean) As String

Dim lStrLen As Integer
Dim lcv As Integer, lCnt As Integer
Dim StartPos As Integer, EndPos As Integer
Dim sVar As String
Dim sPart As String

Dim iSQLCnt As Integer, iSQLLen As Integer
Dim iSELECT As Integer, iFROM As Integer, iWHERE As Integer, iGROUPBY As Integer
Dim iUPDATE As Integer, iINSERT As Integer, iDELETE As Integer

'********************************************
'This class was written by:
'   Karl E. Peterson
'   http://www.mvps.org/vb/
'See the class for more detail. Thank you to
'him for this code, I got it via VBPJ Article
'on string building in ASP.
'********************************************
Dim cSBld As New CStringBuilder





'Chr(34) = "
'Chr(39) = '

On Error GoTo ehHandle
    
        
    '********************************************
    'Clean up String before we begin...
    '********************************************
        'Remove Tab Characters
        sOld = RemoveChar(sOld, CStr(vbTab))
        'Remove Vertical Tab Characters
        'sOld = RemoveChar(sOld, vbVerticalTab)
        'Remove Carriage Returns
        sOld = Replace(sOld, CStr(vbCr), " ")
        'Remove Line Feeds
        sOld = RemoveChar(sOld, CStr(vbLf))
        'Remove extra Spaces
        sOld = Trim(sOld)
        
        
    '********************************************
    'Ready to Rock...
    '********************************************
       
    
    'Replace any quotes with single quotes if desired
    If bFixQuotes = True Then
        sOld = CleanString(sOld)
    End If
    
    'Store original length
    lStrLen = Len(sOld)
    
    'If a variable is given te use it...
    If Len(sVariable) > 0 Then
        sVar = sVariable
        cSBld.Append DIM_STR1 & sVar & DIM_STR2 & vbCrLf & vbCrLf
    Else
        sVar = "strSQL"
    End If
    
    'Place some space between the declare and the code
    cSBld.Append vbCrLf & vbTab
    
    'Set initial values prior to loop
    StartPos = 1
    lCnt = 0
    iSQLCnt = 0
        
    'Essentially, we go through each character iin the string (VB does this nicely).
    'If we reach are character count (iLineCnt) then we make a new line.
    'We do this until we reach the end...
    For lcv = 0 To lStrLen
            lCnt = lCnt + 1
            
            If lcv = 0 Then
                    cSBld.Append sVar & " = "
            End If
            
            If bSQLSmart Then
               If (lCnt = aSQLVar(iSQLCnt)) Then
                    If iSQLCnt = 0 Then
                                       
                    End If
                    iSQLCnt = iSQLCnt + 1
                    
               End If
               
            End If
                        
            If (lCnt = iLineCnt) Or (lcv >= lStrLen) Then
                lCnt = 0
                
                If bContinue Then
                    'Are we at the End
                    If (lcv >= lStrLen) Then
                        cSBld.Append Chr(34) & Mid(sOld, StartPos, (lStrLen - lcv)) & Chr(34)
                    Else
                        cSBld.Append Chr(34) & Mid(sOld, StartPos, iLineCnt) & Chr(34) & CONT_STR
                    End If
                Else
                    'Are we at the End
                    If (lcv >= lStrLen) Then
                        cSBld.Append sVar & " = " & sVar & CONNECT_STR & _
                            Chr(34) & Mid(sOld, StartPos, (lStrLen - lcv)) & Chr(34)
                    Else
                        cSBld.Append sVar & " = " & sVar & CONNECT_STR & _
                            Chr(34) & Mid(sOld, StartPos, iLineCnt) & Chr(34)
                    End If
                
                
                End If
                
                iLineCnt = iLineCnt + 1
                
                If StartPos = 1 Then
                    StartPos = lcv + 2
                Else
                    StartPos = lcv + 1
                End If
                
                cSBld.Append vbCrLf
                                
            End If
            
    Next lcv
    
    'Pass the String Back...
    FormatString = cSBld.ToString
    
    Set cSBld = Nothing
    
    
ExitFunc:
    Exit Function


ehHandle:
    MsgBox "ERROR: " & Err.Number & " - " & Err.Description
    Resume Next


End Function
Function CleanString(szOriginal)
    If szOriginal = "" Then
        CleanString = "NULL"
    Else
        CleanString = Substitute(szOriginal, "'", "''")
        CleanString = Substitute(CleanString, "’", "’’")
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub CmdClearAll_Click()

    If MsgBox("Clear all text boxes?", vbYesNo + vbQuestion, "Clear") = vbYes Then
        txtNewString.Text = ""
        txtOldString.Text = ""
        
    End If
    
    Me.Refresh
    
End Sub

Private Sub cmdFormat_Click()
Dim bContinue As Boolean, sVar As String
Dim bQuote As Boolean, iCnt As Integer, bSQLSmart As Boolean

    bContinue = False
    sVar = ""
    bQuote = False
    bSQLSmart = False
    
    
    If Len(txtOldString.Text) = 0 Then
        MsgBox "No String entered!", vbExclamation
        Exit Sub
    End If
    
'    If ckSQLSmart.Value = vbChecked Then
'       bSQLSmart = True
'    End If
    
    If ckContinue.Value = vbChecked Then
       bContinue = True
    End If
    
    If ckQuotes.Value = vbChecked Then
        bQuote = True
    End If
    
    If ckVarible.Value = vbChecked Then
        sVar = txtVar.Text
    End If
    
    If IsNumeric(txtLineLen.Text) Then
        iCnt = CInt(txtLineLen.Text)
    Else
        iCnt = 50
    End If
        
    'txtNewString.Text = FormatString(txtOldString.Text, False, bContinue, iCnt, sVar, bQuote)
    txtNewString.Text = FormatSQL(txtOldString.Text, True, bContinue, iCnt, sVar, bQuote)

End Sub
Function Substitute(szBuff, szOldString, szNewString)
    Dim iStart
    Dim iEnd
    
    
    ''' Find first substring
    iStart = InStr(1, szBuff, szOldString)
    
    ''' Loop through finding substrings
    Do While iStart <> 0
        ''' Find end of string
        iEnd = iStart + Len(szOldString)
        ''' Concatenate new string
        szBuff = Left(szBuff, iStart - 1) & szNewString & Right(szBuff, Len(szBuff) - iEnd + 1)
        ''' Advance past new string
        iStart = iStart + Len(szNewString)
        ''' Find next occurrence
        iStart = InStr(iStart, szBuff, szOldString)
    Loop
    
    Substitute = szBuff
End Function

Function RemoveChar(sText As String, sChar As String) As String
    Dim iPos As Integer, iStart As Integer
    Dim sTemp As String
    iStart = 1


    Do
        iPos = InStr(iStart, sText, sChar)


        If iPos <> 0 Then
            sTemp = sTemp & Mid(sText, iStart, (iPos - iStart))
            iStart = iPos + 1
        End If
    Loop Until iPos = 0
    sTemp = sTemp & Mid(sText, iStart)
    RemoveChar = sTemp
End Function

Sub SQLVarPos(ByVal sSQL As String)
Dim lcv As Integer
Dim iPos As Integer
Dim iLen As Integer

On Error GoTo ehHandle

    ReDim aSQLVar(3, 1)
    
    Debug.Print sSQL
    
    
    iLen = 0
    '1 SELECT, INSERT, UPDATE, or DELETE
    iPos = InStr(1, UCase(sSQL), SELECT_STR)
    
        
    If iPos = 0 Then
        iPos = InStr(1, UCase(sSQL), INSERT_STR)
        
        If iPos = 0 Then
            iPos = InStr(1, UCase(sSQL), UPDATE_STR)
            
            If iPos = 0 Then
                iPos = InStr(1, UCase(sSQL), DELETE_STR)
            Else
                iLen = Len(UPDATE_STR)
            End If
            
            If iPos <> 0 Then
                iLen = Len(DELETE_STR)
            End If
        Else
            iLen = Len(INSERT_STR)
        End If
    Else
        iLen = Len(SELECT_STR)
    End If
    
    If iPos > 0 Then
        aSQLVar(0, 0) = iPos
        aSQLVar(0, 1) = iLen
        
    Else
        aSQLVar(0, 0) = -1
        aSQLVar(0, 1) = 0
    End If
    
            
    '2 FROM Clause
    iPos = InStr(1, UCase(sSQL), FROM_STR)
    iLen = Len(FROM_STR)
    
    If iPos > 0 Then
        aSQLVar(1, 0) = iPos
        aSQLVar(1, 1) = iLen
        
    Else
        aSQLVar(1, 0) = -1
        aSQLVar(1, 1) = 0
    End If
    
    '3 WHERE Clause
    iPos = InStr(1, UCase(sSQL), WHERE_STR)
    iLen = Len(WHERE_STR)
    
    If iPos > 0 Then
        aSQLVar(2, 0) = iPos
        aSQLVar(2, 1) = iLen
        
    Else
        aSQLVar(2, 0) = -1
        aSQLVar(2, 1) = 0
    End If
            
    '4 GROUP BY Clause
    iPos = InStr(1, UCase(sSQL), GROUPBY_STR)
    iLen = Len(GROUPBY_STR)
    
    If iPos > 0 Then
        aSQLVar(3, 0) = iPos
        aSQLVar(3, 1) = iLen
        
    Else
        aSQLVar(3, 0) = -1
        aSQLVar(3, 1) = 0
    End If
    
SUB_EXIT:
    Exit Sub
    
ehHandle:
    MsgBox "SQLVarPos: " & Err.Number & " - " & Err.Description
    Resume Next
    
    
    
End Sub

Private Sub cmdOK_Click()
Dim hFile As Integer
Dim sFilename As String
Dim iFileName As Integer

    

    If optClipBoard.Value = True Then
        ClipboardCopy txtNewString.Text
        MsgBox "Your code on the Clipboard, Enjoy!", vbExclamation
        
    Else
        
        iFileName = Int((10000 - 100 + 1) * Rnd + 100)
        'obtain the next free file handle from the system
        hFile = FreeFile
        sFilename = App.Path & "\tmp" & iFileName & ".txt"
         
        'open and save the textbox to a file
        Open sFilename For Output As #hFile
            Print #hFile, (txtNewString.Text)
        Close #hFile

        If Err.Number <> 0 Then
            MsgBox "Problem creating temporary file! The disk may be full or read only.", vbExclamation
            Err.Clear
            Exit Sub
        End If
        
        Call Shell("Notepad " & sFilename, vbNormalFocus)
                
        Kill sFilename
        
        MsgBox "Your code is in Notepad, Enjoy!", vbExclamation
        
    End If
    

    Unload Me
    

End Sub
Public Sub ClipboardCopy(Text As String)
'Copies text to the clipboard
On Error GoTo error
    Clipboard.Clear
    Clipboard.SetText Text$
    
Exit Sub

error:  MsgBox Err.Description, vbExclamation, "Error"

End Sub





Private Sub Form_Load()
    MakeLink Label3, Startup
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label3, FormMove
End Sub


Private Sub Form_Resize()
    ResizeForm Me
    
End Sub

Private Sub Label3_Click()
    MakeLink Label3, Click, Me
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeLink Label3, LinkMove
End Sub



Public Sub MakeLink(LabelName As Label, Operation As OpType, Optional FormName As Form)
    Dim Openpage As Integer


    Select Case Operation
        Case LinkMove
        LabelName.ForeColor = 255
        LabelName.FontUnderline = True
        Case Click
        Openpage = ShellExecute(Me.hwnd, "Open", LabelName.Caption, "", App.Path, 1)
        LabelName.ForeColor = 8388736
        Clicked = True
        Case FormMove
        LabelName.FontUnderline = False


        If Not Clicked Then
            LabelName.ForeColor = 16711680
        Else
            LabelName.ForeColor = 8388736
        End If
        Case Startup
        LabelName.ForeColor = 16711680
    End Select
End Sub
        

Public Function FormatSQL(sOld As String, bSQLSmart As Boolean, bContinue As Boolean, _
    iCharCnt As Integer, sVariable As String, bFixQuotes As Boolean) As String

Dim lStrLen As Integer
Dim lcv As Integer, lCnt As Integer
Dim StartPos As Integer, EndPos As Integer
Dim sVar As String
Dim sPart As String
Dim iLineLen As Integer, iLineCnt As Integer
Dim bSQLPart As Boolean
'********************************************
'This class was written by:
'   Karl E. Peterson
'   http://www.mvps.org/vb/
'See the class for more detail. Thank you to
'him for this code, I got it via VBPJ Article
'on string building in ASP.
'********************************************
Dim cSBld As New CStringBuilder


'Chr(34) = "
'Chr(39) = '

On Error GoTo ehHandle
    
        
    '********************************************
    'Clean up String before we begin...
    '********************************************
        
        'Remove Tab Characters
        sOld = RemoveChar(sOld, CStr(vbTab))
        
        'Remove Carriage Returns
        sOld = Replace(sOld, CStr(vbCr), CStr(Chr(32)))
        
        'Remove Line Feeds
        sOld = RemoveChar(sOld, CStr(vbLf))
        
        'Remove extra Spaces
        sOld = Trim(sOld)
        
        
    '********************************************
    'Ready to Rock...
    '********************************************
       
    
    'Replace any quotes with single quotes if desired
    If bFixQuotes = True Then
        sOld = CleanString(sOld)
    End If
    
    'Store original length
    lStrLen = Len(sOld)
    
    'If a variable is given te use it...
    If Len(sVariable) > 0 Then
        sVar = sVariable
        cSBld.Append DIM_STR1 & sVar & DIM_STR2 & vbCrLf & vbCrLf
    Else
        sVar = "strSQL"
    End If
    
    'Set Key SQL Positions in Array
    Call SQLVarPos(sOld)
    
        
    'Place some space between the declare and the code
    cSBld.Append vbCrLf & vbTab
    
    'Set initial values prior to loop
    StartPos = 1
    lCnt = 0
    'iSQLCnt = 0
        
    'Essentially, we go through each character iin the string (VB does this nicely).
    'If we reach are character count (iCharCnt) then we make a new line.
    'We do this until we reach the end...
    
    If lStrLen <= iCharCnt Then
        cSBld.Append sVar & " = "
        cSBld.Append ContinueString(sOld, True)
    Else
        For lcv = 0 To lStrLen
                lCnt = lCnt + 1
                
                If lcv = 0 Then
                        cSBld.Append sVar & " = "
                End If
                
                iLineLen = 0
                bSQLPart = False
                
                'Check if we should to cut here?
                Select Case True
                
                Case aSQLVar(0, 0) = lcv
                    iLineLen = aSQLVar(0, 1)
                    bSQLPart = True
                Case aSQLVar(1, 0) = lcv
                    If lCnt > 1 Then
                        iLineLen = -1
                        
                    Else
                        iLineLen = aSQLVar(1, 1)
                        bSQLPart = True
                    End If
                    
                Case aSQLVar(2, 0) = lcv
                    If lCnt > 1 Then
                        iLineLen = -1
                        
                    Else
                        iLineLen = aSQLVar(2, 1)
                        bSQLPart = True
                    End If
                    
                Case aSQLVar(3, 0) = lcv
                    If lCnt > 1 Then
                        iLineLen = -1
                        
                    Else
                        iLineLen = aSQLVar(3, 1)
                        bSQLPart = True
                    End If
                Case (lCnt = iCharCnt) Or (lcv = lStrLen)
                    
                    If (iCharCnt < (lStrLen - lcv)) Then
                        For iLineLen = 0 To 50
                            Debug.Print Asc(Mid(sOld, (lcv - iLineLen), 1))
                            
                            If Asc(Mid(sOld, (lcv - iLineLen), 1)) = 32 Or _
                                Asc(Mid(sOld, (lcv - iLineLen), 1)) = 44 Then
                                
                                lcv = lcv - iLineLen
                                iLineLen = iCharCnt - iLineLen
                                Exit For
                                
                            End If
                        Next
                        
                    Else
                        'We Should be Done!
                        iLineLen = (lStrLen - (StartPos - 1))
                        
                    End If
                    
                End Select
                
                'This means get whatever is remaining right now
                If iLineLen = -1 Then
                    'lcv = lcv - 1
                    iLineLen = lCnt
                End If
                
                If iLineLen > 0 Then
                    lCnt = 0
                    
                    If bContinue Then
                        cSBld.Append ContinueString(Mid(sOld, StartPos, iLineLen), CBool(((iLineLen + lcv) >= lStrLen)))
                    Else
                        
                        cSBld.Append AppendString(sVar, Mid(sOld, StartPos, iLineLen), CBool(((iLineLen + lcv) >= lStrLen)))
                    End If
                    
                    If CBool(((iLineLen + lcv) >= lStrLen)) = True Then
                        
                        Exit For
                    End If
                    
                    If bSQLPart Then
                        lcv = lcv + iLineLen
                    End If
                    
                    StartPos = lcv
                                        
                    'If StartPos = 1 Then
                    '    StartPos = lcv + 2
                    'Else
                    '    StartPos = lcv + 1
                    'End If
                    
                    cSBld.Append vbCrLf
                             
                    'If CBool(((iLineLen + lcv) >= lStrLen)) = False Then
                    '    lcv = lcv + iLineLen
                    'Else
                    '    Exit For
                    'End If
                             
                End If
                
        Next lcv
    End If
    'Pass the String Back...
    FormatSQL = cSBld.ToString
    
    Set cSBld = Nothing
    
    
ExitFunc:
    Exit Function


ehHandle:
    MsgBox "ERROR: " & Err.Number & " - " & Err.Description
    Resume Next


End Function

Private Function ContinueString(sLine As String, bEnd As Boolean) As String
    
    If bEnd Then
        ContinueString = Chr(34) & sLine & Chr(34)
    Else
        ContinueString = Chr(34) & sLine & Chr(34) & CONT_STR
    End If
    

End Function

Private Function AppendString(sVar As String, sLine As String, bEnd As Boolean) As String

    If bEnd Then
         AppendString = sVar & " = " & sVar & CONNECT_STR & _
                                Chr(34) & sLine & Chr(34)
    Else
         AppendString = sVar & " = " & sVar & CONNECT_STR & _
                                Chr(34) & sLine & Chr(34)
    End If

End Function

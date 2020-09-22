VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConsoleApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CCONSOLE Class
'
' Creates a console window and writes to it
'''''D E C L A R A T I O N S''''''''''''''''''''''''''''''''''''

Private Declare Function AllocConsole Lib "kernel32" () As Long

Private Declare Function FreeConsole Lib "kernel32" () As Long

Private Declare Function GetStdHandle Lib "kernel32" _
(ByVal nStdHandle As Long) As Long

Private Declare Function ReadConsole Lib "kernel32" Alias _
"ReadConsoleA" (ByVal hConsoleInput As Long, _
ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, _
lpNumberOfCharsRead As Long, lpReserved As Any) As Long

Private Declare Function SetConsoleMode Lib "kernel32" (ByVal _
hConsoleOutput As Long, dwMode As Long) As Long

Private Declare Function SetConsoleTextAttribute Lib _
"kernel32" (ByVal hConsoleOutput As Long, ByVal _
wAttributes As Long) As Long

Private Declare Function SetConsoleTitle Lib "kernel32" Alias _
"SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long

Private Declare Function WriteConsole Lib "kernel32" Alias _
"WriteConsoleA" (ByVal hConsoleOutput As Long, _
ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, _
lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

''''C O N S T A N T S'''''''''''''''''''''''''''''''''''''
'I/O handlers for the console window. These are much like the
'hWnd handlers to form windows.

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

Public Enum enmConsoleForegroundColors
    cslForeColor_Blue = &H1
    cslForeColor_Green = &H2
    cslForeColor_Red = &H4
    cslForeColor_Intensity = &H8
End Enum

Public Enum enmConsoleBackgroundColors
    cslBackColor_Blue = &H10
    cslBackColor_Green = &H20
    cslBackColor_Red = &H40
    cslBackColor_Intensity = &H80
End Enum




'For SetConsoleMode (input)
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
'For SetConsoleMode (output)
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

'''''G L O B A L S'''''''''''''''''''''''''''''''''''
Private hConsoleIn As Long 'The console's input handle
Private hConsoleOut As Long 'The console's output handle
Private hConsoleErr As Long 'The console's error handle

Private m_Handle As Long
Private m_Title As String
Private m_LastUserInput As String
Private m_TextBackColor As enmConsoleBackgroundColors
Private m_TextForeColor As enmConsoleForegroundColors

Public Property Get InputHandle() As Long
InputHandle = hConsoleIn
End Property

Public Property Get OutputHandle() As Long
OutputHandle = hConsoleOut
End Property

Public Property Get ErrorHandle() As Long
ErrorHandle = hConsoleErr
End Property

Public Property Get TextForeColor() As enmConsoleForegroundColors
TextForeColor = m_TextForeColor
End Property

Public Property Let TextForeColor(ByVal Value As enmConsoleForegroundColors)
m_TextForeColor = Value

Call SetConsoleTextAttribute(hConsoleOut, Value)
End Property

Public Property Get TextBackColor() As enmConsoleBackgroundColors
TextBackColor = m_TextBackColor
End Property

Public Property Let TextBackColor(ByVal Value As enmConsoleBackgroundColors)
m_TextBackColor = Value
Call SetConsoleTextAttribute(hConsoleOut, Value)
End Property

Public Property Get LastUserInput() As String
LastUserInput = m_LastUserInput
End Property

Public Property Get Title() As String
Title = m_Title
End Property

Public Property Let Title(ByVal Value As String)
m_Title = Value

If hConsoleOut Then
    SetConsoleTitle m_Title
Else
    Err.Raise vbObjectError + 5567, , "Console not initialized"
End If

End Property

Public Sub Create(Optional ByVal ConsoleTitle As String = "")
AllocConsole 'Create a console instance
'Get the console's handle
hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)

m_Title = ConsoleTitle
SetConsoleTitle ConsoleTitle
End Sub

Public Sub WriteText(ByVal Text As String, _
   Optional ByVal Forecolor As enmConsoleForegroundColors = 0, _
   Optional ByVal Backcolor As enmConsoleBackgroundColors = 0)

'==Set the forcolor and backcolor properties==

If Forecolor <> 0 Then TextForeColor = Forecolor

If Backcolor <> 0 Then TextBackColor = Backcolor

'Use the vbCrLf to get to a new line.
ConsolePrint Text
End Sub

Public Sub WriteLine(ByVal Text As String, _
Optional ByVal Forecolor As enmConsoleForegroundColors = 0, _
Optional ByVal Backcolor As enmConsoleBackgroundColors = 0)

Call WriteText(Text & vbCrLf, Forecolor, Backcolor)
End Sub

Public Function GetUserInput() As String
m_LastUserInput = ConsoleRead()

If Not m_LastUserInput = vbNullString Then
    GetUserInput = m_LastUserInput
End If

End Function

Public Sub Destroy()
FreeConsole
End Sub



'Function: ConsolePrint
'
'Summary: Prints the output of a string
'
'Args: String ConsolePrint
'The string to be printed to the console's ouput buffer.
'
'Returns: None
'
'-----------------------------------------------------

Private Sub ConsolePrint(szOut As String)
WriteConsole hConsoleOut, szOut, Len(szOut), vbNull, vbNull
End Sub

'F+F++++++++++++++++++++++++++++++++++++++++++++++++++++
'Function: ConsoleRead
'
'Summary: Gets a line of input from the user.
'
'Args: None
'
'Returns: String ConsoleRead
'The line of input from the user.
'---------------------------------------------------F-F

Private Function ConsoleRead() As String
Dim sUserInput As String * 256
Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
'Trim off the NULL charactors and the CRLF.
ConsoleRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function

Private Sub Class_Terminate()
Call Destroy
End Sub

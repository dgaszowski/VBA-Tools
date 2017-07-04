VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Error Handler Class
'
'

Private Const ERROR_LOG = "Error.log"
Private Const ACCESS = "Acsess"
Private Const EXCEL = "Excel"
Private Const WORD = "WORD"

#Const DEBUG_MODE = True
#Const APPLICATION = ACCESS

#If APPLICATION = ACCESS Then
    Option Compare Database
#Else
    Option Compare Text
#End If

Option Explicit

Private m_objErrorTable As Collection ' Numbers and descriptions of custom errors.
Private m_objCallStack As Collection

Private m_strPath As String
Private m_strLogFileName As String

Public Property Set ErrorTable(ByRef objErrorTable As Collection)
    
    ' Custom error table.
    Set m_objErrorTable = objErrorTable
End Property

Private Property Get ErrorMessage(ByVal lngErrNumber As Long) As String

    ' Custom error table.
    ' Return user defined error description.
    ErrorMessage = m_objErrorTable(CStr(lngErrNumber))
End Property

Public Property Let LogFileName(ByVal strFileName As String)
    
    m_strLogFileName = strFileName
End Property

Public Property Get LogFileName() As String

    LogFileName = m_strLogFileName
End Property

Public Sub raiseException(ByVal lngErrNb As Long)
'
' Raises an exception instead of Err.Raise method.
' Puts the error down on disk.
' Prints debug information.
' Shows error message.
'

    Dim strErrorMessage, strErrorLine As String
    
    If lngErrNb <> 0 Then
        
        ' Is the exception user defined?
        If lngErrNb < 0 Then
            ' It is. Lets use custom error message.
            strErrorMessage = ErrorMessage(lngErrNb)
        Else
            ' Nope. Using standard error message.
            strErrorMessage = Error(lngErrNb)
        End If
        
        strErrorLine = Now & ", " & _
                      "Error nb, description: " & lngErrNb & ", " & strErrorMessage & "." & vbNewLine & vbTab & _
                      "Call stack:" & strPrintCallStack
                      
        ' Save to log file.
        Call writeError(strErrorLine)
        
        #If DEBUG_MODE Then
            Debug.Print strErrorLine
        #End If
        
        strErrorLine = "Error number: " & lngErrNb & vbNewLine & _
                       "Error description: " & strErrorMessage & "." & vbNewLine & vbNewLine & _
                       "Call stack:" & strPrintCallStack & vbNewLine & _
                       "Error has been logged."
        
        MsgBox strErrorLine, vbCritical, "Error"
        
        Stop
    End If
End Sub

Public Function Push(ByVal strMethodName As String) As String
'
' Pushes a method on call stack and returns its name.
' Should be executed at the beginning of every method.
'
    With m_objCallStack
        .Add strMethodName
        Push = m_objCallStack(.Count)
    End With
End Function

Public Function Pop() As String
'
' Removes a method from call stack and returns its name.
' Should be executed at the end of every method upon successfull execution.
'

    With m_objCallStack
        If .Count > 0 Then
            Pop = .Item(.Count)
            .Remove .Count
        End If
    End With
End Function

Private Sub Class_Initialize()
    
    #If APPLICATION = ACCESS Then
        m_strPath = CurrentProject.Path & "\"
    #ElseIf APPLICATION = EXCEL Then
        m_strPath = Workbook.Path & "\"
    #ElseIf APPLICATION = WORD Then
        m_strPath = Document.Path & "\"
    #End If
    
    LogFileName = ERROR_LOG
    
    Set m_objCallStack = New Collection

End Sub

Private Sub Class_Terminate()

    Set m_objCallStack = Nothing
End Sub

Private Function strPrintCallStack() As String
'
' Returns a string containing all methods present at a call stack.
' Includes formatting, like tabs and new lines.
'

    Dim strCall As Variant
    Dim lngCount As Long
    
    strPrintCallStack = vbNewLine
    
    With m_objCallStack
        lngCount = .Count
        
        Do While lngCount > 0
            strCall = CStr(m_objCallStack(lngCount))
            strPrintCallStack = strPrintCallStack & vbTab & lngCount & ") " & strCall & vbNewLine
            lngCount = lngCount - 1
        Loop
    End With
    
    strPrintCallStack = strPrintCallStack + vbNewLine
End Function

Private Sub writeError(ByVal strErrorText As String)
'
' Writes error string to disk.
'

    ' Setting LogFileName to empty string means that the user doesn't want to log errors.
    If LogFileName = "" Then Exit Sub
    
    Dim objFSO As Object
    Dim objLogFile As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Open or create log file.
    Set objLogFile = objFSO.OpenTextFile(m_strPath & LogFileName, 8, True)
    
    objLogFile.WriteLine strErrorText
    objLogFile.Close
    
    Set objLogFile = Nothing
    Set objFSO = Nothing
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#Const DEBUG_MODE = True
#Const APPLICATION = "Access"

Private m_objErrorTable As Collection
Private m_objCallStack As Collection

Private m_strPath As String
Private m_strLogFileName As String

Private m_strLogErrorName As String
Private m_strErrorTitle As String
Private m_strErrorText As String

Public Property Set ErrorTable(ByRef objErrorTable As Collection)
    
    Set m_objErrorTable = objErrorTable
End Property

Public Property Let LogErrorName(ByVal strName As String)
    
    m_strLogErrorName = strName
End Property

Public Property Get LogErrorName() As String

    LogErrorName = m_strLogErrorName
End Property

Public Property Let ErrorTitle(ByVal strTitle As String)
    
    m_strErrorTitle = strTitle
End Property

Public Property Get ErrorTitle() As String

    ErrorTitle = m_strErrorTitle
End Property

Public Property Let ErrorText(ByVal strText As String)

    m_strErrorText = strText
End Property

Public Property Get ErrorText() As String

    ErrorText = m_strErrorText
End Property

Private Property Get ErrorMessage(ByVal lngErrNumber As Long) As String
    
    Error = m_objErrorTable(CStr(lngErrNumber))
End Property

Public Property Let LogFileName(ByVal strFileName As String)
    
    m_strLogFileName = strFileName
End Property

Public Property Get LogFileName() As String

    LogFileName = m_strLogFileName
End Property

Public Sub raiseException(ByVal lngErrNb As Long)
    Dim blnIsClassError As Boolean
    Dim intFileNb As Integer
    Dim strErrorMessage, strErrorLine As String
    
    If lngErrNb <> 0 Then
        
        If lngErrNb < 0 Then
            strErrorMessage = ErrorMessage(lngErrNb)
        Else
            strErrorMessage = Error(lngErrNb)
        End If
        
        strErrorLine = Now & ", " & _
                      "Error nb, description: " & lngErrNb & ", " & strErrorMessage & "." & vbNewLine & vbTab & _
                      "Call stack:" & strPrintCallStack
        
        Call writeError(strErrorLine)
        
        #If DEBUG_MODE Then
            Debug.Print strErrorLine
        #End If
        
        strErrorLine = "Error number: " & lngErrNb & vbNewLine & _
                       "Error description: " & strErrorMessage & vbNewLine & vbNewLine & _
                       "Call stack:" & strPrintCallStack & vbNewLine & _
                       "Error has been logged."
        
        MsgBox strErrorLine, vbCritical, ErrorTitle
        
        Stop
    End If
End Sub

Public Function Push(ByVal strMethodName As String) As String
    
    With m_objCallStack
        .Add strMethodName
        Push = m_objCallStack(.Count)
    End With
End Function

Public Function Pop() As String
    With m_objCallStack
        If .Count > 0 Then
            Pop = .Item(.Count)
            .Remove .Count
        End If
    End With
End Function

Private Sub Class_Initialize()
    
    #If APPLICATION = "Access" Then
        m_strPath = CurrentProject.Path & "\"
    #ElseIf APPLICATION = "Excel" Then
        m_strPath = Workbook.Path & "\"
    #ElseIf APPLICATION = "Word" Then
        m_strPath = Document.Path & "\"
    #End If
    
    ErrorTitle = "Error"
    ErrorText = "Error has occured."
    
    LogFileName = "Error.log"
    
    Set m_objCallStack = New Collection

End Sub

Private Sub Class_Terminate()

    Set m_objCallStack = Nothing
End Sub

Private Function strPrintCallStack() As String
    Dim varCall As Variant
    Dim lngCount As Long
    
    strPrintCallStack = vbNewLine
    
    With m_objCallStack
        lngCount = .Count
        
        Do While lngCount > 0
            varCall = m_objCallStack(lngCount)
            strPrintCallStack = strPrintCallStack & vbTab & lngCount & ") " & varCall & vbNewLine
            lngCount = lngCount - 1
        Loop
    End With
    
    strPrintCallStack = strPrintCallStack + vbNewLine
End Function

Private Sub writeError(ByVal strErrorText As String)
    Dim objFSO As Object
    Dim objLogFile As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objLogFile = objFSO.OpenTextFile(m_strPath & LogFileName, 8)
    
    objLogFile.WriteLine strErrorText
    objLogFile.Close
    
    Set objLogFile = Nothing
    Set objFSO = Nothing
End Sub

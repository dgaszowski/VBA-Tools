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

Private m_objErrorTable As Collection
Private m_objCallStack As Collection

Private m_strLogFileName As String

Private m_strLogErrorName As String
Private m_strErrorTitle As String
Private m_strErrorText As String

Private m_blnDebugMode As Boolean

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

Public Property Let DebugMode(ByVal blnDebug As Boolean)

    m_blnDebugMode = blnDebug
End Property

Public Property Get DebugMode() As Boolean

    DebugMode = m_blnDebugMode
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
        
        intFileNb = FreeFile
        
        Open CurrentProject.Path & "\" & LogFileName For Append As #intFileNb
        
        strErrorLine = Now & ", " & _
                      "Error nb, description: " & lngErrNb & ", " & strErrorMessage & "." & vbNewLine & vbTab & _
                      "Call stack:" & strPrintCallStack
        
        Write #intFileNb, strErrorLine
        Close #intFileNb
        
        If DebugMode Then
            Debug.Print strErrorLine
        End If
        
        MsgBox "Error!", vbCritical
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
    
    ErrorTitle = "Generic error title"
    ErrorText = "Error has occured."
    
    LogFileName = "Error.log"
    
    DebugMode = True
    
    Set m_objCallStack = New Collection

End Sub

Private Sub Class_Terminate()

    Set m_objCallStack = Nothing
End Sub

Private Function strPrintCallStack() As String
    Dim varCall As Variant
    
    strPrintCallStack = vbNewLine
    
    For Each varCall In m_objCallStack
        strPrintCallStack = strPrintCallStack & vbTab & varCall & vbNewLine
    Next
End Function

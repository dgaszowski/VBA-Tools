VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Evaluate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const APP = "Access"
#Const DEBUG_MODE = True

Option Explicit

Const PARSE_MISMATCHED_LEFT_PAR = 1 + vbObjectError
Const PARSE_MISMATCHED_RIGHT_PAR = 2 + vbObjectError
Const PARSE_FORBIDDEN = 3 + vbObjectError

Const EVAL_INSUFFICIENT = 4 + vbObjectError
Const EVAL_TOOMANY = 5 + vbObjectError

Const CALCULATION_CONCAT = 6 + vbObjectError
Const CALCULATION_ADD = 7 + vbObjectError
Const CALCULATION_SUB = 8 + vbObjectError
Const CALCULATION_MULT = 9 + vbObjectError
Const CALCULATION_DIV = 10 + vbObjectError

#If APP = "Access" Then
    Option Compare Database
#Else
    Option Compare Text
#End If

Private ERROR_TABLE As Collection
Private m_objErrorHandler As cls_ErrorHandler

Private m_strInfix As String
Private m_strPostfix As String
Private m_varValue As Variant

Private m_objOutputStack As Collection
Private m_objOperatorStack As Collection
Private m_objValueMapping As Collection
Private m_objValues As Collection

Public Property Set ValueMapping(ByRef objMap As Collection)
    Set m_objValueMapping = objMap
End Property

Public Property Let Infix(ByVal strInfix As String)
    m_strInfix = strInfix
    m_strPostfix = strShuntingYard(m_strInfix)
    m_varValue = varEvaluate(m_strPostfix)
End Property

Public Property Get Infix() As String
    Infix = m_strInfix
End Property

Public Property Get Postfix() As String
    Postfix = m_strPostfix
End Property

Public Property Get Value() As Variant
    Value = m_varValue
End Property

Private Sub Class_Initialize()
    
    Set m_objOutputStack = New Collection
    Set m_objOperatorStack = New Collection
    Set m_objValueMapping = New Collection
    Set m_objValues = New Collection
    
    Set ERROR_TABLE = New Collection
    
    With ERROR_TABLE
        .Add "Mismatched left - ) - parenthesis.", CStr(PARSE_MISMATCHED_LEFT_PAR)
        .Add "Mismatched right - ( - parenthesis.", CStr(PARSE_MISMATCHED_RIGHT_PAR)
        .Add "Forbidden character found in infix string.", CStr(PARSE_FORBIDDEN)
        
        .Add "Evaluation error: number of values in infix expression is insufficient.", CStr(EVAL_INSUFFICIENT)
        .Add "Evaluation error: there is too many values in infix expression.", CStr(EVAL_TOOMANY)
        
        .Add "Calculation error: unable to perform concatenation.", CStr(CALCULATION_CONCAT)
        .Add "Calcualtion error: unable to perform addition.", CStr(CALCULATION_ADD)
        .Add "Calculation error: unable to perform substraction.", CStr(CALCULATION_SUB)
        .Add "Calculation error: unable to perform multiplication.", CStr(CALCULATION_MULT)
        .Add "Calculation error: unable to perform division.", CStr(CALCULATION_DIV)
    End With
    
    Set m_objErrorHandler = New cls_ErrorHandler
    Set m_objErrorHandler.ErrorTable = ERROR_TABLE
End Sub

Private Sub Class_Terminate()
    
    Set m_objOutputStack = Nothing
    Set m_objOperatorStack = Nothing
    Set m_objValueMapping = Nothing
    Set m_objValues = Nothing
    
    Set m_objErrorHandler = Nothing
End Sub

Private Function strShuntingYard(ByVal strInfix As String) As String
    Dim I As Long
    Dim varToken As Variant, strTemp As String, objTokenArray As Collection
    Dim objOperatorStack As Collection
    
    On Error GoTo Error_strShuntingYard
    Call m_objErrorHandler.Push("Function: cls_Evaluate.strShuntingYard")
    
    Set objTokenArray = New Collection
    Set objOperatorStack = New Collection
    
    strInfix = Replace(strInfix, " ", "")
    
    For I = 1 To Len(strInfix)
        strTemp = Mid(strInfix, I, 1)
        
        If blnHasForbiddenChar(strTemp) Then
            Call m_objErrorHandler.raiseException(PARSE_FORBIDDEN)
        End If
        
        If Not blnIsOperator(strTemp) And (strTemp <> ")" And strTemp <> "(") Then
            varToken = varToken & strTemp
        Else
            If varToken <> "" Then
                objTokenArray.Add varToken
                m_objValues.Add varToken
            End If
            
            objTokenArray.Add strTemp
            varToken = ""
        End If
    Next
    
    objTokenArray.Add varToken
    
    For Each varToken In objTokenArray
        
        Select Case varToken
            Case "("
                objOperatorStack.Add varToken
                
            Case ")"
                Do While objOperatorStack.Item(objOperatorStack.Count) <> "("
                    strShuntingYard = strShuntingYard & " " & objOperatorStack.Item(objOperatorStack.Count)
                        
                    Call objOperatorStack.Remove(objOperatorStack.Count)
                    If objOperatorStack.Count = 0 Then
                        Call m_objErrorHandler.raiseException(PARSE_MISMATCHED_LEFT_PAR)
                        Exit Do
                    End If
                Loop
                
                If objOperatorStack.Count > 0 Then objOperatorStack.Remove (objOperatorStack.Count)
                
            Case Else
            
                If Not blnIsOperator(CStr(varToken)) Then
                    strShuntingYard = strShuntingYard & " " & varToken
                Else
                    If objOperatorStack.Count >= 1 Then
                        Do While btPrecedence(objOperatorStack.Item(objOperatorStack.Count)) >= btPrecedence(CStr(varToken))
                    
                            strShuntingYard = strShuntingYard & " " & objOperatorStack.Item(objOperatorStack.Count)
    
                            Call objOperatorStack.Remove(objOperatorStack.Count)
                            If objOperatorStack.Count = 0 Then Exit Do
                        Loop
                    End If
                
                    objOperatorStack.Add varToken
                End If
                
        End Select
    Next
    
    For Each varToken In objOperatorStack
        If varToken = "(" Then
            Call m_objErrorHandler.raiseException(PARSE_MISMATCHED_RIGHT_PAR)
            Exit For
        End If
        
        strShuntingYard = strShuntingYard & " " & varToken
    Next
    
Exit_strShuntingYard:
    Set objTokenArray = Nothing
    Set objOperatorStack = Nothing
    
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_strShuntingYard:
    Call m_objErrorHandler.raiseException(Err.Number)
    
End Function

Private Function varEvaluate(ByVal strPostfix As String) As Variant
    Dim objStack As Collection
    Dim I As Long
    Dim strToken As String
    Dim astrPostfix() As String
    Dim varArg1 As Variant, varArg2 As Variant, varResult As Variant
    
    On Error GoTo Error_varEvaluate
    Call m_objErrorHandler.Push("Function: cls_Evaluate.varEvaluate")
    
    Set objStack = New Collection
    
    strPostfix = Trim(strPostfix)
    astrPostfix = Split(strPostfix, " ")
    
    For I = 0 To UBound(astrPostfix)
        
        strToken = astrPostfix(I)
        
        If Not blnIsOperator(strToken) Then
            objStack.Add varGetValue(strToken)
        Else
            If objStack.Count < 2 Then
                Call m_objErrorHandler.raiseException(EVAL_INSUFFICIENT)
            End If
            
            varArg2 = varGetValue(objStack(objStack.Count))
            varArg1 = varGetValue(objStack(objStack.Count - 1))
            
            Call objStack.Remove(objStack.Count)
            Call objStack.Remove(objStack.Count)
            
            Select Case strToken
                Case "&"
                    varResult = strConcat(varArg1, varArg2)
                Case "+"
                    varResult = varAdd(varArg1, varArg2)
                Case "-"
                    varResult = varSubstract(varArg1, varArg2)
                Case "*"
                    varResult = varMultiply(varArg1, varArg2)
                Case "/"
                    varResult = varDivide(varArg1, varArg2)
            End Select
            
            objStack.Add varResult
        End If
        
    Next
    
    If objStack.Count > 1 Then
        Call m_objErrorHandler.raiseException(EVAL_TOOMANY)
    End If
    
    varEvaluate = objStack(1)
    
Exit_varEvaluate:
    Set objStack = Nothing
    Erase astrPostfix
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_varEvaluate:
    Call m_objErrorHandler.raiseException(Err.Number)
    
End Function

Private Function varGetValue(ByVal varValue As Variant) As Variant
    Dim varTemp As Variant
    
    On Error Resume Next
    
    varTemp = m_objValueMapping(varValue)
    
    If Err.Number <> 0 Then
        varGetValue = varValue
    Else
        varGetValue = varTemp
    End If
    
    On Error GoTo 0
    
End Function

Private Function strConcat(ByVal varArg1 As Variant, varArg2 As Variant) As String
    On Error GoTo Error_strConcat
    
    Call m_objErrorHandler.Push("Function: cls_Evaluate.strConcat")
    
    strConcat = CStr(varArg1) & CStr(varArg2)
    
Exit_strConcat:
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_strConcat:
    Call m_objErrorHandler.raiseException(CALCULATION_CONCAT)
    
End Function

Private Function varAdd(ByVal varArg1 As Variant, varArg2 As Variant) As Variant
    On Error GoTo Error_varAdd
    Call m_objErrorHandler.Push("Function: cls_Evaluate.varAdd")
    
    varAdd = CDbl(varArg1) + CDbl(varArg2)
    
Exit_varAdd:
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_varAdd:
    Call m_objErrorHandler.raiseException(CALCULATION_ADD)
    
End Function

Private Function varSubstract(ByVal varArg1 As Variant, varArg2 As Variant) As Variant
    On Error GoTo Error_varSubstract
    Call m_objErrorHandler.Push("Function: cls_Evaluate.varSubstract")
    
    varSubstract = CDbl(varArg1) - CDbl(varArg2)
    
Exit_varSubstract:
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_varSubstract:
    Call m_objErrorHandler.raiseException(CALCULATION_SUB)
    
End Function

Private Function varMultiply(ByVal varArg1 As Variant, varArg2 As Variant) As Variant
    On Error GoTo Error_varMultiply
    Call m_objErrorHandler.Push("Function: cls_Evaluate.varMultiply")
    
    varMultiply = CDbl(varArg1) * CDbl(varArg2)
    
Exit_varMultiply:
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_varMultiply:
    Call m_objErrorHandler.raiseException(CALCULATION_MULT)
    
End Function

Private Function varDivide(ByVal varArg1 As Variant, varArg2 As Variant) As Variant
    On Error GoTo Error_varDivide
    Call m_objErrorHandler.Push("Function: cls_Evaluate.varDivide")
    
    varDivide = varArg1 / varArg2
    
Exit_varDivide:
    Call m_objErrorHandler.Pop
    Exit Function
    
Error_varDivide:
    Call m_objErrorHandler.raiseException(CALCULATION_DIV)
    
End Function

Private Function blnIsOperator(ByVal strOperator As String) As Boolean

    blnIsOperator = InStr("+-&*/^", strOperator) > 0
End Function

Public Function blnHasForbiddenChar(ByVal strInfix As String) As Boolean
    
    blnHasForbiddenChar = InStr("@#$%_=[{]};:'""""\|<>?~`", strInfix) > 0
End Function

Public Function btPrecedence(ByVal strOperator) As Byte
    If blnIsOperator(strOperator) Then
        btPrecedence = 2 - _
            -(InStr("-+&/*^", strOperator) <> 0) _
            - (InStr("/*^", strOperator) <> 0) _
            - (InStr("^", strOperator) <> 0)
    End If

End Function


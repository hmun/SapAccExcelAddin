Public Class TField
    Public Name As String
    Public Value As String

    Public Function create(pName As String, pValue As String) As TField
        Dim newTField As TField
        newTField = New TField
        newTField.setValues(pName, pValue)
        create = newTField
    End Function

    Public Function setValues(pName As String, pValue As String)
        Name = pName
        Value = pValue
    End Function

    Public Function add(p_Val As Double)
        Dim aVal As Double
        aVal = CDbl(Value)
        aVal = aVal + p_Val
        Value = CStr(aVal)
    End Function

    Public Function subst(p_Val As Double)
        Dim aVal As Double
        aVal = CDbl(Value)
        aVal = aVal - p_Val
        Value = CStr(aVal)
    End Function

    Public Function mul(p_Val As Double)
        Dim aVal As Double
        aVal = CDbl(Value)
        aVal = aVal * p_Val
        Value = CStr(aVal)
    End Function

    Public Function div(p_Val As Double)
        Dim aVal As Double
        aVal = CDbl(Value)
        aVal = aVal / p_Val
        Value = CStr(aVal)
    End Function

End Class

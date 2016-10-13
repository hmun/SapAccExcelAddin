Public Class SAPFormat

    Public Function unpack(val As String, length As Integer) As String
        Dim ZeroStr As String
        If IsNumeric(val) Then
            ZeroStr = "000000000000000000000000000000"
            unpack = Left(ZeroStr, length - Len(val)) & val
        Else
            unpack = val
        End If
    End Function

End Class

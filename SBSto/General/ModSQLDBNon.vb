Module ModSQLDBNon
    Public Function fbuValNum(vriData As Object) As Double
        If IsDBNull(vriData) Or Not IsNumeric(vriData) Then
            fbuValNum = 0
        Else
            fbuValNum = vriData
        End If
    End Function
    Public Function fbuValNull(vriData As Object) As String
        If IsDBNull(vriData) Then
            fbuValNull = "Null"
        Else
            If Trim(vriData) = "" Then
                fbuValNull = "Null"
            Else
                fbuValNull = vriData
            End If
        End If
    End Function
    Public Function fbuValNullStr(vriData As Object) As String
        If IsDBNull(vriData) Then
            fbuValNullStr = "Null"
        Else
            If Trim(vriData) = "" Then
                fbuValNullStr = "Null"
            Else
                fbuValNullStr = "'" & vriData & "'"
            End If
        End If
    End Function
    Public Function fbuValStr(vriData As Object) As String
        If IsDBNull(vriData) Then
            fbuValStr = ""
        Else
            fbuValStr = vriData
        End If
    End Function
    Public Function fbuValEmpty(vriData As Object) As String
        If IsDBNull(vriData) Then
            fbuValEmpty = ""
        Else
            If vriData = "&nbsp;" Then
                fbuValEmpty = ""
            Else
                fbuValEmpty = vriData
            End If
        End If
    End Function
    Public Function fbuAlertScript(vriMsg As String) As String
        Dim vnScript As String
        vnScript = "<script language='javascript'>"
        vnScript += "alert('" & vriMsg & "');"
        vnScript += "</script>"
        fbuAlertScript = vnScript
    End Function
End Module

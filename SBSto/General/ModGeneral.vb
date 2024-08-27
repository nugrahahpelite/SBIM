Imports System.IO
Imports GlobalUtil
Imports System.Net.Mail

Module ModGeneral
    Public Const cbuAppPrefix = "SBReq"
    Public Const cbuAppPrefixName = "SBRequest"

    Public Const cbuActionNorm = 0
    Public Const cbuActionNew = 1
    Public Const cbuActionEdit = 2
    Public Const cbuActionDel = 3
    Public Const cbuActionList = 4
    Public Const cbuActionPreview = 5
    Public Enum enmAction
        ActNormal = 0
        ActNew = 1
        ActEdit = 2
        ActList = 3
        ActPreview = 3
    End Enum

    Public Structure stuSession
        Const Simpan = "SubmitSimpan"
    End Structure

    Public vbuDevelopmentStatus As String
    Public vbuErrorPath As String

    Public Enum EnmApprovalAction
        ManagerApprove = 1
        ManagerReject = 2
        ITManagerApprove = 4
        ITManagerReject = 8
        DoneByITStaff = 16
    End Enum

    Public pbMsgError As String

    Public Function fbuFormatString(vriData As String) As String
        fbuFormatString = Replace(vriData, "'", "''")
    End Function

    Public Function fbuFixString(vriData As String) As String
        fbuFixString = Replace(vriData, "'", "")
    End Function

    Public Function fbuValNumHtml(vriData As Object) As String
        If IsDBNull(vriData) Then
            fbuValNumHtml = "0"
        Else
            If Not IsNumeric(vriData) Then
                fbuValNumHtml = "0"
            Else
                fbuValNumHtml = vriData
            End If
        End If
    End Function

    Public Function fbuValStrHtml(vriData As Object) As String
        If IsDBNull(vriData) Then
            fbuValStrHtml = ""
        Else
            If Trim(vriData) = "&nbsp;" Then
                fbuValStrHtml = ""
            Else
                fbuValStrHtml = vriData
            End If
        End If
    End Function

    Public Function fbuFormatDateEng(vriDateShort As String) As String
        Dim vnReturn As String
        Dim vnDateEng As String
        vnDateEng = LCase(vriDateShort)
        vnDateEng = Replace(Replace(Replace(vnDateEng, "agustus", "August"), "oktober", "october"), "desember", "Desember")
        vnReturn = Replace(Replace(Replace(Replace(vnDateEng, "mei", "May"), "agu", "Aug"), "okt", "Oct"), "des", "Dec")
        Return vnReturn
    End Function

    Public Function fbuFormatDateDMY_To_YMD(vriDate As String) As String
        Dim vnReturn As String
        If Len(vriDate) <> 10 Then
            vnReturn = "01-01-1900"
        Else
            vnReturn = Mid(vriDate, 7, 4) & Mid(vriDate, 4, 2) & Mid(vriDate, 1, 2)
        End If
        Return vnReturn
    End Function

    Public Function fbuFormatDateDMY_To_YMD_Null(vriDate As String) As String
        Dim vnReturn As String
        If Len(vriDate) <> 10 Then
            vnReturn = "Null"
        Else
            vnReturn = "'" & Mid(vriDate, 7, 4) & Mid(vriDate, 4, 2) & Mid(vriDate, 1, 2) & "'"
        End If
        Return vnReturn
    End Function

    Public Sub pbuFillDstHour(vriDst As DropDownList)
        Dim vnDtb As New DataTable()
        vnDtb.Columns.Add("vData")

        For vn = 0 To 23
            vnDtb.Rows.Add(New Object() {Format(vn, "0#")})
        Next

        vriDst.DataSource = vnDtb
        vriDst.DataTextField = "vData"
        vriDst.DataValueField = "vData"
        vriDst.DataBind()
    End Sub

    Public Sub pbuFillDstMinute(vriDst As DropDownList)
        Dim vnDtb As New DataTable()
        vnDtb.Columns.Add("vData")
        For vn = 0 To 59
            vnDtb.Rows.Add(New Object() {Format(vn, "0#")})
        Next

        vriDst.DataSource = vnDtb
        vriDst.DataTextField = "vData"
        vriDst.DataValueField = "vData"
        vriDst.DataBind()
    End Sub
End Module

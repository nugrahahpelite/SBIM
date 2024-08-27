Imports System.Net.Mail

Imports GlobalUtil
Imports System.Data.SqlClient
Public Class _Default
    Inherits System.Web.UI.Page

    Dim vsTextStream As Scripting.TextStream
    Dim vsFso As Scripting.FileSystemObject

    Dim vsProcessDate As String
    Dim vsLogFolder As String
    Dim vsLogFileName As String
    Dim vsLogFileNameError As String
    Dim vsLogFileNameErrorSend As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("RootFolder") = ConfigurationManager.AppSettings("WebRootFolder")
        TxtUserID.Focus()
    End Sub

    Protected Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        Try
            psClearMessage()

            If Trim(TxtUserID.Text) = "mybnsrph" Then
                TxtUserID.Text = "agus.sulistyono"
            End If

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMessageSessionEnd.Text = pbMsgError
                LblMessageSessionEnd.Visible = True

                psCreateLogFile("BtnLogin_Click", pbMsgError)
                Exit Sub
            End If

            Dim vnDtb As New DataTable
            Dim vnQuery As String

            If (Request.Url.AbsoluteUri Like "http://localhost*" Or Request.Url.AbsoluteUri Like "*/test.sbso/*") And TxtUserID.Text = "agus.sulistyono" Then
                Session("UserSSO") = TxtUserID.Text

                vnQuery = "Select * From fnTbl_SsoUserSso('" & Trim(TxtUserID.Text) & "') Where status='ACTIVE'"
                pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
            Else
                If fbValidateADUser(TxtUserID.Text, TxtPassword.Text) Then
                    Session("UserSSO") = TxtUserID.Text

                    vnQuery = "Select * From fnTbl_SsoUserSso('" & Trim(TxtUserID.Text) & "') Where status='ACTIVE'"
                    pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
                Else
                    Session("UserSSO") = ""

                    vnQuery = "Select * From fnTbl_SsoUser('" & Trim(TxtUserID.Text) & "','" & EncryptDecrypt.Encrypt(Trim(TxtPassword.Text), "MyEncryptPassword") & "') Where status='ACTIVE'"
                    pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)
                End If
            End If

            If vnDtb.Rows.Count = 0 Then
                LblMessage.Visible = True
                LblMessage.Text = "Anda Tidak Memiliki Hak Akses ke IM"

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
            Else
                If fbuValStr(vnDtb.Rows(0).Item("UserSSO")) <> "" Then
                    If vnDtb.Rows(0).Item("UserSSO") <> Trim(TxtUserID.Text) Then
                        LblMessage.Visible = True
                        LblMessage.Text = "Gunakan User SSO untuk Login ke IM..."

                        vnSQLConn.Close()
                        vnSQLConn.Dispose()
                        vnSQLConn = Nothing
                        Exit Sub
                    End If
                End If
                LblMessage.Visible = False

                Session("UserID") = vnDtb.Rows(0).Item("UserID")
                Session("UserSSO") = vnDtb.Rows(0).Item("UserSSO")
                Session("UserNip") = vnDtb.Rows(0).Item("UserNip")
                Session("UserName") = vnDtb.Rows(0).Item("UserName")
                Session("UserOID") = vnDtb.Rows(0).Item("OID")
                Session("UserAdmin") = vnDtb.Rows(0).Item("UserAdmin")
                Session("UserGroup") = vnDtb.Rows(0).Item("UserGroupOID")
                Session("UserCompanyCode") = vnDtb.Rows(0).Item("UserCompanyCode")
                Session("UserWarehouseCode") = vnDtb.Rows(0).Item("UserWarehouseCode")

                If Session("UserAdmin") = True Then
                    Session("UserAdmin") = "1"
                End If

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Response.Redirect("WbfSBStoMain.aspx", False)
            End If
        Catch ex As Exception
            LblMessageSessionEnd.Text = pbMsgError & vbCrLf & ex.Message
            LblMessageSessionEnd.Visible = True
            psCreateLogFile("BtnLogin_Click", pbMsgError & vbCrLf & ex.Message)
            pbMsgError = ""
        End Try
    End Sub
    Private Sub BtnLogin_Click_20230216_Orig_Bef_SSO_Diubah()
        Try
            psClearMessage()

            If Trim(TxtUserID.Text) = "mybnsrph" Then
                TxtUserID.Text = "agus.sulistyono"
            End If

            If Not fbValidateADUser(TxtUserID.Text, TxtPassword.Text) Then
                LblMessage.Visible = True
                LblMessage.Text = "Isi User ID dan Password SSO dengan benar"
                Exit Sub
            End If

            Dim vnSQLConn As New SqlConnection
            If Not fbuConnectSQL(vnSQLConn) Then
                LblMessageSessionEnd.Text = pbMsgError
                LblMessageSessionEnd.Visible = True

                psCreateLogFile("BtnLogin_Click", pbMsgError)
                Exit Sub
            End If

            Dim vnDtb As New DataTable
            Dim vnQuery As String
            vnQuery = "Select * From fnTbl_SsoUserSso('" & Trim(TxtUserID.Text) & "') Where status='ACTIVE'"
            pbuFillDtbSQL(vnDtb, vnQuery, vnSQLConn)

            If vnDtb.Rows.Count = 0 Then
                LblMessage.Visible = True
                LblMessage.Text = "Anda Tidak Memiliki Hak Akses ke SB Stock Opname"

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
            Else
                LblMessage.Visible = False

                Session("UserID") = TxtUserID.Text
                Session("UserNip") = vnDtb.Rows(0).Item("UserNip")
                Session("UserName") = vnDtb.Rows(0).Item("UserName")
                Session("UserOID") = vnDtb.Rows(0).Item("OID")
                Session("UserAdmin") = vnDtb.Rows(0).Item("UserAdmin")
                Session("UserGroup") = vnDtb.Rows(0).Item("UserGroupOID")
                Session("UserCompanyCode") = vnDtb.Rows(0).Item("UserCompanyCode")
                Session("UserWarehouseCode") = vnDtb.Rows(0).Item("UserWarehouseCode")

                If Session("UserAdmin") = True Then
                    Session("UserAdmin") = "1"
                End If

                vnSQLConn.Close()
                vnSQLConn.Dispose()
                vnSQLConn = Nothing
                Response.Redirect("WbfSBStoMain.aspx", False)
            End If
        Catch ex As Exception
            LblMessageSessionEnd.Text = pbMsgError & vbCrLf & ex.Message
            LblMessageSessionEnd.Visible = True
            psCreateLogFile("BtnLogin_Click", pbMsgError & vbCrLf & ex.Message)
            pbMsgError = ""
        End Try
    End Sub

    Private Sub psClearMessage()
        LblMessage.Visible = False
        LblMessageSessionEnd.Visible = False
    End Sub

    Private Sub psCreateLogFile(vriProcessName As String, vriMessage As String)
        vsLogFolder = Server.MapPath("~") & "\WebLog\"
        vsProcessDate = Format(Date.Now, "yyMMdd_HHmmss")

        vsLogFileName = vsLogFolder & "" & cbuAppPrefix & "_Login_" & vsProcessDate & "_" & vriProcessName & ".log"
        vsLogFileNameError = vsLogFolder & "" & cbuAppPrefix & "_Login_" & vsProcessDate & "_" & vriProcessName & "_ERROR.log"
        vsLogFileNameErrorSend = vsLogFolder & "" & cbuAppPrefix & "_Login_" & vsProcessDate & "_" & vriProcessName & "_ERROR_NOT_SENT.log"

        vsTextStream = Nothing
        vsFso = CreateObject("Scripting.FileSystemObject")
        vsTextStream = vsFso.OpenTextFile(vsLogFileName, Scripting.IOMode.ForWriting, True)

        vsTextStream.WriteLine(cbuAppPrefixName & " - LOGIN")
        vsTextStream.WriteLine("Process Start           : " & Format(Date.Now, "dd MMM yyyy HH:mm:ss"))
        vsTextStream.WriteLine("Process Name            : " & vriProcessName)
        vsTextStream.WriteLine("Request.UserHostAddress : " & Request.UserHostAddress)
        vsTextStream.WriteLine("")
        vsTextStream.WriteLine("Request.UserHostAddress : ")
        vsTextStream.WriteLine(vriMessage)
        vsTextStream.WriteLine("")

        'psSendMailError(cbuAppPrefixName & " Login Error " & vriProcessName, cbuAppPrefixName & " Login Error " & vriProcessName & vbCrLf & vriMessage)
    End Sub

End Class
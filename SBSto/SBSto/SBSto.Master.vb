Public Class SBSto
    Inherits System.Web.UI.MasterPage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LblAppTitle.Text = "SB WMS " & vbuDevelopmentStatus
        Session("RootFolder") = ConfigurationManager.AppSettings("WebRootFolder")
    End Sub
    Private Sub LnkLogOut_Click(sender As Object, e As EventArgs) Handles LnkLogOut.Click
        Session.Abandon()
        Response.Redirect("~/Default.aspx", False)
    End Sub

End Class
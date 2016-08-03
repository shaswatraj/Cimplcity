Public Class Report
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataGridView1.DataSource = TagExtract.dt
        DataGridView1.DataBind()
    End Sub

   
End Class
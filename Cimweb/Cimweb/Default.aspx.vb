Imports Microsoft.VisualBasic
Imports GefObjectModel
Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
Imports System.Data
Public Class _Default
    Inherits System.Web.UI.Page
    Dim fpath, strfilename As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            fpath = "C:\C_Drive_Backup\Cimproj\screens"
            fpath = txtPath.Text
            If fpath = "" Then
                MsgBox("No path Has Been Selected")
                Exit Sub
            End If
            CheckBoxList1.Items.Clear()
            For Each file As String In IO.Directory.GetFiles(fpath, "*.CIM")
                strfilename = IO.Path.GetFileName(file)
                If IO.Path.GetExtension(file).ToUpper Like "*.CIM" And Not IO.Path.GetFileName(file).ToUpper Like "SMART*" And Not IO.Path.GetFileName(file).ToUpper Like "STATIC*" And Not IO.Path.GetFileName(file).ToUpper Like "*ALARM*" And Not IO.Path.GetFileName(file).ToUpper Like "BANNER*" And Not IO.Path.GetFileName(file).ToUpper Like "*SMART*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAL2BT*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAMOV*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TATRXM*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TATRND*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAPIDC*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAPERM*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAXDEV*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAVD3W*" And Not IO.Path.GetFileName(file).ToUpper Like "COMMON_TAVD2W*" Then
                    CheckBoxList1.Items.Add(strfilename)
                End If

            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Tagextract As New TagExtract
        fpath = txtPath.Text
        Tagextract.screen(CheckBoxList1, DataGridView1, chkAlias, GridView1, fpath)
        Dim meta As New HtmlMeta()
        meta.HttpEquiv = "Refresh"
        meta.Content = "5;url=Report.aspx"
        Me.Page.Controls.Add(meta)
        DataGridView1.DataSource = Tagextract.dt
        DataGridView1.DataBind()
    End Sub

    
    Protected Sub chkSelAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkSelAll.CheckedChanged
        If chkSelAll.Checked = True Then
            For i As Integer = 0 To CheckBoxList1.Items.Count - 1
                'CheckBoxList1.SetItemChecked(i, True)
                CheckBoxList1.Items(i).Selected = True
            Next
        Else
            For i As Integer = 0 To CheckBoxList1.Items.Count - 1
                'CheckBoxList1.SetItemChecked(i, False)
                CheckBoxList1.Items(i).Selected = False
            Next
        End If
    End Sub
End Class
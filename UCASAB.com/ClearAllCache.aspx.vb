
Partial Class ClearAllCache
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim myEventUtil As New EventUtil(True)  'clears the events cache
        BulletedList1.Items.Add(New ListItem("EventCollection cache has been cleared."))
        Dim myExecBoardUtil As New ExecMemberUtil(True) 'clears the exec member util cache
        BulletedList1.Items.Add(New ListItem("ExecMemberCollection cache has been cleared."))

        'clear outputCache on pages & userControls
        Dim paths As New List(Of String)
        With paths
            .Add("/Default.aspx")
            .Add("/Calendar.aspx")
            .Add("/ExecBoard.aspx")
            .Add("/PhotoAlbum.aspx")
            .Add("/rss.aspx")
            .Add("/scroller/scrollcontent.aspx")
            .Add("/UserControls/CommitteesMenuButtonsControl.ascx")
            .Add("/UserControls/ContactControl.ascx")
            .Add("/UserControls/FacebookControl.ascx")
            .Add("/UserControls/MenuControl.ascx")
            .Add("/UserControls/TwitterControl.ascx")
            .Add("/UserControls/W3ValidDisplayControl.ascx")
        End With
        For Each path In paths
            HttpResponse.RemoveOutputCacheItem(path)
            BulletedList1.Items.Add(New ListItem(String.Format("OutputCache for {0} has been cleared.", path)))
        Next

    End Sub
End Class

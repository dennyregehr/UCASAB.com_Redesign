
Partial Class ExecBoard
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Dim execUtil As New ExecMemberUtil
        'Dim members = execUtil.GetBoardMembersForExecBoardPage()

        'For Each member As ExecBoardMembers In members
        '    ph1.Controls.Add(
        '        New LiteralControl(
        '            String.Format("<div class=""ExecBoardMember"">{0}<br/><img src='{1}' /><br />{2}<br /></div>",
        '            member.ExecMember, member.PhotoURL, member.PositionTitle)
        '        )
        '    )
        '    'ph1.Controls.Add(New LiteralControl(member.ToString() & "<br />"))
        'Next

    End Sub
End Class

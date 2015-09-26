Imports System.Data
Imports System.Web.Caching

Public Class ExecMemberUtil

    Dim _execMemberCollectionName As String = "ExecMemberCollection"

    Public Sub New()
    End Sub

    Public Sub New(ByVal ClearCache As Boolean)
        DisposeEventCache()
    End Sub

    Private Sub DisposeEventCache()
        HttpContext.Current.Cache.Remove(_execMemberCollectionName)
    End Sub

    Public Function GetExecMembers() As List(Of ExecBoardMembers)
        Dim execMbrs As New List(Of ExecBoardMembers)
        If HttpContext.Current.Cache(_execMemberCollectionName) Is Nothing Then
            Dim mySqlUtil As New SqlUtil
            Dim ds As DataSet = mySqlUtil.GetExecutiveMembers
            If ds.Tables(0) IsNot Nothing Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    execMbrs.Add(New ExecBoardMembers(dr))
                Next
            End If
            HttpContext.Current.Cache.Insert(_execMemberCollectionName, execMbrs, Nothing, Semester.EndDate, Cache.NoSlidingExpiration, CacheItemPriority.Default, Nothing)
        Else
            execMbrs = CType(HttpContext.Current.Cache(_execMemberCollectionName), List(Of ExecBoardMembers))
        End If
        Return execMbrs
    End Function

    Public Function GetBoardMembersForExecBoardPage() As List(Of ExecBoardMembers)

        'Dim mbrs As New List(Of ExecBoardMembers)
        'Dim memberPositionOrder As String = "2,3,4,9,1,8,7,5,6"
        'For ord As Integer = 0 To memberPositionOrder.Split(",").GetUpperBound(0)
        '    Dim ordNew As Integer = ord
        '    mbrs.Add(Me.GetExecMembers.Find(Function(m) m.Position = memberPositionOrder.Split(",")(ordNew)))
        'Next
        'Return mbrs

        Dim mb = From m In Me.GetExecMembers _
             Order By m.OrderForExecBoardPage _
             Select m
        'Return New List(Of ExecBoardMembers)(mb.ToList())
        Return CType(mb.ToList(), List(Of ExecBoardMembers))

    End Function

    Public Function GetBoardMemberByCommittee(ByVal CommitteeName As String) As ExecBoardMembers
        Return Me.GetExecMembers.Find(Function(m) m.PositionTitle.ToLower = CommitteeName.ToLower)
    End Function

    Public Sub SetImageSwappingAttributes(ByVal MemberImage As Image)
        With MemberImage
            .Attributes.Add("onmouseover", String.Format("javascript: this.src='{0}';", .ImageUrl.Replace("A.jpg", "B.jpg")))
            .Attributes.Add("onmouseout", String.Format("javascript: this.src='{0}';", .ImageUrl))
        End With
    End Sub

End Class

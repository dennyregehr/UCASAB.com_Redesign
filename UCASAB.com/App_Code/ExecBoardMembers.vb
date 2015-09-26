Imports System.Data

<Serializable()> _
Public Class ExecBoardMembers

    Private _position As String
    Public ReadOnly Property Position() As String
        Get
            Return _position
        End Get
    End Property
    Private _execMember As String
    Public ReadOnly Property ExecMember() As String
        Get
            Return _execMember
        End Get
    End Property
    Private _fname As String
    Public ReadOnly Property FirstName() As String
        Get
            Return _fname
        End Get
    End Property
    Private _lname As String
    Public ReadOnly Property LastName() As String
        Get
            Return _lname
        End Get
    End Property
    Private _email As String
    Public ReadOnly Property Email() As String
        Get
            Return _email
        End Get
    End Property
    Private _phone1 As String
    Public ReadOnly Property Phone1() As String
        Get
            Return _phone1
        End Get
    End Property
    Private _phone2 As String
    Public ReadOnly Property Phone2() As String
        Get
            Return _phone2
        End Get
    End Property
    Private _photoURL As String
    Public ReadOnly Property PhotoURL() As String
        Get
            Return _photoURL
        End Get
    End Property

    Private _serviceDates As String
    Public ReadOnly Property ServiceDates() As String
        Get
            Return _serviceDates
        End Get
    End Property
    Private _positionTitle As String
    Public ReadOnly Property PositionTitle() As String
        Get
            Return _positionTitle
        End Get
    End Property
    Private _committeeDescription As String
    Public ReadOnly Property CommitteeDescription() As String
        Get
            Return _committeeDescription
        End Get
    End Property
    Private _orderForExecBoardPage As Integer
    Public ReadOnly Property OrderForExecBoardPage() As Integer
        Get
            Return _orderForExecBoardPage
        End Get
    End Property

    Public Sub New(ByVal ExecBoardMemberDataRow As DataRow)
        With Me
            ._position = SafeData(ExecBoardMemberDataRow("position"))
            ._execMember = SafeData(ExecBoardMemberDataRow("execmember"))
            ._fname = SafeData(ExecBoardMemberDataRow("fname"))
            ._lname = SafeData(ExecBoardMemberDataRow("lname"))
            ._email = SafeData(ExecBoardMemberDataRow("email"))
            ._phone1 = SafeData(ExecBoardMemberDataRow("phone1"))
            ._phone2 = SafeData(ExecBoardMemberDataRow("phone2"))
            ._photoURL = SafeData(ExecBoardMemberDataRow("photourl"))
            ._serviceDates = SafeData(ExecBoardMemberDataRow("servicedates"))
            ._positionTitle = SafeData(ExecBoardMemberDataRow("positionTitle"))
            ._committeeDescription = SafeData(ExecBoardMemberDataRow("committeeDescription"))
            ._orderForExecBoardPage = ExecBoardMemberDataRow("orderForExecBoardPage")
        End With
    End Sub

    Private Function SafeData(ByVal DataItem As Object) As String
        Return IIf(IsDBNull(DataItem), String.Empty, DataItem)
    End Function

    Public Shadows ReadOnly Property ToString()
        Get
            Return String.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}",
                Me.CommitteeDescription, Me.Email, Me.ExecMember, Me.FirstName, Me.LastName,
                Me.OrderForExecBoardPage, Me.Phone1, Me.Phone2, Me.PhotoURL, Me.Position,
                Me.PositionTitle, Me.ServiceDates)
        End Get
    End Property

End Class

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
    Public ReadOnly Property PhotoUrl() As String
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

    Public Sub New(ByVal execBoardMemberDataRow As DataRow)
        With Me
            ._position = SafeData(execBoardMemberDataRow("position"))
            ._execMember = SafeData(execBoardMemberDataRow("execmember"))
            ._fname = SafeData(execBoardMemberDataRow("fname"))
            ._lname = SafeData(execBoardMemberDataRow("lname"))
            ._email = SafeData(execBoardMemberDataRow("email"))
            ._phone1 = SafeData(execBoardMemberDataRow("phone1"))
            ._phone2 = SafeData(execBoardMemberDataRow("phone2"))
            ._photoURL = SafeData(execBoardMemberDataRow("photourl"))
            ._serviceDates = SafeData(execBoardMemberDataRow("servicedates"))
            ._positionTitle = SafeData(execBoardMemberDataRow("positionTitle"))
            ._committeeDescription = SafeData(execBoardMemberDataRow("committeeDescription"))
            ._orderForExecBoardPage = CType(execBoardMemberDataRow("orderForExecBoardPage"), Integer)
        End With
    End Sub

    Private Function SafeData(ByVal dataItem As Object) As String
        Return CType(IIf(IsDBNull(dataItem), String.Empty, dataItem), String)
    End Function

    Public Shadows ReadOnly Property ToString() As String
        Get
            Return String.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}",
                                 CommitteeDescription,
                                 Email,
                                 ExecMember,
                                 FirstName,
                                 LastName,
                                 OrderForExecBoardPage,
                                 Phone1,
                                 Phone2,
                                 PhotoUrl,
                                 Position,
                                 PositionTitle,
                                 ServiceDates)
        End Get
    End Property

End Class

Imports System.Configuration

Public Class Semester

    Private Shared ReadOnly Property FirstDateToDisplayFallEvents As Date
        Get
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FirstDateToDisplayFallEvents")) Then
                Dim dateParts As String() = ConfigurationManager.AppSettings("FirstDateToDisplayFallEvents").Split("/".ToCharArray, StringSplitOptions.None)
                Return New Date(Today.Year, dateParts(0), dateParts(1))
            Else
                Return New Date(Today.Year, 8, 1)
            End If
        End Get
    End Property
    Private Shared ReadOnly Property FirstDateToDisplaySpringEvents As Date
        Get
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FirstDateToDisplaySpringEvents")) Then
                Dim dateParts As String() = ConfigurationManager.AppSettings("FirstDateToDisplaySpringEvents").Split("/".ToCharArray, StringSplitOptions.None)
                Return New Date(Today.Year, dateParts(0), dateParts(1))
            Else
                Return New Date(Today.Year, 1, 1)
            End If
        End Get
    End Property
    Public Shared Function Description() As String
        If Today < FirstDateToDisplayFallEvents Then
            Return "Spring"
        Else
            Return "Fall"
        End If
    End Function
    Public Shared Function StartDate() As Date
        Return IIf(Today < FirstDateToDisplayFallEvents, FirstDateToDisplaySpringEvents, FirstDateToDisplayFallEvents)
    End Function
    Public Shared Function EndDate() As Date
        Return IIf(Today < FirstDateToDisplayFallEvents, FirstDateToDisplayFallEvents.AddDays(-1), FirstDateToDisplaySpringEvents.AddDays(-1).AddYears(1))
    End Function
    Public Shared ReadOnly Property SpringMonths() As Integer()
        Get
            Return New Integer() {1, 2, 3, 4, 5, 6, 7, 8}
        End Get
    End Property
    Public Shared ReadOnly Property FallMonths() As Integer()
        Get
            Return New Integer() {8, 9, 10, 11, 12}
        End Get
    End Property
    Public Shared ReadOnly Property ThisSemestersMonths() As Integer()
        Get
            Return IIf(Today < FirstDateToDisplayFallEvents, SpringMonths, FallMonths)
        End Get
    End Property

End Class

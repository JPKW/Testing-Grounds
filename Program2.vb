Imports System

Module Program2

    Function dts(dtStart As Date, dtEnd As Date) As DateTimeSpan

        Return DateTimeSpan.DateSpan(dtEnd, dtStart)

    End Function


End Module

Public Structure DateTimeSpan

    Private _Years As Integer
    Public ReadOnly Property Years As Integer
        Get
            Return _Years
        End Get
    End Property

    Private _Months As Integer
    Public ReadOnly Property Months As Integer
        Get
            Return _Months
        End Get
    End Property

    Private _Days As Integer
    Public ReadOnly Property Days As Integer
        Get
            Return _Days
        End Get
    End Property

    Private _Hours As Integer
    Public ReadOnly Property Hours As Integer
        Get
            Return _Hours
        End Get
    End Property

    Private _Minutes As Integer
    Public ReadOnly Property Minutes As Integer
        Get
            Return _Minutes
        End Get
    End Property

    Private _Seconds As Integer
    Public ReadOnly Property Seconds As Integer
        Get
            Return _Seconds
        End Get
    End Property

    Private _MilliSeconds As Integer
    Public ReadOnly Property MilliSeconds As Integer
        Get
            Return _MilliSeconds
        End Get
    End Property

    ' the ctor for the result
    Private Sub New(y As Integer, mm As Integer, d As Integer,
                    h As Integer, m As Integer, s As Integer,
                    ms As Integer)
        _Years = y
        _Months = mm
        _Days = d
        _Hours = h
        _Minutes = m
        _Seconds = Seconds
        _MilliSeconds = ms
    End Sub

    ' private time unit tracker when counting
    Private Enum Unit
        Year
        Month
        Day
        Complete
    End Enum

    Public Shared Function DateSpan(dt1 As DateTime,
                                    dt2 As DateTime) As DateTimeSpan
        ' we dont do negatives
        If dt2 < dt1 Then
            Dim tmp = dt1
            dt1 = dt2
            dt2 = tmp
        End If

        Dim thisDT As DateTime = dt1
        Dim years As Integer = 0
        Dim months As Integer = 0
        Dim days As Integer = 0

        Dim level As Unit = Unit.Year
        Dim span As New DateTimeSpan()

        While level <> Unit.Complete
            Select Case level
                ' add a year until it is larger;
                ' then change the "level" to month
                Case Unit.Year
                    If thisDT.AddYears(years + 1) > dt2 Then
                        level = Unit.Month
                        thisDT = thisDT.AddYears(years)
                    Else
                        years += 1
                    End If

                Case Unit.Month
                    If thisDT.AddMonths(months + 1) > dt2 Then
                        level = Unit.Day
                        thisDT = thisDT.AddMonths(months)
                    Else
                        months += 1
                    End If

                Case Unit.Day
                    If thisDT.AddDays(days + 1) > dt2 Then
                        thisDT = thisDT.AddDays(days)
                        Dim thisTS = dt2 - thisDT
                        ' create a new DTS from the values caluclated
                        span = New DateTimeSpan(years, months, days, thisTS.Hours,
                                                thisTS.Minutes, thisTS.Seconds,
                                    thisTS.Milliseconds)
                        level = Unit.Complete
                    Else
                        days += 1
                    End If
            End Select
        End While

        Return span

    End Function

End Structure




Imports System

Module Program
    Sub Main(args As String())
        'Console.WriteLine(strYMDTwoDates("31/07/2019", "15/07/2020"))

        Dim dts As DateTimeSpan = DateTimeSpan.DateSpan("#30/07/2019#", "#15/07/2020#")

        Console.WriteLine("{0} Yrs, {1} Months and {2} Days",
                           dts.Years.ToString, dts.Months.ToString, dts.Days.ToString)

    End Sub

    Function ListOfListTesting() As String

        'Dim rows As New List(Of List(Of String))
        Dim rows As New System.Collections.ArrayList

        'Return Interest Table function (NB: this is the body of the table)
        Dim al As New List(Of String)
        Dim bl As New List(Of Integer)
        Dim cl As New List(Of String)
        Dim dl As New List(Of String)

        al.AddRange({"aa", "ab", "ac"})
        bl.AddRange({1, 2, 3})
        cl.AddRange({"ca", "cb", "cc"})
        dl.AddRange({"da", "db", "dc"})

        rows.Insert(0, {al, bl, cl, dl})

        Dim rowslist As New List(Of Object)

        rowslist.AddRange(rows(0))

        Dim finalrow As New List(Of String)

        finalrow.AddRange(rowslist(2))

        Dim DailyAmount As String
        DailyAmount = finalrow(2)

        Return DailyAmount

    End Function

    Function strYMDTwoDates(D1 As Date, D2 As Date) As String

        Dim dDay As Date

        Dim sYears As String
        Dim sDays As String
        Dim sMonths As String

        sYears = DateDiff(DateInterval.Year, D1, D2)

        dDay = D1.AddYears(sYears)

        sMonths = DateDiff(DateInterval.Month, dDay, D2)

        sDays = DateDiff(DateInterval.Day, dDay, D2)

        Return sYears & " Year " & sMonths & " Months " & sDays & " Day(s)"

    End Function


End Module
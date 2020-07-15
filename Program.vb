Imports System

Module Program
    Sub Main(args As String())
        Console.WriteLine(test)
    End Sub

    Function test() As String

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



End Module
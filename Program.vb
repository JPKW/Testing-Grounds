Imports System

Module Program
    Sub Main(args As String())
        Console.WriteLine(test)
    End Sub

    Function test() As String

        Dim rows = New List(Of List(Of String))
        'Dim rows = New System.Collections.ArrayList

        'Return Interest Table function (NB: this is the body of the table)

        Dim al As New List(Of String)
        Dim bl As New List(Of String)
        Dim cl As New List(Of String)
        Dim dl As New List(Of String)

        al.AddRange({"aa", "ab", "ac"})
        bl.AddRange({"ba", "bb", "bc"})
        cl.AddRange({"ca", "cb", "cc"})
        dl.AddRange({"da", "db", "dc"})

        rows.AddRange({al, bl, cl, dl})

        Dim LastRow As List(Of String)
        LastRow = rows(rows.Count - 1)

        Dim DailyAmount As String
        DailyAmount = LastRow(2)

        Return DailyAmount



    End Function



End Module
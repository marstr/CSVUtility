Imports System.IO
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Utilities.CSV

<TestClass()>
Public Class HeadedCSVWriterTests

    <TestMethod()>
    Public Sub WriteAlignedTuples()
        Dim expectedHead = {"A", "B", "C"}
        Dim expectedRows As New List(Of String())
        expectedRows.Add({"1", "2", "3"})
        expectedRows.Add({"4", "5", "6"})

        Using memStream = New MemoryStream
            Dim writer = New HeadedCSVWriter(memStream, expectedHead)
            For Each row In expectedRows
                writer.WriteTuple(row)
            Next
            writer.Flush()

            memStream.Seek(0L, SeekOrigin.Begin)

            Dim reader = New CSVReader(memStream)
            Dim head = reader.ReadTuple()
            For i = 0 To expectedHead.Length - 1
                Assert.AreEqual(expectedHead(i), head(i))
            Next

            For Each expectedRow In expectedRows
                Dim row = reader.ReadTuple()
                For i = 0 To expectedHead.Length - 1
                    Assert.AreEqual(expectedRow(i), row(i))
                Next
            Next
        End Using
    End Sub

    <TestMethod()>
    <ExpectedException(GetType(ArgumentException), AllowDerivedTypes:=True)>
    Public Sub WriteIllegalyLongTuples()
        Dim expectedHead = {"A", "B", "C"}
        Dim expectedRows As New List(Of String())
        expectedRows.Add({"1", "2", "3", "4"})

        Using memStream = New MemoryStream
            Dim writer = New HeadedCSVWriter(memStream, expectedHead)
            For Each row In expectedRows
                writer.WriteTuple(row)
            Next
        End Using
    End Sub
End Class
Imports System.IO
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Utilities.CSV

<TestClass()>
Public Class HeadedCSVReaderTests

    <TestMethod()>
    Public Sub SimpleStringsColumnMatches()
        Using memStream As New MemoryStream()
            Dim writer = New StreamWriter(memStream)
            writer.Write("a,b,c
1,2,3")
            Dim expected = New Dictionary(Of String, String)
            expected.Add("a", "1")
            expected.Add("b", "2")
            expected.Add("c", "3")

            writer.Flush()
            memStream.Seek(0L, SeekOrigin.Begin)
            Dim reader = New HeadedCSVReader(memStream)
            Assert.AreEqual(3, reader.Header.Length)

            Dim tuple = New Dictionary(Of String, String)
            reader.ReadTuple(tuple)

            Assert.AreEqual(2UI, reader.Position.Row)

            For Each column In expected.Keys
                Assert.IsTrue(reader.Header.Contains(column))
                Assert.AreEqual(expected(column), tuple(column))
            Next

            Assert.IsFalse(reader.ReadTuple(tuple))
        End Using
    End Sub

    <TestMethod()>
    Public Sub ReadBeforeHeaderCheck()
        Using memStream As New MemoryStream()
            Dim writer = New StreamWriter(memStream)
            writer.Write("a,b,c
1,2,3")
            Dim expected = New Dictionary(Of String, String)
            expected.Add("a", "1")
            expected.Add("b", "2")
            expected.Add("c", "3")

            writer.Flush()
            memStream.Seek(0L, SeekOrigin.Begin)
            Dim reader = New HeadedCSVReader(memStream)

            Dim tuple = New Dictionary(Of String, String)
            reader.ReadTuple(tuple)

            Assert.AreEqual(2UI, reader.Position.Row)

            For Each column In expected.Keys
                Assert.IsTrue(reader.Header.Contains(column))
                Assert.AreEqual(expected(column), tuple(column))
            Next

            Assert.IsFalse(reader.ReadTuple(tuple))

            Assert.AreEqual(3, reader.Header.Length)
        End Using
    End Sub

    <TestMethod()>
    Public Sub NoDataAfterHeader()
        Using memStream As New MemoryStream()
            Dim writer = New StreamWriter(memStream)
            writer.Write("A,B,C")
            writer.Flush()
            memStream.Seek(0L, SeekOrigin.Begin)

            Dim reader = New HeadedCSVReader(memStream)
            Assert.AreEqual(3, reader.Header.Length)
            Assert.IsFalse(reader.ReadTuple(Nothing))
        End Using
    End Sub

    <TestMethod()>
    <ExpectedException(GetType(InvalidOperationException), AllowDerivedTypes:=True)>
    Public Sub ReadHeaderAfterProcessing()
        Dim desiredContent = New List(Of String())
        desiredContent.Add({"A", "B", "C"})
        desiredContent.Add({"1", "2", "3"})
        desiredContent.Add({"4", "5", "6"})

        Using memStream As New MemoryStream()
            Dim writer = New CSVWriter(memStream)

            For Each line In desiredContent
                writer.WriteTuple(line)
            Next

            writer.Flush()
            memStream.Seek(0L, SeekOrigin.Begin)

            Dim reader = New HeadedCSVReader(memStream)
            Dim trueHeader = reader.ReadTuple()
            Assert.AreEqual(3, reader.Header.Length)
        End Using
    End Sub

    <TestMethod()>
    Public Sub ReadShortLine()
        Using memStream As New MemoryStream()
            Dim writer = New CSVWriter(memStream)
            writer.WriteTuple({"A", "B", "C"})
            writer.WriteTuple({"1", "2"})
            writer.WriteTuple({"4", "5", "6"})
            writer.Flush()
            memStream.Seek(0L, SeekOrigin.Begin)

            Dim reader = New HeadedCSVReader(memStream)
            Assert.AreEqual(3, reader.Header.Length)
            Dim tuple As IDictionary(Of String, String)

#Disable Warning BC42030
            Assert.IsTrue(reader.ReadTuple(tuple))
#Enable Warning BC42030
        End Using
    End Sub

    <TestMethod()>
    Public Sub ReadLongLine()
        Using memStream As New MemoryStream()
            Dim writer = New CSVWriter(memStream)
            writer.WriteTuple({"A", "B", "C"})
            writer.WriteTuple({"1", "2", "3"})
            writer.WriteTuple({"4", "5", "6", "7"})
            writer.WriteTuple({"8", "9", "10"})
            writer.Flush()
            memStream.Seek(0L, SeekOrigin.Begin)

            Dim reader = New HeadedCSVReader(memStream)
            Assert.AreEqual(3, reader.Header.Length)
            Dim tuple As IDictionary(Of String, String)

#Disable Warning BC42030
            Assert.IsTrue(reader.ReadTuple(tuple))
            Assert.IsFalse(reader.ReadTuple(tuple))
            Assert.IsTrue(reader.ReadTuple(tuple))
#Enable Warning BC42030
        End Using
    End Sub
End Class

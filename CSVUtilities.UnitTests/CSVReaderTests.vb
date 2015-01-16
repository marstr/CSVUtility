Imports System.IO
Imports Utilities.CSV

<TestClass()>
Public Class CSVReaderTests

#If DEBUG Then
    Private Const DEFAULT_TIMEOUT = TestTimeout.Infinite
#Else
    Private Const DEFAULT_TIMEOUT = 2000
#End If

    Private Const ALTERNATE_DELIMITER = ";"c

    <TestInitialize>
    Public Sub Setup()
        Assert.AreNotEqual(ALTERNATE_DELIMITER, DEFAULT_DELIMITER, "For the sake of these tests, the alternate delimiter cannot be equal to the default delimiter. Please change the Alternate.")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub BasicMultiLineCellByCellRead()
        Using input As New CSVReader(GetFile())
            Assert.AreEqual("a", input.ReadCell())
            Assert.AreEqual("b", input.ReadCell())
            Assert.AreEqual("c", input.ReadCell())
            Assert.AreEqual("d", input.ReadCell())
            Assert.AreEqual("e", input.ReadCell())
            Assert.AreEqual("f", input.ReadCell())
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub BasicSingleLineCellByCellRead()
        Dim subject = String.Format("a{0}b{0}c", DEFAULT_DELIMITER)
        Using input As New CSVReader(GetFile(subject))
            Assert.AreEqual("a", input.ReadCell())
            Assert.AreEqual("b", input.ReadCell())
            Assert.AreEqual("c", input.ReadCell())
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub SplitLineCellByCellRead()
        Dim source = String.Format("a{0}""b{1}c""{0}d", DEFAULT_DELIMITER, vbNewLine) 'a,"b{new line}c",d
        Using reader As New CSVReader(GetFile(source))
            Assert.AreEqual("a", reader.ReadCell)
            Assert.AreEqual("b" & vbNewLine & "c", reader.ReadCell)
            Assert.AreEqual("d", reader.ReadCell)
            Assert.IsTrue(reader.EndOfStream)
        End Using
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EmptyFirstEntryCellByCellRead()
        Dim subject = String.Format("{0}b{0}c{0}d", DEFAULT_DELIMITER) ',b,c,d
        Using input As New CSVReader(GetFile(subject))
            Assert.AreEqual("", input.ReadCell())
            Assert.AreEqual("b", input.ReadCell())
            Assert.AreEqual("c", input.ReadCell())
            Assert.AreEqual("d", input.ReadCell())
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EmptyLastEntryCellByCellRead()
        Dim subject = String.Format("a{0}b{0}c{0}", DEFAULT_DELIMITER) 'a,b,c,
        Using input As New CSVReader(GetFile(subject))
            Assert.AreEqual("a", input.ReadCell())
            Assert.AreEqual("b", input.ReadCell())
            Assert.AreEqual("c", input.ReadCell())
            Assert.AreEqual("", input.ReadCell())
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EmptyCellsCellByCellRead()
        Dim subject = String.Format("a{0}{0}{0}{0}e", DEFAULT_DELIMITER) 'a,,,,e
        Using input As New CSVReader(GetFile(subject))
            Assert.AreEqual("a", input.ReadCell())
            Assert.AreEqual("", input.ReadCell())
            Assert.AreEqual("", input.ReadCell())
            Assert.AreEqual("", input.ReadCell())
            Assert.AreEqual("e", input.ReadCell())
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EmptyCellsTupleRead()
        Dim subject = String.Format("a{0}{0}{0}{0}e", DEFAULT_DELIMITER) 'a,,,,e
        Using input As New CSVReader(GetFile(subject))
            Dim tuple = input.ReadTuple
            Assert.AreEqual(5, tuple.Length)
            Assert.AreEqual("a", tuple(0))
            Assert.AreEqual("", tuple(1))
            Assert.AreEqual("", tuple(2))
            Assert.AreEqual("", tuple(3))
            Assert.AreEqual("e", tuple(4))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EmptyLastEntryTupleRead()
        Dim subject = String.Format("a{0}b{0}c{0}d{0}", DEFAULT_DELIMITER) 'a,b,c,d,
        Using input As New CSVReader(GetFile(subject))
            Dim tuple = input.ReadTuple
            Assert.AreEqual(5, tuple.Length)
            Assert.AreEqual("a", tuple(0))
            Assert.AreEqual("b", tuple(1))
            Assert.AreEqual("c", tuple(2))
            Assert.AreEqual("d", tuple(3))
            Assert.AreEqual("", tuple(4))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub BasicMultiLineTupleRead()
        Using input As New CSVReader(GetFile())
            Dim firstTuple = input.ReadTuple
            Dim secondTuple = input.ReadTuple
            Assert.AreEqual(3, firstTuple.Length)
            Assert.AreEqual("a", firstTuple(0))
            Assert.AreEqual("b", firstTuple(1))
            Assert.AreEqual("c", firstTuple(2))
            Assert.AreEqual(3, secondTuple.Length)
            Assert.AreEqual("d", secondTuple(0))
            Assert.AreEqual("e", secondTuple(1))
            Assert.AreEqual("f", secondTuple(2))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub SplitLineCellTupleRead()
        Dim subject = String.Format("a{0}""b{1}c""{0}d", DEFAULT_DELIMITER, vbNewLine)
        Using input As New CSVReader(GetFile(subject))
            Dim firstTuple = input.ReadTuple
            Assert.AreEqual(3, firstTuple.Length)
            Assert.AreEqual("a", firstTuple(0))
            Assert.AreEqual("b" & vbNewLine & "c", firstTuple(1))
            Assert.AreEqual("d", firstTuple(2))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub CustomDelimiterBasicTupleRead()
        Dim subject = String.Format("a{0}b{0}c", ALTERNATE_DELIMITER) ' a;b;c
        Using input As New CSVReader(stream:=GetFile(subject), delimiter:=ALTERNATE_DELIMITER)
            Dim tuple = input.ReadTuple
            Assert.AreEqual(3, tuple.Length)
            Assert.AreEqual("a", tuple(0))
            Assert.AreEqual("b", tuple(1))
            Assert.AreEqual("c", tuple(2))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub CustomDelimiterEscapedValueTupleRead()
        Dim subject = String.Format("a{0}""b{0}c""{0}d", ALTERNATE_DELIMITER) ' a;"b;c";d
        Using input As New CSVReader(GetFile(subject), delimiter:=ALTERNATE_DELIMITER)
            Dim tuple = input.ReadTuple
            Assert.AreEqual(3, tuple.Length)
            Assert.AreEqual("a", tuple(0))
            Assert.AreEqual("b" & ALTERNATE_DELIMITER & "c", tuple(1))
            Assert.AreEqual("d", tuple(2))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub AllTuple_ShowsAllTuples()
        Dim source = String.Format("a{0}b{0}c{1}d{0}e{0}f{1}g{0}h{0}i", ALTERNATE_DELIMITER, vbNewLine)

        Dim expected(3)() As String

        expected(0) = New String() {"a", "b", "c"}
        expected(1) = New String() {"d", "e", "f"}
        expected(2) = New String() {"g", "h", "i"}

        Using input As New CSVReader(GetFile(source), delimiter:=ALTERNATE_DELIMITER)
            Dim count = 0
            Assert.IsFalse(input.EndOfStream)
            For Each tuple In input.AllTuples
                Assert.AreEqual(3, tuple.Length)
                Assert.AreEqual(expected(count)(0), tuple(0))
                Assert.AreEqual(expected(count)(1), tuple(1))
                Assert.AreEqual(expected(count)(2), tuple(2))
                count += 1
            Next
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod>
    Public Sub AllCell_ShowsAllCells()
        Dim source = String.Format("a{0}b{0}c{1}d{0}e{0}f{1}g{0}h{0}i", ALTERNATE_DELIMITER, vbNewLine)

        Dim expected = New String() {"a", "b", "c", "d", "e", "f", "g", "h", "i"}

        Using input As New CSVReader(GetFile(source), delimiter:=ALTERNATE_DELIMITER)
            Dim count = 0
            Assert.IsFalse(input.EndOfStream)
            For Each cell In input.AllCells
                Assert.AreEqual(expected(count), cell)
                count += 1
            Next
            Assert.IsTrue(input.EndOfStream)
        End Using

    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_UnexpectedQuotes1()
        Dim subject = "a""b"
        Using mem = New MemoryStream()
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Write(DEFAULT_DELIMITER)
            writer.Write("abc")
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            Dim result = reader.ReadCell()
        End Using
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_UnexpectedQuotes2()
        Dim subject = "a""""b"
        Using mem = New MemoryStream()
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Write(DEFAULT_DELIMITER)
            writer.Write("abc")
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            Dim result = reader.ReadCell()
        End Using
    End Sub


    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_UnexpectedQuotes3()
        Dim subject = """"
        Using mem = New MemoryStream()
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Write(DEFAULT_DELIMITER)
            writer.Write("abc")
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            reader.ReadCell()
        End Using
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_UnexectedQuotes4()
        Dim reader = New CSVReader(GetFile(""""))
        reader.ReadCell()
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_ExtraCharacters1()
        Dim subject = String.Format("""a{0}b""c{0}d", DEFAULT_DELIMITER) ' "a,b"c,d
        Using mem = New MemoryStream()
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            reader.ReadCell()
        End Using
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_ExtraCharacters2()
        Dim subject = String.Format("a{0}b""""c{0}d", DEFAULT_DELIMITER) ' a,b""c,d
        Using mem = New MemoryStream
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            reader.ReadCell()
            reader.ReadCell()
        End Using
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_UnEscapedQuoteInEscapedCell()
        Dim subject = String.Format("a{0}""bc""d""{0}e", DEFAULT_DELIMITER) ' a,"bc"d",e
        Using mem = New MemoryStream
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            reader.ReadCell()
            reader.ReadCell()
        End Using
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    <ExpectedException(GetType(InvalidDataException))>
    Public Sub ReadCell_TwoUnescapedQuotesInEscapedCell()
        Dim subject = String.Format("a{0}""bc""d""e""{0}e", DEFAULT_DELIMITER) ' a,"bc"d"e"
        Using mem = New MemoryStream
            Dim writer = New StreamWriter(mem)
            writer.Write(subject)
            writer.Flush()
            mem.Seek(0, SeekOrigin.Begin)
            Dim reader = New CSVReader(mem)
            reader.ReadCell()
            reader.ReadCell()
        End Using
    End Sub

    Private Shared ReadOnly Default_GetFile_Contents = String.Format("a{0}b{0}c{1}d{0}e{0}f", DEFAULT_DELIMITER, vbNewLine)
    Private Shared Function GetFile() As Stream
        Return GetFile(Default_GetFile_Contents)
    End Function

    Private Shared Function GetFile(contents As String) As Stream
        Dim mem = New MemoryStream()
        Dim writer = New StreamWriter(mem)
        writer.Write(contents)
        writer.Flush()
        mem.Seek(0, SeekOrigin.Begin)
        Return mem
    End Function

End Class

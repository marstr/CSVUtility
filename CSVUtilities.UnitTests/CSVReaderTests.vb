Imports System.IO
Imports Utilities.CSV

<TestClass()>
Public Class CSVReaderTests

    Private Const DEFAULT_TIMEOUT = 2000

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub BasicMultiLineCellByCellRead()
        Using input As New CSVReader(path:=GetFile())
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
        Using input As New CSVReader(path:=GetFile(<source>a,b,c</source>.Value))
            Assert.AreEqual("a", input.ReadCell())
            Assert.AreEqual("b", input.ReadCell())
            Assert.AreEqual("c", input.ReadCell())
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub SplitLineCellByCellRead()
        Dim source = <source>a,"b<%= vbNewLine %>c",d</source>.Value
        Using reader As New CSVReader(path:=GetFile(source))
            Assert.AreEqual("a", reader.ReadCell)
            Assert.AreEqual("b" & vbNewLine & "c", reader.ReadCell)
            Assert.AreEqual("d", reader.ReadCell)
            Assert.IsTrue(reader.EndOfStream)
        End Using
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EmptyFirstEntryCellByCellRead()
        Using input As New CSVReader(path:=GetFile(<source>,b,c,d</source>.Value))
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
        Using input As New CSVReader(path:=GetFile(<source>a,b,c,</source>.Value))
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
        Using input As New CSVReader(path:=GetFile(<source>a,,,,e</source>.Value))
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
        Using input As New CSVReader(path:=GetFile(<source>a,,,,e</source>.Value))
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
        Using input As New CSVReader(path:=GetFile(<source>a,b,c,d,</source>.Value))
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
        Using input As New CSVReader(path:=GetFile())
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
        Using input As New CSVReader(path:=GetFile(<source>a,"b
c",d</source>.Value))
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
        Using input As New CSVReader(path:=GetFile(<source>a;b;c</source>), delimiter:=";"c)
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
        Using input As New CSVReader(path:=GetFile(<source>a;"b;c";d</source>.Value), delimiter:=";"c)
            Dim tuple = input.ReadTuple
            Assert.AreEqual(3, tuple.Length)
            Assert.AreEqual("a", tuple(0))
            Assert.AreEqual("b;c", tuple(1))
            Assert.AreEqual("d", tuple(2))
            Assert.IsTrue(input.EndOfStream)
        End Using
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub AllTuple_ShowsAllTuples()
        Dim source = <source>a;b;c
d;e;f
g;h;i</source>.Value

        Dim expected(3)() As String

        expected(0) = New String() {"a", "b", "c"}
        expected(1) = New String() {"d", "e", "f"}
        expected(2) = New String() {"g", "h", "i"}

        Using input As New CSVReader(path:=GetFile(source), delimiter:=";"c)
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
        Dim source = <source>a;b;c
d;e;f
g;h;i</source>.Value

        Dim expected = New String() {"a", "b", "c", "d", "e", "f", "g", "h", "i"}

        Using input As New CSVReader(path:=GetFile(source), delimiter:=";"c)
            Dim count = 0
            Assert.IsFalse(input.EndOfStream)
            For Each cell In input.AllCells
                Assert.AreEqual(expected(count), cell)
                count += 1
            Next
            Assert.IsTrue(input.EndOfStream)
        End Using

    End Sub

    Private Shared ReadOnly lockObj As New Object
    Private Shared Function GetFile()
        SyncLock lockObj
            Return GetFile(<source>a,b,c
d,e,f</source>.Value)
        End SyncLock
    End Function

    Private Shared Function GetFile(contents As String)
        SyncLock lockObj
            Dim tempLocation = Path.GetTempFileName
            Dim fileLength = tempLocation.Length
            File.WriteAllText(tempLocation, contents)
            Console.WriteLine("Filename: ""{0}"" Char Count: {1}", tempLocation, fileLength)
            Return tempLocation
        End SyncLock
    End Function

End Class

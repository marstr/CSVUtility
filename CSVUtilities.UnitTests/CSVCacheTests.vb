Imports System.IO
Imports Utilities.CSV

<TestClass>
Public Class CSVCacheTests

#If DEBUG Then
    Private Const DEFAULT_TIMEOUT = TestTimeout.Infinite
#Else
    Private Const DEFAULT_TIMEOUT = 2000
#End If

    Private Shared ReadOnly SLUFF_FILE = Path.GetTempFileName()

    Private Shared ReadOnly PATH_LOCK As New Object

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod, ExpectedException(GetType(ArgumentOutOfRangeException))>
    Public Sub DefaultInitialization()
        Dim subject = New CSVCache(SLUFF_FILE)
        Assert.AreEqual(0UI, subject.Rows)
        Assert.AreEqual(SLUFF_FILE, subject.TargetFile)
        Dim unexpectedColumnCount = subject.ColumnCount(0)
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub TwoDimensionalInitialization()
        Dim simpleContent(1, 1) As String
        simpleContent(0, 0) = "alpha"
        simpleContent(0, 1) = "beta"
        simpleContent(1, 0) = "charlie"
        simpleContent(1, 1) = "delta"
        Dim subject = New CSVCache(SLUFF_FILE, simpleContent)
        Assert.AreEqual(2UI, subject.Rows)
        For i = 0 To 1
            Assert.AreEqual(2UI, subject.ColumnCount(i))
            For j = 0 To 1
                Assert.AreEqual(simpleContent(i, j), subject.ReadCell(i, j))
            Next
        Next
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub JaggedArrayInitialization()
        Dim simpleContent(6)() As String
        simpleContent(0) = New String(2) {"alfa", "bravo", "charlie"}
        simpleContent(1) = New String(1) {"delta", "echo"}
        simpleContent(2) = New String(2) {"foxtrot", "golf", "hotel"}
        simpleContent(3) = New String(5) {"india", "juliett", "kilo", "lima", "mike", "november"}
        simpleContent(4) = New String(2) {"oscar", "papa", "quebec"}
        simpleContent(5) = New String(0) {"romeo"}
        simpleContent(6) = New String(7) {"sierra", "tango", "uniform", "victor", "whiskey", "xray", "yankee", "zulu"}

        'Verify that contents were copied correctly
        Dim subject = New CSVCache(SLUFF_FILE, simpleContent)
        Assert.AreEqual(CType(simpleContent.Length, UInteger), subject.Rows)
        For i = 0 To subject.Rows - 1
            Assert.AreEqual(CType(simpleContent(i).Length, UInteger), subject.ColumnCount(i))
            For j = 0 To subject.ColumnCount(i) - 1
                Assert.AreEqual(simpleContent(i)(j), subject.ReadCell(i, j))
            Next
        Next

        'Verify that the copy that was made was a deep copy
        simpleContent(0)(0) = "new value"
        Assert.AreEqual("alfa", subject.ReadCell(0, 0))
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub TrivialSetReadSequenceNoSerialization()
        Dim subject = New CSVCache(SLUFF_FILE)
        Dim writtenText = "simpleContent"
        Assert.AreEqual(0UI, subject.Rows)
        subject.SetCell(0, 0, writtenText)
        Assert.AreEqual(1UI, subject.Rows)
        Assert.AreEqual(1UI, subject.ColumnCount(0))
        Assert.AreEqual(writtenText, subject.ReadCell(0, 0))
        Assert.AreEqual(1UI, subject.Rows)
        Assert.AreEqual(1UI, subject.ColumnCount(0))
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT), ExpectedException(GetType(ArgumentOutOfRangeException))>
    Public Sub OutOfBoundsReadSmallDifference()
        Dim subject = New CSVCache(SLUFF_FILE)
        subject.ReadCell(0, 0)
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT), ExpectedException(GetType(ArgumentOutOfRangeException))>
    Public Sub OutOfBoundsReadLargeDifference()
        Dim subject = New CSVCache(SLUFF_FILE)
        subject.ReadCell(10000, 10000)
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub TrivialMultipleReadSetNoSerialization()
        Dim subject = New CSVCache(SLUFF_FILE)
        Dim writtenText1 = "simpleContent1"
        Dim writtenText2 = "simpleContent2"
        Assert.AreEqual(0UI, subject.Rows)
        subject.SetCell(0, 0, writtenText1)
        Assert.AreEqual(1UI, subject.Rows)
        Assert.AreEqual(1UI, subject.ColumnCount(0))
        Assert.AreEqual(writtenText1, subject.ReadCell(0, 0))
        subject.SetCell(0, 1, writtenText2)
        Assert.AreEqual(1UI, subject.Rows)
        Assert.AreEqual(2UI, subject.ColumnCount(0))
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub TrivialMultipleReadSetSerialzed()
        Dim desiredPath = String.Empty
        Try
            Dim writtenText = "simpleContent"
            Dim subject As CSVCache
            SyncLock PATH_LOCK
                desiredPath = Path.GetTempFileName()
                subject = New CSVCache(desiredPath)
                Assert.AreEqual(desiredPath, subject.TargetFile)
                subject.SetCell(0, 0, writtenText)
                subject.Save()
            End SyncLock
            Dim serializedContent = File.ReadAllText(subject.TargetFile)
            Assert.AreEqual(writtenText, serializedContent)
        Finally
            File.Delete(desiredPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub TrivialMultilineReadSetSerialized()
        Dim desiredPath = String.Empty
        Dim lineOne = "first line content"
        Dim lineTwo = "second line content"
        Try
            Dim subject As CSVCache
            SyncLock PATH_LOCK
                desiredPath = Path.GetTempFileName()
                subject = New CSVCache(desiredPath)
                Assert.AreEqual(desiredPath, subject.TargetFile)
                subject.SetCell(0, 0, lineOne)
                subject.SetCell(1, 0, lineTwo)
                subject.Save()
            End SyncLock

            Dim serializedContent = File.ReadAllLines(subject.TargetFile)
            Assert.AreEqual(2, serializedContent.Length)
            Assert.AreEqual(lineOne, serializedContent(0))
            Assert.AreEqual(lineTwo, serializedContent(1))
        Finally
            File.Delete(desiredPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub MultilineReadSetSerializedWithHeader()
        Dim desiredPath = String.Empty
        Try
            Dim subject As CSVCache
            Dim columnNames = New String(2) {"First", "Second", "Third"}
            SyncLock PATH_LOCK
                desiredPath = Path.GetTempFileName()
                subject = New CSVCache(desiredPath, columnNames)
                Assert.AreEqual(desiredPath, subject.TargetFile)
                subject.SetCell(0, 0, "alfa")
                subject.SetCell(0, 1, "bravo")
                subject.SetCell(0, 2, "charlie")
                subject.SetCell(1, 0, "delta")
                subject.SetCell(1, 1, "echo")
                subject.SetCell(1, 2, "foxtrot")
                subject.Save()
            End SyncLock
            Dim serializedContent = File.ReadAllLines(desiredPath)
            Assert.AreEqual(3, serializedContent.Length)
            Assert.AreEqual("First,Second,Third", serializedContent(0))
            Assert.AreEqual("alfa,bravo,charlie", serializedContent(1))
            Assert.AreEqual("delta,echo,foxtrot", serializedContent(2))
        Finally
            File.Delete(desiredPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub ForceCacheRowsGrowth()
        Dim message = "content1"
        Dim subject = New CSVCache(SLUFF_FILE)
        For i = 1UI To 100UI
            subject.AppendRow(New String() {message})
            Assert.AreEqual(i, subject.Rows)
        Next
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub ForceCacheRowsShrinkDescending()
        Dim basis = New String(1000, 0) {}
        Dim transform = Function(val) String.Format("content{0}", val)
        For i = 0 To basis.GetUpperBound(0)
            basis(i, 0) = transform(i)
        Next
        Dim subject = New CSVCache(SLUFF_FILE, basis)
        Assert.AreEqual(CType(basis.GetUpperBound(0) + 1, UInteger), subject.Rows)
        For i = basis.GetUpperBound(0) To 0 Step -1
            Assert.AreEqual(CType(basis.GetUpperBound(1) + 1, UInteger), subject.ColumnCount(i))
            Assert.AreEqual(transform(i), subject.ReadCell(i, 0))
            subject.RemoveRow(i)
            Assert.AreEqual(CType(i, UInteger), subject.Rows)
        Next
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub ForceCacheRowsShrinkAscending()
        Dim basis = New String(1000, 0) {}
        Dim transfrom = Function(val) String.Format("content{0}", val)
        For i = 0 To basis.GetUpperBound(0)
            basis(i, 0) = transfrom(i)
        Next
        Dim subject = New CSVCache(SLUFF_FILE, basis)
        Assert.AreEqual(CType(basis.GetUpperBound(0) + 1, UInteger), subject.Rows)
        For i = 0 To basis.GetUpperBound(0)
            Assert.AreEqual(CType(basis.GetUpperBound(1) + 1, UInteger), subject.ColumnCount(0))
            Assert.AreEqual(transfrom(i), subject.ReadCell(0, 0))
            subject.RemoveRow(0)
            Assert.AreEqual(CType(1000 - i, UInteger), subject.Rows)
        Next
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub ForceCacheColumnGrow()

    End Sub
End Class

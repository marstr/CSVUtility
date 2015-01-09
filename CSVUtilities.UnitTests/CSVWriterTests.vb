Imports System.IO
Imports Utilities.CSV

<TestClass()>
Public Class CSVWriterTests
#If DEBUG Then
    Private Const DEFAULT_TIMEOUT = TestTimeout.Infinite
#Else
    Private Const DEFAULT_TIMEOUT = 2000
#End If

    Private Shared ReadOnly TEMP_PATH_LOCK As New Object

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub WriteSingleCell()
        Dim tempPath As String
        SyncLock TEMP_PATH_LOCK
            tempPath = Path.GetTempFileName
            Using output As New CSVWriter(path:=tempPath)
                output.Write(True)
            End Using
        End SyncLock
        Assert.AreEqual("True", File.ReadAllText(tempPath))
        File.Delete(tempPath)
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub WriteMultipleSingleCells()
        Dim tempPath As String
        SyncLock TEMP_PATH_LOCK
            tempPath = Path.GetTempFileName
            Using output As New CSVWriter(path:=tempPath)
                output.Write("Hello")
                output.Write("World")
            End Using
        End SyncLock
        Assert.AreEqual("Hello,World", File.ReadAllText(tempPath))
        File.Delete(tempPath)
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub WriteMultipleCellsSingleTuple()
        Dim tempPath As String
        Dim tuple(2) As String
        tuple(0) = "a"
        tuple(1) = "b"
        tuple(2) = "c"

        SyncLock TEMP_PATH_LOCK
            tempPath = Path.GetTempFileName
            Using output As New CSVWriter(path:=tempPath)
                output.WriteTuple(tuple)
            End Using
        End SyncLock
        Assert.AreEqual("a,b,c", File.ReadAllText(tempPath))
        File.Delete(tempPath)
    End Sub

    <Timeout(DEFAULT_TIMEOUT)>
    <TestMethod()>
    Public Sub WriteMultipleCellTuples()
        Dim tempPath = String.Empty
        Dim firstTuple(2) As String
        Dim secondTuple(2) As String
        Try
            firstTuple(0) = "a"
            firstTuple(1) = "b"
            firstTuple(2) = "c"

            secondTuple(0) = "d"
            secondTuple(1) = "e"
            secondTuple(2) = "f"

            SyncLock TEMP_PATH_LOCK
                tempPath = Path.GetTempFileName
                Using output As New CSVWriter(tempPath)
                    output.WriteTuple(firstTuple)
                    output.WriteTuple(secondTuple)
                End Using
            End SyncLock
            Dim fileLines = File.ReadAllLines(tempPath)
            Assert.AreEqual(2, fileLines.Length)
            Assert.AreEqual("a,b,c", fileLines(0))
            Assert.AreEqual("d,e,f", fileLines(1))
        Finally
            File.Delete(tempPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub WriteMultipleSingleCellTuples()
        Dim tempPath = String.Empty
        Try
            Dim firstContent = "first line content"
            Dim secondContent = "second line content"
            SyncLock TEMP_PATH_LOCK
                tempPath = Path.GetTempFileName
                Using output As New CSVWriter(tempPath)
                    output.WriteTuple(New String() {firstContent})
                    output.WriteTuple(New String() {secondContent})
                End Using
            End SyncLock
            Dim fileLines = File.ReadAllLines(tempPath)
            Assert.AreEqual(2, fileLines.Length)
            Assert.AreEqual(firstContent, fileLines(0))
            Assert.AreEqual(secondContent, fileLines(1))
        Finally
            File.Delete(tempPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub OverloadsForChar()
        Dim tempPath = String.Empty
        Dim fodder As Char
        fodder = Nothing
        Try
            SyncLock TEMP_PATH_LOCK
                tempPath = Path.GetTempFileName
                Using output As New CSVWriter(tempPath)
                    output.Write(fodder)
                    output.WriteLine(fodder)
                    output.WriteTuple(New Char() {fodder})
                End Using
            End SyncLock
            Dim fileContents = File.ReadAllLines(tempPath)
            Assert.AreEqual(2, fileContents.Length)
            Assert.AreEqual(String.Format("{0},{0}", vbNullChar), fileContents(0))
            Assert.AreEqual(vbNullChar, fileContents(1))
        Finally
            File.Delete(tempPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub OverloadsForBoolean()
        Dim tempPath = String.Empty
        Dim fodder As Boolean
        fodder = Nothing
        Try
            SyncLock TEMP_PATH_LOCK
                tempPath = Path.GetTempFileName
                Using output As New CSVWriter(tempPath)
                    output.Write(fodder)
                    output.WriteLine(fodder)
                    output.WriteTuple(New Boolean() {fodder})
                End Using
            End SyncLock
            Dim fileContents = File.ReadAllLines(tempPath)
            Assert.AreEqual(2, fileContents.Length)
            Assert.AreEqual("False,False", fileContents(0))
            Assert.AreEqual("False", fileContents(1))
        Finally
            File.Delete(tempPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub TuplesViaSingleCellAdds()
        Dim tempPath = String.Empty
        Try
            SyncLock TEMP_PATH_LOCK
                tempPath = Path.GetTempFileName
                Using output As New CSVWriter(tempPath)
                    output.Write("alpha")
                    output.Write("beta")
                    output.Write("charlie")
                    output.WriteLine()
                    output.Write("delta")
                    output.Write("epsilon")
                    output.Write("foxtrot")
                End Using
            End SyncLock

            Dim fileContents = File.ReadAllLines(tempPath)
            Assert.AreEqual(2, fileContents.Length)
            Assert.AreEqual("alpha,beta,charlie", fileContents(0))
            Assert.AreEqual("delta,epsilon,foxtrot", fileContents(1))
        Finally
            File.Delete(tempPath)
        End Try
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub ToStringRecquiresNormalization_Cell()
        Dim tempPath As String
        SyncLock TEMP_PATH_LOCK
            tempPath = Path.GetTempFileName
            Using output As New CSVWriter(tempPath)
                output.Write(New ExampleClass1())
            End Using
            Dim observed = File.ReadAllText(tempPath)
            Assert.AreEqual("""I, the object be4 u, use commas to represent myself.""", observed)
        End SyncLock
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub ToStringRecquiresNormalization_Tuple()
        Dim tempPath As String
        SyncLock TEMP_PATH_LOCK
            tempPath = Path.GetTempFileName
            Using output As New CSVWriter(tempPath)
                output.WriteTuple(New Object() {New ExampleClass1(), New ExampleClass2()})
            End Using
            Dim observed = File.ReadAllText(tempPath)
            Assert.AreEqual("""I, the object be4 u, use commas to represent myself."",""The previous object was foolish to use commas. I prefer the """";""""""", observed)
        End SyncLock
    End Sub

    Private Class ExampleClass1
        Public Overrides Function ToString() As String
            Return "I, the object be4 u, use commas to represent myself."
        End Function
    End Class

    Private Class ExampleClass2
        Public Overrides Function ToString() As String
            Return "The previous object was foolish to use commas. I prefer the "";"""
        End Function
    End Class
End Class
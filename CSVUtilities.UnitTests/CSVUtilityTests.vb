Imports System.IO
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Utilities.CSV

<TestClass()> Public Class CSVUtilityTests

    Private Const DEFAULT_TIMEOUT = 2000

#Region "String Normalization & Denormalization"

    Private Shared Sub CheckNormalization(normalized As String, denormalized As String)
        Assert.AreEqual(expected:=normalized,
                        actual:=NormalizeString(denormalized),
                        message:="Normalizing Broken")

        Assert.AreEqual(expected:=denormalized,
                        actual:=DenormalizeString(normalized),
                        message:="Denormalizing Broken")

        Assert.AreEqual(expected:=normalized,
                        actual:=NormalizeString(DenormalizeString(normalized)),
                        message:="Roundtripping Starting from Normalized Broken")

        Assert.AreEqual(expected:=denormalized,
                        actual:=DenormalizeString(NormalizeString(denormalized)),
                        message:="Roundtripping Starting from Denormalized Broken")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub PlainStringNormalization()
        CheckNormalization("a", "a")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub CenterDelimitedStringNormalization()
        CheckNormalization("""a,b""", "a,b")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EndDelimitedStringNormalization()
        CheckNormalization("""ab,""", "ab,")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub StartDelimitedStringNormalization()
        CheckNormalization(""",ab""", ",ab")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub AdjacentDelimitersStringNormalization()
        CheckNormalization("""a,,b""", "a,,b")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub MultipleDelimitersStringNormalization()
        CheckNormalization("""a,b,c""", "a,b,c")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub CenterQuoteNormalization()
        CheckNormalization("""a""""b""", "a""b")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EndQuoteNormalization()
        CheckNormalization("""ab""""""", "ab""")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub StartQuoteNormalization()
        CheckNormalization("""""""ab""", """ab")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub DoubleAdjacentQuotesStringNormalization()
        CheckNormalization("""a""""""""b""", "a""""b")
    End Sub

    <TestMethod()> Public Sub TerminalQuotesStringNormalization()
        CheckNormalization("""""""ab""""""", """ab""")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub LineBreakNormalization()
        CheckNormalization("""a" & vbCrLf & "b""", "a" & vbCrLf & "b")
    End Sub

#End Region

    '#Region "Read/Write Operations"
    '    Shared Sub New()

    '    End Sub

    '    Private Shared simpleFileSource = <source>a,b,c
    'd,e,f</source>.Value

    '    Private Shared Function GetSimpleFile() As CachedCSVFile
    '        Dim tempLocation = Path.GetTempFileName
    '        File.WriteAllText(tempLocation, <source>a,b,c
    'd,e,f</source>.Value)
    '        Return New CachedCSVFile(tempLocation)
    '    End Function

    '    <TestMethod()> Public Sub BasicRead()
    '        Dim csv = GetSimpleFile()
    '        Assert.AreEqual("a", csv.Read(0, 0))
    '        Assert.AreEqual("b", csv.Read(0, 1))
    '        Assert.AreEqual("c", csv.Read(0, 2))
    '        Assert.AreEqual("d", csv.Read(1, 0))
    '        Assert.AreEqual("e", csv.Read(1, 1))
    '        Assert.AreEqual("f", csv.Read(1, 2))
    '    End Sub

    '    <TestMethod()> Public Sub BasicColumnCount()
    '        Dim csv = GetSimpleFile()
    '        Assert.AreEqual(3, csv.Columns)
    '    End Sub

    '    <TestMethod()> Public Sub BasicRowCount()
    '        Dim csv = GetSimpleFile()
    '        Assert.AreEqual(2, csv.Rows)
    '    End Sub
    '#End Region
End Class
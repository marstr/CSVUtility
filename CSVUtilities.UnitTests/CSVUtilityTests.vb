Imports System.IO
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Utilities.CSV

<TestClass()> Public Class CSVUtilityTests

#If DEBUG Then
    Private Const DEFAULT_TIMEOUT = TestTimeout.Infinite
#Else
    Private Const DEFAULT_TIMEOUT = 2000
#End If

    Private Shared Sub CheckNormalization(normalized As String, denormalized As String, Optional delimiter As Char = DEFAULT_DELIMITER)
        Assert.AreEqual(expected:=normalized,
                        actual:=NormalizeString(denormalized, delimiter),
                        message:="Normalizing Broken")

        Assert.AreEqual(expected:=denormalized,
                        actual:=DenormalizeString(normalized),
                        message:="Denormalizing Broken")

        Assert.AreEqual(expected:=normalized,
                        actual:=NormalizeString(DenormalizeString(normalized), delimiter),
                        message:="Roundtripping Starting from Normalized Broken")

        Assert.AreEqual(expected:=denormalized,
                        actual:=DenormalizeString(NormalizeString(denormalized, delimiter)),
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
        CheckNormalization("""a" & DEFAULT_DELIMITER & "b""", "a" & DEFAULT_DELIMITER & "b")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub EndDelimitedStringNormalization()
        CheckNormalization("""ab" & DEFAULT_DELIMITER & """", "ab" & DEFAULT_DELIMITER)
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub StartDelimitedStringNormalization()
        CheckNormalization("""" & DEFAULT_DELIMITER & "ab""", DEFAULT_DELIMITER & "ab")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub AdjacentDelimitersStringNormalization()
        CheckNormalization("""a" & DEFAULT_DELIMITER & DEFAULT_DELIMITER & "b""", "a" & DEFAULT_DELIMITER & DEFAULT_DELIMITER & "b")
    End Sub

    <TestMethod()>
    <Timeout(DEFAULT_TIMEOUT)>
    Public Sub MultipleDelimitersStringNormalization()
        CheckNormalization("""a" & DEFAULT_DELIMITER & "b" & DEFAULT_DELIMITER & "c""", "a" & DEFAULT_DELIMITER & "b" & DEFAULT_DELIMITER & "c")
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

#Region "Non-Default Delimiter"
    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub NonDefaultDelimiter_Basic()
        CheckNormalization("""a;b""", "a;b", ";"c)
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub NonDefaultDelimiter_ValueIncludesDefault()
        Dim subject = String.Format("a{0}b", DEFAULT_DELIMITER)
        CheckNormalization(subject, subject, ";"c)
    End Sub

    <TestMethod, Timeout(DEFAULT_TIMEOUT)>
    Public Sub NonDefaultDelimiter_Whitespace()
        CheckNormalization("""a" & vbTab & "b""", "a" & vbTab & "b", vbTab)
    End Sub
#End Region
End Class
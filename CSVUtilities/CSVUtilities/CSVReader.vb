Imports System.IO
Imports System.Text

Namespace CSV
    Public Class CSVReader
        Inherits StreamReader

        Public ReadOnly Property Delimiter As Char
            Get
                Return _delimiter
            End Get
        End Property
        Private _delimiter As Char

#Region "Constructors"
        Public Sub New(stream As Stream, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream)
            _delimiter = delimiter
        End Sub

        Public Sub New(path As String, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path)
            _delimiter = delimiter
        End Sub

        Public Sub New(stream As Stream, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding)
            _delimiter = delimiter
        End Sub

        Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        Public Sub New(path As String, encoding As Encoding, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, encoding)
            _delimiter = delimiter
        End Sub
        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        Public Sub New(path As String, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
            _delimiter = delimiter
        End Sub

        Public Sub New(path As String, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
            _delimiter = delimiter
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen)
            _delimiter = delimiter
        End Sub
#End Region
        ''' <summary>
        ''' Retrieves the contents of the next previously unread cell.
        ''' </summary>
        ''' <returns>The next cell available in the stream.</returns>
        Public Function ReadCell() As String
            Dim quoteCount = 0
            Dim rawEncounteredText As New StringBuilder
            Dim latest As Char
            Dim encounteredEnd = False
            Do
                If EndOfStream Then
                    encounteredEnd = True
                    Exit Do
                End If
                latest = ChrW(Read())
                rawEncounteredText.Append(latest)
                If latest = """"c Then
                    quoteCount += 1
                End If
            Loop While Not ((latest = Delimiter OrElse latest = ControlChars.Cr OrElse latest = ControlChars.Lf) AndAlso (quoteCount Mod 2 = 0))
            If rawEncounteredText.Length > 0 AndAlso Not encounteredEnd Then
                rawEncounteredText.Remove(rawEncounteredText.Length - 1, 1)
            End If

            Return DenormalizeString(rawEncounteredText.ToString())
        End Function

        ''' <summary>
        ''' Retrieves an entire line of a CSV file.
        ''' </summary>
        ''' <returns>Each of the cells that existed in a logical line of a CSV file.</returns>
        Public Function ReadTuple() As String()
            Dim rawText As String
            Dim quoteCount = 0

            'Combine all physical lines that are 
            Dim accumulator As New StringBuilder
            Do
                rawText = ReadLine()
                accumulator.Append(rawText)
                quoteCount += rawText.Count(Function(c) c = """"c)
                If quoteCount Mod 2 = 0 Then
                    Exit Do
                Else
                    accumulator.AppendLine()
                End If
            Loop
            'Reset State variables
            quoteCount = 0
            rawText = accumulator.ToString

            'Separate Cells
            Dim encountered As New List(Of String)
            Dim tortoise = 0, hare = 0
            While hare < rawText.Length
                Dim latest = rawText(hare)
                If latest = """"c Then
                    quoteCount += 1
                ElseIf latest = Delimiter AndAlso quoteCount Mod 2 = 0 Then
                    Dim cellContents = rawText.Substring(tortoise, hare - tortoise)
                    encountered.Add(DenormalizeString(cellContents))
                    tortoise = hare + 1
                End If
                hare += 1
            End While

            encountered.Add(rawText.Substring(tortoise))

            Return encountered.ToArray
        End Function

        ''' <summary>
        ''' Allows for easy access to each tuple inside of a CSV file.
        ''' </summary>
        ''' <returns>An enumerable that exposes each tuple in a CSV file.</returns>
        ''' <remarks>
        ''' This method does not memoize the entire file, it accesses it as a stream.
        ''' </remarks>
        Public Iterator Function AllTuples() As IEnumerable(Of String())
            While Not EndOfStream
                Yield ReadTuple()
            End While
        End Function

        ''' <summary>
        ''' Allows for easy access to each cell inside of a CSV file.
        ''' </summary>
        ''' <returns>An enumerable that exposes each cell in a CSV file.</returns>
        ''' <remarks>
        ''' This method does not memoize the entire file, it accesses it as a stream.
        ''' </remarks>
        Public Iterator Function AllCells() As IEnumerable(Of String)
            While Not EndOfStream
                Yield ReadCell()
            End While
        End Function
    End Class
End Namespace
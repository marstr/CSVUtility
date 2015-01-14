Imports System.IO
Imports System.Text

Namespace CSV
    ''' <summary>
    ''' Using this class, it is easy to treat a character separated value file as a stream.
    ''' Unlike similiar tools, this extends StreamReader, so all of your intuition on how streams work
    ''' apply to this class as well.
    ''' </summary>
    Public Class CSVReader
        Inherits StreamReader

        ''' <summary>
        ''' The character being used to separate cells on a line.
        ''' </summary>
        Public ReadOnly Property Delimiter As Char
            Get
                Return _delimiter
            End Get
        End Property
        Private _delimiter As Char

        ''' <summary>
        ''' The discrete position of the current cell being read.
        ''' </summary>
        Public ReadOnly Property Position As RasterCursor
            Get
                Return _postion
            End Get
        End Property
        Private ReadOnly _postion As RasterCursor = New RasterCursor

#Region "Constructors"
        'Many of the constructors in this section are fairly redundant and only here in want of a better syntax to express this.

        ''' <summary>
        ''' Instantiates a means of reading a CSV file in a stream based manner.
        ''' </summary>
        ''' <param name="stream">The stream to be read as CSV data.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(stream As Stream, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Instantiates a means of reading a CSV file in a stream based manner.
        ''' </summary>
        ''' <param name="path">The path to a file that should be read as a CSV file.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(path As String, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Instantiates a means of reading a CSV file in a stream based manner, wit the specified byte order mark detection option.
        ''' </summary>
        ''' <param name="stream">The stream to be read as CSV data.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(stream As Stream, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified stream, with the specified character encoding.
        ''' </summary>
        ''' <param name="stream">The stream to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(stream As Stream, encoding As Encoding, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified file name, with the specified byte order mark detection option.
        ''' </summary>
        ''' <param name="path">The complete file path to be read as CSV data.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified stream, with the specified character encoding.
        ''' </summary>
        ''' <param name="path">The complete file path to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(path As String, encoding As Encoding, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, encoding)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified stream, with the specified character encoding.
        ''' </summary>
        ''' <param name="stream">The stream to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified file name, with the specified byte order mark detection option.
        ''' </summary>
        ''' <param name="path">The complete file path to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(path As String, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified stream, with the specified character encoding.
        ''' </summary>
        ''' <param name="stream">The stream to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="bufferSize">The minimum buffer size.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified stream, with the specified character encoding.
        ''' </summary>
        ''' <param name="path">The complete file path to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="bufferSize">The minimum buffer size.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(path As String, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
            _delimiter = delimiter
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the CSVReader class for the specified stream, with the specified character encoding.
        ''' </summary>
        ''' <param name="stream">The stream to be read as CSV data.</param>
        ''' <param name="encoding">The character encoding to use.</param>
        ''' <param name="detectEncodingFromByteOrderMarks">Indicates whether to look for byte order marks at the beginning of the file.</param>
        ''' <param name="bufferSize">The minimum buffer size.</param>
        ''' <param name="leaveOpen"><code>true</code> to leave the stream open after the <see cref="CSVReader"/> object is disposed; otherwise, <code>false</code>.</param>
        ''' <param name="delimiter">The character used to separate cells in a single line of the stream.</param>
        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen)
            _delimiter = delimiter
        End Sub
#End Region
        ''' <summary>
        ''' Retrieves the contents of the next unread cell.
        ''' </summary>
        ''' <returns>The contents of the next cell available in the stream.</returns>
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
            Position.Increment()
            Return DenormalizeString(rawEncounteredText.ToString())
        End Function

        ''' <summary>
        ''' Retrieves the contents of the next unread cell.
        ''' </summary>
        ''' <returns>The contents of the next cell available in the stream.</returns>
        Public Async Function ReadCellAsync() As Task(Of String)
            Return Await Task.Run(AddressOf ReadCell)
        End Function

        ''' <summary>
        ''' Retrieves each entry on an entire line of a CSV file.
        ''' </summary>
        ''' <returns>Each of the cells that existed in a logical line of a CSV file.</returns>
        Public Function ReadTuple() As String()
            Dim rawText As String
            Dim quoteCount = 0

            If EndOfStream Then
                Return Nothing
            End If

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
            Position.NextRow()

            Return encountered.ToArray
        End Function

        ''' <summary>
        ''' Retrieves each entry on an entire line of a CSV file.
        ''' </summary>
        ''' <returns>Each of the cells that existed in a logical line of a CSV file.</returns>
        Public Async Function ReadTupleAsync() As Task(Of String())
            Return Await Task.Run(AddressOf ReadTuple)
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
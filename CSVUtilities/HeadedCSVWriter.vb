Imports System.IO

Namespace CSV
    ''' <summary>
    ''' Writes a CSVFile in a stream based fashion. The first row denotes the names of each column.
    ''' </summary>
    Public Class HeadedCSVWriter
        Inherits CSVWriter
        ''' <summary>
        ''' A mapping between a column and a name.
        ''' </summary>
        Public ReadOnly Property Header As String()
            Get
                Return _header
            End Get
        End Property
        Private ReadOnly _header As String()

        ''' <summary>
        ''' Instantiates a new HeadedCSVWriter.
        ''' </summary>
        ''' <param name="strm">The stream to be written to.</param>
        ''' <param name="header">The names of each column.</param>
        ''' <param name="delimiter">The character that will separate each value in a given line.</param>
        Public Sub New(strm As Stream, header As String(), Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(strm, delimiter)
            _header = header
        End Sub

        ''' <summary>
        ''' Instantiates a new HeadedCSVWriter.
        ''' </summary>
        ''' <param name="path">The complete path to the file to be written.</param>
        ''' <param name="header">The names of each column.</param>
        ''' <param name="delimiter">The character that will separate each value in a given line.</param>
        Public Sub New(path As String, header As String(), Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, delimiter)
            _header = header
        End Sub

        ''' <summary>
        ''' Writes a whole collection of values to a line a CSV file.
        ''' </summary>
        ''' <typeparam name="T">The type of the data to be written.</typeparam>
        ''' <param name="cells">The values to write.</param>
        ''' <param name="conversion">THe means of converting this to a string.</param>
        Protected Overrides Sub WriteTupleHelper(Of T)(cells() As T, Optional conversion As Func(Of T, String) = Nothing)
            If Position.Row = 0 Then
                MyBase.WriteTupleHelper(Header)
            End If
            If Not Position.Column = 0 Then
                WriteLine()
            End If
            If cells.Length <= Header.Length Then
                MyBase.WriteTupleHelper(cells, conversion)
            Else
                Throw New ArgumentException(paramName:="cells", message:="No operation can be performed that would misalign data with its header.")
            End If
        End Sub
    End Class
End Namespace

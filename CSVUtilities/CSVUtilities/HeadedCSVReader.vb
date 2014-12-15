Imports System.IO
Imports System.Text

Namespace CSV
    Public Class HeadedCSVReader
        Inherits CSVReader

        Public ReadOnly Property Header As String()
            Get
                If _Header Is Nothing Then
                    If Position.Row = 0 Then
                        _Header = ReadTuple()
                    Else
                        Throw New InvalidOperationException("Once a stream has already been acted on, a header can no longer be read with integrity.")
                    End If
                End If
                Return _Header.Clone()
            End Get
        End Property
        Private _Header As String()

#Region "Constructors"
        'Many of the constructors in this section are fairly redundant and only here in want of a better syntax to express this.
        Public Sub New(stream As Stream, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(stream, delimiter)
            _Header = header
        End Sub

        Public Sub New(path As String, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(path, delimiter)
            _Header = header
        End Sub

        Public Sub New(stream As Stream, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(stream, detectEncodingFromByteOrderMarks, delimiter)
            _Header = header
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(stream, encoding, delimiter)
            _Header = header
        End Sub

        Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(path, detectEncodingFromByteOrderMarks, delimiter)
            _Header = header
        End Sub

        Public Sub New(path As String, encoding As Encoding, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(path, encoding, delimiter)
            _Header = header
        End Sub
        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, delimiter)
            _Header = header
        End Sub

        Public Sub New(path As String, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, delimiter)
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, delimiter)
            _Header = header
        End Sub

        Public Sub New(path As String, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize, delimiter)
            _Header = header
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean, Optional delimiter As Char = DEFAULT_DELIMITER, Optional header As String() = Nothing)
            MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen, delimiter)
            _Header = header
        End Sub
#End Region

        ''' <summary>
        ''' Fetches the contents of a single row of a CSV file and associates it with the column name that it belonged to.
        ''' </summary>
        ''' <param name="output">A dictionary that will contain the contents of a row after the method has finished executing.</param>
        ''' <returns>True if the operation was successful and </returns>
        Public Overloads Function ReadTuple(ByRef output As Dictionary(Of String, String)) As Boolean
            If output Is Nothing Then
                output = New Dictionary(Of String, String)
            Else
                output.Clear()
            End If

            If (Not Position.Column = 0) OrElse EndOfStream Then
                Return False
            End If

            Dim rawTuple = New LinkedList(Of String)(ReadTuple())
            If Not Header.Length = rawTuple.Count Then
                Return False
            End If

            For Each column In Header
                output.Add(column, rawTuple.FirstOrDefault())
                rawTuple.RemoveFirst()
            Next

            Return True
        End Function
    End Class
End Namespace

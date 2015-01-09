Imports System.IO
Imports System.Text

Namespace CSV
    ''' <summary>
    ''' Captures everything that is about to be written, 
    ''' </summary>
    Public Class CSVWriter
        Inherits StreamWriter

        ''' <summary>
        ''' Stores the next position that should be written to.
        ''' </summary>
        Public Property Position As RasterCursor
            Get
                Return _Position
            End Get
            Protected Set(value As RasterCursor)
                _Position = value
            End Set
        End Property
        Private _Position As New RasterCursor


        ''' <summary>
        ''' The character used to separate columns in a row.
        ''' </summary>
        Public ReadOnly Property Delimiter As Char
            Get
                Return _delimeter
            End Get
        End Property
        Private ReadOnly _delimeter As Char

#Region "Constructors"
        Public Sub New(strm As Stream, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(strm)
            _delimeter = delimiter
        End Sub

        Public Sub New(path As String, Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path)
            _delimeter = delimiter
        End Sub
#End Region
#Region "Helpers"
        ''' <summary>
        ''' Writes the appropiate string to the stream, and handles inserting delimiters correctly.
        ''' </summary>
        ''' <param name="value">A pre-normalized value.</param>
        ''' <remarks>
        ''' The value to write must be normalized to allow for perf-optimizations where possible.
        ''' </remarks>
        Protected Sub WriteHelper(value As String)
            Dim copy As String
            If Position.Column = 0 Then
                copy = value
            Else
                copy = Delimiter & value
            End If
            MyBase.Write(copy)
            'Maintain cursor position
            Position += 1
            If Position.Column = 0 Then
                MyBase.WriteLine()
            End If
        End Sub

        ''' <summary>
        ''' Writes the appropiate string to the stream, and handles inserting delimiters correctly.
        ''' </summary>
        ''' <param name="value">A pre-normalized value.</param>
        ''' <remarks>
        ''' The value to write must be normalized before being passed to this method to allow for perf-optimizations where possible.
        ''' </remarks>
        Protected Async Function WriteHelperAsync(value As String) As Task
            Await Task.Run(Sub() WriteHelper(value))
        End Function

        Protected Overridable Sub WriteTupleHelper(Of T)(cells() As T, Optional conversion As Func(Of T, String) = Nothing)
            conversion = If(conversion, Function(x) x.ToString)
            If Not (Position.Column = 0) Then
                WriteLine()
            End If
            For Each cell In cells
                WriteHelper(conversion(cell))
            Next
        End Sub

        Protected Async Function WriteTupleHelperAsync(Of T)(cells() As T, Optional conversion As Func(Of T, String) = Nothing) As Task
            conversion = If(conversion, Function(x) x.ToString)
            If Not (Position.Column = 0) Then
                Await WriteLineAsync()
            End If
            For Each cell In cells
                Await WriteHelperAsync(conversion(cell))
            Next
        End Function
#End Region
#Region "Write Overloads"
        Public Overrides Sub Write(value As Boolean)
            WriteHelper(value)
        End Sub
        Public Overrides Sub Write(value As Char)
            If value = """"c OrElse value = Delimiter Then
                WriteHelper(NormalizeString(value.ToString, Delimiter))
            Else
                WriteHelper(value)
            End If
        End Sub

        Public Overrides Sub Write(buffer() As Char)
            Dim builder = New StringBuilder
            For Each character In buffer
                builder.Append(character)
            Next
            Write(builder.ToString())
        End Sub

        Public Overrides Sub Write(value As Decimal)
            WriteHelper(value)
        End Sub

        Public Overrides Sub Write(value As Double)
            WriteHelper(value)
        End Sub

        Public Overrides Sub Write(value As Integer)
            WriteHelper(value)
        End Sub

        Public Overrides Sub Write(value As Long)
            WriteHelper(value)
        End Sub

        Public Overrides Sub Write(value As Single)
            WriteHelper(value)
        End Sub

        Public Overrides Sub Write(value As String)
            WriteHelper(NormalizeString(value, Delimiter))
        End Sub
        Public Overrides Sub Write(value As Object)
            WriteHelper(NormalizeString(value.ToString(), Delimiter))
        End Sub

        Public Overrides Sub Write(value As UInteger)
            WriteHelper(value.ToString())
        End Sub

        Public Overrides Sub Write(value As ULong)
            WriteHelper(value.ToString())
        End Sub

        Public Overrides Sub Write(format As String, arg0 As Object)
            Dim realizedValue = String.Format(format, arg0)
            WriteHelper(NormalizeString(realizedValue, Delimiter))
        End Sub

        Public Overrides Sub Write(format As String, arg0 As Object, arg1 As Object)
            Dim copy = String.Format(format, arg0, arg1)
            WriteHelper(NormalizeString(copy, Delimiter))
        End Sub

        Public Overrides Sub Write(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
            Dim copy = String.Format(format, arg0, arg1, arg2)
            WriteHelper(NormalizeString(format, Delimiter))
        End Sub

        Public Overrides Sub Write(buffer() As Char, index As Integer, count As Integer)
            Dim accumulator As New StringBuilder
            For i = index To index + count - 1
                accumulator.Append(buffer(i))
            Next
            WriteHelper(NormalizeString(accumulator.ToString, Delimiter))
        End Sub

        Public Overrides Sub Write(format As String, ParamArray arg() As Object)
            Dim copy = String.Format(format, arg)
            WriteHelper(NormalizeString(copy, Delimiter))
        End Sub

        Public Overloads Sub Write(cells() As String)
            For Each cell In cells
                WriteHelper(NormalizeString(cell, Delimiter))
            Next
        End Sub
#End Region
#Region "WriteAsync Overloads"
        Public Overrides Async Function WriteAsync(buffer() As Char, index As Integer, count As Integer) As Task
            Dim accumulator As New StringBuilder
            For i = index To index + count - 1
                accumulator.Append(buffer(i))
            Next
            Await WriteHelperAsync(NormalizeString(accumulator.ToString, Delimiter))
        End Function

        Public Overrides Async Function WriteAsync(value As Char) As Task
            Await WriteHelperAsync(NormalizeString(value, Delimiter))
        End Function

        Public Overrides Async Function WriteAsync(value As String) As Task
            Await WriteHelperAsync(NormalizeString(value, Delimiter))
        End Function

#End Region
#Region "WriteTuple Overloads"
        Public Sub WriteTuple(cells() As Boolean)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As Char)
            WriteTupleHelper(cells, Function(chr) NormalizeString(chr, Delimiter))
        End Sub

        Public Sub WriteTuple(cells()() As Char)
            WriteTupleHelper(cells, Function(arr() As Char)
                                        Dim accumulator As New StringBuilder
                                        For Each entry In arr
                                            accumulator.Append(entry)
                                        Next
                                        Return NormalizeString(accumulator.ToString, Delimiter)
                                    End Function)
        End Sub

        Public Sub WriteTuple(cells() As Decimal)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As Double)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As Integer)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As Long)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells())
            WriteTupleHelper(cells, Function(x) NormalizeString(x.ToString, Delimiter))
        End Sub

        Public Sub WriteTuple(cells() As Single)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As UInteger)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As ULong)
            WriteTupleHelper(cells)
        End Sub

        Public Sub WriteTuple(cells() As String)
            WriteTupleHelper(cells, Function(str)
                                        If str Is Nothing Then
                                            Return String.Empty
                                        Else
                                            Return NormalizeString(str, Delimiter)
                                        End If
                                    End Function)
        End Sub
#End Region
#Region "WriteTupleAsync"
        Public Async Function WriteTupleAsync(cells() As Char) As Task
            Await WriteTupleHelperAsync(cells)
        End Function
        Public Async Function WriteTupleAsync(cells()() As Char) As Task
            Await WriteTupleHelperAsync(cells, Function(arr() As Char)
                                                   Dim accumulator As New StringBuilder
                                                   For Each entry In arr
                                                       accumulator.Append(entry)
                                                   Next
                                                   Return NormalizeString(accumulator.ToString, Delimiter)
                                               End Function)
        End Function

        Public Async Function WriteTupleAsync(cells() As Decimal) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As Double) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As Integer) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As Long) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As Object) As Task
            Await WriteTupleHelperAsync(cells, Function(x) NormalizeString(x.ToString, Delimiter))
        End Function

        Public Async Function WriteTupleAsync(cells() As Single) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As UInteger) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As ULong) As Task
            Await WriteTupleHelperAsync(cells)
        End Function

        Public Async Function WriteTupleAsync(cells() As String) As Task
            Await WriteTupleHelperAsync(cells, Function(str) NormalizeString(str, Delimiter))
        End Function
#End Region
#Region "WriteLine Overloads"
        Public Overrides Sub WriteLine()
            MyBase.Write(NewLine)
            Position.NextRow()
        End Sub
        Public Overrides Sub WriteLine(buffer() As Char)
            Dim builder = New StringBuilder
            For Each letter In buffer
                builder.Append(letter)
            Next
            WriteHelper(NormalizeString(builder.ToString(), Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Char)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Boolean)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(buffer() As Char, index As Integer, count As Integer)
            Dim builder = New StringBuilder
            For i = 0 To count - 1
                builder.Append(buffer(index + i))
            Next
            WriteHelper(NormalizeString(builder.ToString(), Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Decimal)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Double)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Integer)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Long)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Object)
            WriteHelper(NormalizeString(value.ToString(), Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As Single)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As String)
            WriteHelper(NormalizeString(value, Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As UInteger)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(value As ULong)
            WriteHelper(value)
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(format As String, arg0 As Object)
            WriteHelper(NormalizeString(String.Format(format, arg0), Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(format As String, arg0 As Object, arg1 As Object)
            WriteHelper(NormalizeString(String.Format(format, arg0, arg1), Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
            WriteHelper(NormalizeString(String.Format(format, arg0, arg1, arg2), Delimiter))
            WriteLine()
        End Sub

        Public Overrides Sub WriteLine(format As String, ParamArray arg() As Object)
            WriteHelper(NormalizeString(String.Format(format, arg), Delimiter))
            WriteLine()
        End Sub
#End Region
    End Class
End Namespace

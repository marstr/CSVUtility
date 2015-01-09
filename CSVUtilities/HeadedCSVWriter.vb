Imports System.IO

Namespace CSV
    Public Class HeadedCSVWriter
        Inherits CSVWriter

        Public ReadOnly Property Header As String()
            Get
                Return _header
            End Get
        End Property
        Private ReadOnly _header As String()

        Public Sub New(strm As Stream, header As String(), Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(strm, delimiter)
            _header = header

        End Sub

        Public Sub New(path As String, header As String(), Optional delimiter As Char = DEFAULT_DELIMITER)
            MyBase.New(path, delimiter)
            _header = header
        End Sub

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

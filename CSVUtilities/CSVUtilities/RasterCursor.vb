''' <summary>
''' A simple mechanism for working with a rasterizing cursor.
''' </summary>
Public Class RasterCursor

    Public Property CurrentPosition As Position
        Get
            Return New Position With {.Row = Row, .Column = Column}
        End Get
        Set(value As Position)
            Row = value.Row
            Column = value.Column
        End Set
    End Property


    Public Property Row As UInteger
        Get
            Return _row
        End Get
        Set(value As UInteger)
            If value < 0 Then
                Throw New IndexOutOfRangeException("Cursor cannot be placed before start of raster.")
            End If
            _row = value
        End Set
    End Property
    Private _row As UInteger

    Public Property Column As UInteger
        Get
            Return _column
        End Get
        Set(value As UInteger)
            If value < 0 OrElse ((Not ColumnCount = 0) AndAlso value > ColumnCount) Then
                Throw New IndexOutOfRangeException("Cursor cannot be placed outside of the raster.")
            Else
                _column = value
            End If
        End Set
    End Property
    Private _column As UInteger

    Public ReadOnly Property ColumnCount As UInteger
        Get
            Return _columnCount
        End Get
    End Property
    Private ReadOnly _columnCount As UInteger


    Public Sub New(Optional columnCount As UInteger = Nothing)
        MyClass.New(New Position With {.Row = 0, .Column = 0}, columnCount)
    End Sub

    Public Sub New(row As UInteger, col As UInteger, Optional columnCount As UInteger = Nothing)
        MyClass.New(New Position With {.Row = row, .Column = col}, columnCount)
    End Sub

    Public Sub New(pos As Position, Optional columnCount As UInteger = Nothing)
        CurrentPosition = pos
        _columnCount = columnCount
    End Sub

    Public Sub New(original As RasterCursor)
        _columnCount = original.ColumnCount

    End Sub

    Public Sub Increment(Optional magnitude As UInteger = 1UI)
        If ColumnCount = 0 Then
            Column += 1
        Else
            Dim proposed = Column + magnitude
            Row += proposed / ColumnCount
            Column += proposed Mod ColumnCount
        End If
    End Sub

    Public Sub Decrement(Optional magnitude As UInteger = 1UI)
        Dim proposed = Column - magnitude

        If proposed >= 0 Then
            Column = Row
        ElseIf ColumnCount = 0 Then
            Throw New InvalidOperationException("When the number of columns in the raster has not been specified, decrement cannot wrap around rows.")
        Else
            Row -= Math.Abs(proposed) / ColumnCount
            Column -= Math.Abs(proposed) Mod ColumnCount
        End If
    End Sub

    Public Sub NextRow()
        Row += 1
        Column = 0
    End Sub

    Public Shared Operator +(cursor As RasterCursor, magnitude As Integer)
        cursor.Increment(magnitude)
        Return cursor
    End Operator

    Public Shared Operator -(cursor As RasterCursor, magnitude As Integer)
        cursor.Decrement(magnitude)
        Return cursor
    End Operator

    Public Structure Position
        Public Property Row As UInteger
        Public Property Column As UInteger

        Public Shared Operator =(ByVal leftSide As Position, ByVal rightSide As Position)
            Return leftSide.Column = rightSide.Column AndAlso leftSide.Row = rightSide.Row
        End Operator

        Public Shared Operator <>(ByVal leftSide As Position, ByVal rightSide As Position)
            Return leftSide = rightSide
        End Operator

        Public Shared Operator <(ByVal leftSide As Position, ByVal rightSide As Position)
            Throw New NotImplementedException
        End Operator

        Public Shared Operator >(ByVal leftSide As Position, ByVal rightSide As Position)
            Throw New NotImplementedException
        End Operator
    End Structure
End Class

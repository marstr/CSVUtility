Namespace CSV
    ''' <summary>
    ''' An easy way to interact with small sets of data that are or will become stored in a CSV file.
    ''' </summary>
    ''' <remarks>
    ''' This class is not an alternative to <seealso cref="CSVWriter"/> or <seealso cref="CSVReader"/>, if data needs to be treated as if it were a stream do not use <see cref="CSVCache"/>.
    ''' </remarks>
    Public Class CSVCache

        Private Const DEFAULT_CACHE_DIMENSION = 4UI

        Private m_cache()() As String

        Private m_columns() As UInteger

        ''' <summary>
        ''' The number of rows that are currently represented in the cache.
        ''' </summary>
        Public Property Rows As UInteger
            Get
                Return _observedRows
            End Get
            Protected Set(value As UInteger)
                _observedRows = value
            End Set
        End Property
        Private _observedRows As UInteger

        ''' <summary>
        ''' The path to the file where serialization and deserialization of the cache should occur.
        ''' </summary>
        Public Property TargetFile As String

        ''' <summary>
        ''' The character that is printed to separate columns in a row.
        ''' </summary>
        Public Property Delimiter As Char

        ''' <summary>
        ''' A row of values that should be ignored during interactions, but still serialized when written.
        ''' </summary>
        Public ReadOnly Property Header As String()
            Get
                Return _header
            End Get
        End Property
        Private ReadOnly _header As String()

        ''' <summary>
        ''' Initializes an empty <see cref="CSVCache"/>.
        ''' </summary>
        ''' <param name="targetFile">The file that will be used to store the cache.</param>
        ''' <param name="header">A set of values to prepend or ignore each time the file is serialized or deserialized respectively.</param>
        Public Sub New(targetFile As String, Optional header As String() = Nothing)
            Me.TargetFile = targetFile
            Delimiter = DEFAULT_DELIMITER
            _header = If(header Is Nothing, Nothing, header.Clone())
            m_columns = New UInteger(DEFAULT_CACHE_DIMENSION) {}
            m_cache = New String(DEFAULT_CACHE_DIMENSION)() {}
            Rows = 0
            For i = 0UI To DEFAULT_CACHE_DIMENSION
                m_columns(i) = 0
                m_cache(i) = New String(DEFAULT_CACHE_DIMENSION) {}
            Next
        End Sub

        ''' <summary>
        ''' Initializes a <see cref="CSVCache"/> with copied values from <paramref name="array"/>.
        ''' </summary>
        ''' <param name="targetFile">The file underlying this cache.</param>
        ''' <param name="array">The values that should initially be held.</param>
        ''' <param name="header">A set of values to prepend or ignore each time the file is serialized or deserialized respectively.</param>
        Public Sub New(targetFile As String, array()() As String, Optional header As String() = Nothing)
            Me.TargetFile = targetFile
            Delimiter = DEFAULT_DELIMITER
            _header = If(header Is Nothing, Nothing, header.Clone())
            Rows = array.Length
            Dim rowCount As UInteger = Math.Max(Rows, DEFAULT_CACHE_DIMENSION)
            m_columns = (New UInteger(rowCount) {})
            m_cache = New String(rowCount)() {}
            For i = 0 To array.Length - 1
                Dim rowLength As Integer = array(i).Length
                m_cache(i) = (New String(Math.Max(rowLength, DEFAULT_CACHE_DIMENSION)) {})
                m_columns(i) = rowLength
                For j = 0 To rowLength - 1
                    m_cache(i)(j) = array(i)(j)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Initializes a <see cref="CSVCache"/> with copied values from <paramref name="array"/>.
        ''' </summary>
        ''' <param name="targetFile">The file underlying this cache.</param>
        ''' <param name="array">The values that should be initially held in the cache.</param>
        ''' <param name="header">A set of values to prepend or ignore each time the file is serialized or deserialized respectively.</param>
        Public Sub New(targetFile As String, array(,) As String, Optional header As String() = Nothing)
            Me.TargetFile = targetFile
            _header = If(header Is Nothing, Nothing, header.Clone())
            Delimiter = DEFAULT_DELIMITER

            Dim firstDimensionSize As Integer = array.GetUpperBound(0)
            Dim secondDimensionSize As Integer = array.GetUpperBound(1)
            m_columns = New UInteger(firstDimensionSize) {}
            m_cache = (New String(firstDimensionSize)() {})
            Rows = firstDimensionSize + 1
            For i = 0 To firstDimensionSize
                m_cache(i) = New String(secondDimensionSize) {}
                m_columns(i) = secondDimensionSize + 1
                For j = 0 To secondDimensionSize
                    m_cache(i)(j) = array(i, j)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Replaces the contents on the underlying file with the current values in the <see cref="CSVCache"/>.
        ''' </summary>
        Public Sub Save()
            Using writer As New CSVWriter(path:=TargetFile, delimiter:=Delimiter)
                If Header IsNot Nothing Then
                    writer.WriteTuple(Header)
                End If
                For i = 0 To Rows - 1
                    Dim observedColumns As Integer = m_columns(i) - 1
                    Dim pureContents = New String(observedColumns) {}
                    For j = 0 To observedColumns
                        pureContents(j) = m_cache(i)(j)
                    Next
                    writer.WriteTuple(pureContents)
                Next
            End Using
        End Sub

        ''' <summary>
        ''' Replaces the contents of the underlying file with the current values in the <see cref="CSVCache"/>.
        ''' </summary>
        ''' <returns>An awaitable <seealso cref="Task"/>.</returns>
        Public Async Function SaveAsync() As Task
            Await Task.Run(AddressOf Save)
        End Function

        ''' <summary>
        ''' Replaces the current contents of the <see cref="CSVCache"/> with the values stored in the underlying file.
        ''' </summary>
        Public Sub Refresh()
            Using reader As New CSVReader(path:=TargetFile, delimiter:=Delimiter)
                If Header IsNot Nothing Then
                    reader.ReadTuple()
                End If
                For Each row In reader.AllTuples
                    AppendRow(row)
                Next
            End Using
        End Sub

        ''' <summary>
        ''' Retrieves the value stored in a particular location.
        ''' </summary>
        ''' <param name="row">The vertical position of the the cell to be read.</param>
        ''' <param name="col">The horizontal position of the cell to be read.</param>
        ''' <returns>A textual representation of the value stored at a position.</returns>
        Public Function ReadCell(row As UInteger, col As UInteger) As String
            If row >= Rows OrElse col >= m_columns(row) Then
                Throw New ArgumentOutOfRangeException()
            End If
            Return m_cache(row)(col)
        End Function

        ''' <summary>
        ''' Sets a cell with some content in the <see cref="CSVCache"/>
        ''' </summary>
        ''' <param name="row">The vertical position of the cell to be set.</param>
        ''' <param name="col">The horizontal position of the cell to be set.</param>
        ''' <param name="value">The contents that the cell should have.</param>
        ''' <remarks>When <paramref name="value"/> is null, it's value will be converted to <seealso cref="String"/>.Empty</remarks>
        Public Sub SetCell(row As UInteger, col As UInteger, value As String)
            GrowCache(row, col)
            If row >= Rows Then
                Rows = row + 1
            End If
            If col >= m_columns(row) Then
                m_columns(row) = col + 1
            End If
            m_cache(row)(col) = If(value, String.Empty)
        End Sub

        ''' <summary>
        ''' Deletes a cell and shifts all values behind it on the same row of it to the left one column.
        ''' When this method is called on an empty Row, that row is deleted and all following rows are moved up.
        ''' </summary>
        ''' <param name="row">The vertical position of the cell to be removed.</param>
        ''' <param name="col">The horizontal position of the cell to be removed.</param>
        Public Sub RemoveCell(row As UInteger, col As UInteger)
            If row > Rows Then 'Separating this from the first clause prevents an ArgumentOutOfRangeException
                Exit Sub
            ElseIf col = 0 AndAlso m_columns(row) = 0 Then
                RemoveRow(row)
            ElseIf col >= m_columns(row) Then
                Exit Sub
            Else
                For i = col To m_columns(row) - 2
                    m_cache(row)(i) = m_cache(row)(i + 1)
                Next
                m_columns(row) -= 1
                ShrinkRow(row)
            End If
        End Sub

        ''' <summary>
        ''' Adds a row between existing rows, or at the end of the <see cref="CSVCache"/>.
        ''' </summary>
        ''' <param name="index">The location of the new row.</param>
        ''' <param name="values">The values to insert at the given location.</param>
        Public Sub InsertRow(index As UInteger, values As IEnumerable(Of String))
            Dim numValues = values.Count()
            Dim i As UInteger
            If index > Rows Then
                Throw New ArgumentOutOfRangeException()
            ElseIf index = Rows Then
                GrowCache(Rows, Math.Max(numValues - 1, 0))
            Else
                GrowCache(Rows, 0)
                i = Rows
                While i > index
                    m_cache(i) = m_cache(i - 1)
                    m_columns(i) = m_columns(i - 1)
                    i -= 1
                End While
                m_cache(index) = (New String(Math.Max(numValues, DEFAULT_CACHE_DIMENSION)) {})
            End If
            i = 0
            m_columns(index) = numValues
            For Each entry In values
                m_cache(index)(i) = entry
                i += 1
            Next
            Rows += 1
        End Sub

        ''' <summary>
        ''' Inserts an empty row in a given location.
        ''' </summary>
        ''' <param name="index"></param>
        Public Sub InsertRow(index As UInteger)
            InsertRow(index, Enumerable.Empty(Of String))
        End Sub

        ''' <summary>
        ''' Replaces a row with the provided values.
        ''' </summary>
        ''' <param name="index">The index of the row to update.</param>
        ''' <param name="values"></param>
        ''' <remarks>When null is provided for <paramref name="values"/> </remarks>
        Public Sub SetRow(index As UInteger, Optional values As IEnumerable(Of String) = Nothing)
            If index > Rows Then
                Throw New ArgumentOutOfRangeException()
            ElseIf index = Rows Then
                InsertRow(index, values)
            Else
                If values Is Nothing Then
                    values = Enumerable.Empty(Of String)
                End If

                m_columns(index) = values.Count()
                m_cache(index) = New String(Math.Max(CType(values.Count(), UInteger), DEFAULT_CACHE_DIMENSION)) {}
                Dim i = 0
                For Each entry In values 'For Each is used to prevent perf hits where argument values is not well suited for random access.
                    m_cache(index)(i) = entry
                    i += 1
                Next
            End If
        End Sub

        ''' <summary>
        ''' Adds a collection of cells to the end of the <see cref="CSVCache"/>.
        ''' </summary>
        ''' <param name="values">The values to be appended.</param>
        Public Sub AppendRow(values As IEnumerable(Of String))
            InsertRow(Rows, values)
        End Sub

        ''' <summary>
        ''' Adds an empty row to the end of the <see cref="CSVCache"/>.
        ''' </summary>
        Public Sub AppendRow()
            InsertRow(Rows)
        End Sub

        ''' <summary>
        ''' Deletes a row by shifting all rows after it.
        ''' </summary>
        ''' <param name="row"></param>
        Public Sub RemoveRow(row As UInteger)
            If row >= Rows Then
                Throw New ArgumentOutOfRangeException()
            End If
            For i = row To Rows - 2
                m_cache(i) = m_cache(i + 1)
                m_columns(i) = m_columns(i + 1)
            Next
            m_cache(Rows - 1) = New String(DEFAULT_CACHE_DIMENSION) {}
            m_columns(Rows - 1) = 0
            Rows -= 1
            ShrinkCache()
        End Sub

        Public Sub ClearRow(row As UInteger)
            If row >= Rows Then
                Throw New ArgumentOutOfRangeException()
            End If
            m_cache(row) = New String(DEFAULT_CACHE_DIMENSION) {}
            m_columns(row) = 0
        End Sub

        ''' <summary>
        ''' Deletes the contents of a cell, but preserves the notion that there is a cell there.
        ''' </summary>
        ''' <param name="row">The vertical position of the cell to be cleared.</param>
        ''' <param name="col">The horizontal position of the cell to be cleared.</param>
        Public Sub ClearCell(row As UInteger, col As UInteger)
            If row >= Rows OrElse col >= m_columns(row) Then
                Throw New ArgumentOutOfRangeException()
            End If
            m_cache(row)(col) = String.Empty
        End Sub

        ''' <summary>
        ''' Counts the number of columns that are in a particular row.
        ''' </summary>
        ''' <param name="row">The vertical position of the row that should be evaluated.</param>
        ''' <returns>The number of columns in a row.</returns>
        Public Function ColumnCount(row As UInteger) As UInteger
            If row >= Rows Then
                Throw New ArgumentOutOfRangeException()
            End If
            Return m_columns(row)
        End Function

        ''' <summary>
        ''' Expand the cache so that it can accomodate data such that it was added at the given location.
        ''' </summary>
        ''' <param name="row">The largest index of a row that currently needs to be supported.</param>
        ''' <param name="col">The largest index of a column that currently needs to be supported.</param>
        Private Sub GrowCache(row As UInteger, col As UInteger)
            'Grow the number of rows if necessary
            If row >= m_cache.Length Then
                Dim origCache = m_cache
                Dim origCols = m_columns
                Dim newRowCount = origCache.Length

                While newRowCount < row
                    newRowCount *= 2
                End While

                m_cache = New String(newRowCount)() {}
                m_columns = New UInteger(newRowCount) {}
                For i = 0 To origCache.Length - 1
                    m_cache(i) = origCache(i)
                    m_columns(i) = origCols(i)
                Next
                For i = origCache.Length To m_cache.Length - 1
                    m_cache(i) = New String(DEFAULT_CACHE_DIMENSION) {}
                    m_columns(i) = 0UI
                Next
            End If

            'Grow a particular row so that it can accomodate the column.
            If col >= m_cache(row).Length Then
                Dim origRow = m_cache(row)
                Dim columnCountPrime = origRow.Length

                While columnCountPrime < col
                    columnCountPrime *= 2
                End While

                m_cache(row) = New String(columnCountPrime) {}

                For i = 0 To origRow.Length - 1
                    m_cache(row)(i) = origRow(i)
                Next
                For i = origRow.Length To m_cache(row).Length - 1
                    m_cache(row)(i) = Nothing
                Next
            End If
        End Sub

        ''' <summary>
        ''' Reduces the size of the cache to the smallest size that is still accomodating of all of the present values.
        ''' </summary>
        Private Sub ShrinkCache()
            If Rows < (m_cache.Length / 2) Then
                Dim origCache = m_cache
                Dim origCols = m_columns
                Dim newRowCount As Integer = m_cache.Length / 2
                m_cache = New String(newRowCount)() {}
                m_columns = New UInteger(newRowCount) {}
                For i = 0 To m_cache.Length - 1
                    m_cache(i) = origCache(i)
                    m_columns(i) = origCols(i)
                Next
            End If
        End Sub

        Private Sub ShrinkRow(row As UInteger)
            If row >= Rows Then
                Throw New ArgumentOutOfRangeException()
            End If
            If m_columns(row) < m_cache(row).Length / 2 Then
                Dim origCache = m_cache(row)
                m_cache(row) = New String(m_cache(row).Length / 2) {}
                For j = 0 To m_cache(row).Length
                    m_cache(row)(j) = origCache(j)
                Next
            End If
        End Sub
    End Class
End Namespace

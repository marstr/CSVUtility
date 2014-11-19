﻿Imports System.Text.RegularExpressions

Namespace CSV
    ''' <summary>
    ''' This is a bucket to hold common definitions between the CSVWriter and the CSVReader.
    ''' </summary>
    Public Module CSVUtilities

        Public Const DEFAULT_DELIMITER = ","c

        ''' <summary>
        ''' Adds a characters to a string, so that it can be written safely to a CSV file.
        ''' </summary>
        ''' <param name="content">The message to be escaped.</param>
        ''' <param name="delimiter">The character that is used to terminate cells.</param>
        ''' <returns>A safe string</returns>
        Public Function NormalizeString(content As String, Optional delimiter As String = DEFAULT_DELIMITER) As String
            content = content.Replace("""", """""") 'Replace single quotes with double quotes
            If escapeRequired.IsMatch(content) Or content.Contains(delimiter) Then
                content = """" & content & """" 'Surround message with quotes
            End If
            Return content
        End Function
        Private ReadOnly escapeRequired As New Regex("(?(\r\f)\r\f|[\""\,\n])")

        ''' <summary>
        ''' Removes any characters from a string that were added to make it safe to be written to a CSV file.
        ''' </summary>
        ''' <param name="content">The escaped message to be traslated.</param>
        ''' <returns>The orignal and unescaped message.</returns>
        Public Function DenormalizeString(content As String) As String
            If Not String.IsNullOrEmpty(content) Then
                If content.First().Equals(""""c) Then
                    content = content.Substring(1, content.Length - 2)
                End If
                content = content.Replace("""""", """") 'Replace double quotes with single quotes
            End If
            Return content
        End Function
    End Module
End Namespace

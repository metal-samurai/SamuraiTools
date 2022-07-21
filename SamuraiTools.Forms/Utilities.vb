Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Namespace Forms
    Public Module Utilities

        'credit to zhen yang for the original c# code
        <Extension>
        Public Sub PasteFromClipboard(ByVal grid As System.Windows.Forms.DataGridView)
            Dim o As DataObject = CType(Clipboard.GetDataObject(), DataObject)

            If o.GetDataPresent(DataFormats.Text) Then
                Dim rowOfInterest As Integer = grid.CurrentCell.RowIndex
                Dim selectedRows As String() = System.Text.RegularExpressions.Regex.Split(o.GetData(DataFormats.Text).ToString().TrimEnd("\r\n".ToCharArray()), "\r\n")

                If selectedRows Is Nothing OrElse selectedRows.Length = 0 Then
                    Return
                End If

                For Each row As String In selectedRows
                    If rowOfInterest >= grid.Rows.Count Then
                        Exit For
                    End If

                    Dim data As String() = System.Text.RegularExpressions.Regex.Split(row, "\t")
                    Dim col As Integer = grid.CurrentCell.ColumnIndex

                    For Each ob As String In data
                        If col >= grid.Columns.Count Then
                            Exit For
                        End If

                        If ob IsNot Nothing Then
                            grid(col, rowOfInterest).Value = Convert.ChangeType(ob, grid(col, rowOfInterest).ValueType)
                            col += 1
                        End If
                    Next

                    rowOfInterest += 1
                Next
            End If
        End Sub
    End Module
End Namespace

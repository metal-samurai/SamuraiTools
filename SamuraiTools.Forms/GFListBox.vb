Imports System.Windows.Forms

'to do:
'figure out why horizontal scrollbars don't show when not wrapping text
Namespace Forms
    'inspired by code from stackoverflow user Kosmos and Zuoliu Ding
    Public Class GFListBox
        Inherits System.Windows.Forms.ListBox

        Public ReadOnly Property SeparatorPen As System.Drawing.Pen
        Public Property WrapText As Boolean

        'used to determine HorizontalExtent property
        Protected maxItemWidth As Long

        Public Sub New()
            MyBase.New()

            SeparatorPen = New System.Drawing.Pen(System.Drawing.Color.Black, 0)
            SeparatorPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid

            WrapText = False
        End Sub

        Protected Overridable Function GetDisplayString(item As Object) As String
            If Me.DisplayMember = String.Empty Then
                Return item.ToString()
            Else
                Return Interaction.CallByName(item, Me.DisplayMember, CallType.Get, Nothing)
            End If
        End Function

        Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
            MyBase.OnDrawItem(e)

            e.DrawBackground()
            e.DrawFocusRectangle()

            If Me.Items.Count > 0 And e.Index > -1 Then
                Dim displayString As String = GetDisplayString(Me.Items(e.Index))

                If e.Index > 0 And SeparatorPen.Width > 0 Then
                    'the pen "straddles" the y coordinate when drawn, so i want that coordinate to increase by half the pen's width so as to contain it within the same rectangle as the text
                    Dim y As Long = e.Bounds.Location.Y + (SeparatorPen.Width \ 2)

                    e.Graphics.DrawLine(SeparatorPen, e.Bounds.Location.X, y, e.Bounds.Location.X + e.Bounds.Width, y)

                    e.Graphics.DrawString(displayString, e.Font, New System.Drawing.SolidBrush(e.ForeColor), New System.Drawing.Rectangle(e.Bounds.Location.X, e.Bounds.Location.Y + SeparatorPen.Width, e.Bounds.Width, e.Bounds.Height - SeparatorPen.Width))
                Else
                    e.Graphics.DrawString(displayString, e.Font, New System.Drawing.SolidBrush(e.ForeColor), e.Bounds)
                End If
            End If
        End Sub

        Protected Overrides Sub OnMeasureItem(e As MeasureItemEventArgs)
            MyBase.OnMeasureItem(e)

            If Me.Items.Count > 0 And e.Index > -1 Then
                Dim displayString As String = GetDisplayString(Me.Items(e.Index))
                Dim itemSize As System.Drawing.SizeF

                If WrapText Then
                    itemSize = e.Graphics.MeasureString(displayString, Me.Font, Me.Width)
                    e.ItemHeight = Convert.ToInt32(itemSize.Height)

                    Me.HorizontalExtent = 0
                Else
                    itemSize = e.Graphics.MeasureString(displayString, Me.Font)

                    If e.Index = 0 Or itemSize.Width > maxItemWidth Then
                        maxItemWidth = itemSize.Width
                    End If

                    Me.HorizontalExtent = maxItemWidth
                End If

                If e.Index > 0 Then
                    e.ItemHeight += SeparatorPen.Width
                End If
            End If
        End Sub
    End Class
End Namespace

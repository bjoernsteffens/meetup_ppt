' https://msdn.microsoft.com/en-us/library/office/ee814734(v=office.14).aspx
' https://msdn.microsoft.com/en-us/library/office/ff744590.aspx    
' https://msdn.microsoft.com/en-us/library/office/ff744177.aspx
' https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx
' https://msdn.microsoft.com/en-us/library/office/ff746629.aspx


Sub fixTableFormat()

    For Each Slide In ActivePresentation.Slides
        
        For Each Shape In Slide.Shapes
           
             If Shape.Type = msoTable Then

                Set tbl = Shape.Table
                tbl.ApplyStyle "{7DF18680-E054-41AD-8BC1-D1AEF772440D}", False
                tbl.FirstRow = True
                tbl.HorizBanding = True
                tbl.LastRow = True
                tbl.LastCol = True

                'Prepare to swap colours around of we are looking at Parature Tickets
                'because then a decrease is good.
                If (tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Ticket Cause") Then
                    'MsgBox ("This is what triggers the arrow colour to change")
                    changeArrowColour = True
                End If
                
                For tblRows = 1 To tbl.Rows.Count
                    
                    tbl.Rows(tblRows).Height = 17
                    For tblCols = 1 To tbl.Columns.Count

                        'MsgBox (tbl.Cell(Rows, Cols).Type)
                        tbl.Columns(tblCols).Width = 55
                        With tbl.Cell(tblRows, tblCols).Shape.TextFrame.TextRange
                            '.Font.Name = "Arial"
                            .Font.Size = 10
                        End With
                        
                        ' Dont touch the last line => Green on Blue bad
                        If (tblRows < tbl.Rows.Count) Then
                            oTxt = tbl.Cell(tblRows, tblCols).Shape.TextFrame.TextRange.Text
                            ' Find the ASCII value of the up arrow   => 8593
                            ' Find the ASCII value of the down arrow => 8595
                            ' MsgBox (AscW(oTxt))
                            If (InStr(1, oTxt, ChrW(8595), vbTextCompare) > 0) Then
                                If (changeArrowColour) Then
                                    ' Decrease in tickets is Good => Green
                                    tbl.Cell(tblRows, tblCols).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(112, 173, 71)
                                Else
                                    ' Decrease in revenue is bad => Red
                                    tbl.Cell(tblRows, tblCols).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(221, 80, 68)
                                End If
                            End If

                            If (InStr(1, oTxt, ChrW(8593), vbTextCompare) > 0) Then
                                If (changeArrowColour) Then
                                    ' Red
                                    tbl.Cell(tblRows, tblCols).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(221, 80, 68)
                                Else
                                    ' Green
                                    tbl.Cell(tblRows, tblCols).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(112, 173, 71)
                                End If

                            End If
                        End If

                    Next

                Next

                ' Fix First Column Witdh
                tbl.Columns(1).Width = 100
                ' Center the table on the slide
                Shape.Left = (ActivePresentation.PageSetup.SlideWidth / 2) - (Shape.Width / 2)
                Shape.Top = (ActivePresentation.PageSetup.SlideHeight / 2) - (Shape.Height / 2)
                'MsgBox (Shape.Width)
                
            End If

        Next

    Next
    
End Sub


↑
UPWARDS ARROW
Unicode: U+2191, UTF-8: E2 86 91




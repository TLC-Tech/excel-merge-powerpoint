Attribute VB_Name = "GeneratePowerPointMacro"
Sub GeneratePowerPoint()

    ' Instantiate powerpoint vars
    Dim newPowerPoint As PowerPoint.Application
    Dim activeSlide As PowerPoint.Slide
    
    ' Open Powerpoint
    On Error Resume Next
    Set newPowerPoint = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    
    ' Make a new powerpoint project
    If newPowerPoint Is Nothing Then
        Set newPowerPoint = New PowerPoint.Application
    End If
    
    ' Give it a blank presentation
    If newPowerPoint.Presentations.Count = 0 Then
        newPowerPoint.Presentations.Add
    End If
    
    ' Make sure its visible to the user
    newPowerPoint.Visible = True

    ' Now look at the current excel file and find the last row
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).row
    
    ' Vars were going to use for the loop
    Dim rng As Range
    Dim row As Range

    Set rng = Range("A1:B" & CStr(lastRow))
    
    ' Loop through rows, treat each row as a new powerpoint slide
    For Each row In rng.Rows
    
        ' Make a new slide
        newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
        ' Go to the new slide in ppt
        newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
        ' Set last slide as active slide for programming on
        Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)
      
        ' Add a textbox with the Lexeme
        activeSlide.Shapes(1).TextFrame.TextRange.Text = row.Cells(1)
        ' Add a textbox with the Translation
        activeSlide.Shapes(2).TextFrame.TextRange.Text = row.Cells(2)
        ' Remove bullet from Translation textbox
        activeSlide.Shapes(2).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
        ' Align the slides to center horizontally
        activeSlide.Shapes(1).TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
        activeSlide.Shapes(2).TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
        ' Slide the text boxes down to center vertically
        activeSlide.Shapes(1).Top = 150
        activeSlide.Shapes(2).Top = 300
        
    Next row
    
    ' Set the background color of the slides so theyre not blinding white
    With newPowerPoint.ActivePresentation.Slides.Range
        .FollowMasterBackground = False
        .Background.Fill.Solid
        .Background.Fill.ForeColor.RGB = RGB(230, 230, 230)
    End With


End Sub

Sub RandomizeAll
    Dim i As Integer
    Dim nextSlides As Integer
    Dim count As Integer
    
    nextSlides = ActivePresentation.Slides.count
    
    For i = 1 To ActivePresentation.Slides.count
        count = Int((i * Rnd) + 1)
        ActiveWindow.ViewType = ppViewSlideSorter
        ActivePresentation.Slides(count).Select
        ActiveWindow.Selection.Cut
        ActivePresentation.Slides(nextSlides - 1).Select
        ActiveWindow.View.Paste
    Next
End Sub


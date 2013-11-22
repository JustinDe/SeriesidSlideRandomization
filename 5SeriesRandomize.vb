Sub Random5()
    
    'For x = 1 To 4 (cycles 4 times for for slides)
    'rng = Int((HighestNumber - LowestNumber + 1) * Rnd) + LowestNumber
    'ActivePresentation.Slides(rng).MoveTo LowestNumber
    'Next (part of for loop)
    
    For x = 1 To 4
        rng = Int((5 - 2 + 1) * Rnd) + 2
        ActivePresentation.Slides(rng).MoveTo 2
    Next
    
    For y = 1 To 4
        rng = Int((10 - 7 + 1) * Rnd) + 7
        ActivePresentation.Slides(rng).MoveTo 7
    Next
    
    For z = 1 To 4
        rng = Int((15 - 12 + 1) * Rnd) + 12
        ActivePresentation.Slides(rng).MoveTo 12
    Next

    rng2 = Int((5 - 1 + 1) * Rnd) + 1
    
    For w = 1 To rng2
        For v = 0 To 4
            ActivePresentation.Slides(1).Select
            ActiveWindow.Selection.Cut
            ActivePresentation.Slides(ActivePresentation.Slides.Count).Select
            ActiveWindow.View.Paste
        Next
    Next

End Sub


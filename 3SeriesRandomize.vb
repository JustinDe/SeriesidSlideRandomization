Sub Random3()
    
    'For x = 1 To 4 (cycles 4 times for for slides)
    'rng = Int((HighestNumber - LowestNumber + 1) * Rnd) + LowestNumber
    'ActivePresentation.Slides(rng).MoveTo LowestNumber
    'Next (part of for loop)
    
    For x = 1 To 3
        rng = Int((3 - 2 + 1) * Rnd) + 2
        ActivePresentation.Slides(rng).MoveTo 2
    Next
    
    For y = 1 To 3
        rng = Int((6 - 5 + 1) * Rnd) + 5
        ActivePresentation.Slides(rng).MoveTo 5
    Next
    
    For z = 1 To 3
        rng = Int((9 - 8 + 1) * Rnd) + 8
        ActivePresentation.Slides(rng).MoveTo 8
    Next

    rng2 = Int((6 - 1 + 1) * Rnd) + 1
    
    For w = 1 To rng2
        For v = 0 To 2
            ActivePresentation.Slides(1).Select
            ActiveWindow.Selection.Cut
            ActivePresentation.Slides(ActivePresentation.Slides.Count).Select
            ActiveWindow.View.Paste
        Next
    Next

End Sub



# VBA - word

`alt F11` `->` `Insert` `->` `Module`.  

# In đậm "Ví dụ mẫu 1:"

```sh
Sub BoldViDuMau()

    Dim rng As Range
    Dim count As Long

    Set rng = ActiveDocument.Content

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting

        .Text = "Ví dụ mẫu [0-9]{1,}:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True

        Do While .Execute
            rng.Font.Bold = True
            count = count + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With

    MsgBox "Đã hoàn thành!" & vbCrLf & _
           "Đã bôi đậm " & count & " cụm 'Ví dụ mẫu'.", _
           vbInformation, "Hoàn tất"

End Sub
```

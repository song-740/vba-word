# VBA - word

`alt F11` `->` `Insert` `->` `Module`.  

# 1. In đậm "Ví dụ mẫu 1:"

```sh
Sub BoldViDuMau()

    Dim rng As Range
    Dim demSoLan As Long
    Dim mauTim As String

    demSoLan = 0

    ' Dung ChrW de dung ky tu co dau, tranh loi hien thi dau trong VBA Editor
    mauTim = "V" & ChrW(237) & " d" & ChrW(7909) & " m" & ChrW(7851) & "u [0-9]{1,}:"

    Set rng = ActiveDocument.Content

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting

        .Text = mauTim
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True

        Do While .Execute
            rng.Font.Bold = True
            demSoLan = demSoLan + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With

    MsgBox "Da hoan thanh!" & vbCrLf & _
           "Da boi dam " & demSoLan & " cum tu khoa da duyet.", _
           vbInformation, "Hoan tat"

End Sub
```

# 2. "Hướng dẫn giải" - In nghiêng, gạch chân, căn giữa

```sh
Sub DinhDangHuongDanGiai()
    Dim para As Paragraph
    Dim rng As Range
    Dim tuKhoa As String
    Dim demSoDong As Integer
    
    ' Dung ChrW de tranh loi font khong hien dau khi go truc tiep
    tuKhoa = "H" & ChrW(432) & ChrW(7899) & "ng d" & ChrW(7851) & "n gi" & ChrW(7843) & "i"
    demSoDong = 0
    
    ' Duyet qua tat ca cac doan van trong tai lieu
    For Each para In ActiveDocument.Paragraphs
        If InStr(1, para.Range.Text, tuKhoa, vbTextCompare) > 0 Then
            Set rng = para.Range
            
            ' Can giua
            rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            ' In nghieng
            rng.Font.Italic = True
            
            ' Gach chan
            rng.Font.Underline = wdUnderlineSingle
            
            demSoDong = demSoDong + 1
        End If
    Next para
    
    MsgBox "Da dinh dang xong! Tong so dong chua tu khoa da duyet: " & demSoDong, vbInformation
End Sub
```

# 3. `a) b) c)` bôi đậm

```sh
Sub BoiDamChuThuong()
    Dim demSoLan As Integer
    Dim rng As Range
    
    demSoLan = 0
    
    ' Tao vung tim kiem la toan bo tai lieu
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Text = "<[a-z]\)"       ' Tim chu thuong dung dau ngoac, o dau tu
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            rng.Font.Bold = True
            demSoLan = demSoLan + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    MsgBox "Da boi dam xong! Tong so cum 'a)' den 'z)' da duyet: " & demSoLan, vbInformation
End Sub
```

# 4. Bôi đậm "Câu 2 [204850]:"

## 4.1. Cơ bản

```sh
Sub BoiDamCauSo()
    Dim demSoLan As Integer
    Dim rng As Range
    Dim mauTim As String
    
    demSoLan = 0
    
    ' Dung ChrW(226) de dung ky tu "a" co mu, tranh loi hien thi dau trong VBA Editor
    mauTim = "C" & ChrW(226) & "u [0-9]{1,} \[[0-9]{1,}\]:"
    
    ' Tao vung tim kiem la toan bo tai lieu
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Text = mauTim
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            rng.Font.Bold = True
            demSoLan = demSoLan + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    MsgBox "Da boi dam xong! Tong so cum 'Cau x [y]:' da duyet: " & demSoLan, vbInformation
End Sub
```

## 4.2. Nâng cao

```sh
Sub BoiDamCauSoTrichSGK()
    Dim demSoLan As Integer
    Dim rng As Range
    Dim mauTim As String
    Dim dsTu(1 To 4) As String
    Dim i As Integer
    
    demSoLan = 0
    
    ' Dung ChrW de dung ky tu co dau, tranh loi hien thi trong VBA Editor
    dsTu(1) = "Tr" & ChrW(237) & "ch SGK C" & ChrW(249) & "ng Kh" & ChrW(225) & "m Ph" & ChrW(225)
    dsTu(2) = "Tr" & ChrW(237) & "ch SGK Ch" & ChrW(226) & "n Tr" & ChrW(7901) & "i S" & ChrW(225) & "ng T" & ChrW(7841) & "o"
    dsTu(3) = "Tr" & ChrW(237) & "ch SGK C" & ChrW(225) & "nh Di" & ChrW(7873) & "u"
    dsTu(4) = "Tr" & ChrW(237) & "ch SGK K" & ChrW(7871) & "t N" & ChrW(7889) & "i Tri Th" & ChrW(7913) & "c"
    
    ' Duyet qua tung cum z, tim rieng cho moi cum
    For i = 1 To 4
        Set rng = ActiveDocument.Content
        
        mauTim = "C" & ChrW(226) & "u [0-9]{1,} \[[0-9]{1,}\] \[" & dsTu(i) & "\]:"
        
        With rng.Find
            .ClearFormatting
            .Text = mauTim
            .MatchWildcards = True
            .MatchCase = False   ' Khong phan biet hoa/thuong de bat moi bien the viet hoa
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                rng.Font.Bold = True
                demSoLan = demSoLan + 1
                rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    
    MsgBox "Da boi dam xong! Tong so cum 'Cau x [y] [z]:' da duyet: " & demSoLan, vbInformation
End Sub
```

# 5. Bôi đậm "1." "2."

```sh
Sub BoiDamSoDauCham()

    Dim rng As Range
    Dim demSoLan As Long
    Dim mauTim As String

    demSoLan = 0

    ' Mau tim: mot hoac nhieu chu so dung o dau tu, theo sau la dau cham
    mauTim = "<[0-9]{1,}\."

    Set rng = ActiveDocument.Content

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting

        .Text = mauTim
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True

        Do While .Execute
            rng.Font.Bold = True
            demSoLan = demSoLan + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With

    MsgBox "Da hoan thanh!" & vbCrLf & _
           "Da boi dam " & demSoLan & " cum so co dau cham da duyet.", _
           vbInformation, "Hoan tat"

End Sub
```

# 6. Thiết lập header "facebook.com"

```sh
Sub ThietLapHeader()

    Dim doc As Document
    Dim hdr As HeaderFooter
    Dim rng As Range
    Dim tenNguoi As String
    Dim gachNgang As String
    Dim linkFb As String
    Dim soDienThoai As String
    
    Set doc = ActiveDocument
    
    ' Dung ChrW de dung ky tu co dau, tranh loi hien thi dau trong VBA Editor
    tenNguoi = ChrW(272) & "inh V" & ChrW(259) & "n T" & ChrW(249) & "ng "
    gachNgang = ChrW(8211) & " "
    linkFb = "facebook.com/tung.dvan.777 "
    soDienThoai = "0987.640.315"
    
    ' Lay Header cua Section dau tien (loai Primary - trang thuong)
    Set hdr = doc.Sections(1).Headers(wdHeaderFooterPrimary)
    
    ' Xoa noi dung cu trong Header (neu co)
    hdr.Range.Delete
    
    ' Dat khoang cach Header from Top = 0.8 cm
    doc.Sections(1).PageSetup.HeaderDistance = CentimetersToPoints(0.8)
    
    ' 1. Nhap ten (dam)
    Set rng = hdr.Range
    rng.Text = tenNguoi
    rng.Font.Bold = True
    rng.Font.Size = 12
    
    ' 2. Nhap gach ngang + link facebook (khong dam)
    Set rng = hdr.Range
    rng.Collapse wdCollapseEnd
    rng.InsertAfter gachNgang & linkFb
    rng.Font.Bold = False
    rng.Font.Size = 12
    
    ' 3. Nhap gach ngang (khong dam)
    Set rng = hdr.Range
    rng.Collapse wdCollapseEnd
    rng.InsertAfter gachNgang
    rng.Font.Bold = False
    rng.Font.Size = 12
    
    ' 4. Nhap so dien thoai (dam)
    Set rng = hdr.Range
    rng.Collapse wdCollapseEnd
    rng.InsertAfter soDienThoai
    rng.Font.Bold = True
    rng.Font.Size = 12
    
    ' Can phai cho toan bo doan Header
    hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    ' Dam bao khong co border nao
    hdr.Range.ParagraphFormat.Borders.Enable = False
    
    MsgBox "Da thiet lap Header xong!", vbInformation, "Hoan tat"

End Sub
```


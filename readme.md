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
    doc.Sections(1).PageSetup.HeaderDistance = CentimetersToPoints(0.5)
    
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

# 7. Watermark


```sh
Sub TaoWatermark()

    Dim doc As Document
    Dim sec As Section
    Dim hdrRange As Range
    Dim shp As Shape
    Dim noiDungWM As String
    Dim i As Integer
    
    Set doc = ActiveDocument
    noiDungWM = "Tung Dvan"
    
    ' Ap dung watermark cho tat ca cac Section trong tai lieu
    For Each sec In doc.Sections
    
        ' Xoa watermark cu neu co (tranh chong lap khi chay lai)
        For i = sec.Headers(wdHeaderFooterPrimary).Shapes.Count To 1 Step -1
            If InStr(1, sec.Headers(wdHeaderFooterPrimary).Shapes(i).Name, "PowerPlusWaterMarkObject") > 0 Then
                sec.Headers(wdHeaderFooterPrimary).Shapes(i).Delete
            End If
        Next i
        
        ' Tao Shape dang WordArt/TextEffect lam watermark
        Set shp = sec.Headers(wdHeaderFooterPrimary).Shapes.AddTextEffect( _
            PresetTextEffect:=msoTextEffect1, _
            Text:=noiDungWM, _
            FontName:="Times New Roman", _
            FontSize:=1, _
            FontBold:=False, _
            FontItalic:=False, _
            Left:=0, _
            Top:=0)
        
        With shp
            .Name = "PowerPlusWaterMarkObject" & Format(Now, "hhmmss") & i
            
            ' Kich thuoc khung watermark (pt)
            .Width = CentimetersToPoints(20.29)   ' tuong duong 575.15pt
            .Height = CentimetersToPoints(4.51)   ' tuong duong 127.8pt
            
            ' Xoay 315 do
            .Rotation = 315
            
            ' Mau bac (silver), khong vien
            .Fill.ForeColor.RGB = RGB(192, 192, 192)   ' Silver
            .Fill.Transparency = 0.5                    ' Do mo 50%
            .Line.Visible = msoFalse
            
            ' Khong khoa ty le, cho phep di chuyen tu do trong khung
            .LockAspectRatio = msoFalse
            
            ' Canh giua trang theo chieu ngang va doc (dua theo margin)
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeCenter
            .Top = wdShapeCenter
            
            ' Dua watermark ra sau noi dung van ban
            .WrapFormat.AllowOverlap = True
            .ZOrder msoSendBehindText
            
            ' Cho phep watermark hien thi ca khi khong o che do Header
            .LayoutInCell = False
        End With
        
    Next sec
    
    MsgBox "Da tao watermark xong!", vbInformation, "Hoan tat"

End Sub
```

# 8. Page setup

```sh
Sub ThietLapPageSetup()

    Dim doc As Document
    
    Set doc = ActiveDocument
    
    With doc.PageSetup
        ' Can le
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(1.5)
        .LeftMargin = CentimetersToPoints(1.5)
        .RightMargin = CentimetersToPoints(1.5)
        
        ' Gutter
        .Gutter = CentimetersToPoints(0)
        .GutterPos = wdGutterPosLeft
        
        ' Huong trang: Portrait
        .Orientation = wdOrientPortrait
        
        ' Che do "Normal" (khong mirror margins, khong book fold)
        .MirrorMargins = False
        .BookFoldPrinting = False
    End With
    
    MsgBox "Da thiet lap Page Setup xong!", vbInformation, "Hoan tat"

End Sub
```

# 9. Số trang

```sh
Sub ChenSoTrang()

    Dim doc As Document
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim rng As Range
    
    Set doc = ActiveDocument
    
    For Each sec In doc.Sections
    
        ' Dat Footer from Bottom = 0.5 cm
        sec.PageSetup.FooterDistance = CentimetersToPoints(0.5)
        
        Set ftr = sec.Footers(wdHeaderFooterPrimary)
        
        ' Xoa noi dung cu trong Footer (neu co)
        ftr.Range.Delete
        
        Set rng = ftr.Range
        
        ' Chen field so trang
        rng.Fields.Add Range:=rng, Type:=wdFieldPage
        
        ' Can giua
        ftr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        ' Dinh dang font: Times New Roman, co 12
        ftr.Range.Font.Name = "Times New Roman"
        ftr.Range.Font.Size = 12
        ftr.Range.Font.Bold = False
        ftr.Range.Font.Italic = False
        
    Next sec
    
    MsgBox "Da chen so trang xong!", vbInformation, "Hoan tat"

End Sub
```

# 10. AutoFit Windows

```sh
Sub BatAutoFitWindow()

    Dim tbl As Table
    Dim demSoBang As Long
    
    demSoBang = 0
    
    For Each tbl In ActiveDocument.Tables
        tbl.AutoFitBehavior wdAutoFitWindow
        demSoBang = demSoBang + 1
    Next tbl
    
    MsgBox "Da hoan thanh!" & vbCrLf & _
           "Da bat AutoFit Window cho " & demSoBang & " bang da duyet.", _
           vbInformation, "Hoan tat"

End Sub
```

# 11. In đậm "Đáp số:"

```sh
Sub InDamDapSo()

    Dim rng As Range
    Dim demSoLan As Long
    Dim mauTim As String

    demSoLan = 0

    ' Dung ChrW de dung ky tu co dau, tranh loi hien thi dau trong VBA Editor
    mauTim = ChrW(272) & ChrW(225) & "p s" & ChrW(7889) & ":"

    Set rng = ActiveDocument.Content

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting

        .Text = mauTim
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
        .MatchCase = False   ' Khong phan biet in hoa/thuong

        Do While .Execute
            rng.Font.Bold = True
            demSoLan = demSoLan + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With

    MsgBox "Da hoan thanh!" & vbCrLf & _
           "Da in dam " & demSoLan & " cum 'Dap so:' da duyet.", _
           vbInformation, "Hoan tat"

End Sub
```


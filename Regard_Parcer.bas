Attribute VB_Name = "Regard"
Public data_arr
Public card_number%

Public Sub Regard_Parcer()
Dim url$

    Application.StatusBar = "Парсится Regard - 0%"

    url = "https://www.regard.ru/catalog/filter/?id=NDAzNzs2NywyMDM4NywyMTM3NSwyNDA1OSwyNDA2MCwyNDA2MSwyNDI0OSwyNDI1MC" & _
    "wyNDI2OSwyNDM1NiwyNDc0NiwyNDc4OCwyNTAwMSwyNTAwMiwyNTIzNSwyNTkwOSwyNzAyOCwyOTIwMw=="
    Get_Data (url)
    
    Лист5.Activate
    Лист5.Rows("3:1000").Delete
    Лист5.Range(Cells(2, 1), Cells(2, 7)).Clear
    Лист5.Range(Cells(2, 1), Cells(UBound(data_arr) + 2, 7)) = data_arr
    With Лист5.Range(Cells(2, 4), Cells(UBound(data_arr) + 2, 4))
        .NumberFormat = "0"
        .Value = .Value
    End With
    Application.StatusBar = False
End Sub

Function Get_Data(ByVal url$)
Dim html_arr$()
Dim cards_total%, page_count%, card_count%
Dim html_cut$
On Error Resume Next

    html_str = Get_HTML_TXT(url)
    html_cut = Mid(html_str, InStr(1, html_str, "Найдено:") + 9, 50)
    cards_total = Left(html_cut, InStr(1, html_cut, ")") - 1)   'Количество карт
    ReDim data_arr(cards_total - 1, 6)
    
    If InStr(1, html_str, "class=""left"">...<") <> 0 Then                'Количество страниц
        html_cut = Mid(html_str, InStr(1, html_str, "class=""left"">...<") + 25, 500)
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, ">"))
        page_count = Left(html_cut, InStr(1, html_cut, "<") - 1)
    Else
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "class=""curr"""))
        html_cut = Left(html_cut, InStr(1, html_cut, "class=""right"" id=""sel-cont"">"))
        html_arr = Split(html_cut, "href=")
        page_count = UBound(html_arr) + 1
    End If
    
    html_arr = Split(html_str, """block""")
    card_number = 0
    Processing_Data html_arr
    Application.StatusBar = "Парсится Regard - " & CInt(1 / (page_count) * 100) & "%"
    
    If page_count > 1 Then
        For page = 2 To page_count
            html_str = Get_HTML_TXT(url & "&page=" & page)
            html_arr = Split(html_str, """block""")
            Processing_Data html_arr
            Application.StatusBar = "Парсится Regard - " & CInt(page / (page_count) * 100) & "%"
        Next page
    End If
    
End Function

Function Get_HTML_TXT(ByVal url$) As String
On Error Resume Next
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", url, False
        '.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:101.0) Gecko/20100101 Firefox/101.0"
        '.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8"
        .send
        Do: DoEvents: Loop Until .readyState = 4
        Get_HTML_TXT = .responseText
    End With
End Function

Sub Processing_Data(ByVal html_arr)
Dim html_cut$, html_str$
Dim card_number_in_page%

    For card_number_in_page = 1 To UBound(html_arr):
        html_str = html_arr(card_number_in_page)
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "data-brand=") - 11)
        data_arr(card_number, 4) = Left(html_cut, InStr(1, html_cut, """") - 1)                                    'Vendor
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "alt=") - 4)
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        data_arr(card_number, 0) = Left(html_cut, InStr(1, html_cut, " ") - 1)                                     'GPU Manufacturer
        
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        data_arr(card_number, 1) = Left(html_cut, InStr(1, html_cut, " ") - 1)
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        data_arr(card_number, 1) = data_arr(card_number, 1) + " " & Left(html_cut, InStr(1, html_cut, " ") - 1)
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        If Left(html_cut, InStr(1, html_cut, " ") - 1) = "Super" Or _
        Left(html_cut, InStr(1, html_cut, " ") - 1) = "XT" Or _
        Left(html_cut, InStr(1, html_cut, " ") - 1) = "Ti" Then
            data_arr(card_number, 1) = data_arr(card_number, 1) + Left(html_cut, InStr(1, html_cut, " ") - 1)
        End If
        data_arr(card_number, 1) = GPU_Replacer(data_arr(card_number, 1))                                          'GPU
        
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        html_cut = Left(html_cut, InStr(1, html_cut, """") - 1)
        If (Left(html_cut, InStr(1, html_cut, " ") - 1)) = (data_arr(card_number, 4)) Then
            html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        End If
        data_arr(card_number, 5) = GPU_Replacer(html_cut)                                                          'Model
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "Gb ") + 3)
        html_cut = Trim(html_cut)
        data_arr(card_number, 2) = Left(html_cut, InStr(1, html_cut, "Gb") - 1) & " Gb"                            'Memory
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "data-price=") - 11)
        data_arr(card_number, 3) = Left(html_cut, InStr(1, html_cut, """") - 1)                                    'Price
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "href=") - 5)
        data_arr(card_number, 6) = "https://www.regard.ru" & Left(html_cut, InStr(1, html_cut, """") - 1)          'Link

        card_number = card_number + 1
    Next card_number_in_page
        
End Sub

Function GPU_Replacer(ByVal gpu As String) As String
    If InStr(1, gpu, "GeForce") <> 0 Then gpu = Replace(gpu, "GeForce ", "") _
    Else gpu = Replace(gpu, "Radeon ", "")
    If InStr(1, gpu, " SUPER") <> 0 Then gpu = Replace(gpu, " SUPER", "S")
    If InStr(1, gpu, "Super") <> 0 Then gpu = Replace(gpu, "Super", "S")
    If InStr(1, gpu, "SUPER") <> 0 Then gpu = Replace(gpu, "SUPER", "S")
    If InStr(1, gpu, " Super") <> 0 Then gpu = Replace(gpu, " Super", "S")
    If InStr(1, gpu, " Ti") <> 0 Then gpu = Replace(gpu, " Ti", "TI")
    If InStr(1, gpu, " XT") <> 0 Then gpu = Replace(gpu, " XT", "XT")
    GPU_Replacer = gpu
End Function

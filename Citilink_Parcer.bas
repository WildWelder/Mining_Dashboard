Attribute VB_Name = "Citilink"
Public data_arr
Public card_number%

Public Sub Citilink_Parcer()
Dim url$
    
    Application.StatusBar = "Парсится Citilink - 0%"
    
    url = "https://www.citilink.ru/catalog/videokarty/?f=discount.any%2Crating.any%" & _
    "2C9368_29amdd1d1radeond1rxd16600%" & _
    "2C9368_29amdd1d1radeond1rxd16600xt%" & _
    "2C9368_29amdd1d1radeond1rxd16700xt%" & _
    "2C9368_29amdd1d1radeond1rxd16800%" & _
    "2C9368_29amdd1d1radeond1rxd16800xt%" & _
    "2C9368_29amdd1d1radeond1rxd16900xt%" & _
    "2C9368_29nvidiad1d1geforced1gtxd11660super%" & _
    "2C9368_29nvidiad1d1geforced1rtxd12060super%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13060%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13060ti%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13070%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13070ti%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13080%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13080ti%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13090%" & _
    "2C9368_29nvidiad1d1geforced1rtxd13090ti"

    Get_Data (url)
    
    Лист4.Activate
    Лист4.Rows("3:1000").Delete
    Лист4.Range(Cells(2, 1), Cells(2, 7)).Clear
    Лист4.Range(Cells(2, 1), Cells(UBound(data_arr) + 2, 7)) = data_arr
    With Лист4.Range(Cells(2, 4), Cells(UBound(data_arr) + 2, 4))
        .NumberFormat = "0"
        .Value = .Value
    End With
    
End Sub

Function Get_Data(ByVal url$)
Dim html_arr$()
Dim cards_total%, page_count%, card_count%
Dim html_cut$, html_str$
On Error Resume Next

    html_str = Get_HTML_TXT(url)
    html_cut = Mid(html_str, InStr(1, html_str, "js--Subcategory__count") + 57, 100)
    cards_total = Left(html_cut, InStr(1, html_cut, " ") - 1)   'Количество карт
    ReDim data_arr(cards_total - 1, 6)
    
    If InStr(1, html_str, "_page_last") = 0 Then                'Количество страниц
        html_arr = Split(html_str, "_page_next")
        html_cut = html_arr(UBound(html_arr))
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "data-page=") - 10)
        page_count = Left(html_cut, InStr(1, html_cut, """") - 1)
    Else
        html_cut = Mid(html_str, InStr(1, html_str, "_page_last") + 82, 100)
        page_count = Left(html_cut, InStr(1, html_cut, """") - 1)
    End If
    
    html_arr = Split(html_str, "ProductCardInWishlist")
    card_number = 0
    Processing_Data html_arr
    Application.StatusBar = "Парсится Citilink - " & CInt(1 / (page_count) * 100) & "%"
    
    If page_count > 1 Then
        For page = 2 To page_count
            html_str = Get_HTML_TXT(url & "&p=" & page)
            html_arr = Split(html_str, "ProductCardInWishlist")
            Processing_Data html_arr
            Application.StatusBar = "Парсится Citilink - " & CInt(page / (page_count) * 100) & "%"
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
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "Видеочипсет") - 133)
        data_arr(card_number, 0) = Left(html_cut, InStr(1, html_cut, " ") - 1)                                     'GPU Manufacturer
        
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "  "))
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, " "))
        html_cut = Left(html_cut, InStr(1, html_cut, "  ") - 2)
        data_arr(card_number, 1) = GPU_Replacer(html_cut)                                                          'GPU
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "Объем видеопамяти") - 139)
        data_arr(card_number, 2) = Left(html_cut, InStr(1, html_cut, " ") - 1) & " Gb"                             'Memory
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "price") - 11)
        data_arr(card_number, 3) = Left(html_cut, InStr(1, html_cut, ",") - 1)                                     'Price
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "brandName") - 21)
        data_arr(card_number, 4) = Left(html_cut, InStr(1, html_cut, "&") - 1)                                     'Vendor
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "shortName"))
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, ", ") - 2)
        html_cut = Left(html_cut, InStr(1, html_cut, "&") - 1)
        data_arr(card_number, 5) = GPU_Replacer(html_cut)                                                          'Model
        
        html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "href=") - 5)
        data_arr(card_number, 6) = "https://www.citilink.ru" & Left(html_cut, InStr(1, html_cut, """") - 1)        'Link

        card_number = card_number + 1
    Next card_number_in_page
        
End Sub

Function GPU_Replacer(ByVal gpu As String) As String
    If InStr(1, gpu, "GeForce") <> 0 Then gpu = Replace(gpu, "GeForce ", "") _
    Else gpu = Replace(gpu, "Radeon ", "")
    If InStr(1, gpu, " SUPER") <> 0 Then gpu = Replace(gpu, " SUPER", "S")
    If InStr(1, gpu, "SUPER") <> 0 Then gpu = Replace(gpu, "SUPER", "S")
    If InStr(1, gpu, " Super") <> 0 Then gpu = Replace(gpu, " Super", "S")
    If InStr(1, gpu, " Ti") <> 0 Then gpu = Replace(gpu, " Ti", "TI")
    If InStr(1, gpu, " XT") <> 0 Then gpu = Replace(gpu, " XT", "XT")
    GPU_Replacer = gpu
End Function

Attribute VB_Name = "DNS"
Option Explicit
Public page_count%

Public Sub DNS_Parcer()
Dim links_arr, data_arr
Dim url$

    Application.StatusBar = "Парсится DNS - 0%"

    url = "https://www.dns-shop.ru/catalog/17a89aab16404e77/videokarty/" & _
    "?f[mv]=1dn4f0-udtje-13n3m1-udtf8-145iin-uiykt-v7hg2-zyyhm-1cwhi2-m2gz2-xv2j6-wq5qn-wq5qd-11xign-n20i4-170r3q-19tve8"
    links_arr = Get_Links(url)
    data_arr = Get_Data(links_arr)
    
    Лист3.Activate
    Лист3.Rows("3:1000").Delete
    Лист3.Range(Cells(2, 1), Cells(2, 7)).Clear
    Лист3.Range(Cells(2, 1), Cells(UBound(data_arr) + 2, 7)) = data_arr
    With Лист3.Range(Cells(2, 4), Cells(UBound(data_arr) + 2, 4))
        .NumberFormat = "0"
        .Value = .Value
    End With
    Application.StatusBar = False
End Sub

Function Get_HTML_TXT(ByVal url As String) As String
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

Function Get_Links(ByVal url As String) As Variant
Dim page%, link_number%, cards_count%, link_number_in_page%
Dim links_arr$(), html_arr$()
Dim html_str$, html_cut$, url_page$
    
    html_str = Get_HTML_TXT(url)
    
    html_cut = Right(html_str, Len(html_str) - InStr(1, html_str, "products-count") - 15)
    cards_count = Left(html_cut, InStr(1, html_cut, " ") - 1)          'общее количество карт
    ReDim links_arr(cards_count - 1)
    
    link_number = 0
    html_arr = Split(html_str, "_name")         'считывание ссылок из главной-первой страницы
    For link_number_in_page = 1 To UBound(html_arr):
        html_cut = Right(html_arr(link_number_in_page), Len(html_arr(link_number_in_page)) - InStr(1, html_arr(link_number_in_page), "=") - 1)
        links_arr(link_number) = "https://www.dns-shop.ru" & Left(html_cut, InStr(1, html_cut, ">") - 2)
        link_number = link_number + 1
    Next link_number_in_page
    
    html_cut = Mid(html_str, InStr(1, html_str, "page-link_last") - 200, 200)
    html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "data-page-number") - 19)
    page_count = Left(html_cut, InStr(1, html_cut, """") - 1)           'количество страниц
    Application.StatusBar = "Парсится DNS - " & CInt(1 / (cards_count + page_count) * 100) & "%"
    
    If page_count > 1 Then
        For page = 2 To page_count
            url_page = url & "&p=" & page
            html_str = Get_HTML_TXT(url_page)
            html_arr = Split(html_str, "_name")         'считывание ссылок на карты из страницы
            For link_number_in_page = 1 To UBound(html_arr):
                html_cut = Right(html_arr(link_number_in_page), Len(html_arr(link_number_in_page)) - InStr(1, html_arr(link_number_in_page), "=") - 1)
                links_arr(link_number) = "https://www.dns-shop.ru" & Left(html_cut, InStr(1, html_cut, ">") - 2)
                link_number = link_number + 1
            Next link_number_in_page
            Application.StatusBar = "Парсится DNS - " & CInt(page / (cards_count + page_count) * 100) & "%"
        Next page
    End If
    
    Get_Links = links_arr
    
End Function

Function Get_Data(ByVal links_arr As Variant) As Variant
Dim gpu_manufacturer$, gpu$, vendor$, price$, model$, html_cut$, html_str$
Dim data_arr$()
Dim link_number%
On Error Resume Next
ReDim data_arr(UBound(links_arr), 6)
    For link_number = 0 To UBound(links_arr):
        html_str = Get_HTML_TXT(links_arr(link_number))
        
        html_cut = Mid(html_str, InStr(1, html_str, "Микроархитектура") + 117, 100)
        data_arr(link_number, 0) = Left(html_cut, InStr(1, html_cut, " ") - 1)          'GPU Manufacturer
        
        html_cut = Mid(html_str, InStr(1, html_str, "Графический процессор") + 122, 200)
        gpu = Left(html_cut, InStr(1, html_cut, "<") - 1)
        data_arr(link_number, 1) = GPU_Replacer(gpu)                                    'GPU
        
        html_cut = Mid(html_str, InStr(html_str, "Объем видеопамяти") + 118, 200)
        data_arr(link_number, 2) = Left(html_cut, InStr(1, html_cut, "<") - 4) & " Gb"  'Memory
        
        html_cut = Right(html_str, Len(html_str) - InStr(html_str, "price"":") - 6)
        data_arr(link_number, 3) = Left(html_cut, InStr(1, html_cut, ",") - 1)          'Price
    
        html_cut = Mid(html_str, InStr(1, html_str, "Модель <") + 107, 200)
        data_arr(link_number, 4) = Left(html_cut, InStr(1, html_cut, " ") - 1)          'Vendor
        
        model = Mid(html_cut, InStr(1, html_cut, " ") + 1, InStr(1, html_cut, "<") - InStr(1, html_cut, " ") - 1)
        data_arr(link_number, 5) = GPU_Replacer(model)                                  'Model

        data_arr(link_number, 6) = links_arr(link_number)                               'Link
        
        Application.StatusBar = "Парсится DNS - " & CInt((page_count + link_number + 1) / (UBound(links_arr) + 1 + page_count) * 100) & "%"
    Next link_number
    
    Get_Data = data_arr
    
End Function

Function GPU_Replacer(ByVal gpu As String) As String
    If InStr(1, gpu, "GeForce") <> 0 Then gpu = Replace(gpu, "GeForce ", "") _
    Else gpu = Replace(gpu, "Radeon ", "")
    If InStr(1, gpu, " SUPER") <> 0 Then gpu = Replace(gpu, " SUPER", "S")
    If InStr(1, gpu, " Super") <> 0 Then gpu = Replace(gpu, " Super", "S")
    If InStr(1, gpu, " Ti") <> 0 Then gpu = Replace(gpu, " Ti", "TI")
    If InStr(1, gpu, " XT") <> 0 Then gpu = Replace(gpu, " XT", "XT")
    GPU_Replacer = gpu
End Function

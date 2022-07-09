Attribute VB_Name = "Command"
Sub Update_All()

Citilink_Parcer
DNS_Parcer
'OnlineTrade_Parcer
Regard_Parcer
WhatToMine_Parcer

End Sub

Sub Get_BTC_ExRate()
Dim html_str$
    html_str = Get_HTML_TXT("https://coinmarketcap.com/currencies/bitcoin/")
    html_str = Right(html_str, Len(html_str) - InStr(1, html_str, "priceValue ") - 19)
    BTC_ExRate = Format(CDbl(Replace(Left(html_str, InStr(1, html_str, "<") - 1), ",", "")), "#.0")
    Лист8.Cells(1, 6) = BTC_ExRate
    ActiveWorkbook.RefreshAll
End Sub

Sub Get_USD_ExRate()
Dim html_str$
    html_str = Get_HTML_TXT("https://ru.investing.com/currencies/usd-rub")
    html_str = Right(html_str, Len(html_str) - InStr(1, html_str, "price-last") - 11)
    USD_ExRate = Format(CDbl(Replace(Left(html_str, InStr(1, html_str, "<") - 1), ",", ".")), "#.0")
    Лист8.Cells(1, 4) = USD_ExRate
    ActiveWorkbook.RefreshAll
End Sub

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

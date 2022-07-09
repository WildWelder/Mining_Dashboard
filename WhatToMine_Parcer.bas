Attribute VB_Name = "WhatToMine"
Option Explicit
Public algo_data_arr$(), data_arr$()
Public BTC_ExRate#, USD_ExRate#
Public data_number%

Public Sub WhatToMine_Parcer()
Dim data_3090ti$(), data_3090$(), data_3080ti$(), data_3080$(), data_3070ti$(), data_3070$(), data_3060ti$(), data_3060$(), _
data_3050$(), data_2060s$(), data_1660s$(), data_6900xt$(), data_6800xt$(), data_6800$(), data_6700xt$(), data_6600xt$(), data_6600$(), hashrate_data$()
Dim nonempty_count%
Dim url$
    
    Application.StatusBar = "Парсится WhatToMine - 0%"
    url = "https://whattomine.com/coins.json?eth=true&factor[eth_hr]=100&e4g=true&factor[e4g_hr]=100" & _
    "&eqb=true&factor[eqb_hr]=100&al=true&factor[al_hr]=100&ops=true&factor[ops_hr]=100" & _
    "&zlh=true&factor[zlh_hr]=100&kpw=true&factor[kpw_hr]=100&sort=Profit24&revenue=24h"
    Get_Algo_Data (url)
    
    Application.StatusBar = "Парсится WhatToMine - 5%"
    Get_BTC_ExRate
    Get_USD_ExRate
    Application.StatusBar = "Парсится WhatToMine - 10%"
    
    data_3090ti = Get_Card_Data("https://whattomine.com/gpus/74-nvidia-geforce-rtx-3090-ti", "RTX 3090TI")
    Application.StatusBar = "Парсится WhatToMine - 15%"
    data_3090 = Get_Card_Data("https://whattomine.com/gpus/49-nvidia-geforce-rtx-3090", "RTX 3090")
    Application.StatusBar = "Парсится WhatToMine - 20%"
    data_3080ti = Get_Card_Data("https://whattomine.com/gpus/61-nvidia-geforce-rtx-3080-ti", "RTX 3080TI")
    Application.StatusBar = "Парсится WhatToMine - 25%"
    data_3080 = Get_Card_Data("https://whattomine.com/gpus/46-nvidia-geforce-rtx-3080", "RTX 3080")
    Application.StatusBar = "Парсится WhatToMine - 30%"
    data_3070ti = Get_Card_Data("https://whattomine.com/gpus/62-nvidia-geforce-rtx-3070-ti", "RTX 3070TI")
    Application.StatusBar = "Парсится WhatToMine - 35%"
    data_3070 = Get_Card_Data("https://whattomine.com/gpus/48-nvidia-geforce-rtx-3070", "RTX 3070")
    Application.StatusBar = "Парсится WhatToMine - 40%"
    data_3060ti = Get_Card_Data("https://whattomine.com/gpus/52-nvidia-geforce-rtx-3060-ti", "RTX 3060TI")
    Application.StatusBar = "Парсится WhatToMine - 45%"
    data_3060 = Get_Card_Data("https://whattomine.com/gpus/58-nvidia-geforce-rtx-3060", "RTX 3060")
    Application.StatusBar = "Парсится WhatToMine - 50%"
    data_3050 = Get_Card_Data("https://whattomine.com/gpus/73-nvidia-geforce-rtx-3050", "RTX 3050")
    Application.StatusBar = "Парсится WhatToMine - 55%"
    data_2060s = Get_Card_Data("https://whattomine.com/gpus/54-nvidia-geforce-rtx-2060-super", "RTX 2060S")
    Application.StatusBar = "Парсится WhatToMine - 60%"
    data_1660s = Get_Card_Data("https://whattomine.com/gpus/53-nvidia-geforce-gtx-1660-super", "GTX 1660S")
    Application.StatusBar = "Парсится WhatToMine - 65%"
    data_6900xt = Get_Card_Data("https://whattomine.com/gpus/57-amd-radeon-rx-6900-xt", "RX 6900XT")
    Application.StatusBar = "Парсится WhatToMine - 70%"
    data_6800xt = Get_Card_Data("https://whattomine.com/gpus/50-amd-radeon-rx-6800-xt", "RX 6800XT")
    Application.StatusBar = "Парсится WhatToMine - 75%"
    data_6800 = Get_Card_Data("https://whattomine.com/gpus/51-amd-radeon-rx-6800", "RX 6800")
    Application.StatusBar = "Парсится WhatToMine - 80%"
    data_6700xt = Get_Card_Data("https://whattomine.com/gpus/59-amd-radeon-rx-6700-xt", "RX 6700XT")
    Application.StatusBar = "Парсится WhatToMine - 85%"
    data_6600xt = Get_Card_Data("https://whattomine.com/gpus/67-amd-radeon-rx-6600-xt", "RX 6600XT")
    Application.StatusBar = "Парсится WhatToMine - 90%"
    data_6600 = Get_Card_Data("https://whattomine.com/gpus/68-amd-radeon-rx-6600", "RX 6600")
    Application.StatusBar = "Парсится WhatToMine - 95%"
    
    data_number = 0
    ReDim data_arr(17 * 7, 6)
    Data_Transfer (data_3090ti)
    Data_Transfer (data_3090)
    Data_Transfer (data_3080ti)
    Data_Transfer (data_3080)
    Data_Transfer (data_3070ti)
    Data_Transfer (data_3070)
    Data_Transfer (data_3060ti)
    Data_Transfer (data_3060)
    Data_Transfer (data_3050)
    Data_Transfer (data_2060s)
    Data_Transfer (data_1660s)
    Data_Transfer (data_6900xt)
    Data_Transfer (data_6800xt)
    Data_Transfer (data_6800)
    Data_Transfer (data_6700xt)
    Data_Transfer (data_6600xt)
    Data_Transfer (data_6600)
    
    hashrate_data = Data_Filter()
    
    Application.StatusBar = "Парсится WhatToMine - 100%"
    Лист1.Activate
    Лист1.Rows("3:1000").Delete
    Лист1.Range(Cells(2, 1), Cells(2, 5)).Clear
    Лист1.Range(Cells(2, 1), Cells(UBound(algo_data_arr) + 2, 5)) = algo_data_arr
    With Лист1.Range(Cells(2, 4), Cells(UBound(algo_data_arr) + 2, 5))
        .NumberFormat = "0.00"
        .Value = .Value
    End With
    Лист1.Range(Cells(2, 5), Cells(UBound(algo_data_arr) + 2, 5)).NumberFormat = "0.0000000"
    
    Лист2.Activate
    Лист2.Rows("3:1000").Delete
    Лист2.Range(Cells(2, 1), Cells(2, 9)).Clear
    Лист2.Range(Cells(2, 1), Cells(UBound(hashrate_data) + 2, 7)) = hashrate_data
    With Лист2.Range(Cells(2, 5), Cells(UBound(hashrate_data) + 2, 7))
        .NumberFormat = "0"
        .Value = .Value
    End With
    Лист2.Range(Cells(2, 7), Cells(UBound(hashrate_data) + 2, 7)).NumberFormat = "0.0000000"
    
    Лист8.Activate
    Worksheets("Profit").ListObjects(1).Refresh

    Application.StatusBar = False
End Sub

Function Get_Card_Data(ByVal url$, ByVal card_name$)
Dim html_arr$(), card_data_arr$()
Dim html_cut$, html_str$
Dim algo_number%

    html_str = Get_HTML_TXT(url)
    html_arr = Split(html_str, "class=""list-group-item")
    ReDim card_data_arr(UBound(html_arr) - 2, 3)
    
    For algo_number = 2 To UBound(html_arr):
    
        html_cut = Right(html_arr(algo_number), Len(html_arr(algo_number)) - InStr(1, html_arr(algo_number), ">") - 1)
        card_data_arr(algo_number - 2, 0) = card_name                                           'GPU
        
        If Left(html_cut, InStr(1, html_cut, "<") - 2) = "Equihash (150,5)" Then
            card_data_arr(algo_number - 2, 1) = "BeamHashIII"
        ElseIf Left(html_cut, InStr(1, html_cut, "<") - 2) = "Ethash4G" Then
            card_data_arr(algo_number - 2, 1) = "Etchash"
        Else
            card_data_arr(algo_number - 2, 1) = Left(html_cut, InStr(1, html_cut, "<") - 2)     'Algorythm
        End If
        
        If InStr(1, html_cut, ">Linux<") = 0 Then
            html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "text-end") - 10)
        Else
            html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "Linux") - 12)
        End If
            card_data_arr(algo_number - 2, 2) = Left(html_cut, InStr(1, html_cut, "/s") - 4)    'HashRate
            html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "@ ") - 1)
            card_data_arr(algo_number - 2, 3) = Left(html_cut, InStr(1, html_cut, "W") - 1)     'Power
        
    Next algo_number
    Get_Card_Data = card_data_arr
End Function

Function Data_Transfer(ByVal arr)
Dim arr_number%, algo_number%

    For algo_number = 0 To UBound(algo_data_arr)
        For arr_number = 0 To UBound(arr)
            If arr(arr_number, 1) = algo_data_arr(algo_number, 2) Then
                data_arr(data_number, 0) = arr(arr_number, 0)                                                'GPU
                data_arr(data_number, 1) = arr(arr_number, 1)                                                'Algorythm
                data_arr(data_number, 2) = algo_data_arr(algo_number, 0)                                     'Coin
                data_arr(data_number, 3) = algo_data_arr(algo_number, 1)                                     'Tag
                data_arr(data_number, 4) = arr(arr_number, 2)                                                'HashRate
                data_arr(data_number, 5) = arr(arr_number, 3)                                                'Power
                data_arr(data_number, 6) = algo_data_arr(algo_number, 4) / 100 * data_arr(data_number, 4)    'BTC Revenue
                data_number = data_number + 1
            End If
        Next arr_number
    Next algo_number
End Function

Function Data_Filter()
Dim temp_arr() As String
Dim temp_number%

    ReDim temp_arr(data_number - 1, 6)
    temp_number = 0
    For data_number = 0 To UBound(data_arr)
        If data_arr(data_number, 0) <> "" Then
            temp_arr(temp_number, 0) = data_arr(data_number, 0)                                              'GPU
            temp_arr(temp_number, 1) = data_arr(data_number, 1)                                              'Algorythm
            temp_arr(temp_number, 2) = data_arr(data_number, 2)                                              'Coin
            temp_arr(temp_number, 3) = data_arr(data_number, 3)                                              'Tag
            temp_arr(temp_number, 4) = data_arr(data_number, 4)                                              'HashRate
            temp_arr(temp_number, 5) = data_arr(data_number, 5)                                              'Power
            temp_arr(temp_number, 6) = data_arr(data_number, 6)                                              'BTC Revenue
            temp_number = temp_number + 1
        End If
    Next data_number
    Data_Filter = temp_arr
End Function

Function Get_Algo_Data(ByVal url$)
Dim html_str$, html_cut$
Dim html_arr$()
Dim algo_number%, algo_number_in_arr%

html_str = Get_HTML_TXT(url)
html_str = Right(html_str, Len(html_str) - 11)
html_arr = Split(html_str, "},""")
algo_number_in_arr = 0

ReDim algo_data_arr(6, 4)
For algo_number = 0 To UBound(html_arr)
    html_cut = Left(html_arr(algo_number), InStr(1, html_arr(algo_number), """") - 1)
    If html_cut = "Ethereum" Or html_cut = "Ravencoin" Or html_cut = "Beam" Or html_cut = "Conflux" _
    Or html_cut = "Ergo" Or html_cut = "Flux" Or html_cut = "EthereumClassic" Then
        algo_data_arr(algo_number_in_arr, 0) = html_cut                                                                                'Coin
        html_cut = Right(html_arr(algo_number), Len(html_arr(algo_number)) - InStr(1, html_arr(algo_number), "tag") - 5)
        algo_data_arr(algo_number_in_arr, 1) = Left(html_cut, InStr(1, html_cut, """") - 1)                                            'Tag
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "algorithm") - 11)
        algo_data_arr(algo_number_in_arr, 2) = Left(html_cut, InStr(1, html_cut, """") - 1)                                            'Algorithm
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "estimated_rewards24") - 21)
        algo_data_arr(algo_number_in_arr, 3) = Left(html_cut, InStr(1, html_cut, """") - 1)                                            'Est.Reward
        html_cut = Right(html_cut, Len(html_cut) - InStr(1, html_cut, "btc_revenue24") - 15)
        algo_data_arr(algo_number_in_arr, 4) = Left(html_cut, InStr(1, html_cut, """") - 1)                                            'BTC Revenue
        algo_number_in_arr = algo_number_in_arr + 1
    End If
Next algo_number

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

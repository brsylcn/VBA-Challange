Sub itWasDamnHard()
        
        'worksheet ac (I GOOGLED IT LOL)
        
        Dim ws As Worksheet
    

        'LOOP KULLAN
        
        For Each ws In Worksheets
    

            'BASLIK BELIRLEME
            'WS KOYARSAN HERYERE EKLIYOR
            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
    

            'TICKER SYMBOLE DIM AC
            
            Dim tickerSymbol As String
    

            'TOTALE DIM AC VE 0 A ESITLE KI LOOPA GIRINCE SONLANSIN
            
            Dim total_vol As Double
            total_vol = 0
    

            'TRACKER AC
            
            Dim counter As Long
            counter = 2
    

            'OPENA DA DIM AC VE 0 LA
            
            Dim yearOpen As Double
            yearOpen = 0
    

            'CLOSE DIM
            
            Dim yearClose As Double
            yearClose = 0
            
            'AYNI SEYI BURADA DA YAP
            
            Dim yearChange As Double
            yearChange = 0
    

            'AYNI SEY
            
            Dim percentChange As Double
            percentChange = 0
    

            'ROWLARA EKLEMEK ICIN LOOP AC
            
            Dim lastrow As Long
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

            'TICKER SEMBOL LOOPU
            
            For i = 2 To lastrow
                
                'OPEN YEAR PRICE GRAP ETMEK ICN
                
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    

                    yearOpen = ws.Cells(i, 3).Value
    

                End If
    

                'STOCK VOLUMELARIN HER BIR ROW ICIN HESAPLANMASI
                
                total_vol = total_vol + ws.Cells(i, 7)
    

                'TICKER DEGISIYOR MU DIYE KONDISYON EKLENMESI
                
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    

                    'TICKER SEMBOLUNUN MOVE EDILMESI
                    
                    ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
    

                    'OZET TABLOSUNA GONDERMEK
                    
                    ws.Cells(counter, 12).Value = total_vol
    

                    'YILLIK KAPANIS
                    
                    yearClose = ws.Cells(i, 6).Value
    

                    'DEGISIMIN HESAPLANMASI
                    
                    yearChange = yearClose - yearOpen
                    ws.Cells(counter, 10).Value = yearChange
    

                    'NEGATIV POZITIV DEGISIMLERIN BOYANMASI ICIN
                    
                    If yearChange >= 0 Then
                        ws.Cells(counter, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(counter, 10).Interior.ColorIndex = 3
                    End If
    

                    
                    'YUZDELIK DEGISIMIN HESAPLANMASI VE SUMMARY E GONDERILMESI
                    
                    If yearOpen = 0 And yearClose = 0 Then
                        
                        percentChange = 0
                        ws.Cells(counter, 11).Value = percentChange
                        ws.Cells(counter, 11).NumberFormat = "0.00%"
                    
                    ElseIf yearOpen = 0 Then
                      
                        
                        Dim percentNS As String
                        percentNS = "New Stock"
                        ws.Cells(counter, 11).Value = percentChange
                    
                    Else
                        
                        percentChange = yearChange / yearOpen
                        ws.Cells(counter, 11).Value = percentChange
                        ws.Cells(counter, 11).NumberFormat = "0.00%"
                    
                    End If
    

                    counter = counter + 1
    

                    totalVol = 0
                    yearOpen = 0
                    yearClose = 0
                    yearChange = 0
                    percentChange = 0
                    
                End If
            Next i
    

            'BEST VE WORSTLERIN BELIRLENMESI VE BASLIKLARIN EKLENMESI
            
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
    

            
            lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    

            'BEST VE WORSTLER ICIN DIMLERIN ACILMASI
            
            Dim bestStock As String
            Dim bestValue As Double
    

            'BEST ILK OLANA ESITLENIR
            
            bestValue = ws.Cells(2, 11).Value
    

            Dim worstStock As String
            Dim worstValue As Double
    

            'WORST ILK OLANA ESITLENIR
            
            worstValue = ws.Cells(2, 11).Value
    

            Dim mostVolStock As String
            Dim mostVolValue As Double
    

            'MOST U ILK OLANA ESITLE
            
            mostVolValue = ws.Cells(2, 12).Value
    

            'OZETTE BAKMAK ICIN LOOP AC
            
            For j = 2 To lastrow
    

                'BES I BELIRLEMEK ICIN CONDITION BELIRLE
                
                If ws.Cells(j, 11).Value > bestValue Then
                    
                    bestValue = ws.Cells(j, 11).Value
                    
                    bestStock = ws.Cells(j, 9).Value
                
                End If
    

                'WORST U BELIRLEMEK ICIN CONDITION BELIRLE
                
                If ws.Cells(j, 11).Value < worstValue Then
                    
                    worstValue = ws.Cells(j, 11).Value
                    
                    worstStock = ws.Cells(j, 9).Value
                
                End If
    

                'GREATEST ICIN CONDITION BELIRLE
                
                If ws.Cells(j, 12).Value > mostVolValue Then
                    
                    mostVolValue = ws.Cells(j, 12).Value
                    
                    mostVolStock = ws.Cells(j, 9).Value
                
                End If
    

            Next j
    

            'PERFORMANS TABLOSUNA EKLE HERSEYI
            
            ws.Cells(2, 16).Value = bestStock
            ws.Cells(2, 17).Value = bestValue
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 16).Value = worstStock
            ws.Cells(3, 17).Value = worstValue
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(4, 16).Value = mostVolStock
            ws.Cells(4, 17).Value = mostVolValue
    

            'AUTOFIT COLUMNS
            
            ws.Columns("I:L").EntireColumn.AutoFit
            ws.Columns("O:Q").EntireColumn.AutoFit
    

        Next ws
    

    End Sub




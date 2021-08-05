Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Module Module1
   Sub Main()
      Dim aExcelApp As New Excel.Application
      Dim aExcelWrkbook As Excel.Workbook = aExcelApp.Workbooks.Add
      For Each brand As String In {"WTB", "TLD", "TIM", "SHI", "SET", "SEL", "ROC", "RAC", "PNT", "MIN", "MAX", "KEN", "FSA", "BLACKSPIRE", "GROUP4", "GROUP3", "GROUP2", "GROUP1"}
         'For Each brand As String In {"BEL", "GIR", "EAS"}
         Console.WriteLine(brand)
         Dim brandItems As New List(Of ItemExtraInfo)
         brandItems = GetVendorItems(brand)
         Dim c As Integer = 1
         Dim aExcelWrkSheet As Excel.Worksheet = aExcelWrkbook.Worksheets.Add
         With aExcelWrkSheet
            .Name = brand
            .Range("A1").Value = "Item Number"
            .Range("B1").Value = "Vendor Item Number"
            .Range("C1").Value = "Description"
            .Range("D1").Value = "Status"
            .Range("E1").Value = "Inventory"
            .Range("F1").Value = "Wharehouse Location"
            .Range("G1").Value = "Weeks of Stock"
            .Range("H1").Value = "Price"
            .Range("I1").Value = "MSRP"
            .Range("J1").Value = "Cost"
            .Range("K1").Value = "Vendor"
            .Range("L1").Value = "Lowest Cost"
            .Range("M1").Value = "Lowest Cost Vendor"
            .Range("N1").Value = "PO"
            .Range("O1").Value = "DUE DATE"
            .Range("P1").Value = "QTY DUE"
            .Range("Q1").Value = "Year Sales"
            .Range("R1").Value = "Year Sales Qty"
            .Range("S1").Value = "Six Month Sales"
            .Range("T1").Value = "Six Month Sales Qty"
            .Range("U1").Value = "One Month Sales"
            .Range("V1").Value = "One Month Sales Qty"
            .Range("W1").Value = "Recommended Buy"
            .Range("X1").Value = "Order Qty"
            .Range("X1").ColumnWidth = 12
            .Range("A1:X1").Font.Bold = True
            .Range("A1:X1").Interior.ColorIndex = 34
            .Range("A1:X1").Interior.Pattern = Excel.XlPattern.xlPatternSolid
            For Each i As ItemExtraInfo In brandItems
               c += 1
               .Range("A" & c).Value = i.itemNo
               .Range("B" & c).Value = i.vendorItemNo
               .Range("C" & c).Value = i.itemDesc
               .Range("D" & c).Value = i.status
               .Range("E" & c).Value = i.InStock
               .Range("F" & c).Value = i.WhseLoc
               .Range("G" & c).Value = i.WeeskLeftWithCommingPOAndStock.ToString("f2")
               .Range("H" & c).Value = i.price.ToString("f2")
               .Range("I" & c).Value = i.msrp.ToString("f2")
               .Range("J" & c).Value = i.cost.ToString("f2")
               .Range("K" & c).Value = i.vendor
               .Range("L" & c).Value = i.lowestcost.ToString("f2")
               .Range("M" & c).Value = i.lowestcostvendor
               .Range("N" & c).Value = i.PONumbers
               .Range("O" & c).Value = i.PODueDates
               .Range("P" & c).Value = i.TotalQtyDue
               .Range("Q" & c).Value = i.yearsales.ToString("f2")
               .Range("R" & c).Value = i.yearsalesnumber
               .Range("S" & c).Value = i.sixmothsales.ToString("f2")
               .Range("T" & c).Value = i.sixmothsalesnumber
               .Range("U" & c).Value = i.monthsales.ToString("f2")
               .Range("V" & c).Value = i.monthsalesnumber
               .Range("W" & c).Value = i.RecommendedBuy
            Next
            With .Range("A1:X" & c)
               'Fix sort after adding columns
               .Sort(.Range("W2"), Excel.XlSortOrder.xlDescending, , , , , , Excel.XlYesNoGuess.xlGuess, 1, False, Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, , )
            End With
            .Range("A1").EntireColumn.AutoFit()
            .Range("B1").EntireColumn.AutoFit()
            .Range("C1").EntireColumn.AutoFit()
            .Range("D1").EntireColumn.AutoFit()
            .Range("E1").EntireColumn.AutoFit()
            .Range("F1").EntireColumn.AutoFit()
            .Range("G1").EntireColumn.AutoFit()
            .Range("H1").EntireColumn.AutoFit()
            .Range("I1").EntireColumn.AutoFit()
            .Range("J1").EntireColumn.AutoFit()
            .Range("K1").EntireColumn.AutoFit()
            .Range("L1").EntireColumn.AutoFit()
            .Range("M1").EntireColumn.AutoFit()
            .Range("N1").EntireColumn.AutoFit()
            .Range("O1").EntireColumn.AutoFit()
            .Range("P1").EntireColumn.AutoFit()
            .Range("Q1").EntireColumn.AutoFit()
            .Range("R1").EntireColumn.AutoFit()
            .Range("S1").EntireColumn.AutoFit()
            .Range("T1").EntireColumn.AutoFit()
            .Range("U1").EntireColumn.AutoFit()
            .Range("V1").EntireColumn.AutoFit()
            .Range("W1").EntireColumn.AutoFit()
            .Range("A2").Select()
         End With
         aExcelApp.ActiveWindow.FreezePanes = True
      Next
      aExcelWrkbook.Sheets.Item("Sheet1").delete()
      aExcelWrkbook.Sheets.Item("Sheet2").delete()
      aExcelWrkbook.Sheets.Item("Sheet3").delete()
      Dim tmpFilePath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\PREORDERS_REPORT." & Now.ToString("yyyyMMdd") & ".xls"
      If File.Exists(tmpFilePath) Then File.Move(tmpFilePath, tmpFilePath.Replace(".xls", "." & Now.Ticks & ".xls"))
      aExcelWrkbook.SaveAs(tmpFilePath)
      aExcelWrkbook.Close()
      aExcelApp.Quit()
      System.Runtime.InteropServices.Marshal.ReleaseComObject(aExcelWrkbook)
      System.Runtime.InteropServices.Marshal.ReleaseComObject(aExcelApp)
      aExcelWrkbook = Nothing
      aExcelApp = Nothing
      GC.Collect()
      Dim Message As New MailMessage
      With Message
            .From = New MailAddress(My.Computer.Name & "@ecommerce.com")
            .Subject = "Preorders Report"
            .To.Add("donovan@ecommerce.com")
            .To.Add("tony@ecommerce.com")
            .To.Add("paul@ecommerce.com")
            .CC.Add("federico@ecommerce.com")
            .Body = "Report Attached"
         .Attachments.Add(New Attachment(tmpFilePath))
      End With
      Dim SMTPClient As New SmtpClient("SMTP")
      Try
         SMTPClient.Send(Message)
      Catch ex As Exception
         Console.WriteLine(ex.Message)
      End Try
   End Sub
   Function GetVendorItems(ByVal VendorID As String) As List(Of ItemExtraInfo)
      Dim tmpResult As New List(Of ItemExtraInfo)
      Dim NewQueryString As String = ""
      Select Case VendorID
         Case "BLACKSPIRE", "BLACKBURN"
            NewQueryString = QueryStringBLA(VendorID)
         Case "GROUP1"
            NewQueryString = QueryStringGroup("'HAY','SUN','ANS','MAN'")
         Case "GROUP2"
            NewQueryString = QueryStringGroup("'AVD','TRU','SRA'")
         Case "GROUP3"
            NewQueryString = QueryStringGroup("'AZO','CAC','LIZ'")
         Case "GROUP4"
            NewQueryString = QueryStringGroup("'BEL','GIR','EAS'")
         Case Else
            NewQueryString = QueryString(VendorID)
      End Select
        Using conn As New SqlConnection("Data Source=ECOM-DB1;Initial Catalog=ECOMLIVE;UID=ssss;PWD=ssssss")
            Dim cmd As New SqlCommand(NewQueryString, conn)
            Try
                conn.Open()
                Dim r As SqlDataReader = cmd.ExecuteReader
                If r.HasRows Then
                    Do While r.Read
                        Dim avg52 As Decimal = 0
                        Dim avg26 As Decimal = 0
                        Dim avg13 As Decimal = 0
                        Dim avg8 As Decimal = 0
                        Dim avg4 As Decimal = 0
                        Dim lastweek As Decimal = 0
                        If Not IsDBNull(r.Item("AVG52WKS")) Then
                            avg52 = r.Item("AVG52WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG26WKS")) Then
                            avg26 = r.Item("AVG26WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG13WKS")) Then
                            avg13 = r.Item("AVG13WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG8WKS")) Then
                            avg8 = r.Item("AVG8WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG4WKS")) Then
                            avg4 = r.Item("AVG4WKS")
                        End If
                        If Not IsDBNull(r.Item("LASTWK")) Then
                            lastweek = r.Item("LASTWK")
                        End If
                        Dim higestavg As Decimal = 0
                        If avg52 > avg26 And avg52 > avg13 And avg52 > avg8 And avg52 > avg4 And avg52 > lastweek Then
                            higestavg = avg52
                        ElseIf avg26 > avg13 And avg26 > avg8 And avg26 > avg4 And avg26 > lastweek Then
                            higestavg = avg26
                        ElseIf avg13 > avg8 And avg13 > avg4 And avg13 > lastweek Then
                            higestavg = avg13
                        ElseIf avg8 > avg4 And avg8 > lastweek Then
                            higestavg = avg8
                        ElseIf avg4 > lastweek Then
                            higestavg = avg4
                        Else
                            higestavg = lastweek
                        End If
                        Dim itemsneededtolast6months As Decimal = higestavg * 26
                        Dim itemsInStock As Integer = r.Item("AVAILABLEINV")
                        Dim itemsarrivingsoon As Integer = NumberOFItemsArrivingWithin6Months(r)
                        Dim i As New ItemExtraInfo
                        With i
                            .edpno = r.Item("EDPNO")
                            .itemNo = Trim(r.Item("ITEMNO"))
                            .itemDesc = Trim(r.Item("DESCRIPTION"))
                            .status = Trim(r.Item("STATUS"))
                            If Not IsDBNull(r.Item("VENDORNO")) Then
                                .vendor = Trim(r.Item("VENDORNO"))
                            Else
                                .vendor = ""
                            End If
                            If Not IsDBNull(r.Item("VENDORITEMNO")) Then
                                .vendorItemNo = Trim(r.Item("VENDORITEMNO"))
                            Else
                                .vendorItemNo = ""
                            End If
                            If Not IsDBNull(r.Item("DEFAULT_COST")) Then
                                .cost = r.Item("DEFAULT_COST")
                            Else
                                .cost = 0
                            End If

                            If Not IsDBNull(r.Item("PRICE")) Then
                                .price = r.Item("PRICE")
                            Else
                                .price = 0
                            End If
                            If Not IsDBNull(r.Item("MSRP")) Then
                                .msrp = r.Item("MSRP")
                            Else
                                .msrp = 0
                            End If
                            .InStock = itemsInStock
                            .yearsales = r.Item("YEARSALES")
                            .yearsalesnumber = r.Item("YEARSALESNUMBER")
                            .sixmothsales = r.Item("SIXMOTHSALES")
                            .sixmothsalesnumber = r.Item("SIXMOTHSALESNUMBER")
                            .monthsales = r.Item("MOTHSALES")
                            .monthsalesnumber = r.Item("MOTHSALESNUMBER")
                            .WhseLoc = Trim(r.Item("WAREHOUSELOCS_001"))
                            .ExpectedWithin6Months = itemsarrivingsoon
                            .TotalExpected = r.Item("TOTALONORDER")
                            .PONumber1 = r.Item("PONUMBERS_001")
                            .PODueDate1 = r.Item("EXPECTEDDATE_001")
                            .POQty1 = r.Item("NEXTQTY_001")
                            .PONumber2 = r.Item("PONUMBERS_002")
                            .PODueDate2 = r.Item("EXPECTEDDATE_002")
                            .POQty2 = r.Item("NEXTQTY_002")
                            .PONumber3 = r.Item("PONUMBERS_003")
                            .PODueDate3 = r.Item("EXPECTEDDATE_003")
                            .POQty3 = r.Item("NEXTQTY_003")
                            .PONumber4 = r.Item("PONUMBERS_004")
                            .PODueDate4 = r.Item("EXPECTEDDATE_004")
                            .POQty4 = r.Item("NEXTQTY_004")
                            .PONumber5 = r.Item("PONUMBERS_005")
                            .PODueDate5 = r.Item("EXPECTEDDATE_005")
                            .POQty5 = r.Item("NEXTQTY_005")
                            .HighestAvgWk = higestavg
                            .SixMonthNeed = itemsneededtolast6months
                            .SetLowestCostVendor()
                        End With
                        tmpResult.Add(i)
                    Loop
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
        Return tmpResult
   End Function

   Function GetVendorItemsOLD(ByVal VendorID As String) As List(Of ItemExtraInfo)
      Dim tmpResult As New List(Of ItemExtraInfo)
      Dim NewQueryString As String = ""
      Select Case VendorID
         Case "BLACKSPIRE", "BLACKBURN"
            NewQueryString = QueryStringBLA(VendorID)
         Case Else
            NewQueryString = QueryString(VendorID)
      End Select
        Using conn As New SqlConnection("Data Source=ECOM-DB1;Initial Catalog=ECOMLIVE;UID=ssss;PWD=ssssss")
            Dim cmd As New SqlCommand(NewQueryString, conn)
            Try
                conn.Open()
                Dim r As SqlDataReader = cmd.ExecuteReader
                If r.HasRows Then
                    'Console.WriteLine(r.RecordsAffected)
                    Do While r.Read
                        Dim avg52 As Decimal = 0
                        Dim avg26 As Decimal = 0
                        Dim avg13 As Decimal = 0
                        Dim avg8 As Decimal = 0
                        Dim avg4 As Decimal = 0
                        Dim lastweek As Decimal = 0
                        If Not IsDBNull(r.Item("AVG52WKS")) Then
                            avg52 = r.Item("AVG52WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG26WKS")) Then
                            avg26 = r.Item("AVG26WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG13WKS")) Then
                            avg13 = r.Item("AVG13WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG8WKS")) Then
                            avg8 = r.Item("AVG8WKS")
                        End If
                        If Not IsDBNull(r.Item("AVG4WKS")) Then
                            avg4 = r.Item("AVG4WKS")
                        End If
                        If Not IsDBNull(r.Item("LASTWK")) Then
                            lastweek = r.Item("LASTWK")
                        End If
                        Dim higestavg As Decimal = 0
                        If avg52 > avg26 And avg52 > avg13 And avg52 > avg8 And avg52 > avg4 And avg52 > lastweek Then
                            higestavg = avg52
                        ElseIf avg26 > avg13 And avg26 > avg8 And avg26 > avg4 And avg26 > lastweek Then
                            higestavg = avg26
                        ElseIf avg13 > avg8 And avg13 > avg4 And avg13 > lastweek Then
                            higestavg = avg13
                        ElseIf avg8 > avg4 And avg8 > lastweek Then
                            higestavg = avg8
                        ElseIf avg4 > lastweek Then
                            higestavg = avg4
                        Else
                            higestavg = lastweek
                        End If
                        Dim itemsneededtolast6months As Decimal = higestavg * 26
                        Dim itemsInStock As Integer = r.Item("AVAILABLEINV")
                        Dim itemsarrivingsoon As Integer = NumberOFItemsArrivingWithin6Months(r)
                        Dim i As New ItemExtraInfo
                        With i
                            .itemNo = Trim(r.Item("ITEMNO"))
                            .itemDesc = Trim(r.Item("DESCRIPTION"))
                            .status = Trim(r.Item("STATUS"))
                            If Not IsDBNull(r.Item("DEFAULT_COST")) Then
                                .cost = r.Item("DEFAULT_COST")
                            Else
                                .cost = 0
                            End If
                            '.msrp = r.Item("MSRP")
                            If Not IsDBNull(r.Item("PRICE")) Then
                                .price = r.Item("PRICE")
                            Else
                                .price = 0
                            End If
                            '.buyercode = r.Item("BUYERCODE")
                            .InStock = itemsInStock
                            .ExpectedWithin6Months = itemsarrivingsoon
                            .TotalExpected = r.Item("TOTALONORDER")
                            .PONumber1 = r.Item("PONUMBERS_001")
                            .PODueDate1 = r.Item("EXPECTEDDATE_001")
                            .POQty1 = r.Item("NEXTQTY_001")
                            .PONumber2 = r.Item("PONUMBERS_002")
                            .PODueDate2 = r.Item("EXPECTEDDATE_002")
                            .POQty2 = r.Item("NEXTQTY_002")
                            .PONumber3 = r.Item("PONUMBERS_003")
                            .PODueDate3 = r.Item("EXPECTEDDATE_003")
                            .POQty3 = r.Item("NEXTQTY_003")
                            .PONumber4 = r.Item("PONUMBERS_004")
                            .PODueDate4 = r.Item("EXPECTEDDATE_004")
                            .POQty4 = r.Item("NEXTQTY_004")
                            .PONumber5 = r.Item("PONUMBERS_005")
                            .PODueDate5 = r.Item("EXPECTEDDATE_005")
                            .POQty5 = r.Item("NEXTQTY_005")
                            .HighestAvgWk = higestavg
                            .SixMonthNeed = itemsneededtolast6months
                            'Console.WriteLine(.itemNo)
                        End With
                        tmpResult.Add(i)
                    Loop
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Using
        Return tmpResult
   End Function
   Function NumberOFItemsArrivingWithin6Months(ByVal r As SqlDataReader) As Integer
      Dim tmpResult As Integer = 0
      For i As Integer = 1 To 5
         Dim tmpdate As Integer = r.Item("EXPECTEDDATE_00" & i)
         If tmpdate > 0 And tmpdate < Convert.ToInt32(Now.AddMonths(6).ToString("yyyyMMdd")) Then
            tmpResult += r.Item("NEXTQTY_00" & i)
         End If
      Next
      Return tmpResult
   End Function
   Function QueryString(ByVal VendorID As String) As String
      Return "SELECT EDPNO,ITEMNO,DESCRIPTION,STATUS,AVAILABLEINV,WAREHOUSELOCS_001,VENDORNO,CAST(PRICE AS MONEY) / 100 AS PRICE," & _
"DEFAULT_COST = CASE WHEN (SELECT COUNT(*) FROM VENDORITEMS V WHERE PREFERENCE = '00' AND V.EDPNO = I.EDPNO) = 1 " & _
    "THEN (SELECT TOP 1 CAST(DOLLARCOST AS MONEY) / 10000 FROM VENDORITEMS V WHERE PREFERENCE = '00' AND V.EDPNO = I.EDPNO) " & _
    "ELSE (SELECT TOP 1 CAST(DOLLARCOST AS MONEY) / 10000 FROM VENDORITEMS V WHERE PREFERENCE = '01' AND V.EDPNO = I.EDPNO) END," & _
"VENDORITEMNO = (SELECT MAX(ITEMNO) FROM VENDORITEMS WHERE VENDORNO = I.VENDORNO AND EDPNO = I.EDPNO )," & _
"CAST(CAST(CAST((SUBSTRING(FLAGS,37,4))AS VARBINARY)AS BIGINT)AS MONEY)/100 AS MSRP," & _
"(NEXTQTY_001+NEXTQTY_002+NEXTQTY_003+NEXTQTY_004+NEXTQTY_005) AS TOTALONORDER," & _
"PONUMBERS_001,PONUMBERS_002,PONUMBERS_003,PONUMBERS_004,PONUMBERS_005,EXPECTEDDATE_001,EXPECTEDDATE_002," & _
"EXPECTEDDATE_003,EXPECTEDDATE_004,EXPECTEDDATE_005,NEXTQTY_001,NEXTQTY_002,NEXTQTY_003,NEXTQTY_004,NEXTQTY_005," & _
"AVG52WKS = (SELECT SUM(ITEMQTYS) / 52 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365)," & _
"AVG26WKS = (SELECT SUM(ITEMQTYS) / 26 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182)," & _
"AVG13WKS = (SELECT SUM(ITEMQTYS) / 13 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -91)," & _
"AVG8WKS = (SELECT SUM(ITEMQTYS) / 8 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -56)," & _
"AVG4WKS = (SELECT SUM(ITEMQTYS) / 4 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -28)," & _
"LASTWK = (SELECT SUM(ITEMQTYS) FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -7), " & _
"YEARSALES = ISNULL((SELECT SUM(EXTPRICES) FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365),0), " & _
"YEARSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365 ),0), " & _
"SIXMOTHSALES = ISNULL((SELECT SUM(EXTPRICES) FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182),0), " & _
"SIXMOTHSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182 ),0), " & _
"MOTHSALES = ISNULL((SELECT SUM(EXTPRICES)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -30 ),0), " & _
"MOTHSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -30 ),0) " & _
"FROM ITEMMAST I WHERE SUBSTRING(UPPER(ITEMNO),5,3) = '" & VendorID & "' AND UPPER(I.STATUS) IN ('R1','A1') ORDER BY ITEMNO"
   End Function
   Function QueryStringGroup(ByVal Group As String) As String
      Return "SELECT EDPNO,ITEMNO,DESCRIPTION,STATUS,AVAILABLEINV,WAREHOUSELOCS_001,VENDORNO,CAST(PRICE AS MONEY) / 100 AS PRICE," & _
"DEFAULT_COST = CASE WHEN (SELECT COUNT(*) FROM VENDORITEMS V WHERE PREFERENCE = '00' AND V.EDPNO = I.EDPNO) = 1 " & _
    "THEN (SELECT TOP 1 CAST(DOLLARCOST AS MONEY) / 10000 FROM VENDORITEMS V WHERE PREFERENCE = '00' AND V.EDPNO = I.EDPNO) " & _
    "ELSE (SELECT TOP 1 CAST(DOLLARCOST AS MONEY) / 10000 FROM VENDORITEMS V WHERE PREFERENCE = '01' AND V.EDPNO = I.EDPNO) END," & _
"VENDORITEMNO = (SELECT MAX(ITEMNO) FROM VENDORITEMS WHERE VENDORNO = I.VENDORNO AND EDPNO = I.EDPNO )," & _
"CAST(CAST(CAST((SUBSTRING(FLAGS,37,4))AS VARBINARY)AS BIGINT)AS MONEY)/100 AS MSRP," & _
"(NEXTQTY_001+NEXTQTY_002+NEXTQTY_003+NEXTQTY_004+NEXTQTY_005) AS TOTALONORDER," & _
"PONUMBERS_001,PONUMBERS_002,PONUMBERS_003,PONUMBERS_004,PONUMBERS_005,EXPECTEDDATE_001,EXPECTEDDATE_002," & _
"EXPECTEDDATE_003,EXPECTEDDATE_004,EXPECTEDDATE_005,NEXTQTY_001,NEXTQTY_002,NEXTQTY_003,NEXTQTY_004,NEXTQTY_005," & _
"AVG52WKS = (SELECT SUM(ITEMQTYS) / 52 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365)," & _
"AVG26WKS = (SELECT SUM(ITEMQTYS) / 26 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182)," & _
"AVG13WKS = (SELECT SUM(ITEMQTYS) / 13 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -91)," & _
"AVG8WKS = (SELECT SUM(ITEMQTYS) / 8 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -56)," & _
"AVG4WKS = (SELECT SUM(ITEMQTYS) / 4 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -28)," & _
"LASTWK = (SELECT SUM(ITEMQTYS) FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -7), " & _
"YEARSALES = ISNULL((SELECT SUM(EXTPRICES) FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365),0), " & _
"YEARSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365 ),0), " & _
"SIXMOTHSALES = ISNULL((SELECT SUM(EXTPRICES) FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182),0), " & _
"SIXMOTHSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182 ),0), " & _
"MOTHSALES = ISNULL((SELECT SUM(EXTPRICES)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -30 ),0), " & _
"MOTHSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -30 ),0) " & _
"FROM ITEMMAST I WHERE SUBSTRING(UPPER(ITEMNO),5,3) IN (" & Group & ") AND UPPER(I.STATUS) IN ('R1','A1') ORDER BY ITEMNO"
   End Function
   Function QueryStringBLA(ByVal VendorID As String) As String
      Return "SELECT EDPNO,ITEMNO,DESCRIPTION,STATUS,AVAILABLEINV,WAREHOUSELOCS_001,VENDORNO,CAST(PRICE AS MONEY) / 100 AS PRICE," & _
"DEFAULT_COST = CASE WHEN (SELECT COUNT(*) FROM VENDORITEMS V WHERE PREFERENCE = '00' AND V.EDPNO = I.EDPNO) = 1 " & _
    "THEN (SELECT TOP 1 CAST(DOLLARCOST AS MONEY) / 10000 FROM VENDORITEMS V WHERE PREFERENCE = '00' AND V.EDPNO = I.EDPNO) " & _
    "ELSE (SELECT TOP 1 CAST(DOLLARCOST AS MONEY) / 10000 FROM VENDORITEMS V WHERE PREFERENCE = '01' AND V.EDPNO = I.EDPNO) END," & _
"VENDORITEMNO = (SELECT MAX(ITEMNO) FROM VENDORITEMS WHERE VENDORNO = I.VENDORNO AND EDPNO = I.EDPNO )," & _
"CAST(CAST(CAST((SUBSTRING(FLAGS,37,4))AS VARBINARY)AS BIGINT)AS MONEY)/100 AS MSRP," & _
"(NEXTQTY_001+NEXTQTY_002+NEXTQTY_003+NEXTQTY_004+NEXTQTY_005) AS TOTALONORDER," & _
"PONUMBERS_001,PONUMBERS_002,PONUMBERS_003,PONUMBERS_004,PONUMBERS_005,EXPECTEDDATE_001,EXPECTEDDATE_002," & _
"EXPECTEDDATE_003,EXPECTEDDATE_004,EXPECTEDDATE_005,NEXTQTY_001,NEXTQTY_002,NEXTQTY_003,NEXTQTY_004,NEXTQTY_005," & _
"AVG52WKS = (SELECT SUM(ITEMQTYS) / 52 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365)," & _
"AVG26WKS = (SELECT SUM(ITEMQTYS) / 26 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182)," & _
"AVG13WKS = (SELECT SUM(ITEMQTYS) / 13 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -91)," & _
"AVG8WKS = (SELECT SUM(ITEMQTYS) / 8 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -56)," & _
"AVG4WKS = (SELECT SUM(ITEMQTYS) / 4 FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -28)," & _
"LASTWK = (SELECT SUM(ITEMQTYS) FROM VW_ORDERSUBHEAD WHERE EDPNO = I.EDPNO AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -7), " & _
"YEARSALES = ISNULL((SELECT SUM(EXTPRICES) FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365),0), " & _
"YEARSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -365 ),0), " & _
"SIXMOTHSALES = ISNULL((SELECT SUM(EXTPRICES) FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182),0), " & _
"SIXMOTHSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -182 ),0), " & _
"MOTHSALES = ISNULL((SELECT SUM(EXTPRICES)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -30 ),0), " & _
"MOTHSALESNUMBER = ISNULL((SELECT SUM(ITEMQTYS)FROM VW_ORDERSUBHEAD W WHERE W.EDPNO = I.EDPNO AND BS = 'S' AND " & _
"DATEDIFF(d,LEFT(Convert(Varchar, GetDate(), 120),10), STUFF(STUFF(SHIPDATE,5,0,'-'),8,0,'-')) >= -30 ),0) " & _
"FROM ITEMMAST I WHERE UPPER(DESCRIPTION) LIKE '%" & VendorID & "%' AND UPPER(I.STATUS) IN ('R1','A1') ORDER BY ITEMNO"
   End Function
   Structure ItemExtraInfo
      Dim edpno As Integer
      Dim itemNo As String
      Dim itemDesc As String
      Dim status As String
      Dim vendor As String
      Dim vendorItemNo As String
      Dim cost As Decimal
      Dim lowestcostvendor As String
      Dim lowestcost As Decimal
      Dim msrp As Decimal
      Dim price As Decimal
      Dim buyercode As String
      Dim InStock As Integer
      Dim WhseLoc As String
      Dim ExpectedWithin6Months As Integer
      Dim TotalExpected As Integer
      Dim PONumber1 As String
      Dim PODueDate1 As String
      Dim POQty1 As String
      Dim PONumber2 As String
      Dim PODueDate2 As String
      Dim POQty2 As String
      Dim PONumber3 As String
      Dim PODueDate3 As String
      Dim POQty3 As String
      Dim PONumber4 As String
      Dim PODueDate4 As String
      Dim POQty4 As String
      Dim PONumber5 As String
      Dim PODueDate5 As String
      Dim POQty5 As String
      Dim yearsales As Decimal
      Dim yearsalesnumber As Decimal
      Dim sixmothsales As Decimal
      Dim sixmothsalesnumber As Decimal
      Dim monthsales As Decimal
      Dim monthsalesnumber As Decimal
      Dim HighestAvgWk As Decimal
      Dim SixMonthNeed As Decimal
      Function TotalQtyDue() As String
         If TotalExpected > 0 Then
            Return TotalExpected
         Else
            Return ""
         End If
      End Function
      Function Margin() As Decimal
         If price > 0 Then
            Return ((price - cost) / price) * 100
         Else
            Return 0
         End If
      End Function
      Function PONumbers() As String
         Dim tmpResult As String = ""
         If PONumber1.Trim.Length > 5 Then tmpResult &= PONumber1.Trim & " "
         If PONumber2.Trim.Length > 5 Then tmpResult &= PONumber2.Trim & " "
         If PONumber3.Trim.Length > 5 Then tmpResult &= PONumber3.Trim & " "
         If PONumber4.Trim.Length > 5 Then tmpResult &= PONumber4.Trim & " "
         If PONumber5.Trim.Length > 5 Then tmpResult &= PONumber5.Trim & " "
         Return tmpResult
      End Function
      Function PODueDates() As String
         Dim tmpResult As String = ""
         If PODueDate1.Trim.Length > 5 And Not PODueDate1.Contains("00000000") Then tmpResult &= PODueDate1.Trim & " "
         If PODueDate2.Trim.Length > 5 And Not PODueDate2.Contains("00000000") Then tmpResult &= PODueDate2.Trim & " "
         If PODueDate3.Trim.Length > 5 And Not PODueDate3.Contains("00000000") Then tmpResult &= PODueDate3.Trim & " "
         If PODueDate4.Trim.Length > 5 And Not PODueDate4.Contains("00000000") Then tmpResult &= PODueDate4.Trim & " "
         If PODueDate5.Trim.Length > 5 And Not PODueDate5.Contains("00000000") Then tmpResult &= PODueDate5.Trim & " "
         Return tmpResult
      End Function
      Function InPlusArrivingWithin6Months() As Integer
         Return InStock + ExpectedWithin6Months
      End Function
      Function RecommendedBuy() As Integer
         If SixMonthNeed > InPlusArrivingWithin6Months() Then
            Return SixMonthNeed - InPlusArrivingWithin6Months()
         Else
            Return 0
         End If
      End Function
      Function MonthsLeftInStock() As Decimal
         If HighestAvgWk > 0 Then
            Return (InStock / HighestAvgWk) / 4.334
         Else
            Return 0
         End If
      End Function
      Function WeeksLeftInStock() As Decimal
         If HighestAvgWk > 0 Then
            Return (InStock / HighestAvgWk)
         Else
            Return 0
         End If
      End Function
      Function MonthsLeftWithCommingPOAndStock() As Decimal
         If HighestAvgWk > 0 Then
            Return ((InStock + ExpectedWithin6Months) / HighestAvgWk) / 4.334
         Else
            Return 0
         End If
      End Function
      Function WeeskLeftWithCommingPOAndStock() As Decimal
         If HighestAvgWk > 0 Then
            Return ((InStock + ExpectedWithin6Months) / HighestAvgWk)
         Else
            Return 0
         End If
      End Function
      Sub SetLowestCostVendor()
         lowestcost = cost
         lowestcostvendor = vendor
         Dim QueryString As String = "SELECT VENDORNO,CAST(DOLLARCOST AS MONEY)/10000 AS COST FROM VENDORITEMS WHERE EDPNO =" & edpno
            Using conn As New SqlConnection("Data Source=ECOM-DB1;Initial Catalog=ECOMLIVE;UID=ssss;PWD=sssss")
                Dim cmd As New SqlCommand(QueryString, conn)
                Try
                    conn.Open()
                    Dim r As SqlDataReader = cmd.ExecuteReader
                    If r.HasRows Then
                        While r.Read
                            Dim tmpc As Decimal = r.Item("COST")
                            If tmpc < lowestcost Then
                                lowestcost = tmpc
                                lowestcostvendor = Trim(r.Item("VENDORNO"))
                            End If
                        End While
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End Using
        End Sub
   End Structure
End Module


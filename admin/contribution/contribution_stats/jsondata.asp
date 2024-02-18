<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"--> 
<!-- #include virtual="/lib/classes/contribution/contributionCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
    Dim oJson, clsCMeachulLogTotal, arrList, i, mode, syear, smonth, eyear, emonth

    mode = request("mode")
   	syear     = requestcheckvar(request("sY"),4)
	smonth     = requestcheckvar(request("sM"),2)
   	eyear     = requestcheckvar(request("eY"),4)
	emonth     = requestcheckvar(request("eM"),2)

    If len(smonth) = 1 Then
        smonth = "0"&smonth
    End If
    If len(emonth) = 1 Then
        emonth = "0"&emonth
    End If    

    If mode = "totalcontribution" Then
        set clsCMeachulLogTotal = new CMeachulLog
        clsCMeachulLogTotal.FstartYearMonth = syear&"-"&smonth
        clsCMeachulLogTotal.FendYearMonth = eyear&"-"&emonth        
        clsCMeachulLogTotal.ForderGubun = "ASC"
        arrList = clsCMeachulLogTotal.fnGetTotalContributionProfit

        SET oJson = jsObject()
        SET oJson("contributionProfitData") = jsArray()

        IF isArray(arrList) then
            For i=0 to ubound(arrList,2)
                SET oJson("contributionProfitData")(NULL) = jsObject()
                    oJson("contributionProfitData")(NULL)("idx")                        = arrList(0,i)
                    oJson("contributionProfitData")(NULL)("YYYYMM")                     = arrList(1,i)
                    oJson("contributionProfitData")(NULL)("totalPurchase")              = Formatnumber(arrList(2,i),0)
                    oJson("contributionProfitData")(NULL)("totalPurchaseIncome")        = Formatnumber(arrList(3,i),0)
                    oJson("contributionProfitData")(NULL)("bonusCoupon")                = Formatnumber(arrList(4,i),0)
                    oJson("contributionProfitData")(NULL)("handllingAmount")            = Formatnumber(arrList(5,i),0)
                    oJson("contributionProfitData")(NULL)("handllingAmountIncome")      = Formatnumber(arrList(6,i),0)
                    oJson("contributionProfitData")(NULL)("productQuantity")            = Formatnumber(arrList(7,i),0)
                    oJson("contributionProfitData")(NULL)("numberOfOrders")             = Formatnumber(arrList(8,i),0)
                    If ISNULL(arrList(9,i)) Then
                        oJson("contributionProfitData")(NULL)("variableCost1")              = 0
                    Else
                        oJson("contributionProfitData")(NULL)("variableCost1")              = Formatnumber(arrList(9,i),0)
                    End If
                    If ISNULL(arrList(10,i)) Then
                        oJson("contributionProfitData")(NULL)("variableCost2")              = 0
                    Else
                        oJson("contributionProfitData")(NULL)("variableCost2")              = Formatnumber(arrList(10,i),0)
                    End If
                    If ISNULL(arrList(11,i)) Then
                        oJson("contributionProfitData")(NULL)("contributionProfit1")        = 0
                    Else
                        oJson("contributionProfitData")(NULL)("contributionProfit1")        = Formatnumber(arrList(11,i),1)
                    End If
                    If ISNULL(arrList(12,i)) Then
                        oJson("contributionProfitData")(NULL)("contributionProfit2")        = 0
                    Else
                        oJson("contributionProfitData")(NULL)("contributionProfit2")        = Formatnumber(arrList(12,i),1)
                    End If
                    oJson("contributionProfitData")(NULL)("totalPurchaseRate")          = Formatnumber(arrList(13,i),1)
                    oJson("contributionProfitData")(NULL)("bonusCouponRate")            = Formatnumber(arrList(14,i),1)
                    oJson("contributionProfitData")(NULL)("handllingAmountRate")        = Formatnumber(arrList(15,i),1)
                    If ISNULL(arrList(16,i)) Then
                        oJson("contributionProfitData")(NULL)("variableCostRate")           = 0
                    Else
                        oJson("contributionProfitData")(NULL)("variableCostRate")           = Formatnumber(arrList(16,i),1)
                    End If
                    If ISNULL(arrList(17,i)) Then
                        oJson("contributionProfitData")(NULL)("contributionProfitRate")     = 0
                    Else
                        oJson("contributionProfitData")(NULL)("contributionProfitRate")     = Formatnumber(arrList(17,i),1)                    
                    End If
                    oJson("contributionProfitData")(NULL)("regdate")                    = arrList(18,i)
            Next
        End If
        oJson.flush
        Set oJson = Nothing
        Set clsCMeachulLogTotal = Nothing

    ElseIf mode = "ContributionMarginStructure" Then
        set clsCMeachulLogTotal = new CMeachulLog
        clsCMeachulLogTotal.FstartYearMonth = syear&"-"&smonth
        clsCMeachulLogTotal.FendYearMonth = eyear&"-"&emonth        
        clsCMeachulLogTotal.ForderGubun = "ASC"
        arrList = clsCMeachulLogTotal.fnGetTotalContributionProfit

        '{} 선언
        SET oJson = jsObject()
        
        IF isArray(arrList) then
            'chart 선언{}
            SET oJson("chart") = jsObject()
            'json key : theme, value : fusion 정의
            oJson("chart")("theme") = "fusion"
            oJson("chart")("xAxisName") = "년월"
            oJson("chart")("caption") = "공헌이익율구조"
            oJson("chart")("subCaption") = "단위:%"
            
            'categories 선언[]
            SET oJson("categories") = jsArray()
            SET oJson("categories")(NULL) = jsObject()

            'category 배열 선언
            SET oJson("categories")(NULL)("category") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("categories")(NULL)("category")(NULL) = jsObject()
                oJson("categories")(NULL)("category")(NULL)("label") = arrList(1,i)
            Next

            'dataset 선언
            SET oJson("dataset") = jsArray()

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "구매총액수익율"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(Formatnumber(arrList(13,i),1))
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "보너스쿠폰율"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(Formatnumber(arrList(14,i),1))
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "취급액수익율"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)            
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(Formatnumber(arrList(15,i),1))
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "변동비율"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                If ISNULL(arrList(16,i)) Then
                    oJson("dataset")(NULL)("data")(NULL)("value") = 0
                Else
                    oJson("dataset")(NULL)("data")(NULL)("value") = CStr(Formatnumber(arrList(16,i),1))
                End If
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "공헌이익율"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                If ISNULL(arrList(17,i)) Then
                    oJson("dataset")(NULL)("data")(NULL)("value") = 0
                Else
                    oJson("dataset")(NULL)("data")(NULL)("value") = CStr(Formatnumber(arrList(17,i),1))
                End If
            Next
        End If 
        oJson.flush

        Set oJson = Nothing
        Set clsCMeachulLogTotal = Nothing
    ElseIf mode = "TotalPurchase" Then
        set clsCMeachulLogTotal = new CMeachulLog
        clsCMeachulLogTotal.FstartYearMonth = syear&"-"&smonth
        clsCMeachulLogTotal.FendYearMonth = eyear&"-"&emonth        
        clsCMeachulLogTotal.ForderGubun = "ASC"
        arrList = clsCMeachulLogTotal.fnGetTotalContributionProfit

        '{} 선언
        SET oJson = jsObject()
        
        IF isArray(arrList) then
            'chart 선언{}
            SET oJson("chart") = jsObject()
            'json key : theme, value : fusion 정의
            oJson("chart")("theme") = "fusion"
            oJson("chart")("xAxisName") = "년월"
            oJson("chart")("caption") = "구매총액"            
            oJson("chart")("subCaption") = "단위:₩"
            
            'categories 선언[]
            SET oJson("categories") = jsArray()
            SET oJson("categories")(NULL) = jsObject()

            'category 배열 선언
            SET oJson("categories")(NULL)("category") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("categories")(NULL)("category")(NULL) = jsObject()
                oJson("categories")(NULL)("category")(NULL)("label") = arrList(1,i)
            Next

            'dataset 선언
            SET oJson("dataset") = jsArray()

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "구매총액"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(2,i))
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "보너스쿠폰"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(4,i))
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "구매총액수익"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)            
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(3,i))
            Next
        End If 
        oJson.flush

        Set oJson = Nothing
        Set clsCMeachulLogTotal = Nothing        

    ElseIf mode = "HandlingAmount" Then
        set clsCMeachulLogTotal = new CMeachulLog
        clsCMeachulLogTotal.FstartYearMonth = syear&"-"&smonth
        clsCMeachulLogTotal.FendYearMonth = eyear&"-"&emonth        
        clsCMeachulLogTotal.ForderGubun = "ASC"
        arrList = clsCMeachulLogTotal.fnGetTotalContributionProfit

        '{} 선언
        SET oJson = jsObject()
        
        IF isArray(arrList) then
            'chart 선언{}
            SET oJson("chart") = jsObject()
            'json key : theme, value : fusion 정의
            oJson("chart")("theme") = "fusion"
            oJson("chart")("xAxisName") = "년월"
            oJson("chart")("caption") = "취급액"            
            oJson("chart")("subCaption") = "단위:₩"
            
            'categories 선언[]
            SET oJson("categories") = jsArray()
            SET oJson("categories")(NULL) = jsObject()

            'category 배열 선언
            SET oJson("categories")(NULL)("category") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("categories")(NULL)("category")(NULL) = jsObject()
                oJson("categories")(NULL)("category")(NULL)("label") = arrList(1,i)
            Next

            'dataset 선언
            SET oJson("dataset") = jsArray()

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "취급액"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(5,i))
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "취급액수익"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(6,i))
            Next

        End If 
        oJson.flush

        Set oJson = Nothing
        Set clsCMeachulLogTotal = Nothing

    ElseIf mode = "VariableCost" Then
        set clsCMeachulLogTotal = new CMeachulLog
        clsCMeachulLogTotal.FstartYearMonth = syear&"-"&smonth
        clsCMeachulLogTotal.FendYearMonth = eyear&"-"&emonth        
        clsCMeachulLogTotal.ForderGubun = "ASC"
        arrList = clsCMeachulLogTotal.fnGetTotalContributionProfit

        '{} 선언
        SET oJson = jsObject()
        
        IF isArray(arrList) then
            'chart 선언{}
            SET oJson("chart") = jsObject()
            'json key : theme, value : fusion 정의
            oJson("chart")("theme") = "fusion"
            oJson("chart")("xAxisName") = "년월"
            oJson("chart")("caption") = "변동비"            
            oJson("chart")("subCaption") = "단위:₩"
            
            'categories 선언[]
            SET oJson("categories") = jsArray()
            SET oJson("categories")(NULL) = jsObject()

            'category 배열 선언
            SET oJson("categories")(NULL)("category") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("categories")(NULL)("category")(NULL) = jsObject()
                oJson("categories")(NULL)("category")(NULL)("label") = arrList(1,i)
            Next

            'dataset 선언
            SET oJson("dataset") = jsArray()

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "변동비1(물류,수수료)"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                If ISNULL(arrList(9,i)) Then
                    oJson("dataset")(NULL)("data")(NULL)("value") = 0
                Else
                    oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(9,i))
                End If
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "변동비2(판촉비)"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                If ISNULL(arrList(10,i)) Then                
                    oJson("dataset")(NULL)("data")(NULL)("value") = 0
                Else
                    oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(10,i))
                End If
            Next

        End If 
        oJson.flush

        Set oJson = Nothing
        Set clsCMeachulLogTotal = Nothing

    ElseIf mode = "ContributionProfit" Then
        set clsCMeachulLogTotal = new CMeachulLog
        clsCMeachulLogTotal.FstartYearMonth = syear&"-"&smonth
        clsCMeachulLogTotal.FendYearMonth = eyear&"-"&emonth        
        clsCMeachulLogTotal.ForderGubun = "ASC"
        arrList = clsCMeachulLogTotal.fnGetTotalContributionProfit

        '{} 선언
        SET oJson = jsObject()
        
        IF isArray(arrList) then
            'chart 선언{}
            SET oJson("chart") = jsObject()
            'json key : theme, value : fusion 정의
            oJson("chart")("theme") = "fusion"
            oJson("chart")("xAxisName") = "년월"
            oJson("chart")("caption") = "공헌이익"            
            oJson("chart")("subCaption") = "단위:₩"
            
            'categories 선언[]
            SET oJson("categories") = jsArray()
            SET oJson("categories")(NULL) = jsObject()

            'category 배열 선언
            SET oJson("categories")(NULL)("category") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("categories")(NULL)("category")(NULL) = jsObject()
                oJson("categories")(NULL)("category")(NULL)("label") = arrList(1,i)
            Next

            'dataset 선언
            SET oJson("dataset") = jsArray()

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "공헌이익1"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                If ISNULL(arrList(11,i)) Then
                    oJson("dataset")(NULL)("data")(NULL)("value") = 0
                Else
                    oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(11,i))
                End If
            Next

            SET oJson("dataset")(NULL) = jsObject()
            oJson("dataset")(NULL)("seriesname") = "공헌이익2"
            Set oJson("dataset")(NULL)("data") = jsArray()
            For i=0 to ubound(arrList,2)
                SET oJson("dataset")(NULL)("data")(NULL) = jsObject()
                If ISNULL(arrList(12,i)) Then
                    oJson("dataset")(NULL)("data")(NULL)("value") = 0
                Else
                    oJson("dataset")(NULL)("data")(NULL)("value") = CStr(arrList(12,i))
                End If
            Next

        End If 
        oJson.flush

        Set oJson = Nothing
        Set clsCMeachulLogTotal = Nothing                
    End If
%> 
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" --> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
function fnGetMenuGrpIDx()
    Dim retVal : retVal = 2  ''각종쿼리
    Dim imenupos : imenupos = request("menupos")

    if Application("Svr_Info")="Dev" then
        if (imenupos="6086") then
            retVal = 2 ''각종쿼리
        elseif (imenupos="6087") then
            retVal = 3 ''각종검토
        elseif (imenupos="6085") then
            retVal = 1 ''정산검토
        end if
    else
        if (imenupos="4116") then
            retVal = 2 ''각종쿼리
        elseif (imenupos="4117") then
            retVal = 3 ''각종검토
        elseif (imenupos="4115") then
            retVal = 1 ''정산검토
        end if
    end if

    fnGetMenuGrpIDx = retVal
end function

Class CSimpleQueryParamItem

    Public FqryIdx
    Public Fparamidx
    Public Fparamname
    Public Fparamtype
    Public Fparamdirection
    Public Fparamlength
    Public Fisoptional
    Public Fdefaultval
    Public Fparamtitle
    Public Fparamboxtype
    Public FparamSelectOpt


    Public FStoredparamVal

    public function getRequestParam()
        dim bparam1, bparam2

        if FparamBoxtype="yyyymm" then
           bparam1 = requestCheckVar(request("yyyy"&cStr(oQryParam.FQryParamList(i).Fparamidx)),4)
           bparam2 = requestCheckVar(request("mm"&cStr(oQryParam.FQryParamList(i).Fparamidx)),2)

           if (bparam1<>"") and (bparam2<>"") then
                getRequestParam = bparam1&"-"&bparam2
           end if
        ' elseif FparamBoxtype="yyyymmdd" then
        '     getRequestParam = requestCheckVar(request(iparamname),oQryParam.FQryParamList(i).Fparamlength)
        else
            getRequestParam = requestCheckVar(request(iparamname),oQryParam.FQryParamList(i).Fparamlength)
        end if
    end function



    public function getDefaultVAL()
        getDefaultVAL = ""
        if (Fdefaultval="") then Exit function

        if FparamBoxtype="yyyymm" then
            getDefaultVAL = LEFT(dateadd("m",Fdefaultval,now()),7)
        elseif FparamBoxtype="yyyymmdd" then
            getDefaultVAL = LEFT(dateadd("d",Fdefaultval,now()),10)
        else
            getDefaultVAL = Fdefaultval
        end if
    end function

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

End Class

Class CSimpleQueryItem
	Public FYyyymmdd
	Public FMwOrderCnt
	Public FPcOrderCnt
	Public FAppOrderCnt
	Public FMwSumAmount
	Public FPcSumAmount
	Public FAppSumAmount

End Class

Class CSimpleQuery
	Public FOneItem
	Public FItemList()
    Public FQryParamList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

    Public FRectQryidx
    public FParamCount

	Public FRectSDate
	Public FRectEDate
    Public FRectChannel
    public FRectReportType
    Public FRectOrderType
    Public FRectGroupType
    Public FRectDateGbn
    Public FRectAddParam1
    Public FRectAddParam2
    Public FRectchkvs


    public FRectHH

	Private Sub Class_Initialize()
		redim FItemList(0)
        redim FQryParamList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

        FParamCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

    function getQueryParamArr()
        dim ret, strSql, arrVal, i
        strSql = "exec [db_dataSummary].[dbo].[usp_TEN_simpleQueryParamListGet] "&FRectQryidx&""

        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly

        FParamCount = db3_rsget.RecordCount
        if FParamCount<1 then FParamCount=0

        redim preserve FQryParamList(FParamCount)
        If not db3_rsget.EOF Then
            do until db3_rsget.eof
                set FQryParamList(i) = new CSimpleQueryParamItem

                FQryParamList(i).FqryIdx             = db3_rsget("qryIdx")
                FQryParamList(i).Fparamidx           = db3_rsget("paramidx")
                FQryParamList(i).Fparamname          = db3_rsget("paramname")
                FQryParamList(i).Fparamtype          = db3_rsget("paramtype")
                FQryParamList(i).Fparamdirection     = db3_rsget("paramdirection")
                FQryParamList(i).Fparamlength        = db3_rsget("paramlength")
                FQryParamList(i).Fisoptional         = db3_rsget("isoptional")
                FQryParamList(i).Fdefaultval         = db3_rsget("defaultval")
                FQryParamList(i).Fparamtitle         = db3_rsget("paramtitle")
                FQryParamList(i).Fparamboxtype       = db3_rsget("paramboxtype")
                FQryParamList(i).FparamSelectOpt     = db3_rsget("paramSelectOpt")
                db3_rsget.moveNext
                i=i+1
            loop
        End If
        db3_rsget.Close

    end function

    public function ExecSimpleQuery(byref explain, byref colRows, byref retRows, byref returnValue)
        Dim strSql, exeType

        Dim i_dbNm, i_procName : exeType=-1
        ExecSimpleQuery = FALSE


        strSql = "EXEC  [db_dataSummary].[dbo].[usp_TEN_simpleQuery_GetOne] "&FRectQryidx
        db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
        If not db3_rsget.EOF Then
            i_dbNm = db3_rsget("dbNm")
            i_procName = db3_rsget("procName")
            exeType = db3_rsget("exeType")
            explain = db3_rsget("explain")
        End If
        db3_rsget.Close

        if (i_dbNm="") then
            Exit function
        end if

        Dim objRs, fld
        Dim objCmd, i
        Set objCmd = Server.CreateObject("ADODB.COMMAND")
            if (i_dbNm="ten") then
                objCmd.ActiveConnection = dbget
            elseif (i_dbNm="anal") then
                objCmd.ActiveConnection = dbAnalget
            elseif (i_dbNm="dw") then
                objCmd.ActiveConnection = dbSTSget
            else ''datamart
                objCmd.ActiveConnection = db3_dbget
            end if

            objCmd.CommandType = adCmdStoredProc
            objCmd.CommandText = i_procName
            objCmd.CommandTimeout=120

            if (exeType=1) then
                objCmd.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
            end if


            for i=0 to FParamCount-1
                ''rw FQryParamList(i).Fparamtype&"|"&FQryParamList(i).FStoredparamVal
                if (FQryParamList(i).Fparamtype=200 or FQryParamList(i).Fparamtype=201 Or FQryParamList(i).Fparamtype=129 Or FQryParamList(i).Fparamtype=130 Or FQryParamList(i).Fparamtype=202) then
                    objCmd.Parameters.Append objCmd.CreateParameter("@"&FQryParamList(i).Fparamname, FQryParamList(i).Fparamtype, FQryParamList(i).Fparamdirection, FQryParamList(i).Fparamlength, FQryParamList(i).FStoredparamVal)
                else
                    objCmd.Parameters.Append objCmd.CreateParameter("@"&FQryParamList(i).Fparamname, FQryParamList(i).Fparamtype, FQryParamList(i).Fparamdirection, , FQryParamList(i).FStoredparamVal)
                end if
                ' objCmd.Parameters.Append objCmd.CreateParameter("@jyyyymm", adVarchar, adParamInput, 7, "aaa")
                ' objCmd.Parameters.Append objCmd.CreateParameter("@chktype", adInteger, adParamInput, , 1)
                ' objCmd.Parameters.Append .CreateParameter("@chktype", adVarchar, adParamInput, 1, CHKIIF(mode="delexceptbrand","D",""))
                ' objCmd.Parameters.Append .CreateParameter("@retErrText", adVarchar, adParamOutput, 100, retErrText)
            next

            if (exeType=1) then
                objCmd.Execute, , adExecuteNoRecords
                returnValue = objCmd.Parameters("RETURN_VALUE").Value
                'retErrText  = objCmd.Parameters("@retErrText").Value
            else

                set objRs=Server.CreateObject("ADODB.recordset")

                objRs.CursorLocation = adUseClient
                objRs.Open objCmd, , adOpenForwardOnly, adLockReadOnly
                if NOT objRs.Eof then
                    colRows = Array()
                    For Each fld In objRs.Fields
                        reDim Preserve colRows(UBound(colRows) + 1)
                        colRows(UBound(colRows))=fld.Name

                    Next

                    retRows = objRs.getRows()

                end if
                objRs.Close
            end if
        Set objCmd = nothing
        ExecSimpleQuery = True
    end function



    public function getSimpleQuery(byref colRows)
        Dim strSql

	    if (vReportType="top100item") then
	        FRectEDate = DateAdd("d",-1,FRectEDate)
	        strSql = " exec [db_statistics].[dbo].[usp_TEN_Statistics_BestItem_Multi] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectOrderType&"',"&FPageSize&",'"&FRectchkvs&"'"
	    else
            'strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_daily_by_pType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"
        end if

        if (strSql="") then Exit function
'  rw strSql
'response.end
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then

		    colRows = Array()
		    For Each fld In rsSTSget.Fields
		        reDim Preserve colRows(UBound(colRows) + 1)
                colRows(UBound(colRows))=fld.Name

            Next

			getSimpleReportStatistics = rsSTSget.getRows()
		End If
		rsSTSget.Close
	end function

	public function getSimpleReport(byref colRows)
	    Dim strSql

	    if (vReportType="bestitemcoupon") then
	        if (FRectAddParam1<>"") then
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Sales_ByCoupon] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"','"&FRectAddParam1&"'"
	        else
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"','"&FRectAddParam2&"'"
	        end if
	    elseif (vReportType="salesitemcpnbyuserlevel") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Sales_ByLevel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectAddParam1&"'"
	    elseif (vReportType="itemcpnevalwithsales") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Eval_Vs_Spend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectAddParam1&"','"&FRectOrderType&"'"
	    elseif (vReportType="outmallsales") then
	        if (FRectAddParam1<>"") then
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Outmall_Sales_BySitename] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectOrderType&"','"&FRectAddParam1&"'"
	        else
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Outmall_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectOrderType&"'"
	        end if
	    elseif (vReportType="evtsubscript") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Event_subscript] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectAddParam1&"'"
	    elseif (vReportType="rdsitesales") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_rdsite_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectOrderType&"'"
	    elseif (vReportType="newitembycate") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_NewItem_ByCate] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"'"
	    elseif (vReportType="newitembybrandcate") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_NewItem_ByBrandCate] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"'"
        elseif (vReportType="dealsales") then
	        if (FRectAddParam1<>"") then
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Deal_Sales_ByDealCode] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"','"&FRectAddParam1&"'"
	        else
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Deal_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"'"
	        end if
	    else
            'strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_daily_by_pType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"
        end if

        if (strSql="") then Exit function

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then

		    colRows = Array()
		    For Each fld In rsAnalget.Fields
		        reDim Preserve colRows(UBound(colRows) + 1)
                colRows(UBound(colRows))=fld.Name

            Next

			getSimpleReport = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function


End Class


''-------------------------------------------

function drawSimpleQuerySelectBox(iboxname,iselname,iqryGrp)
    dim ret, strSql, arrVal, i
    strSql = "exec [db_dataSummary].[dbo].[usp_TEN_simpleQueryListGet] "&iqryGrp&""

    db3_rsget.CursorLocation = adUseClient
    db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
	If not db3_rsget.EOF Then
		arrVal = db3_rsget.getRows()
	End If
	db3_rsget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(CStr(iselname)=CStr(arrVal(0,i)),"selected","")&">"&arrVal(1,i)&"</option>"
	    next
	    ret = ret&"</select>"
	end if
	response.write ret
end function













function drawConversionChannelSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">ALL</option>"
    ret = ret&"<option value='pc' "&CHKIIF(iselname="pc","selected","")&">WEB</option>"
    ret = ret&"<option value='mw' "&CHKIIF(iselname="mw","selected","")&">MOB</option>"
    ret = ret&"<option value='app' "&CHKIIF(iselname="app","selected","")&">APP</option>"
    ret = ret&"</select>"

    response.write ret
end function


function drawConversionTypeGroupSelectBox(iboxname,iselname, iptype)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_EVT].[dbo].[sp_TEN_conversion_get_comm_code] '"&iptype&"', ''"

    rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
	If not rsAnalget.EOF Then
		arrVal = rsAnalget.getRows()
	End If
	rsAnalget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(0,i)&"</option>"
	    next
	    ret = ret&"</select>"
	end if
	response.write ret
end function

function drawConversionTypeGroupSelectBox2(iboxname,iselname,iptype,idepth,iupvalue)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_EVT].[dbo].[sp_TEN_Conversion_get_comm_code_depth] '"&iptype&"',"&idepth&",'"&iupvalue&"'"
    rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
	If not rsAnalget.EOF Then
		arrVal = rsAnalget.getRows()
	End If
	rsAnalget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(1,i)&"</option>"
	    next

	    if (iselname="UNKNOWN") then
	        ret = ret&"<option value='UNKNOWN' selected>UNKNOWN</option>"
	    else
	        ret = ret&"<option value='UNKNOWN' >UNKNOWN</option>"
	    end if
	    ret = ret&"</select>"
	end if
	response.write ret
end function


function drawPreTimeSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='3' "&CHKIIF(iselname="3","selected","")&">최근3시간</option>"
    ret = ret&"<option value='6' "&CHKIIF(iselname="6","selected","")&">최근6시간</option>"
    ret = ret&"<option value='12' "&CHKIIF(iselname="12","selected","")&">최근12시간</option>"
    ret = ret&"<option value='24' "&CHKIIF(iselname="24","selected","")&">최근24시간</option>"
    ret = ret&"<option value='48' "&CHKIIF(iselname="48","selected","")&">최근48시간</option>"
    ret = ret&"</select>"

    response.write ret
end function

function getpTypeName(ipType)
    select CASE ipType
        CASE "pRtr"
            getpTypeName = "검색어"
        CASE "pBtr"
            getpTypeName = "브랜드"
        CASE "pEtr"
            getpTypeName = "이벤트"
        CASE "pCtr"
            getpTypeName = "카테고리"
        CASE "rc"
            getpTypeName = "상품추천"
        CASE "gaparam"
            getpTypeName = "gaparam"
        CASE ELSE
            getpTypeName = ".."
    end Select
end function
%>

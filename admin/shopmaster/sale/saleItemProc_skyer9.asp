<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 상품등록
' History : 2008.04.04 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim sMode,sCode, dSDate, dEDate
Dim strSql,addSql, addSqlDB
Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype,Keyword,CouponYn, dispCate
Dim egCode,eCode,itemidarr,sType, i
Dim ErrStr : ErrStr = ""
Dim objCmd,iResult
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus
sMode     = requestCheckVar(Request("mode"),1)

if not(C_ADMIN_AUTH) then
	response.write "<script>alert('\n\n\n테스트중입니다.\n\n\n')</script>"
	response.write "테스트중입니다."
	dbget.close()
	response.end
end if

dim oitem

SELECT Case sMode
	Case "I"	'할인상품 추가
		itemidarr = Request("itemidarr")

		sType 		=  Request("sType")

		IF sType = "all" THEN
			'검색된 모든 내용 insert  처리
			makerid     = request("makerid")
			itemid      = request("itemid")
			itemname    = request("itemname")
			Keyword		= request("Keyword")

			sellyn      = request("sellyn")
			usingyn     = request("usingyn")
			danjongyn   = request("danjongyn")
			limityn     = request("limityn")
			mwdiv       = request("mwdiv")
			deliverytype= request("deliverytype")
			sailyn      = request("sailyn")
			CouponYn	= request("CouponYn")

			cdl = request("cdl")
			cdm = request("cdm")
			cds = request("cds")
			dispCate = requestCheckvar(request("disp"),16)

			set oitem = new CItem

			oitem.FPageSize         = 1001
			oitem.FCurrPage         = 1
			oitem.FRectMakerid      = makerid
			oitem.FRectItemid       = itemid
			oitem.FRectItemName     = itemname
			oitem.FRectKeyword		= keyword

			oitem.FRectSellYN       = sellyn
			oitem.FRectIsUsing      = usingyn
			oitem.FRectDanjongyn    = danjongyn
			oitem.FRectLimityn      = limityn
			oitem.FRectMWDiv        = mwdiv
			oitem.FRectDeliveryType = deliverytype
			oitem.FRectSailYn       = sailyn
			oitem.FRectCouponYn		= couponyn

			oitem.FRectCate_Large   = cdl
			oitem.FRectCate_Mid     = cdm
			oitem.FRectCate_Small   = cds
			oitem.FRectDispCate		= dispCate
			''oitem.FRectSortDiv = SortDiv

			oitem.GetItemList
		end if

		response.write "<script>alert('\n\n\n테스트중입니다.\n\n\n')</script>"
		response.write "테스트중입니다."
		dbget.close()
		response.end


		sCode 		= requestCheckVar(Request("sC"),10)
		eCode 		= requestCheckVar(Request("eC"),10)
		egCode 		= Request("egC")	: if egCode = "" then egCode = 0















		addSql = ""
		addSqlDB = ""

		IF sType = "all" THEN '검색된 모든 내용 insert  처리

			'// 추가 쿼리
			if (makerid <> "") then addSql = addSql & " and i.makerid='" & makerid & "'"

			if (itemid <> "") then
				dim iA ,arrTemp,arrItemid

				itemid = replace(itemid,chr(13),"") '상품코드검색 엔터로(2013.12.24)
				arrTemp = Split(itemid,chr(10))

				iA = 0
				do while iA <= ubound(arrTemp)

					if trim(arrTemp(iA))<>"" then
						'상품코드 유효성 검사(2008.08.04;허진원)
						if Not(isNumeric(trim(arrTemp(iA)))) then
							Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
							dbget.close()	:	response.End
						else
							arrItemid = arrItemid & trim(arrTemp(iA)) & ","
						end if
					end if
					iA = iA + 1
				loop

				itemid = left(arrItemid,len(arrItemid)-1)
				addSql = addSql & " and i.itemid in (" + itemid + ")"
				itemidarr = itemid
			end if

			if (itemname <> "") then addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"

			if (Keyword <> "") then
				addSqlDB = addSqlDB + " Join [db_item].[dbo].tbl_item_contents Ct  on i.itemid=Ct.itemid "
            	addSql = addSql & " and Ct.keywords like '%" + Keyword + "%'"
        	end if

			if (sellyn <> "") then addSql = addSql & " and i.sellyn='" + sellyn + "'"
			if (usingyn <> "") then addSql = addSql & " and i.isusing='" + usingyn + "'"

			if danjongyn="SN" then
				addSql = addSql + " and i.danjongyn<>'Y'"
				addSql = addSql + " and i.danjongyn<>'M'"
			elseif danjongyn<>"" then
				addSql = addSql + " and i.danjongyn='" + danjongyn + "'"
			end if

			if limityn="Y0" then
				addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
			elseif limityn<>"" then
				addSql = addSql + " and i.limityn='" + limityn + "'"
			end if

			if mwdiv="MW" then
				addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
			elseif mwdiv<>"" then
				addSql = addSql + " and i.mwdiv='" + mwdiv + "'"
			end if

			if cdl<>"" then addSql = addSql + " and i.cate_large='" + cdl + "'"
			if cdm<>"" then addSql = addSql + " and i.cate_mid='" + cdm + "'"
			if cds<>"" then addSql = addSql + " and i.cate_small='" + cds + "'"
			if sailyn<>"" then addSql = addSql + " and i.sailyn='" + sailyn + "'"
			if deliverytype <> "" then addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
			if CouponYn<>"" then  addSql = addSql + " and i.itemCouponyn='" + CouponYn + "'"

		ELSE
			addSql = addSql & " and i.itemid in ("&trim(itemidarr)&")"
		END IF

		if eCode <> "" then
			addSqlDB = addSqlDB +  " , [db_event].[dbo].[tbl_eventitem] c "
			addSql = addSql + " and i.itemid = c.itemid and c.evt_code = "&eCode&" and c.evtgroup_code ="&egCode
		end if


		'- 할인중이지 않은 상품에 한해서 (2013.06.21; MD팀 요청에 의해 자체 할인중인 상품은 추가 가능)
		'- 추가하려는 할인기간내에 할인예정이지 않은 상품에 한해서
		dim iSaleRate, iSaleMargin, iSaleMarginValue

		'- 추가하려는 할인정보의 기간 확인
		strSql = " SELECT sale_startdate, sale_enddate, sale_rate, sale_margin, sale_marginvalue, sale_status from [db_event].[dbo].tbl_sale where sale_code= "&sCode
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			dSDate = rsget("sale_startdate")
			dEDate = rsget("sale_enddate")
			iSaleRate = rsget("sale_rate")
			iSaleMargin = rsget("sale_margin")
			iSaleMarginValue = rsget("sale_marginvalue")
			saleStatus	= rsget("sale_status")
		End IF
		rsget.Close

		dim strStatus, arrList,intLoop

		IF itemidarr <> "" THEN
			strSql = "SELECT TOP 1000  b.itemid, a.sale_code, a.sale_status "&_
				"   FROM  [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b "&_
				"   WHERE  a.sale_code = b.sale_code and (( a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"'"&_
				"	 and a.sale_using =1 and a.sale_status <> 8 and  b.saleitem_status <> 8 ) or (a.sale_code = "&sCode&")) and b.itemid in ("&itemidarr&")"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				arrList = rsget.getRows()
			End IF
			rsget.Close

			If isArray(arrList) THEN
				For intLoop =0 To UBound(arrList,2)
					Select Case arrList(2,intLoop)
						Case 6
							strStatus = "할인중"
						Case 7
							strStatus = "할인예정"
						Case 0
							strStatus = "등록대기"
					End Select

					ErrStr = ErrStr + "할인코드 : " + CStr(arrList(1,intLoop)) + " - 상품번호 : " + CStr(arrList(0,intLoop)) +" "+ strStatus + " \n"
				Next
			END IF
		END IF

		Dim iChkCount,sqlStr
	 		sqlStr = "SELECT  count(i.itemid) FROM  [db_item].[dbo].tbl_item as i " &addSqlDB
	 		sqlStr = sqlStr &" WHERE i.itemid not in (select itemid from [db_event].[dbo].tbl_saleItem where sale_code="+sCode+") "+addSql
	 		rsget.Open sqlStr, dbget
			IF not rsget.EOF THEN
				iChkCount = rsget(0)
			END IF
			rsget.close
			IF iChkCount>1000 THEN
					%>
				<script language="javascript">
				<!--
				alert("상품은 최대 1000건까지 가능합니다. 조건을 다시 설정해주세요");
				self.location.href ="about:blank";
				//-->
				</script>
			<%
			response.end
			END IF
		' ', orgsailprice, orgsailsuplycash, orgsailyn) "
		strSql = "INSERT INTO [db_event].[dbo].[tbl_saleItem]([sale_code], [itemid], [saleItem_status], [saleprice],[salesupplycash])"
		'strSql = strSql&" SELECT "&sCode&", i.itemid, 7, convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100))"
		strSql = strSql&" SELECT "&sCode&", i.itemid, 7, round(convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100)), -1, 1)"
		Select Case iSaleMargin
			Case 1		'동일마진
				strSql = strSql&" ,convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100))- convert(int,(i.orgprice-(i.orgprice*"&iSaleRate&"/100))*(100-convert(float,convert(int,i.orgsuplycash/i.orgprice*10000)/100))/100)"
			Case 2		'업체부담
				strSql = strSql&" ,convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100)) - (i.orgprice- i.orgsuplycash)"
			Case 3		'반반부담
				strSql = strSql&" , i.orgsuplycash - Convert(int, (i.orgprice-(i.orgprice-(i.orgprice*"&iSaleRate&"/100)))/2)"
			Case 4		'10x10부담
				strSql = strSql&" , i.orgsuplycash "
			Case 5		'직접설정
				strSql = strSql&" , convert(int,i.orgprice-(i.orgprice*"&iSaleRate&"/100)) - convert(int, (i.orgprice-(i.orgprice*"&iSaleRate&"/100))*convert(float,"&iSaleMarginValue&")/100) "
		End Select

		'strSql = strSql&"	, i.sailprice, i.sailsuplycash, i.sailyn "
		strSql = strSql&"	FROM [db_item].[dbo].tbl_item i "&addSqlDB
		''''strSql = strSql&"   WHERE i.sailyn ='N' and "			'(2013.06.21; MD팀 요청에 의해 자체 할인중인 상품은 추가 가능)
		strSql = strSql&" Where i.itemid not in "
		strSql = strSql&" (select b.itemid from [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b"
		strSql = strSql&" 	where a.sale_code = b.sale_code and (( a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"'"
		strSql = strSql&		"	 and a.sale_using =1 and a.sale_status <> 8  and  b.saleitem_status <> 8 ) or (a.sale_code = "&sCode&")) )"&addSql
		'response.write strSql
		dbget.execute strSql

		IF Err.Number <> 0 THEN
	       Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank"
			dbget.close()	:	response.End
		END IF

		IF saleStatus = 6 THEN
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime] ("&sCode&",'"&sMode&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
			 iResult = objCmd(0).Value
			 Set objCmd = nothing
			IF iResult <> 1 THEN
		   			dbget.RollBackTrans
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		       		dbget.close()	:	response.End
			End IF
		END IF
%>
	<script type="text/javascript">
	<!--
		<%
		if ErrStr<>"" then
			ErrStr = ErrStr + "\n\n 할인은 중복설정 불가능합니다. 할인상품을 제외한 상품이 추가됩니다."
			response.write "alert('" + ErrStr + "')"
		end if
		%>
		location.href ="about:blank";
		parent.history.go(0);
		//parent.location.reload();
	//-->
	</script>
<%
		dbget.close()	:	response.End
	Case "U"	'할인 선택상품 수정
	Dim  dissellprice,disbuyprice,arrsaleItemStatus,saleStatus, tmpsaleItemStatus
		sCode = requestCheckVar(Request("sC"),10)
		iCurrpage 	= request("iC")
		itemid 		= split(request("itemid"),",")
		dissellprice= split(request("iDSPrice"),",")
		disbuyprice = split(request("iDBPrice"),",")
		arrsaleItemStatus	=split(request("saleItemStatus"),",")
		saleStatus	=requestCheckVar(Request("saleStatus"),4)

		dbget.beginTrans
		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then

				if Cint(trim(arrsaleItemStatus(i))) = 6 then '오픈중 상태일때 값 변경시 상태값 오픈예정으로 변경처리
					arrsaleItemStatus(i) = 7
				end if

				IF trim(arrsaleItemStatus(i)) = 9 THEN	'원가로
					strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&_
							" SET saleitem_status =9, lastupdate = getdate()"&_
							" WHERE itemid = "&trim(itemid(i)) &_
							"	and sale_code=" & sCode
				ELSE
					strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&_
							" SET saleprice = "&trim(dissellprice(i))&", salesupplycash="&trim(disbuyprice(i))&", saleitem_status ="&arrsaleItemStatus(i)&", lastupdate = getdate()"&_
							" WHERE itemid = "&trim(itemid(i)) &_
							"	and sale_code=" & sCode
				END IF
					dbget.execute strSql

				IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		       		dbget.close()	:	response.End
				End IF

				IF Cint(saleStatus) = 6 or Cint(saleStatus) = 9 or Cint(saleStatus) = 8 THEN	'오픈, 종료, 종료예정 상태일때만 바로 적용
					Set objCmd = Server.CreateObject("ADODB.COMMAND")
					With objCmd
					.ActiveConnection = dbget
					.CommandType = adCmdText
					.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime] ("&sCode&",'"&sMode&"')}"
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With
				    iResult = objCmd(0).Value
				    Set objCmd = nothing

				    IF iResult <> 1 THEN
		   			dbget.RollBackTrans
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		       		dbget.close()	:	response.End
					End IF
				END IF

			end if
		next

		dbget.CommitTrans
		response.redirect("saleReg.asp?menupos="&menupos&"&sC="&sCode&"&eC="&eCode&"&iC="&iCurrpage)
	dbget.close()	:	response.End
	Case "D"	'할인상품 삭제
		sCode = requestCheckVar(Request("sC"),10)
		itemid 		= split(request("itemid"),",")

		dbget.beginTrans
		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then
			strSql ="UPDATE [db_event].[dbo].[tbl_saleItem] "&_
					" SET saleItem_using = 0 , lastupdate=getdate()"&_
					" WHERE itemid = "&trim(itemid(i))
			dbget.execute strSql

				IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		       		dbget.close()	:	response.End
				End IF
			end if
		next

		dbget.CommitTrans
		response.redirect("saleReg.asp?menupos="&menupos&"&sC="&sCode&"&eC="&eCode)
	dbget.close()	:	response.End
	Case "P"	'상품테이블 적용
		sCode = requestCheckVar(Request("sC"),10)
		iCurrpage 	= request("iC")
		IF sCode = "" THEN
			Alert_return("파라미터값에 문제가 있습니다.")
		     dbget.close()	:	response.End
		END IF

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime] ("&sCode&",'"&sMode&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    iResult = objCmd(0).Value
	    Set objCmd = nothing

	'검색어 체크--------------------------------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'검색어
	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	 iCurrpage 		= requestCheckVar(Request("iC"),10)			'현재 페이지 번호
	 ssStatus		= requestCheckVar(Request("sstatus"),10)	'검색 상태
 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus
 	'--------------------------------------------------------------

		IF iResult <> 1 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		    dbget.close()	:	response.End
		END IF
			response.redirect("saleList.asp?menupos="&menupos&"&"&strParm)
	dbget.close()	:	response.End
	CASE Else
	Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")
       dbget.close()	:	response.End
END SELECT


%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

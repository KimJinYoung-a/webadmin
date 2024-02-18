<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim sMode,discountKey, stDT, edDT, discountpro, discountbuyRule, discountbuyPro, discountstatus
Dim strSql,addSql
Dim itemid, sailyn, itemidarr, sType, i
Dim ErrStr : ErrStr = ""
Dim objCmd,iResult
Dim sDate,sSdate, page, strParm, ssStatus, assignedRow
Dim tensalecode, tenSaleStDT, tenSaleEdDT, validCnt
sMode     = requestCheckVar(Request("mode"),1)
tensalecode = requestCheckVar(Request("tensalecode"),10)

function FnSalePriceTouch()
    Dim sqlStr
    sqlStr = " exec db_item.dbo.sp_TEN_Kaffa_Sale_touch"
    dbget.Execute sqlStr

end function

Select Case sMode
	Case "I"	'할인상품 추가
		itemidarr		= Request("itemidarr")
		sType 			= Request("sType")
		discountKey		= Request("discountKey")

		'- 추가하려는 할인정보의 기간 확인
		strSql = ""
		strSql = strSql & " SELECT stDT, convert(varchar(19),edDT,21) as edDT, discountpro, discountbuyRule, discountbuyPro FROM db_item.dbo.tbl_kaffa_Discount_List where discountKey= '"&discountKey&"'"
		rsget.Open strSql,dbget
		If not rsget.EOF Then
			stDT				= rsget("stDT")
			edDT				= rsget("edDT")
			discountpro			= rsget("discountpro")
			discountbuyRule 	= rsget("discountbuyRule")
			discountbuyPro		= rsget("discountbuyPro")
		End IF
		rsget.Close

		Dim strStatus, arrList,intLoop
		If itemidarr <> "" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 100 b.itemid, a.discountKey, a.opendate " & VBCRLF
			strSql = strSql & " FROM db_item.dbo.tbl_kaffa_Discount_List a, db_item.dbo.tbl_kaffa_Discount_Item b "& VBCRLF
			strSql = strSql & " WHERE  a.discountKey = b.discountKey and a.stDT <= '"&edDT&"' and a.edDT >= '"&stDT&"'"& VBCRLF
			strSql = strSql & "	and a.expireddate is NULL"
			strSql = strSql & "	and b.expireddate is NULL"
			strSql = strSql & "	and b.itemid in ("&itemidarr&")"
			rsget.Open strSql,dbget
			If not rsget.EOF Then
				arrList = rsget.getRows()
			End IF
			rsget.Close

			If isArray(arrList) Then
				For intLoop = 0 To UBound(arrList, 2)
				    if isNULL(arrList(2, intLoop)) then
					    strStatus = "등록대기"
					elseif not isNULL(arrList(2, intLoop)) then
					    strStatus = "할인중"
					end if
					ErrStr = ErrStr + "할인코드 : " + CStr(arrList(1,intLoop)) + " - 상품번호 : " + CStr(arrList(0,intLoop)) +" "+ strStatus + " \n"
				Next
			END IF
		END IF

        strSql = "delete D"
        strSql = strSql & " from db_item.dbo.tbl_kaffa_Discount_Item D"
        strSql = strSql & " where D.discountKey="&discountKey
        strSql = strSql & " and D.expiredDate is Not NULL"
        strSql = strSql & " and D.itemid in ("&itemidarr&") "
        dbget.execute strSql

		strSql = ""
		strSql = strSql & " INSERT INTO db_item.dbo.tbl_kaffa_Discount_Item "
		strSql = strSql & " (discountKey, itemid, discountPrice, discountbuyMoney, regdate, regUserID) "
		strSql = strSql & " SELECT '"&discountKey&"', i.itemid, convert(int,i.orgprice-(i.orgprice*"&discountpro&"/100))"
		Select Case discountbuyRule
			Case 0		'매입가지정::기본 원매입가
				strSql = strSql&" , oi.buycash "
			Case 1		'판매가의N%
				strSql = strSql&" , convert(int,i.orgprice-(i.orgprice*"&discountpro&"/100)) - convert(int, (i.orgprice-(i.orgprice*"&discountpro&"/100))*convert(float,"&discountbuyPro&")/100) "
		End Select
		strSql = strSql & " , getdate(), '"&session("ssBctId")&"' "
		strSql = strSql & " FROM db_item.dbo.tbl_item_multiLang_price i "
		strSql = strSql & "     join db_item.dbo.tbl_item oi"
		strSql = strSql & "     on i.itemid=oi.itemid"
		strSql = strSql & " WHERE i.sitename = 'CHNWEB' "
		strSql = strSql & " and i.itemid in ("&itemidarr&") "
		strSql = strSql & " and i.itemid not in "
		strSql = strSql & " (SELECT b.itemid from db_item.dbo.tbl_kaffa_Discount_List a, db_item.dbo.tbl_kaffa_Discount_Item b "
		strSql = strSql & " WHERE a.discountKey = b.discountKey and a.stDT <= '"&edDT&"' and a.edDT >= '"&stDT&"' "
		strSql = strSql & "	and a.expireddate is NULL and b.expiredDate is NULL ) "

		dbget.execute strSql
		If Err.Number <> 0 Then
	       Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank"
			dbget.close()	:	response.End
		End If

		call FnSalePriceTouch()
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
		discountKey = requestCheckVar(Request("discountKey"),10)
		page 	= request("page")
		itemid 		= split(request("itemid"),",")
		dissellprice= split(request("iDSPrice"),",")
		disbuyprice = split(request("iDBPrice"),",")
		arrsaleItemStatus	=split(request("saleItemStatus"),",")
		saleStatus	=requestCheckVar(Request("saleStatus"),4)
		menupos  = request("menupos")
		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then

'' 상태와 상관없음
''				if Cint(trim(arrsaleItemStatus(i))) = 6 then '오픈중 상태일때 값 변경시 상태값 오픈예정으로 변경처리
''					arrsaleItemStatus(i) = 7
''				end if
''
''				IF trim(arrsaleItemStatus(i)) = 9 THEN
''					strSql ="UPDATE db_item.dbo.tbl_kaffa_Discount_Item "&_
''							" SET lastupdate = getdate()"&_
''							" WHERE itemid = "&trim(itemid(i))
''				ELSE
''					strSql ="UPDATE db_item.dbo.tbl_kaffa_Discount_Item "&_
''							" SET discountPrice = "&trim(dissellprice(i))&", discountbuyMoney="&trim(disbuyprice(i))&", lastupdate = getdate()"&_
''							" WHERE itemid = "&trim(itemid(i))
''				END IF
''					dbget.execute strSql

                strSql ="UPDATE db_item.dbo.tbl_kaffa_Discount_Item "&VbCRLF
				strSql = strSql & " SET discountPrice = "&trim(dissellprice(i))&VbCRLF
				strSql = strSql & ", discountbuyMoney="&trim(disbuyprice(i))&VbCRLF
				strSql = strSql & ", lastupdate = getdate()"&VbCRLF
				strSql = strSql & " WHERE itemid = "&trim(itemid(i))
                dbget.execute strSql
				IF Err.Number <> 0 THEN
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		       		dbget.close()	:	response.End
				End IF
			end if
		next

		call FnSalePriceTouch()
		response.redirect("saleItemReg.asp?menupos="&menupos&"&discountKey="&discountKey&"&page="&page)
	dbget.close()	:	response.End
	Case "D"	'할인상품 삭제
		discountKey = requestCheckVar(Request("discountKey"),10)
		itemid 		= split(request("itemid"),",")
		page 	= request("page")
		menupos  = request("menupos")
		For i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then
			strSql ="UPDATE db_item.dbo.tbl_kaffa_Discount_Item "&VbCRLF
			strSql =strSql&" SET lastupdate=getdate()"&VbCRLF
			strSql =strSql&" , expiredDate=getdate()"&VbCRLF
			strSql =strSql&" WHERE itemid = "&trim(itemid(i))
			dbget.execute strSql

				IF Err.Number <> 0 THEN
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		       		dbget.close()	:	response.End
			End IF
			End If
		Next
		call FnSalePriceTouch()
		response.redirect("saleItemReg.asp?menupos="&menupos&"&discountKey="&discountKey&"&page="&page)
	dbget.close()	:	response.End
	CASE "T" '텐코드로 등록
	    validCnt = 0

	    strSql ="select convert(varchar(19),sale_startdate,21) as sale_startdate"
	    strSql = strSql & " ,convert(varchar(10),sale_enddate,21) + ' 23:59:59' as sale_enddate"&VBCRLF
        strSql = strSql & " from db_event.dbo.tbl_sale S"&VBCRLF
        strSql = strSql & " where sale_code="&tensalecode&""&VBCRLF

	    rsget.Open strSql,dbget
		If not rsget.EOF Then
			tenSaleStDT = rsget("sale_startdate")
			tenSaleEdDT = rsget("sale_enddate")
		End IF
		rsget.Close


	    strSql = " select count(*) as validCnt"&VBCRLF
        strSql = strSql & " from db_event.dbo.tbl_sale S"&VBCRLF
    	strSql = strSql & " Join db_event.dbo.tbl_saleItem SI "&VBCRLF
    	strSql = strSql & " on s.sale_code=SI.sale_code"&VBCRLF
    	strSql = strSql & " Join db_item.dbo.tbl_kaffa_reg_Item K"&VBCRLF
    	strSql = strSql & " on SI.itemid=K.itemid"&VBCRLF
    	strSql = strSql & " Join db_Item.dbo.tbl_item i"&VBCRLF
    	strSql = strSql & " on SI.itemid=i.itemid"&VBCRLF
    	strSql = strSql & " Join db_item.dbo.tbl_item_multiLang_price P"&VBCRLF
    	strSql = strSql & " on P.sitename='CHNWEB'"&VBCRLF
    	strSql = strSql & " and P.currencyUnit='WON'"&VBCRLF
    	strSql = strSql & " and P.itemid=SI.itemid"&VBCRLF
    	strSql = strSql & " and P.orgPrice=i.OrgPrice"&VBCRLF
    	strSql = strSql & " where S.sale_status in (6,7)"&VBCRLF
    	strSql = strSql & " and S.sale_using=1"&VBCRLF
    	strSql = strSql & " and S.sale_code="&tensalecode&""&VBCRLF
    	strSql = strSql & " and SI.itemid not in ("&VBCRLF
    	strSql = strSql & " 	select itemid from db_item.dbo.tbl_kaffa_Discount_List DL"&VBCRLF
    	strSql = strSql & " 		Join db_item.dbo.tbl_kaffa_Discount_Item DI"&VBCRLF
    	strSql = strSql & " 		on DL.discountKey=DI.discountKey"&VBCRLF
    	strSql = strSql & " 		and DI.expireddate is NULL"&VBCRLF
    	strSql = strSql & " 	where  DL.expireddate is NULL"&VBCRLF
    	strSql = strSql & " 	and ((DL.STDT<='"&tenSaleStDT&"' and DL.EDDT>='"&tenSaleStDT&"')"&VBCRLF
    	strSql = strSql & " 		or (DL.STDT<='"&tenSaleEdDT&"' and DL.EDDT>='"&tenSaleEdDT&"')"&VBCRLF
   	strSql = strSql & " 	)"&VBCRLF
    	strSql = strSql & " )"&VBCRLF
	    rsget.Open strSql,dbget
		If not rsget.EOF Then
		    validCnt = rsget("validCnt")
		End IF
		rsget.Close

		if (validCnt<1) then
		    Alert_return("등록 가능상품이 없습니다.")
		    dbget.close()	:	response.End
		end if

        ''세일 등록 By Ten SaleCode
        strSql = "exec db_item.[dbo].[sp_Ten_Kaffa_SaleReg_By_TenSaleCode] "&tensalecode&",'"&session("ssBctId")&"'"

		dbget.execute strSql,assignedRow

        call FnSalePriceTouch()
		response.write "<script>alert('저장 되었습니다.');opener.location.reload();window.close();</script>"
		dbget.close()	:	response.End
END SELECT
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

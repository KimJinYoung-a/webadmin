<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : 오프라인 할인 관리
' History : 2010.12.01 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/sale_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<%
Dim sMode , strSql ,sale_shopmarginvalue , sale_shopmargin , osale , copyshopid , sale_code ,tmpmessage ,ErrStr
Dim sCode, eCode,iGroupCode, ssName, dSDay, dEDay, isRate, isMargin, isStatus,isUsing , shopid
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus,sOpenDate,isMValue, Err_contractwrong
dim sale_startdate,sale_enddate,sale_rate,sale_margin,sale_marginValue,sale_status ,strStatus ,point_rate
dim onlySameMargin
	sMode     = requestCheckVar(Request("sM"),10)
	sCode     = requestCheckVar(Request("sC"),10)
	eCode     = requestCheckVar(Request("eC"),10)
	copyshopid     = requestCheckVar(Request("copyshopid"),32)
	ssName			= html2db(requestCheckVar(Request.Form("sSN"),64))
	dSDay 			= requestCheckVar(Request.Form("sSD"),10)
	dEDay			= requestCheckVar(Request.Form("sED"),10)
	isRate			= requestCheckVar(Request.Form("iSR"),10)
	isMargin		= requestCheckVar(Request.Form("salemargin"),10)
	sale_shopmargin = requestCheckVar(Request.Form("shopsalemargin"),10)
	isStatus		= requestCheckVar(Request.Form("salestatus"),10)
	iGroupCode		= requestCheckVar(Request.Form("selG"),10)
	isUsing			= requestCheckVar(Request.Form("sSU"),1)
	sOpenDate		= requestCheckVar(Request.Form("sOD"),30)
	isMValue		= requestCheckVar(Request.Form("isMV"),10)
	sale_shopmarginvalue		= requestCheckVar(Request.Form("sale_shopmarginvalue"),10)
	shopid		= requestCheckVar(Request.Form("shopid"),32)
	point_rate     = requestCheckVar(Request("point_rate"),10)
	onlySameMargin	= requestCheckVar(Request("sOnlySameMargin"),10)

IF point_rate = "" THEN point_rate = 0
IF eCode ="" THEN eCode = 0
IF iGroupCode ="" THEN iGroupCode = 0
IF isRate = "" then	isRate = 0
IF isMValue = "" THEN isMValue =0
if isStatus = "" then isStatus = 0

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'/이벤트를 통한 할인의 경우, 이벤트의 등록대는 해당 매장과, 할인에 등록대는 해당대는 매장이 같아야 댐
if eCode <> 0 then
	strSql = "select top 1 shopid from db_shop.dbo.tbl_event_off" + vbcrlf
	strSql = strSql & " where evt_code = "&eCode&""

	'response.write sql &"<br>"
	rsget.open strSql,dbget,1
		if not(rsget.bof or rsget.eof) then
			if shopid <> rsget("shopid") then
				response.write "<script>"
				response.write "	alert('이벤트를 통한 할인의 경우 이벤트의 등록대는 해당 매장과 할인에 등록대는 해당대는 매장이 같아야 합니다');"
				response.write "	history.go(-1);"
				response.write "</script>"
				response.end	dbget.close()
			end if
		end if
	rsget.close
end if
strSql = ""

Select Case sMode

	'/다른매장에 할인복사
	case "copyshop"

		if sCode = "" then
			response.write "<script>"
			response.write "	alert('할인코드가 없습니다');"
			response.write " 	history.go(-1);"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if copyshopid = "" then
			response.write "<script>alert('동일 적용 매장이 없습니다'); history.go(-1);</script>"
			dbget.close()	:	response.End
		end if

		set osale = new csale_list
		osale.frectsale_code = sCode
		osale.getsaledetail

		if osale.ftotalcount > 0 then
			shopid	= osale.foneitem.fshopid

			if shopid = copyshopid then
				response.write "<script>"
				response.write "	alert('할인 대상 매장과 복사할 적용매장이 같습니다.\n적용매장을 다시 선택해 주세요');"
				response.write "	history.go(-1);"
				response.write "</script>"
				dbget.close()	:	response.End
			end if
		end if

		'//계약이 틀린 브랜드가 있을경우 팅겨낸다. 마진률만 틀릴경우 해당 매장에 맞게 재계산 되서 등록함.
		strSql = "select top 100"
		strSql = strSql & " t.sale_code, t.makerid, t.comm_cd, t.defaultmargin, t.defaultsuplymargin"
		strSql = strSql & " , g.comm_cd, g.defaultmargin, g.defaultsuplymargin"
		strSql = strSql & " from ("
		strSql = strSql & " 	select"
		strSql = strSql & " 	a.sale_code, ii.makerid, sd.comm_cd, sd.defaultmargin, sd.defaultsuplymargin"
		strSql = strSql & " 	from [db_shop].[dbo].tbl_sale_off a"
		strSql = strSql & " 	join [db_shop].[dbo].[tbl_saleitem_off] b"
		strSql = strSql & " 		on a.sale_code = b.sale_code"
		strSql = strSql & " 	join [db_shop].dbo.tbl_shop_item ii"
		strSql = strSql & " 		on b.itemid=ii.shopitemid"
		strSql = strSql & " 		and b.itemgubun=ii.itemgubun"
		strSql = strSql & " 		and b.itemoption=ii.itemoption"
		strSql = strSql & " 	left join db_shop.dbo.tbl_shop_designer sd"
		strSql = strSql & " 		on a.shopid=sd.shopid"
		strSql = strSql & " 		and ii.makerid=sd.makerid"
		strSql = strSql & " 	where a.sale_code="&sCode&""
		strSql = strSql & " 	group by a.sale_code, ii.makerid, sd.comm_cd, sd.defaultmargin, sd.defaultsuplymargin"
		strSql = strSql & " ) as t"
		strSql = strSql & " left join ("
		strSql = strSql & " 	select"
		strSql = strSql & " 	a.sale_code, ii.makerid, sd.comm_cd, sd.defaultmargin, sd.defaultsuplymargin"
		strSql = strSql & " 	from [db_shop].[dbo].tbl_sale_off a"
		strSql = strSql & " 	join [db_shop].[dbo].[tbl_saleitem_off] b"
		strSql = strSql & " 		on a.sale_code = b.sale_code"
		strSql = strSql & " 	join [db_shop].dbo.tbl_shop_item ii"
		strSql = strSql & " 		on b.itemid=ii.shopitemid"
		strSql = strSql & " 		and b.itemgubun=ii.itemgubun"
		strSql = strSql & " 		and b.itemoption=ii.itemoption"
		strSql = strSql & " 	left join db_shop.dbo.tbl_shop_designer sd"
		strSql = strSql & " 		on sd.shopid='"&copyshopid&"'"
		strSql = strSql & " 		and ii.makerid=sd.makerid"
		strSql = strSql & " 	where a.sale_code="&sCode&""
		strSql = strSql & " 	group by a.sale_code, ii.makerid, sd.comm_cd, sd.defaultmargin, sd.defaultsuplymargin"
		strSql = strSql & " ) as g"
		strSql = strSql & " 	on t.sale_code=g.sale_code"
		strSql = strSql & " 	and t.makerid=g.makerid"
		strSql = strSql & " where isnull(t.comm_cd,'')<>isnull(g.comm_cd,'')"

		'response.write strSql &"<Br>"
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			do until rsget.EOF

			Err_contractwrong = Err_contractwrong + "브랜드ID : " + CStr(rsget("makerid")) + " \n"

			rsget.movenext
			loop
		End IF
		rsget.Close

		if Err_contractwrong <> "" then
			response.write "<script>"
			response.write "	alert('매장간 계약이 틀린 브랜드가 있습니다. 확인하시고, 다시 시도 하세요.\n\n"& Err_contractwrong &"');"
			response.write "	history.go(-1);"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		dbget.beginTrans

		'/할인 master등록
		strSql = "INSERT INTO [db_shop].[dbo].[tbl_sale_off] ([sale_name], [sale_rate], point_rate, [sale_margin], [evt_code]"
		strSql = strSql & " , [evtgroup_code], [sale_startdate], [sale_enddate], [sale_status], [adminid]" + vbcrlf
		strSql = strSql & " , [opendate],[lastupdate],sale_marginvalue ,shopid,sale_shopmargin,sale_shopmarginvalue)" + vbcrlf
		strSql = strSql & " 	select sale_name , sale_rate, point_rate ,sale_margin,evt_code,evtgroup_code,sale_startdate,sale_enddate" + vbcrlf
		strSql = strSql & " 	,(case when sale_status = 6 then 7 else sale_status end) as sale_status,'"&session("ssBctId")&"'" + vbcrlf
		strSql = strSql & " 	,opendate,getdate(),sale_marginvalue,'"&copyshopid&"',sale_shopmargin,sale_shopmarginvalue" + vbcrlf
		strSql = strSql & " 	from [db_shop].[dbo].[tbl_sale_off]" + vbcrlf
		strSql = strSql & " 	where sale_code = "&sCode&""

		'response.write strSql &"<br>"
		dbget.execute strSql

		set osale = new csale_list

		'/위에서 등록된 가장 최근 할인정보를 가져온다
		osale.getsalenew()

		if osale.ftotalcount > 0 then
			sale_startdate = osale.foneitem.fsale_startdate
			sale_enddate = osale.foneitem.fsale_enddate
			sale_rate = osale.foneitem.fsale_rate
			point_rate = osale.foneitem.fpoint_rate
			sale_margin = osale.foneitem.fsale_margin
			sale_marginValue = osale.foneitem.fsale_marginValue
			sale_status	= osale.foneitem.fsale_status
			sale_code	= osale.foneitem.fsale_code
			shopid	= osale.foneitem.fshopid
			sale_shopmargin	= osale.foneitem.fsale_shopmargin
			sale_shopmarginvalue = osale.foneitem.fsale_shopmarginvalue
		end if

		'/현재할인과 다른할인과의 상품 중복체크.
		'/중복상품이 있을시 얼럿 띄우고 팅겨낸다. 중복 처리가 되어 버릴경우, 포스에서 할인 꼬여서 인생 피곤해짐..
		strSql = "SELECT distinct b.itemid, a.sale_code, a.sale_status,b.itemoption" + vbcrlf
		strSql = strSql & " FROM [db_shop].[dbo].tbl_sale_off a" + vbcrlf
		strSql = strSql & " join [db_shop].[dbo].[tbl_saleitem_off] b " + vbcrlf
		strSql = strSql & " 	on a.sale_code = b.sale_code " + vbcrlf
		strSql = strSql & " left join (" + vbcrlf
		strSql = strSql & " 	SELECT d.itemid , d.itemgubun , d.itemoption ,c.shopid" + vbcrlf
		strSql = strSql & " 	FROM [db_shop].[dbo].tbl_sale_off c" + vbcrlf
		strSql = strSql & " 	join [db_shop].[dbo].[tbl_saleitem_off] d " + vbcrlf
		strSql = strSql & " 		on c.sale_code = d.sale_code " + vbcrlf
		strSql = strSql & " 	WHERE c.sale_code = "&sCode&"" + vbcrlf
		strSql = strSql & " ) as t " + vbcrlf
		strSql = strSql & " 	on b.itemid=t.itemid"
		strSql = strSql & " 	and b.itemgubun=t.itemgubun"
		strSql = strSql & " 	and b.itemoption=t.itemoption"
		strSql = strSql & " WHERE a.sale_startdate <= '"&sale_enddate&"'"
		strSql = strSql & " and a.sale_enddate >= '"&sale_startdate&"'" + vbcrlf
		strSql = strSql & " and a.sale_using =1"
		strSql = strSql & " and a.sale_status <> 8"
		strSql = strSql & " and b.saleitem_status <> 8"
		strSql = strSql & " and a.shopid = '"&shopid&"'" + vbcrlf
		strSql = strSql & " and t.itemid is not null" + vbcrlf

		'response.write strSql &"<Br>"
		rsget.Open strSql,dbget

		IF not rsget.EOF THEN
			do until rsget.EOF
			IF rsget("sale_status") = 6 THEN
				strStatus = "할인중"
			ELSEIF rsget("sale_status") = 7 THEN
				strStatus = "할인예정"
			ELSEIF rsget("sale_status") = 0 THEN
				strStatus = "등록대기"
			END IF

			ErrStr = ErrStr + "할인코드 : " + CStr(rsget("sale_code")) + " - 상품번호 : " + CStr(rsget("itemid")) +" ("&CStr(rsget("itemoption"))&") "+ strStatus + " \n"

			rsget.movenext
			loop
		End IF

		rsget.Close

		strSql = "INSERT INTO [db_shop].[dbo].[tbl_saleItem_off]" + vbcrlf
		strSql = strSql & " ([sale_code], [itemid],itemgubun , itemoption, [saleItem_status], [saleprice],[salesupplycash]" + vbcrlf
		strSql = strSql & " ,saleshopsupplycash,lastadminid ,point_rate, orgcomm_cd)" + vbcrlf
		strSql = strSql & " 	SELECT "&sale_code&", i.shopitemid,i.itemgubun,i.itemoption, 7" + vbcrlf

		if (onlySameMargin = "Y") then
			strSql = strSql & "		,db_shop.dbo.uf_GetItemPriceCutting( r.saleprice )" + vbcrlf
			strSql = strSql & "		, db_shop.dbo.uf_GetItemPriceCutting( r.salesupplycash )" + vbcrlf
			strSql = strSql & "		, db_shop.dbo.uf_GetItemPriceCutting( r.saleshopsupplycash )" + vbcrlf
		else
			strSql = strSql & " 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100) )" + vbcrlf

			'/매입마진
			'동일마진
			IF sale_margin = 1 THEN
				strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))- convert(int,(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*(100-convert(float,convert(int,i.shopsuplycash/i.orgsellprice*10000)/100))/100) else i.shopsuplycash end) )" + vbcrlf

			'업체부담
			ELSEIF sale_margin = 2 THEN
				'strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - (i.orgsellprice- i.shopsuplycash) else i.shopsuplycash end) )" + vbcrlf

			'반반부담
			ELSEIF 	sale_margin = 3 THEN
				strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf

			'10x10부담
			ELSEIF 	sale_margin = 4 THEN
				strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopsuplycash )" + vbcrlf

			'직접설정
			ELSEIF sale_margin = 5 THEN
				strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - convert(int, (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*convert(float,"&sale_marginValue&")/100) else i.shopsuplycash end) )" + vbcrlf

			'업체특정반반부담/나머지텐바이텐부담
			ELSEIF sale_margin = 6 THEN
				'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf

			'업체특정,출고특정,텐바이텐특정반반부담/나머지텐바이텐부담
			ELSEIF sale_margin = 7 THEN
				'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
			END IF

			'/샾공급마진
			'동일마진
			IF sale_shopmargin = 1 THEN
				strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))- convert(int,(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*(100-convert(float,convert(int,i.shopbuyprice/i.orgsellprice*10000)/100))/100) else i.shopbuyprice end) )" + vbcrlf

			'업체부담
			ELSEIF sale_shopmargin = 2 THEN
				'strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - (i.orgsellprice- i.shopbuyprice) else i.shopbuyprice end) )" + vbcrlf

			'반반부담
			ELSEIF 	sale_shopmargin = 3 THEN
				strSql = strSql&"	 ,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf

			'10x10부담
			ELSEIF 	sale_shopmargin = 4 THEN
				strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopbuyprice )" + vbcrlf

			'직접설정
			ELSEIF 	sale_shopmargin = 5 THEN
				strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - convert(int, (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*convert(float,"&sale_shopmarginvalue&")/100) else i.shopbuyprice end) )" + vbcrlf

			'업체특정반반부담/나머지매장부담
			ELSEIF sale_shopmargin = 6 THEN
				'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf

			'업체특정,출고특정,텐바이텐특정반반부담/나머지텐바이텐부담
			ELSEIF sale_shopmargin = 7 THEN
				'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
			END IF
		end if

		strSql = strSql & "		,'"&session("ssBctId")&"', r.point_rate, i.comm_cd"
		strSql = strSql & " 	from ("
		strSql = strSql & " 		select" + vbcrlf
		strSql = strSql & " 		ii.shopitemprice , ii.makerid, ii.shopitemname , ii.shopitemid ,ii.itemgubun ,ii.itemoption,sdd.comm_cd" + vbcrlf
		strSql = strSql & " 		,(CASE" + vbcrlf
		strSql = strSql & " 			when sdd.comm_cd in ('B012','B013','B011') and ii.shopsuplycash=0" + vbcrlf
		strSql = strSql & " 				THEN convert(int,ii.shopitemprice*(100-IsNULL(sdd.defaultmargin,100))/100)" + vbcrlf
		strSql = strSql & " 			ELSE ii.shopsuplycash END) as 'shopsuplycash'" + vbcrlf
		strSql = strSql & " 		,(CASE" + vbcrlf
		strSql = strSql & " 			when sdd.comm_cd in ('B012','B013','B011') and ii.shopbuyprice=0" + vbcrlf
		strSql = strSql & " 				THEN convert(int,ii.shopitemprice*(100-IsNULL(sdd.defaultsuplymargin,100))/100)" + vbcrlf
		strSql = strSql & " 			ELSE ii.shopbuyprice END) as 'shopbuyprice'" + vbcrlf
		strSql = strSql & " 		,ii.orgsellprice ,sdd.shopid" + vbcrlf
		strSql = strSql & " 		from [db_shop].dbo.tbl_shop_item ii" + vbcrlf

		if (onlySameMargin = "Y") then

			'// 계약 및 마진 동일한 상품만
			strSql = strSql & " join ( "
			strSql = strSql & " 	select "
			strSql = strSql & " 		sio.itemgubun, sio.itemid, sio.itemoption "
			strSql = strSql & " 	from "
			strSql = strSql & " 		[db_shop].[dbo].[tbl_saleItem_off] sio "
			strSql = strSql & " 		join [db_shop].[dbo].[tbl_sale_off] so on sio.sale_code = so.sale_code "
			strSql = strSql & " 		join [db_shop].[dbo].[tbl_shop_item] si "
			strSql = strSql & " 		on "
			strSql = strSql & " 			1 = 1 "
			strSql = strSql & " 			AND si.itemgubun = sio.itemgubun "
			strSql = strSql & " 			and si.shopitemid = sio.itemid "
			strSql = strSql & " 			AND si.itemoption = sio.itemoption "
			strSql = strSql & " 		join db_shop.dbo.tbl_shop_designer sdA "
			strSql = strSql & " 		ON "
			strSql = strSql & " 			1 = 1 "
			strSql = strSql & " 			and so.shopid = sdA.shopid "
			strSql = strSql & " 			AND si.makerid = sdA.makerid "
			strSql = strSql & " 		join db_shop.dbo.tbl_shop_designer sdB "
			strSql = strSql & " 		ON "
			strSql = strSql & " 			1 = 1 "
			strSql = strSql & " 			and '" & copyshopid & "' = sdB.shopid "
			strSql = strSql & " 			AND si.makerid = sdB.makerid "
			strSql = strSql & " 	where "
			strSql = strSql & " 		1 = 1 "
			strSql = strSql & " 		and sio.sale_code = " & sCode
			strSql = strSql & " 		and sdA.defaultmargin = sdB.defaultmargin "
			strSql = strSql & " 		and sdA.defaultsuplymargin = sdB.defaultsuplymargin "
			strSql = strSql & " 		and sdA.comm_cd = sdB.comm_cd			 "
			strSql = strSql & " 	group by sio.itemgubun, sio.itemid, sio.itemoption "
			strSql = strSql & " ) TT "
			strSql = strSql & " on "
			strSql = strSql & " 	1 = 1 "
			strSql = strSql & " 	and ii.itemgubun = TT.itemgubun "
			strSql = strSql & " 	and ii.shopitemid = TT.itemid "
			strSql = strSql & " 	and ii.itemoption = TT.itemoption "

		end if

		strSql = strSql & " 		join db_shop.dbo.tbl_shop_designer sdd" + vbcrlf
		strSql = strSql & " 			on sdd.shopid = '"&shopid&"'" + vbcrlf
		strSql = strSql & " 			and ii.makerid=sdd.makerid" + vbcrlf
		strSql = strSql & " 			and ii.isusing='Y'" + vbcrlf
		strSql = strSql & " 		where ii.orgsellprice = ii.shopitemprice" + vbcrlf		'/상품할인은 제낌(원칙상 이중할인은 안됨)
		strSql = strSql & "		) as i" + vbcrlf
		strSql = strSql & " 	join (" + vbcrlf
		strSql = strSql & " 		SELECT" + vbcrlf
		strSql = strSql & " 		d.itemid , d.itemgubun , d.itemoption, d.saleprice ,d.salesupplycash ,d.saleshopsupplycash" + vbcrlf
		strSql = strSql & " 		, d.point_rate" + vbcrlf
		strSql = strSql & " 		FROM [db_shop].[dbo].tbl_sale_off c" + vbcrlf
		strSql = strSql & " 		join [db_shop].[dbo].[tbl_saleitem_off] d " + vbcrlf
		strSql = strSql & " 			on c.sale_code = d.sale_code " + vbcrlf
		strSql = strSql & " 		WHERE c.sale_code = "&sCode&"" + vbcrlf
		strSql = strSql & " 	) as r " + vbcrlf
		strSql = strSql & " 		on i.shopitemid=r.itemid"
		strSql = strSql & " 		and i.itemgubun=r.itemgubun"
		strSql = strSql & " 		and i.itemoption=r.itemoption" + vbcrlf
		strSql = strSql & " 	left join (" + vbcrlf
		strSql = strSql & " 		select b.itemid ,b.itemgubun , b.itemoption ,a.shopid" + vbcrlf
		strSql = strSql & " 		from [db_shop].[dbo].tbl_sale_off a" + vbcrlf
		strSql = strSql & " 		join [db_shop].[dbo].[tbl_saleitem_off] b" + vbcrlf
		strSql = strSql & " 			on a.sale_code = b.sale_code" + vbcrlf
		strSql = strSql & " 		where a.sale_startdate <= '"&sale_enddate&"'"
		strSql = strSql & " 		and a.sale_enddate >= '"&sale_startdate&"'" + vbcrlf
		strSql = strSql & " 		and a.sale_using = 1"
		strSql = strSql & " 		and a.sale_status <> 8"
		strSql = strSql & " 		and b.saleitem_status not in (8,9)"
		strSql = strSql & " 		and a.shopid = '"&shopid&"'" + vbcrlf
		strSql = strSql & " 	) as t" + vbcrlf
		strSql = strSql & " 		on i.shopitemid = t.itemid" + vbcrlf
		strSql = strSql & "			and i.itemgubun = t.itemgubun" + vbcrlf
		strSql = strSql & "			and i.itemoption = t.itemoption"
		strSql = strSql & "			and i.shopid = t.shopid "
		strSql = strSql & " 	WHERE"
		strSql = strSql & "		i.shopitemprice > 0" + vbcrlf
		strSql = strSql & " 	and t.itemid is null"  '/현재 할인테이블에 등록 되어 있는 상품 제낌

		''response.write strSql &"<Br>"
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank"
			dbget.close()	:	response.End
		END IF

		dbget.CommitTrans

		tmpmessage = "OK (판매금액이 0이하인 상품은 제외되고 등록됩니다)"
		if ErrStr<>"" then
			tmpmessage = tmpmessage & "\n\n중복설정된 상품은 제외되고 등록됩니다.\n" & ErrStr
		end if
%>
		<script langauge="javascript">
			alert('<%= tmpmessage %>');
			location.href='<%=refer%>';
			//history.go(-1);
		</script>
<%
		dbget.close()	:	response.End

	'/할인등록
	Case "I"

		'/할인예정 상태일때
		IF isStatus = "7" THEN
			if sOpenDate = "" then
				'sOpenDate = "getdate()"
			else
				sOpenDate = "convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
			end if
		END IF

		IF sOpenDate = "" THEN sOpenDate = "NULL"

		strSql = "INSERT INTO [db_shop].[dbo].[tbl_sale_off] ([sale_name], [sale_rate], point_rate, [sale_margin], [evt_code], [evtgroup_code], [sale_startdate]" &_
				" , [sale_enddate], [sale_status], [adminid], [opendate],[lastupdate],sale_marginvalue ,shopid,sale_shopmargin,sale_shopmarginvalue)"&_
				" Values ('"&ssName&"',"&isRate&","&point_rate&","&isMargin&","&eCode&","&iGroupCode&",'"&dSDay&"','"&dEDay&"',"&isStatus&",'"&session("ssBctId")&"'" &_
				" ,"&sOpenDate&",getdate(),"&isMValue&" ,'"&shopid&"',"&sale_shopmargin&" ,'"&sale_shopmarginvalue&"')"

		'response.write strSql &"<br>"
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
	       dbget.close()	:	response.End
		END IF

		IF eCode = 0 THEN eCode = ""
		response.redirect("saleList.asp?menupos="&menupos&"&eC="&eCode)
		dbget.close()	:	response.End

	'/할인수정
	Case "U"
		Dim strAdd : strAdd = ""

		IF isStatus ="7" AND sOpenDate="" THEN
			strAdd = " , [opendate] = getdate()"
		END IF

		'검색어 체크--------------------------------------------------------------
		 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
		 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'검색어
		 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
		 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
		 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
		 iCurrpage 		= requestCheckVar(Request("iC"),10)			'현재 페이지 번호
		 ssStatus		= requestCheckVar(Request("sstatus"),10)	'검색 상태
	 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus&"&shopid="&shopid
	 	'--------------------------------------------------------------

		'/현재할인과 다른할인과의 상품 중복체크.
		'/중복상품이 있을시 얼럿 띄우고 팅겨낸다. 중복 처리가 되어 버릴경우, 포스에서 할인 꼬여서 인생 피곤해짐..
		strSql = "SELECT TOP 100"
		strSql = strSql & " b.itemid, b.itemgubun, b.itemoption ,a.sale_code, a.sale_status"
		strSql = strSql & " FROM [db_shop].[dbo].tbl_sale_off a"
		strSql = strSql & " join [db_shop].[dbo].[tbl_saleitem_off] b"
		strSql = strSql & " 	on a.sale_code = b.sale_code"
		strSql = strSql & " left join ("
		strSql = strSql & " 	select"
		strSql = strSql & " 	tb.itemid, tb.itemgubun, tb.itemoption ,ta.sale_code ,ta.shopid"
		strSql = strSql & " 	FROM [db_shop].[dbo].tbl_sale_off ta"
		strSql = strSql & " 	join [db_shop].[dbo].[tbl_saleitem_off] tb"
		strSql = strSql & " 		on ta.sale_code = tb.sale_code"
		strSql = strSql & " 	WHERE "
		strSql = strSql & "  	tb.saleitem_status not in (8,9)"		'/8:종료 9:종료예정
		strSql = strSql & " 	and ta.sale_code = "&sCode&""
		strSql = strSql & " ) as t"
		strSql = strSql & " 	on a.shopid = t.shopid"
		strSql = strSql & " 	and b.itemgubun = t.itemgubun"
		strSql = strSql & " 	and b.itemid = t.itemid"
		strSql = strSql & " 	and b.itemoption = t.itemoption"
		strSql = strSql & " WHERE a.sale_startdate <= '"&dEDay&"'"
		strSql = strSql & " and a.sale_enddate >= '"&dSDay&"'"
		strSql = strSql & " and a.sale_using =1"
		strSql = strSql & " and a.sale_status <> 8"
		strSql = strSql & " and b.saleitem_status not in (8,9)"		'/8:종료 9:종료예정
		strSql = strSql & " and a.sale_code <> "&sCode&""
		strSql = strSql & " and t.shopid is not null"

		'response.write strSql &"<Br>"
		rsget.Open strSql,dbget

		IF not rsget.EOF THEN
			IF rsget("sale_status") = 6 THEN
				strStatus = "할인중"
			ELSEIF rsget("sale_status") = 7 THEN
				strStatus = "할인예정"
			ELSEIF rsget("sale_status") = 0 THEN
				strStatus = "등록대기"
			END IF

			ErrStr = ErrStr + "할인코드 : " + CStr(rsget("sale_code")) + " - 상품번호 : " + CStr(rsget("itemid")) +" "+ strStatus + " \n"
		End IF

		rsget.Close

		if ErrStr<>"" then
			ErrStr = ErrStr + "\n할인기간내에 타 할인과 중복 상품이 있습니다.\n확인후 다시 시도 하세요"
%>
			<script langauge="javascript">
				alert('<%=ErrStr%>');
				location.href='<%=refer%>';
			</script>
<%
			dbget.close()	:	response.End
		end if

		strSql ="UPDATE  [db_shop].[dbo].[tbl_sale_off]  SET "&_
				" sale_name='"&ssName&"', sale_rate="&isRate&",point_rate ="&point_rate&", sale_margin= "&isMargin&",evt_code= "&eCode&_
				" ,evtgroup_code="&iGroupCode&",sale_startdate= '"&dSDay&"',sale_enddate='"&dEDay&"',sale_status="&isStatus&",sale_using='"&isUsing&"'"&_
				" ,sale_marginvalue = "&isMValue&", adminid='"&session("ssBctId")&"' , lastupdate =getdate(),shopid = '"&shopid&"'"&_
				" ,sale_shopmargin="&sale_shopmargin&" , sale_shopmarginvalue="&sale_shopmarginvalue&" "&strAdd&_
				" WHERE sale_code = "&sCode

		'response.write strSql &"<br>"
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
	       dbget.close()	:	response.End
		END IF

		IF eCode = 0 THEN eCode = ""
		response.redirect("saleList.asp?menupos="&menupos&"&"&strParm)
		dbget.close()	:	response.End

	'/실시간 할인 전체 적용
	Case "realall"

		call offitemsaleSet_all

		response.write "<script language='javascript'>"
		response.write "	alert('OK');"
		response.write "	opener.location.reload();"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close()	:	response.End

	CASE Else
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")
	       dbget.close()	:	response.End
End Select

set osale = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2012.01.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
if C_IS_OWN_SHOP or C_IS_SHOP then
	IS_HIDE_BUYCASH = True
end if

dim mode,yyyymmdd,baljuid,targetid
dim reguser, divcode, vatinclude
dim comment, targetname, baljuname, regname
dim waitflag, statecd
dim masteridx, cwflag
dim uniqregdate, errMSG
masteridx = requestCheckVar(request("masteridx"),10)
waitflag  = requestCheckVar(request("waitflag"),10)
mode = requestCheckVar(request("mode"),32)
yyyymmdd = requestCheckVar(request("yyyymmdd"),30)
baljuid = requestCheckVar(request("baljuid"),32)
targetid = requestCheckVar(request("targetid"),32)
cwflag = requestCheckVar(request("cwflag"),10)
reguser = requestCheckVar(request("reguser"),32)
divcode = requestCheckVar(request("divcode"),3)
vatinclude = requestCheckVar(request("vatinclude"),1)
comment = html2db(request("comment"))
targetname = requestCheckVar(html2db(request("targetname")),64)
baljuname = requestCheckVar(html2db(request("baljuname")),64)
regname = requestCheckVar(html2db(request("regname")),64)

statecd = requestCheckVar(request("statecd"),1)

uniqregdate = requestCheckVar(request("uniqregdate"),30)

dim idx
idx = requestCheckVar(request("idx"),10)

dim itemgubunarr, itemarr, itemoptionarr
dim sellcasharr,suplycasharr,buycasharr,itemnoarr,designerarr
itemgubunarr = request("itemgubunarr")
itemarr = request("itemarr")
itemoptionarr = request("itemoptionarr")

sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
buycasharr = request("buycasharr")
itemnoarr  = request("itemnoarr")
designerarr = request("designerarr")

dim itemgubun, itemid, itemoption
dim sellcash, suplycash, buycash, baljuitemno

itemgubun = replace(request("itemgubun"),"|","")
itemid		= replace(request("itemid"),"|","")
itemoption	= replace(request("itemoption"),"|","")
sellcash	= replace(request("sellcash"),"|","")
suplycash	= replace(request("suplycash"),"|","")
buycash		= replace(request("buycash"),"|","")
baljuitemno	= replace(request("baljuitemno"),"|","")

dim i,cnt,sqlStr

dim refer
refer = request.ServerVariables("HTTP_REFERER")
dim brandlist

dim iid,baljucode
dim currmasterState

if mode="addshopjumun" then
	if comment <> "" then
		if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	'// ========================================================================
	if (uniqregdate <> "") then
		'// 등록자 아이디 + 시간을 가지고 중복입력 체크
		sqlStr = "select top 1 idx from db_storage.dbo.tbl_ordersheet_master "
		sqlStr = sqlStr + " where regdate = '" + CStr(uniqregdate) + "' and baljuid = '" + CStr(baljuid) + "' "

		errMSG = ""
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			errMSG = "이미 저장내역이 저장되었습니다.(중복입력)"
		end if
		rsget.close

		if (errMSG <> "") then
			response.write "<script type='text/javascript'>alert('" + CStr(errMSG) + "');</script>"
			response.write errMSG
			dbget.close()	:	response.End
		end if
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("targetid") = targetid
	rsget("targetname") = targetname
	rsget("baljuid") = baljuid
	rsget("baljuname") = baljuname
	rsget("reguser") = reguser
	rsget("regname") = regname
	rsget("divcode") = divcode

	rsget("vatinclude") = vatinclude
	rsget("scheduledate") = yyyymmdd
	if (waitflag<>"") then
		rsget("statecd") = " " ''작성중.
	else
		rsget("statecd") = "0" ''주문접수
	end if

	rsget("comment") = comment

	rsget.update
	iid = rsget("idx")
	rsget.close

    IF (C_IS_OWN_SHOP) then
        baljucode = "JJ" + Format00(6,Right(CStr(iid),6))
    ELSE
	    baljucode = "SJ" + Format00(6,Right(CStr(iid),6))
    END IF

	if targetid="10x10" then
		targetname = "텐바이텐"
	else
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = db2Html(rsget("socname_kor"))
		end if
		rsget.close
	end if

    if baljuid="10x10" then
		baljuname = "텐바이텐"
	else
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + baljuid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			baljuname = db2Html(rsget("socname_kor"))
		end if
		rsget.close
	end if

	if getcwflag(baljuid,"B013") = "1" then
		cwflag = cwflag
	else
		cwflag = "0"
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " ,targetname='" + html2Db(targetname) + "'" + VbCrlf
	sqlStr = sqlStr + " ,baljuname='" + html2Db(baljuname) + "'" + VbCrlf
	sqlStr = sqlStr + " ,cwflag='" + cwflag + "'" + VbCrlf

	if (uniqregdate <> "") then
		sqlStr = sqlStr + " ,regdate='" + CStr(uniqregdate) + "' " + VBCrlf
	end if

	sqlStr = sqlStr + " where idx=" + CStr(iid)

	'response.write sqlStr &"<Br>"
	dbget.execute sqlStr

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	buycasharr = Left(buycasharr,Len(buycasharr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	buycasharr = split(buycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
		sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(iid)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "''," + vbCrlf
		sqlStr = sqlStr + "''," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(buycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'0')"

		rsget.Open sqlStr, dbget, 1
	next

	''if C_IS_OWN_SHOP or C_IS_SHOP then
		sqlStr = " IF EXISTS(select top 1 idx from [db_storage].[dbo].tbl_ordersheet_detail where masteridx = " + CStr(iid)  + " and buycash < 0) "
		sqlStr = sqlStr + " BEGIN "
		sqlStr = sqlStr + " 	update d "
		''sqlStr = sqlStr + " 	set d.sellcash = T.sellcash, d.suplycash = (case when T.suplycash < T.buycash then T.buycash else T.suplycash end), d.buycash = T.buycash "
		sqlStr = sqlStr + " 	set d.buycash = T.buycash "
		sqlStr = sqlStr + " 	FROM "
		sqlStr = sqlStr + " 		[db_storage].[dbo].[tbl_ordersheet_detail] d "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select "
		sqlStr = sqlStr + " 				d.masteridx, d.itemgubun, d.itemid, d.itemoption "
		sqlStr = sqlStr + " 				, s.shopitemprice as sellcash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopbuyprice <> 0 then s.shopbuyprice "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - (35 - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - (m.defaultmargin - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) <> 0 then Round(s.shopitemprice * (100.0 - m.defaultsuplymargin)/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as suplycash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopsuplycash <> 0 then s.shopsuplycash "
		sqlStr = sqlStr + " 						when IsNull(i.mwdiv, '') = 'M' and IsNull(i.buycash, 0) <> 0 and IsNull(m.comm_cd,'') <> 'B012' and IsNull(m.comm_cd,'') <> 'B022' then Round(IsNull(i.buycash,0),0) + Round(IsNull(o.optaddprice,0),0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - 35)/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - IsNull(m.defaultmargin,0))/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as buycash "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				[db_storage].[dbo].[tbl_ordersheet_detail] d "
		sqlStr = sqlStr + " 				join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and d.masteridx = " + CStr(iid)  + " "
		sqlStr = sqlStr + " 					and d.itemgubun = s.itemgubun "
		sqlStr = sqlStr + " 					and d.itemid = s.shopitemid "
		sqlStr = sqlStr + " 					and d.itemoption = s.itemoption "
		sqlStr = sqlStr + " 				left join [db_shop].[dbo].tbl_shop_designer m "
		sqlStr = sqlStr + " 				on	 "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and m.shopid = '" & baljuid & "' "
		sqlStr = sqlStr + " 					and m.makerid = s.makerid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun = '10' "
		sqlStr = sqlStr + " 					and s.shopitemid = i.itemid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun='10' "
		sqlStr = sqlStr + " 					and s.shopitemid = o.itemid "
		sqlStr = sqlStr + " 					and s.itemoption=o.itemoption "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and d.masteridx = T.masteridx "
		sqlStr = sqlStr + " 			and d.itemgubun = T.itemgubun "
		sqlStr = sqlStr + " 			and d.itemid = T.itemid "
		sqlStr = sqlStr + " 			and d.itemoption = T.itemoption "
		sqlStr = sqlStr + " 	WHERE "
		sqlStr = sqlStr + " 		d.buycash < 0 "
		sqlStr = sqlStr + " END "
		rsget.Open sqlStr, dbget, 1
	''end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " ,itemoptionname=[db_shop].[dbo].tbl_shop_item.shopitemoptionname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1

	'주문서 작성하면 상품재고에 표시
    sqlStr = " exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RegJumunJupsuTouch] " & CStr(iid) & " "
    dbget.Execute sqlStr

    '//기주문 업데이트
	PreOrderUpdateByBrand_off iid,targetid,baljuid

elseif mode="shopjumunitemadd" then
	dim itemAlreadyExists

	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>" ") then
		response.write "<script type='text/javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr, dbget, 1
		itemAlreadyExists = rsget("cnt")>0
	rsget.close

	if itemAlreadyExists then
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " set baljuitemno = baljuitemno + " + baljuitemno  + vbCrlf
		sqlStr = sqlStr + " ,realitemno = realitemno + " + baljuitemno  + vbCrlf
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
		sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
		sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
		rsget.Open sqlStr, dbget, 1
	else
		sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
		sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
		sqlStr = sqlStr + " select top 1 "
		sqlStr = sqlStr + " " + CStr(masteridx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + itemgubun + "'," + vbCrlf
		sqlStr = sqlStr + "makerid," + vbCrlf
		sqlStr = sqlStr + "" + itemid + "," + vbCrlf
		sqlStr = sqlStr + "'" + itemoption + "'," + vbCrlf
		sqlStr = sqlStr + "shopitemname," + vbCrlf
		sqlStr = sqlStr + "shopitemoptionname," + vbCrlf
		sqlStr = sqlStr + "" + sellcash + "," + vbCrlf
		sqlStr = sqlStr + "" + suplycash + "," + vbCrlf
		sqlStr = sqlStr + "" + buycash + "," + vbCrlf
		sqlStr = sqlStr + "" + baljuitemno + "," + vbCrlf
		sqlStr = sqlStr + "" + baljuitemno + "," + vbCrlf
		sqlStr = sqlStr + "'0'"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
		sqlStr = sqlStr + " where shopitemid=" + itemid
		sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
		rsget.Open sqlStr, dbget, 1
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1

elseif mode="shopjumunitemaddarr" then
'response.write itemarr

	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd, baljuid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		baljuid = rsget("baljuid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>" ") then
		response.write "<script type='text/javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if


	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	buycasharr = Left(buycasharr,Len(buycasharr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	buycasharr = split(buycasharr,"|")
	itemnoarr = split(itemnoarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + vbCrlf
		sqlStr = sqlStr + " and itemid=" + requestCheckVar(itemarr(i),10) + vbCrlf
		sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'"

		rsget.Open sqlStr, dbget, 1
			itemAlreadyExists = rsget("cnt")>0
		rsget.close

		if itemAlreadyExists then
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " set baljuitemno = baljuitemno + " + requestCheckVar(itemnoarr(i),10)  + vbCrlf
			sqlStr = sqlStr + " ,realitemno = realitemno + " + requestCheckVar(itemnoarr(i),10)  + vbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
			sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + vbCrlf
			sqlStr = sqlStr + " and itemid=" + requestCheckVar(itemarr(i),10) + vbCrlf
			sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'"
'response.write sqlStr
			rsget.Open sqlStr, dbget, 1
		else
			sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
			sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
			sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
			sqlStr = sqlStr + " select top 1 "
			sqlStr = sqlStr + " " + CStr(masteridx)  + "," + vbCrlf
			sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
			sqlStr = sqlStr + "makerid," + vbCrlf
			sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
			sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
			sqlStr = sqlStr + "shopitemname," + vbCrlf
			sqlStr = sqlStr + "shopitemoptionname," + vbCrlf
			sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
			sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
			sqlStr = sqlStr + "" + requestCheckVar(buycasharr(i),20) + "," + vbCrlf
			sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
			sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
			sqlStr = sqlStr + "'0'"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
			sqlStr = sqlStr + " where shopitemid=" + requestCheckVar(itemarr(i),10)
			sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"
			sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'"
'response.write sqlStr
			rsget.Open sqlStr, dbget, 1
		end if
	next

	''if C_IS_OWN_SHOP or C_IS_SHOP then
		sqlStr = " IF EXISTS(select top 1 idx from [db_storage].[dbo].tbl_ordersheet_detail where masteridx = " + CStr(masteridx)  + " and buycash < 0) "
		sqlStr = sqlStr + " BEGIN "
		sqlStr = sqlStr + " 	update d "
		''sqlStr = sqlStr + " 	set d.sellcash = T.sellcash, d.suplycash = (case when T.suplycash < T.buycash then T.buycash else T.suplycash end), d.buycash = T.buycash "
		sqlStr = sqlStr + " 	set d.buycash = T.buycash "
		sqlStr = sqlStr + " 	FROM "
		sqlStr = sqlStr + " 		[db_storage].[dbo].[tbl_ordersheet_detail] d "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select "
		sqlStr = sqlStr + " 				d.masteridx, d.itemgubun, d.itemid, d.itemoption "
		sqlStr = sqlStr + " 				, s.shopitemprice as sellcash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopbuyprice <> 0 then s.shopbuyprice "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - (35 - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) = 0 and IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - (m.defaultmargin - 5))/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultsuplymargin,0) <> 0 then Round(s.shopitemprice * (100.0 - m.defaultsuplymargin)/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as suplycash "
		sqlStr = sqlStr + " 				, (case "
		sqlStr = sqlStr + " 						when s.shopsuplycash <> 0 then s.shopsuplycash "
		sqlStr = sqlStr + " 						when IsNull(i.mwdiv, '') = 'M' and IsNull(i.buycash, 0) <> 0 and IsNull(m.comm_cd,'') <> 'B012' and IsNull(m.comm_cd,'') <> 'B022' then Round(IsNull(i.buycash,0),0) + Round(IsNull(o.optaddprice,0),0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) = 0 then Round(s.shopitemprice * (100.0 - 35)/100, 0) "
		sqlStr = sqlStr + " 						when IsNull(m.defaultmargin,0) <> 0 then Round(s.shopitemprice * (100.0 - IsNull(m.defaultmargin,0))/100, 0) "
		sqlStr = sqlStr + " 						else s.shopitemprice end) as buycash "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				[db_storage].[dbo].[tbl_ordersheet_detail] d "
		sqlStr = sqlStr + " 				join [db_shop].[dbo].tbl_shop_item s "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and d.masteridx = " + CStr(masteridx)  + " "
		sqlStr = sqlStr + " 					and d.itemgubun = s.itemgubun "
		sqlStr = sqlStr + " 					and d.itemid = s.shopitemid "
		sqlStr = sqlStr + " 					and d.itemoption = s.itemoption "
		sqlStr = sqlStr + " 				left join [db_shop].[dbo].tbl_shop_designer m "
		sqlStr = sqlStr + " 				on	 "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and m.shopid = '" & baljuid & "' "
		sqlStr = sqlStr + " 					and m.makerid = s.makerid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun = '10' "
		sqlStr = sqlStr + " 					and s.shopitemid = i.itemid "
		sqlStr = sqlStr + " 				left join [db_item].[dbo].tbl_item_option o "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					1 = 1 "
		sqlStr = sqlStr + " 					and s.itemgubun='10' "
		sqlStr = sqlStr + " 					and s.shopitemid = o.itemid "
		sqlStr = sqlStr + " 					and s.itemoption=o.itemoption "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and d.masteridx = T.masteridx "
		sqlStr = sqlStr + " 			and d.itemgubun = T.itemgubun "
		sqlStr = sqlStr + " 			and d.itemid = T.itemid "
		sqlStr = sqlStr + " 			and d.itemoption = T.itemoption "
		sqlStr = sqlStr + " 	WHERE "
		sqlStr = sqlStr + " 		d.buycash < 0 "
		sqlStr = sqlStr + " END "
		rsget.Open sqlStr, dbget, 1
	''end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1

elseif mode="modimaster" then
	if comment <> "" then
		if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set comment='" + comment + "'"  + vbCrlf
	sqlStr = sqlStr + " ,scheduledate='" + yyyymmdd + "'" + vbCrlf
	if (statecd=" ") or (statecd="0") then
	    sqlStr = sqlStr + " ,statecd='" + statecd + "'" + vbCrlf
	end if
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

elseif (mode="jupsuchange") then
    sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
    sqlStr = sqlStr + " set statecd='0'" + vbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(masteridx) + vbCrlf
    sqlStr = sqlStr + " and statecd=' '"
    dbget.Execute sqlStr

    '//기주문 업데이트
	PreOrderUpdateByBrand_off masteridx,"10x10",""

elseif mode="delmaster" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

    ''기주문수량 업데이트
    PreOrderUpdateByBrand_off masteridx,"10x10",""

elseif mode="modidetail" then
	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>" ") then
		response.write "<script type='text/javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set baljuitemno = " + baljuitemno  + vbCrlf
	sqlStr = sqlStr + " ,realitemno = " + baljuitemno  + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1

elseif mode="deldetail" then
	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>" ") then
		response.write "<script type='text/javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
		brandlist = Left(brandlist,255)
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
end if

if  mode="addshopjumun" then
	refer = "/common/offshop/shop_jumunlist.asp" ''?menupos=497"
elseif mode="delmaster" then
	refer = "/common/offshop/shop_jumunlist.asp" ''?menupos=497"
end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 주문 처리
' History : 2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
' 사용안하는듯
response.end

dim mode,yyyymmdd,baljuid,targetid ,reguser, divcode, vatinclude ,comment, targetname, baljuname, regname , orderno
dim idx , masteridx , itemgubunarr, itemarr, itemoptionarr , sellcasharr,suplycasharr,buycasharr,itemnoarr,designerarr
dim itemgubun, itemid, itemoption , sellcash, suplycash, buycash, baljuitemno , i,cnt,sqlStr ,iid,baljucode
	masteridx = requestCheckVar(request("masteridx"),10)
	mode = requestCheckVar(request("mode"),32)
	yyyymmdd = requestCheckVar(request("yyyymmdd"),10)
	baljuid = requestCheckVar(request("baljuid"),32)
	targetid = requestCheckVar(request("targetid"),32)
	reguser = requestCheckVar(request("reguser"),32)
	divcode = requestCheckVar(request("divcode"),3)
	vatinclude = requestCheckVar(request("vatinclude"),1)
	comment = html2db(request("comment"))
	targetname = requestCheckVar(html2db(request("targetname")),64)
	baljuname = requestCheckVar(html2db(request("baljuname")),64)
	regname = requestCheckVar(html2db(request("regname")),64)
	idx = requestCheckVar(request("idx"),10)
	itemgubunarr = request("itemgubunarr")
	itemarr = request("itemarr")
	itemoptionarr = request("itemoptionarr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	buycasharr = request("buycasharr")
	itemnoarr  = request("itemnoarr")
	designerarr = request("designerarr")
	itemgubun = replace(request("itemgubun"),"|","")
	itemid		= replace(request("itemid"),"|","")
	itemoption	= replace(request("itemoption"),"|","")
	sellcash	= replace(request("sellcash"),"|","")
	suplycash	= replace(request("suplycash"),"|","")
	buycash		= replace(request("buycash"),"|","")
	baljuitemno	= replace(request("baljuitemno"),"|","")
	orderno	= requestCheckVar(request("orderno"),16)

dim refer
refer = request.ServerVariables("HTTP_REFERER")
dim brandlist

if mode="addshopjumun" then
	if comment <> "" then
		if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
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
	rsget("statecd") = "0"
	rsget("comment") = comment

	rsget.update
	iid = rsget("idx")
	rsget.close

	baljucode = "JJ" + Format00(6,Right(CStr(iid),6))

	if targetid="10x10" then
		targetname = "텐바이텐"
	else
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = rsget("socname_kor")
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
			baljuname = rsget("socname_kor")
		end if
		rsget.close
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " ,targetname='" + targetname + "'" + VbCrlf
	sqlStr = sqlStr + " ,baljuname='" + baljuname + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1

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

elseif mode="shopjumunitemadd" then
	dim itemAlreadyExists
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
			sqlStr = sqlStr + "convert(varchar(64),shopitemname)," + vbCrlf
			sqlStr = sqlStr + "convert(varchar(64),shopitemoptionname)," + vbCrlf
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
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

elseif mode="delmaster" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

elseif mode="modidetail" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set baljuitemno = " + baljuitemno  + vbCrlf
	sqlStr = sqlStr + " ,realitemno = " + baljuitemno  + vbCrlf
	''sqlStr = sqlStr + " ,suplycash = " + suplycash  + vbCrlf
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

elseif mode="jumundel" then
    rw "사용중지"
    dbget.Close() : response.end

'	sqlStr = " update db_shop.dbo.tbl_shopjumun_master set" + vbCrlf
'	sqlStr = sqlStr & " cancelyn='Y'" + vbCrlf
'	sqlStr = sqlStr & " where orderno='"&orderno&"'"
'
'	'response.write sqlStr &"<br>"
'	dbget.execute sqlStr
'
'	response.write "<script type='text/javascript'>alert('OK'); opener.location.reload(); self.close();</script>"
'	response.end
elseif mode="jflagdel" then
    ''mode="jumundel" 대신 사용
    sqlStr = " update D " + vbCrlf
	sqlStr = sqlStr & " set jcomm_cd='B000'" + vbCrlf
	sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_detail D"
	sqlStr = sqlStr & " where orderno='"&orderno&"'"&VbCRLF
	sqlStr = sqlStr & " and idx="&idx&VbCRLF
	sqlStr = sqlStr & " and jcomm_cd is NULL"

	dbget.execute sqlStr
	response.write "<script type='text/javascript'>alert('OK'); opener.location.reload(); self.close();</script>"
    dbget.Close() : response.end
elseif mode="jflagedit" then
	if (request("jcommcd")<>"") then
		sqlStr = " update D " + vbCrlf
		sqlStr = sqlStr & " set jcomm_cd='"&request("jcommcd")&"'" + vbCrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_detail D"
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_master M"
		sqlStr = sqlStr & " 	on m.orderno=d.orderno"
		sqlStr = sqlStr & " where d.orderno='"&orderno&"'"&VbCRLF
		sqlStr = sqlStr & " and d.idx="&idx&VbCRLF
		sqlStr = sqlStr & " and d.jcomm_cd is NULL"
		sqlStr = sqlStr & " and m.shopid in ('streetshop011', 'streetshop999')"

		dbget.execute sqlStr
	end if
	response.write "<script type='text/javascript'>alert('OK'); opener.location.reload(); self.close();</script>"
    dbget.Close() : response.end
end if

if  mode="addshopjumun" then
	refer = "/admin/offshop/offjumunlist.asp?menupos=552"
end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<%

dim ipgomasterid


dim mode,yyyymmdd,baljuid,targetid
dim reguser, divcode, vatinclude
dim comment, targetname, baljuname, regname, statecd
dim opage, ourl
dim masteridx
dim beasongdate, songjangdiv, songjangname, songjangno
dim ipgodate
dim scheduledt, finishuser, finishname, checkusersn, rackipgousersn
dim ojbaljucode, ojstatecd, HTTP_Object, SiteURL

masteridx = request("masteridx")
opage = request("opage")
ourl = request("ourl")

mode = request("mode")
yyyymmdd = request("yyyymmdd")
baljuid = request("baljuid")
targetid = request("targetid")

reguser = request("reguser")
divcode = request("divcode")
vatinclude = request("vatinclude")
comment = html2db(request("comment"))
targetname = html2db(request("targetname"))
baljuname = html2db(request("baljuname"))
regname = html2db(request("regname"))

statecd = request("statecd")
beasongdate = request("beasongdate")
songjangdiv = request("songjangdiv")
songjangname = html2db(request("songjangname"))
songjangno = request("songjangno")
ipgodate = request("ipgodate")

scheduledt = request("scheduledt")
finishuser = request("finishuser")
finishname = html2db(request("finishname"))

checkusersn = request("checkusersn")
rackipgousersn = request("rackipgousersn")

ojbaljucode = request("ojbaljucode")

dim idx
idx = request("idx")

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
dim realitemno

itemgubun = replace(request("itemgubun"),"|","")
itemid		= replace(request("itemid"),"|","")
itemoption	= replace(request("itemoption"),"|","")
sellcash	= replace(request("sellcash"),"|","")
suplycash	= replace(request("suplycash"),"|","")
buycash		= replace(request("buycash"),"|","")
baljuitemno	= replace(request("baljuitemno"),"|","")
realitemno  = replace(request("realitemno"),"|","")

dim detailidx ,dtstat
detailidx= request("detailidx")
dtstat = request("dtstat")

dim i,cnt,sqlStr

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim iid,baljucode
dim itemAlreadyExists
dim brandlist

dim chk, tmp

dim currmasterState

if mode="addshopjumun" then

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

	baljucode = "OJ" + Format00(6,Right(CStr(iid),6))

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

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " ,targetname='" + targetname + "'" + VbCrlf
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
		sqlStr = sqlStr + "'" + itemgubunarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + designerarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "" + itemarr(i) + "," + vbCrlf
		sqlStr = sqlStr + "'" + itemoptionarr(i) + "'," + vbCrlf
		sqlStr = sqlStr + "''," + vbCrlf
		sqlStr = sqlStr + "''," + vbCrlf
		sqlStr = sqlStr + "" + sellcasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + "0" + "," + vbCrlf
		sqlStr = sqlStr + "" + buycasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
		sqlStr = sqlStr + "'0')"

		rsget.Open sqlStr, dbget, 1
	next

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set itemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
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


	PreOrderUpdateByBrand(targetid)

elseif mode="shopjumunitemadd" then

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
		sqlStr = sqlStr + "'10'," + vbCrlf
		sqlStr = sqlStr + "makerid," + vbCrlf
		sqlStr = sqlStr + "" + itemid + "," + vbCrlf
		sqlStr = sqlStr + "'" + itemoption + "'," + vbCrlf
		sqlStr = sqlStr + "itemname," + vbCrlf
		sqlStr = sqlStr + "optionname," + vbCrlf
		sqlStr = sqlStr + "" + sellcash + "," + vbCrlf
		sqlStr = sqlStr + "" + suplycash + "," + vbCrlf
		sqlStr = sqlStr + "" + buycash + "," + vbCrlf
		sqlStr = sqlStr + "" + baljuitemno + "," + vbCrlf
		sqlStr = sqlStr + "" + baljuitemno + "," + vbCrlf
		sqlStr = sqlStr + "'0'"
		sqlStr = sqlStr + " from ("
		sqlStr = sqlStr + " select i.*, v.optionname from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on v.itemid=i.itemid"
		sqlStr = sqlStr + " and v.itemoption='" + itemoption + "'"
		sqlStr = sqlStr + " where i.itemid=" + itemid
		sqlStr = sqlStr + " ) as T"

'response.write sqlStr
		rsget.Open sqlStr, dbget, 1


	end if

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



elseif mode="shopjumunitemaddarr" then

	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>"1") then
		response.write "<script language='javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
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
		sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'" + vbCrlf
		sqlStr = sqlStr + " and itemid=" + itemarr(i) + vbCrlf
		sqlStr = sqlStr + " and itemoption='" + itemoptionarr(i) + "'"

		rsget.Open sqlStr, dbget, 1
			itemAlreadyExists = rsget("cnt")>0
		rsget.close

		if itemAlreadyExists then
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " set baljuitemno = baljuitemno + " + itemnoarr(i)  + vbCrlf
			sqlStr = sqlStr + " ,realitemno = realitemno + " + itemnoarr(i)  + vbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
			sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'" + vbCrlf
			sqlStr = sqlStr + " and itemid=" + itemarr(i) + vbCrlf
			sqlStr = sqlStr + " and itemoption='" + itemoptionarr(i) + "'"
'response.write sqlStr
			rsget.Open sqlStr, dbget, 1
		else
			sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
			sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
			sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
			sqlStr = sqlStr + " select top 1 "
			sqlStr = sqlStr + " " + CStr(masteridx)  + "," + vbCrlf
			sqlStr = sqlStr + "'" + itemgubunarr(i) + "'," + vbCrlf
			sqlStr = sqlStr + "makerid," + vbCrlf
			sqlStr = sqlStr + "" + itemarr(i) + "," + vbCrlf
			sqlStr = sqlStr + "'" + itemoptionarr(i) + "'," + vbCrlf
			sqlStr = sqlStr + "itemname," + vbCrlf
			sqlStr = sqlStr + "optionname," + vbCrlf
			sqlStr = sqlStr + "" + sellcasharr(i) + "," + vbCrlf
			sqlStr = sqlStr + "0," + vbCrlf
			sqlStr = sqlStr + "" + buycasharr(i) + "," + vbCrlf
			sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
			sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
			sqlStr = sqlStr + "'0'"
			sqlStr = sqlStr + " from ("
			sqlStr = sqlStr + " select i.*, v.optionname from [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on v.itemid=i.itemid"
			sqlStr = sqlStr + " and v.itemoption='" + itemoptionarr(i) + "'"
			sqlStr = sqlStr + " where i.itemid=" + itemarr(i)
			sqlStr = sqlStr + " ) as T"
'response.write sqlStr
			rsget.Open sqlStr, dbget, 1
		end if
	next

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

	PreOrderUpdateByBrand(targetid)

elseif mode="modimaster" then
	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set comment='" + comment + "'"  + vbCrlf
	sqlStr = sqlStr + " ,scheduledate='" + yyyymmdd + "'" + vbCrlf
	sqlStr = sqlStr + " ,statecd='" + statecd + "'" + vbCrlf

	if beasongdate<>"" then
		sqlStr = sqlStr + " ,beasongdate='" + beasongdate + "'" + vbCrlf
	end if

	if songjangdiv<>"" then
		sqlStr = sqlStr + " ,songjangdiv='" + songjangdiv + "'" + vbCrlf
	end if

 	if songjangno<>"" then
		sqlStr = sqlStr + " ,songjangno='" + songjangno + "'" + vbCrlf
	end if

 	if songjangname<>"" and songjangname<>"선택" then
		sqlStr = sqlStr + " ,songjangname='" + songjangname + "'" + vbCrlf
	end if

	if divcode<>"" then
		sqlStr = sqlStr + " ,divcode='" + divcode + "'" + vbCrlf
	end if

	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
''response.write sqlStr
	rsget.Open sqlStr, dbget, 1

	PreOrderUpdateByBrand(targetid)

elseif mode="franupcheipgofinish" then

	if (divcode="101") or (divcode="111") then
		''가맹점용 개별매입, 가맹점용 개별위탁 인경우 입고완료로만 진행.
	elseif (divcode="121") then
		''[온라인위탁재고->가맹점용위탁] 인경우 온라인 내역에 출고로 잡히고 가맹점으로 위탁입고됩니다. 입고 확정

		'1.온라인 출고 가능내역인지 확인 itemgubun start with 10

		sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and itemgubun<>'10'"
		sqlStr = sqlStr + " and deldt is null"
		rsget.Open sqlStr,dbget,1
			itemAlreadyExists = rsget("cnt")>0
		rsget.Close

		if itemAlreadyExists then
			response.write "<script>alert('온라인에서 출고할 수 없는 아이템이 있습니다. 작업이 취소되었습니다.');</script>"
			response.write "<script>location.replace('" + refer + "');</script>"
			dbget.close()	:	response.End
		end if

		'1.온라인 출고 마스타
		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		''출고처
		rsget("socid") = "streetshop800"  '-가맹점 대표
		rsget("chargeid") = reguser
		rsget("divcode") = "006"
		rsget("vatcode") = "008"
		rsget("comment") = ""
		rsget("ipchulflag") = "S"

		rsget.update
		iid = rsget("id")
		rsget.close

		baljucode = "SO" + Format00(6,Right(CStr(iid),6))

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		rsget.Open sqlStr,dbget,1

		'''2.온라인 출고 디테일 입력
		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash,"
		sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),i.buycash,'W',d.itemgubun,d.itemname,d.itemoptionname"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VBCrlf
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and d.itemgubun='10'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and deldt is null"
		rsget.Open sqlStr,dbget,1

		'''2.온라인 출고 마스타 업데이트
		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set executedt='" + ipgodate + "'" + VBCrlf
		sqlStr = sqlStr + " ,scheduledt='" + ipgodate + "'" + VBCrlf
		sqlStr = sqlStr + " ,totalsellcash=T.totsell" + VBCrlf
		sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
		sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
		sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
		sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
		sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
		sqlStr = sqlStr + " where mastercode='"  + CStr(baljucode) + "'" + vbCrlf
		sqlStr = sqlStr + " and deldt is null" + vbCrlf
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where id=" + CStr(iid)
		rsget.Open sqlStr,dbget,1

	elseif (divcode="131") then
		''[온라인위탁재고->가맹점용매입] 인경우 온라인 내역에 반품으로 잡히고 가맹점으로 매입입고됩니다. 입고 확정
		'1.온라인 출고 가능내역인지 확인 itemgubun start with 10

		sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and itemgubun<>'10'"
		sqlStr = sqlStr + " and deldt is null"
		rsget.Open sqlStr,dbget,1
			itemAlreadyExists = rsget("cnt")>0
		rsget.Close

		if itemAlreadyExists then
			response.write "<script>alert('온라인에서 출고할 수 없는 아이템이 있습니다. 작업이 취소되었습니다.');</script>"
			response.write "<script>location.replace('" + refer + "');</script>"
			dbget.close()	:	response.End
		end if

		'1.온라인 출고 마스타
		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		''업체ID
		rsget("socid") = targetid
		rsget("chargeid") = reguser
		rsget("divcode") = "002"
		rsget("vatcode") = "008"
		rsget("comment") = ""
		rsget("ipchulflag") = "I"

		rsget.update
		iid = rsget("id")
		rsget.close

		baljucode = "ST" + Format00(6,Right(CStr(iid),6))

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 디테일 입력
		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname) " + VBCrlf
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, i.sellcash, i.buycash,"
		sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),0,'W',d.itemgubun,d.itemname,d.itemoptionname" + VBCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d," + VBCrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VBCrlf
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and d.itemgubun='10'" + VBCrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + VBCrlf
		sqlStr = sqlStr + " and d.deldt is null" + VBCrlf
		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 마스타 업데이트
		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set executedt='" + ipgodate + "'" + VBCrlf
		sqlStr = sqlStr + " ,scheduledt='" + ipgodate + "'" + VBCrlf
		sqlStr = sqlStr + " ,totalsellcash=T.totsell" + VBCrlf
		sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
		sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
		sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
		sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
		sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
		sqlStr = sqlStr + " where mastercode='"  + CStr(baljucode) + "'" + vbCrlf
		sqlStr = sqlStr + " and deldt is null" + vbCrlf
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where id=" + CStr(iid)
		rsget.Open sqlStr,dbget,1
	elseif (divcode="201") then
		''[온라인매입재고->가맹점용매입] 인경우 온라인 내역에 출고로 잡히고 가맹점으로 매입입고됩니다. 입고 확정
		'1.온라인 출고 가능내역인지 확인 itemgubun start with 10

		sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and itemgubun<>'10'"
		sqlStr = sqlStr + " and deldt is null"
		rsget.Open sqlStr,dbget,1
			itemAlreadyExists = rsget("cnt")>0
		rsget.Close

		if itemAlreadyExists then
			response.write "<script>alert('온라인에서 출고할 수 없는 아이템이 있습니다. 작업이 취소되었습니다.');</script>"
			response.write "<script>location.replace('" + refer + "');</script>"
			dbget.close()	:	response.End
		end if

		'1.온라인 출고 마스타
		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		''출고처
		rsget("socid") = "streetshop800"  '-가맹점 대표
		rsget("chargeid") = reguser
		rsget("divcode") = "006"
		rsget("vatcode") = "008"
		rsget("comment") = ""
		rsget("ipchulflag") = "S"

		rsget.update
		iid = rsget("id")
		rsget.close

		baljucode = "SO" + Format00(6,Right(CStr(iid),6))

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		rsget.Open sqlStr,dbget,1

		'''2.온라인 출고 디테일 입력
		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash,"
		sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),i.buycash,'M',d.itemgubun,d.itemname,d.itemoptionname"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VBCrlf
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and d.itemgubun='10'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and d.deldt is null"
		rsget.Open sqlStr,dbget,1

		'''2.온라인 출고 마스타 업데이트
		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set executedt='" + ipgodate + "'" + VBCrlf
		sqlStr = sqlStr + " ,scheduledt='" + ipgodate + "'" + VBCrlf
		sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totsell,0)" + VBCrlf
		sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + VBCrlf
		sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totbuy,0)" + VBCrlf
		sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
		sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
		sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
		sqlStr = sqlStr + " where mastercode='"  + CStr(baljucode) + "'" + vbCrlf
		sqlStr = sqlStr + " and deldt is null" + vbCrlf
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " where id=" + CStr(iid)
		rsget.Open sqlStr,dbget,1
	else
		response.write "<script>alert('구분코드 없음." + divcode + "')</script>"
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set statecd='9'" + vbCrlf
	if ipgodate<>"" then
		sqlStr = sqlStr + " ,ipgodate='" + ipgodate + "'" + vbCrlf
	end if

	if baljucode<>"" then
		sqlStr = sqlStr + " ,alinkcode='" + baljucode + "'" + vbCrlf
	end if

	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1


elseif mode="delmaster" then
	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>"1") then
		response.write "<script language='javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

	PreOrderUpdateByBrand(targetid)

elseif mode="modidetail" then
	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>"1") then
		response.write "<script language='javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if


	i = Request.Form("itemgubun").Count
	redim chk(i)
	redim itemgubun(i)
	redim itemid(i)
	redim itemoption(i)
	redim sellcash(i)
	redim buycash(i)
	redim suplycash(i)
	redim baljuitemno(i)
	redim realitemno(i)
	redim comment(i)

	for i = 0 to Request.Form("itemgubun").Count - 1
		if (Request.Form("chk").Count >= (i + 1)) then
			chk(i) = Request.Form("chk")(i + 1)
		else
			chk(i) = ""
		end if

		itemgubun(i) = Request.Form("itemgubun")(i + 1)
		itemid(i) = Request.Form("itemid")(i + 1)
		itemoption(i) = Request.Form("itemoption")(i + 1)
		sellcash(i) = Request.Form("sellcash")(i + 1)
		buycash(i) = Request.Form("buycash")(i + 1)
		suplycash(i) = Request.Form("suplycash")(i + 1)
		baljuitemno(i) = Request.Form("baljuitemno")(i + 1)
		realitemno(i) = Request.Form("realitemno")(i + 1)
		comment(i) = Request.Form("comment")(i + 1)
	next

	for i=0 to UBound(chk) - 1
		tmp = trim(chk(i))
		if (tmp <> "") then
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " set baljuitemno = " + trim(baljuitemno(CInt(tmp)))  + vbCrlf
			sqlStr = sqlStr + " ,realitemno = " + trim(realitemno(CInt(tmp)))  + vbCrlf
			sqlStr = sqlStr + " ,comment = '" + trim(comment(CInt(tmp))) + "'" + vbCrlf

			if trim(sellcash(CInt(tmp)))<>"" then
				sqlStr = sqlStr + " ,sellcash = " + trim(sellcash(CInt(tmp))) + "" + vbCrlf
			end if
			if trim(suplycash(CInt(tmp)))<>"" then
				sqlStr = sqlStr + " ,suplycash = " + trim(suplycash(CInt(tmp))) + "" + vbCrlf
			end if
			if trim(buycash(CInt(tmp)))<>"" then
				sqlStr = sqlStr + " ,buycash = " + trim(buycash(CInt(tmp))) + "" + vbCrlf
			end if

			sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
			sqlStr = sqlStr + " and itemgubun='" + trim(itemgubun(CInt(tmp))) + "'" + vbCrlf
			sqlStr = sqlStr + " and itemid=" + trim(itemid(CInt(tmp))) + vbCrlf
			sqlStr = sqlStr + " and itemoption='" + trim(itemoption(CInt(tmp))) + "'"
			'response.write sqlStr
			rsget.Open sqlStr, dbget, 1
		end if
	next

	'dbget.close()	:	response.End

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

	PreOrderUpdateByBrand(targetid)

elseif mode="modidetail2" then
	''주문 접수 상태인지 체크

	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>"1") then
		response.write "<script language='javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if


		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " set baljuitemno = " + trim(baljuitemno)  + vbCrlf
		sqlStr = sqlStr + " ,realitemno = " + trim(realitemno)  + vbCrlf
		sqlStr = sqlStr + " ,comment = '" + trim(comment) + "'" + vbCrlf

		if trim(sellcash)<>"" then
			sqlStr = sqlStr + " ,sellcash = " + trim(sellcash) + "" + vbCrlf
		end if
		if trim(suplycash)<>"" then
			sqlStr = sqlStr + " ,suplycash = " + trim(suplycash) + "" + vbCrlf
		end if
		if trim(buycash)<>"" then
			sqlStr = sqlStr + " ,buycash = " + trim(buycash) + "" + vbCrlf
		end if

		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
		sqlStr = sqlStr + " and itemgubun='" + trim(itemgubun) + "'" + vbCrlf
		sqlStr = sqlStr + " and itemid=" + trim(itemid) + vbCrlf
		sqlStr = sqlStr + " and itemoption='" + trim(itemoption) + "'"
		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1



	'dbget.close()	:	response.End

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

	dim detail_status,dtcnt

    sqlStr = " select count(*) as cnt from [db_storage].[dbo].tbl_ordersheet_detail_log where detail_idx='" & CStr(detailidx) & "'"

    rsget.open sqlStr,dbget,1

    if not rsget.eof then
    	dtcnt = rsget("cnt")
    end if
    rsget.close


    if dtstat="" then
    	detail_status=""

	elseif dtstat="ipt" then
		detail_status= "직접입력"

	elseif dtstat="so" then
		detail_status ="단종"

	elseif dtstat="sso" then
		detail_status ="일시품절"

	end if

    if dtcnt>0 then
    	sqlStr =" update [db_storage].[dbo].tbl_ordersheet_detail_log " &_
    			" set detail_status='" & detail_status & "'" &_
    			" ,detail_description ='" & comment & "'" &_
    			" where detail_idx='" & CStr(detailidx) & "'"
    else
    	sqlStr =" insert into db_storage.dbo.tbl_ordersheet_detail_log(detail_idx,detail_status,detail_description) " &_
    			" values('" & CStr(detailidx) & "','" & detail_status & "','" & comment & "') "
    end if

    	dbget.execute(sqlStr)

	PreOrderUpdateByBrand(targetid)

elseif mode="deldetail2" then



	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>"1") then
		response.write "<script language='javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and itemgubun='" + trim(itemgubun) + "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" + trim(itemid) + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + trim(itemoption) + "'"
	'response.write sqlStr
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

	sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail_log where detail_idx='" & CStr(detailidx) & "'"

    dbget.execute(sqlStr)


	PreOrderUpdateByBrand(targetid)

elseif mode="deldetail" then
	''주문 접수 상태인지 체크
	sqlStr = "select top 1 statecd, targetid from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
		currmasterState = rsget("statecd")
		targetid = rsget("targetid")
	rsget.close

	if (currmasterState<>"0") and (currmasterState<>"1") then
		response.write "<script language='javascript'>"
		response.write "alert('Error !! \n\n주문 접수 상태에서만 수정 가능합니다.');"
		response.write "location.replace('" +  refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if


	i = Request.Form("itemgubun").Count
	redim chk(i)
	redim itemgubun(i)
	redim itemid(i)
	redim itemoption(i)
	redim sellcash(i)
	redim buycash(i)
	redim suplycash(i)
	redim baljuitemno(i)
	redim realitemno(i)
	redim comment(i)

	for i = 1 to Request.Form("itemgubun").Count
		if (Request.Form("chk").Count >= i) then
			chk(i - 1) = Request.Form("chk")(i)
		else
			chk(i - 1) = ""
		end if

		itemgubun(i - 1) = Request.Form("itemgubun")(i)
		itemid(i - 1) = Request.Form("itemid")(i)
		itemoption(i - 1) = Request.Form("itemoption")(i)
		sellcash(i - 1) = Request.Form("sellcash")(i)
		buycash(i - 1) = Request.Form("buycash")(i)
		suplycash(i - 1) = Request.Form("suplycash")(i)
		baljuitemno(i - 1) = Request.Form("baljuitemno")(i)
		realitemno(i - 1) = Request.Form("realitemno")(i)
		comment(i - 1) = Request.Form("comment")(i)
	next

	for i=0 to UBound(chk) - 1
		tmp = trim(chk(i))
		if (tmp <> "") then
			sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
			sqlStr = sqlStr + " and itemgubun='" + trim(itemgubun(CInt(tmp))) + "'" + vbCrlf
			sqlStr = sqlStr + " and itemid=" + trim(itemid(CInt(tmp))) + vbCrlf
			sqlStr = sqlStr + " and itemoption='" + trim(itemoption(CInt(tmp))) + "'"
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

	PreOrderUpdateByBrand(targetid)
elseif mode="justnext" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set statecd='9'"
	sqlStr = sqlStr + " , ipgodate='" + ipgodate + "'"
	sqlStr = sqlStr + " , finishuser='" + finishuser + "'"
	sqlStr = sqlStr + " , finishname='" + finishname + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
elseif mode="savennext" then

	if (comment = "") then
		comment = CStr(ojbaljucode) + " 주문 자동입고처리"
	else
		comment = CStr(ojbaljucode) + " 주문 자동입고처리" + vbCrLf + vbCrLf + comment
	end if

	'// ========================================================================
	sqlStr = " select statecd from [db_storage].[dbo].tbl_ordersheet_master "
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	ojstatecd = ""
	rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			ojstatecd = rsget("statecd")
		end if
	rsget.close

	if (ojstatecd = "9") then
		response.write "<script>alert('이미 입고내역이 저장되었습니다.(중복입력)');</script>"
		response.write "이미 입고내역이 저장되었습니다.(중복입력)"
		dbget.close()	:	response.End
	else
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
		sqlStr = sqlStr + " set statecd='9'"
		sqlStr = sqlStr + " where idx=" + CStr(masteridx)
		rsget.Open sqlStr, dbget, 1
	end if

	'// ========================================================================
	'1.온라인 입고 마스타
	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	''업체ID
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = finishuser
	rsget("checkusersn") = checkusersn
	rsget("rackipgousersn") = rackipgousersn
	rsget("chargename") = finishname
	rsget("divcode") = divcode ''001-매입, 002-위탁
	rsget("vatcode") = "008"
	rsget("comment") = comment
	rsget("ipchulflag") = "I"

	rsget.update
	iid = rsget("id")
	rsget.close
	ipgomasterid = iid

	baljucode = "ST" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)

	rsget.Open sqlStr,dbget,1

	'''2.온라인 입고 디테일 입력 : 기존 매입구분 잘못입력되어있었음.
	sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
	sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
	sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid,comment) " + VBCrlf
	sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.buycash,"
	sqlStr = sqlStr + " d.realitemno, getdate(),getdate(),0,i.mwdiv,d.itemgubun,d.itemname,d.itemoptionname,d.makerid,convert(varchar(32), (g.detail_status + ' ' + g.detail_description))" + VBCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d" + VBCrlf
	sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + "     on d.itemgubun='10' and d.itemid=i.itemid"
	sqlStr = sqlStr + " 	left join db_storage.dbo.tbl_ordersheet_detail_log g "
	sqlStr = sqlStr + "		on d.idx= g.detail_idx "
	sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and d.deldt is null" + VBCrlf
	rsget.Open sqlStr,dbget,1

    '// 업배상품 입고시 센터매입구분 적용, skyer9, 2021-01-26
	sqlStr = " update d " + vbCrLf
	sqlStr = sqlStr + " set d.mwgubun = si.centermwdiv " + vbCrLf
	sqlStr = sqlStr + " from " + vbCrLf
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_detail d " + vbCrLf
	sqlStr = sqlStr + " 	join [db_shop].[dbo].[tbl_shop_item] si " + vbCrLf
	sqlStr = sqlStr + " 	on " + vbCrLf
	sqlStr = sqlStr + " 		1 = 1 " + vbCrLf
	sqlStr = sqlStr + " 		and d.iitemgubun = si.itemgubun " + vbCrLf
	sqlStr = sqlStr + " 		and si.shopitemid = d.itemid " + vbCrLf
	sqlStr = sqlStr + " 		and si.itemoption = d.itemoption " + vbCrLf
	sqlStr = sqlStr + " where " + vbCrLf
	sqlStr = sqlStr + " 	1 = 1 " + vbCrLf
	sqlStr = sqlStr + " 	and d.mastercode = '" & baljucode & "' " + vbCrLf
	'sqlStr = sqlStr + " 	and d.mwgubun = 'U' " + vbCrLf		' 2021.04.09 한용민(이문재이사님 요청 주석처리함)
	sqlStr = sqlStr + " 	and si.centermwdiv is not NULL " + vbCrLf
    sqlStr = sqlStr + " 	and d.deldt is null" + VBCrlf
    rsget.Open sqlStr,dbget,1

	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set executedt='" + ipgodate + "'" + VBCrlf
	sqlStr = sqlStr + " ,scheduledt='" + scheduledt + "'" + VBCrlf
	sqlStr = sqlStr + " ,totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(baljucode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

    'TODO : 신규등록된 입고에 대해, 서머리정보를 업데이트한다.
   QuickUpdateNewIpgoDetailSummary baljucode, false


	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set statecd='9'"
	sqlStr = sqlStr + " , ipgodate='" + ipgodate + "'"
	sqlStr = sqlStr + " , finishuser='" + finishuser + "'"
	sqlStr = sqlStr + " , finishname='" + finishname + "'"
	sqlStr = sqlStr + " , blinkcode='" + baljucode + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1

	''기주문수량 업데이트
	if (masteridx <> 194488) then
		PreOrderUpdateByBrand(targetid)
	end if

    '// AGV에 상품정보 전송
    Set HTTP_Object = Server.CreateObject("MSXML2.ServerXMLHTTP")

    IF application("Svr_Info")="Dev" THEN
        SiteURL = "http://testwapi.10x10.co.kr/agv/api.asp?mode=senditeminfo&ordertype=ipgo&baljucode=" & baljucode
    else
        SiteURL = "http://wapi.10x10.co.kr/agv/api.asp?mode=senditeminfo&ordertype=ipgo&baljucode=" & baljucode
    end if

    With HTTP_Object
        .SetTimeouts 30000, 30000, 30000, 30000
        .Open "POST", SiteURL, False
        .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .Send ""
        .WaitForResponse 60
    End With

    Set HTTP_Object = Nothing
end if



if  mode="addshopjumun" then
	refer = "/admin/newstorage/orderlist.asp?menupos=536"
elseif (mode="justnext") then

	response.write "<script language='javascript'>"
	response.write "alert('저장 되었습니다.');"
	response.write "window.close();"
	response.write "opener.document.location.reload();"
	response.write "</script>"
	dbget.close()	:	response.End
elseif (mode="savennext") then
	response.write "<script language='javascript'>"
	response.write "alert('저장 되었습니다.\r\n한정 수량을 처리해 주세요.');"
	'response.write "window.resizeTo(900,600);location.replace('poplimitcheck.asp?idx=" + Cstr(masteridx) + "');"
	response.write "window.resizeTo(900,600);location.replace('poplimitcheckipgoNew.asp?idx=" + Cstr(ipgomasterid) + "');"
	response.write "</script>"
	dbget.close()	:	response.End

'elseif mode="modimaster" then
'	if ourl<>"" then
'		refer = "/admin/newstorage/" + ourl + "?menupos=530&page=" + opage
'	else
'		refer = "/admin/newstorage/orderlist.asp?menupos=536&page=" + opage
'	end if
end if
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

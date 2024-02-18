<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 물류주문서 처리
' History : 2009.04.07 서동석 생성
'			2012.01.11 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode,yyyymmdd,baljuid,targetid ,oshopid,ostatecd,odesinger
dim reguser, divcode, vatinclude ,comment, targetname, baljuname, regname, statecd
dim opage, ourl ,masteridx ,beasongdate, songjangdiv, songjangname, songjangno ,idx
dim ipgodate , datestr, orgbaljucode , shopid ,alinkcode ,waitflag, limitflag
dim iid,baljucode ,itemAlreadyExists ,brandlist , itemexists, obaljucode ,regidx
dim itemgubunarr, itemarr, itemoptionarr ,i,cnt,sqlStr
dim sellcasharr,suplycasharr,buycasharr,itemnoarr,designerarr,detailidxarr,baljuitemnoarr,realitemnoarr,commentarr
dim ipgoflagarr, defaultmaginflagarr, buymaginflagarr, suplymaginflagarr
dim cpbaljuid, cwFlag

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
	orgbaljucode = request("orgbaljucode")
	statecd = request("statecd")
	beasongdate = request("beasongdate")
	songjangdiv = request("songjangdiv")
	songjangname = html2db(request("songjangname"))
	songjangno = request("songjangno")
	ipgodate = request("ipgodate")
	datestr = request("datestr")
	shopid = request("shopid")
	alinkcode = request("alinkcode")
	oshopid = request("oshopid")
	ostatecd = request("ostatecd")
	odesinger = request("odesinger")
	idx = request("idx")

	''작성중인경우.
	waitflag = request("waitflag")
	limitflag = request("limitflag")

	itemgubunarr = request("itemgubunarr")
	itemarr = request("itemarr")
	itemoptionarr = request("itemoptionarr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	buycasharr = request("buycasharr")
	itemnoarr  = request("itemnoarr")
	designerarr = request("designerarr")
	detailidxarr = request("detailidxarr")
	baljuitemnoarr = request("baljuitemnoarr")
	realitemnoarr = request("realitemnoarr")
	commentarr = html2db(request("commentarr"))
	ipgoflagarr = request("ipgoflagarr")
	defaultmaginflagarr = request("defaultmaginflagarr")
	buymaginflagarr     = request("buymaginflagarr")
	suplymaginflagarr   = request("suplymaginflagarr")
cpbaljuid= request("cpbaljuid")

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

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

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

	if (waitflag<>"") then
		rsget("statecd") = " " ''작성중.
	else
		rsget("statecd") = "0" ''주문접수
	end if

	rsget("comment") = comment

	rsget.update
	iid = rsget("idx")
	rsget.close

	baljucode = "SJ" + Format00(6,Right(CStr(iid),6))

	if targetid="10x10" then
		targetname = "텐바이텐"
	else
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = db2html(rsget("socname_kor"))
		end if
		rsget.close
	end if

	if baljuname="" then
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + baljuid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			baljuname = db2html(rsget("socname_kor"))
		end if
		rsget.close
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " ,targetname='" + html2db(targetname) + "'" + VbCrlf
	sqlStr = sqlStr + " ,baljuname='" + html2db(baljuname) + "'" + VbCrlf
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
		sqlStr = sqlStr + "" + suplycasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + buycasharr(i) + "," + vbCrlf
		sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
		if (waitflag<>"") then
			sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
		else
			sqlStr = sqlStr + "0," + vbCrlf
		end if
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
elseif mode="modeshopjumunarr" then
	'itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	'itemarr = Left(itemarr,Len(itemarr)-1)
	'itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	'sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	'suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	'buycasharr = Left(buycasharr,Len(buycasharr)-1)
	'detailidxarr = Left(detailidxarr,Len(detailidxarr)-1)

	'baljuitemnoarr = Left(baljuitemnoarr,Len(baljuitemnoarr)-1)
	'realitemnoarr = Left(realitemnoarr,Len(realitemnoarr)-1)
	'commentarr = Left(commentarr,Len(commentarr)-1)


	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	buycasharr = split(buycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	detailidxarr = split(detailidxarr,"|")

	baljuitemnoarr = split(baljuitemnoarr,"|")
	realitemnoarr = split(realitemnoarr,"|")
	commentarr = split(commentarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		if Trim(itemarr(i)<>"") then
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " set sellcash=" + sellcasharr(i)
			sqlStr = sqlStr + " , buycash=" + buycasharr(i)
			sqlStr = sqlStr + " , suplycash=" + suplycasharr(i)
			sqlStr = sqlStr + " , baljuitemno=" + baljuitemnoarr(i)
			sqlStr = sqlStr + " , realitemno=" + realitemnoarr(i)
			sqlStr = sqlStr + " , comment='" + commentarr(i) + "'"
			sqlStr = sqlStr + " where idx=" + detailidxarr(i)

			rsget.Open sqlStr, dbget, 1
		end if
	next

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " ,itemoptionname=[db_shop].[dbo].tbl_shop_item.shopitemoptionname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
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
elseif mode="modeshopjumunmasterdetail" then
	''edit master
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

	''edit detail
	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	buycasharr = split(buycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	detailidxarr = split(detailidxarr,"|")

	baljuitemnoarr = split(baljuitemnoarr,"|")
	realitemnoarr = split(realitemnoarr,"|")
	commentarr = split(commentarr,"|")
	ipgoflagarr = split(ipgoflagarr,"|")
	defaultmaginflagarr = split(defaultmaginflagarr,"|")
	buymaginflagarr = split(buymaginflagarr,"|")
	suplymaginflagarr = split(suplymaginflagarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		if Trim(itemarr(i)<>"") then
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
			sqlStr = sqlStr + " set sellcash=" + sellcasharr(i)
			sqlStr = sqlStr + " , buycash=" + buycasharr(i)
			sqlStr = sqlStr + " , suplycash=" + suplycasharr(i)
			sqlStr = sqlStr + " , baljuitemno=" + baljuitemnoarr(i)
			sqlStr = sqlStr + " , realitemno=" + realitemnoarr(i)
			sqlStr = sqlStr + " , comment='" + commentarr(i) + "'"
			sqlStr = sqlStr + " , ipgoflag='" + ipgoflagarr(i) + "'"
			sqlStr = sqlStr + " , defaultmaginflag='" + defaultmaginflagarr(i) + "'"
			sqlStr = sqlStr + " , buymaginflag='" + buymaginflagarr(i) + "'"
			sqlStr = sqlStr + " , suplymaginflag='" + suplymaginflagarr(i) + "'"
			sqlStr = sqlStr + " where idx=" + detailidxarr(i)

			rsget.Open sqlStr, dbget, 1
		end if
	next

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " ,itemoptionname=[db_shop].[dbo].tbl_shop_item.shopitemoptionname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
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
elseif mode="delshopjumunarr" then
	if Right(detailidxarr,1)="," then
		detailidxarr = Left(detailidxarr,Len(detailidxarr)-1)
	end if


	if Trim(detailidxarr<>"") then
		sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " where idx in (" + detailidxarr + ")"
'response.write sqlStr
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
	end if
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
		'sqlStr = sqlStr + " ,realitemno = realitemno + " + baljuitemno  + vbCrlf
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
		sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
		sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
		rsget.Open sqlStr, dbget, 1
'response.write sqlStr
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
		sqlStr = sqlStr + "0," + vbCrlf
		sqlStr = sqlStr + "'0'"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
		sqlStr = sqlStr + " where shopitemid=" + itemid
		sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
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
			'sqlStr = sqlStr + " ,realitemno = realitemno + " + itemnoarr(i)  + vbCrlf
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
			sqlStr = sqlStr + "shopitemname," + vbCrlf
			sqlStr = sqlStr + "shopitemoptionname," + vbCrlf
			sqlStr = sqlStr + "" + sellcasharr(i) + "," + vbCrlf
			sqlStr = sqlStr + "" + suplycasharr(i) + "," + vbCrlf
			sqlStr = sqlStr + "" + buycasharr(i) + "," + vbCrlf
			sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
			sqlStr = sqlStr + "0," + vbCrlf
			sqlStr = sqlStr + "'0'"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
			sqlStr = sqlStr + " where shopitemid=" + itemarr(i)
			sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'"
			sqlStr = sqlStr + " and itemoption='" + itemoptionarr(i) + "'"
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
elseif mode="modimaster" then
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

	rsget.Open sqlStr, dbget, 1
elseif mode="franupcheipgofinish" then

	if (divcode="101") then
		''가맹점용 개별매입 - 입고리스트에 가맹점 입고로 잡음(801)
		'''입고 가능여부 체크
		'sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		'sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		'sqlStr = sqlStr + " and itemgubun<>'10'"
		'sqlStr = sqlStr + " and deldt is null"
		'rsget.Open sqlStr,dbget,1
		'	itemAlreadyExists = rsget("cnt")>0
		'rsget.Close

		'if itemAlreadyExists then
		'	response.write "<script>alert('온라인에서 출고할 수 없는 아이템이 있습니다. 작업이 취소되었습니다.');</script>"
		'	response.write "<script>location.replace('" + refer + "');</script>"
		'	dbget.close()	:	response.End
		'end if

		'1.온라인 입고 마스타
		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		rsget("socid") = targetid  '' - 입고브랜드
		rsget("chargeid") = reguser
		rsget("divcode") = "801"   '' - 가맹용입고
		rsget("vatcode") = "008"
		rsget("comment") = comment + VBCRLF + "가맹점 개별매입 주문서 " + orgbaljucode + " 입고처리"
		rsget("ipchulflag") = "I"

		rsget.update
		iid = rsget("id")
		rsget.close

		baljucode = "ST" + Format00(6,Right(CStr(iid),6))


		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = rsget("socname_kor")
		end if
		rsget.close


		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " ,socname='" + targetname + "'" + VBCrlf
		sqlStr = sqlStr + " ,chargename='" + regname + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 디테일 입력
		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.buycash,"
		sqlStr = sqlStr + " d.realitemno, getdate(),getdate(),d.buycash,'M',d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and deldt is null"
		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 마스타 업데이트
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

	elseif (divcode="111") then
		'' 가맹점용 개별위탁 - 입고리스트에 일반 위탁으로 잡음(002)
		'''입고 가능여부 체크
		'sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		'sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		'sqlStr = sqlStr + " and itemgubun<>'10'"
		'sqlStr = sqlStr + " and deldt is null"
		'rsget.Open sqlStr,dbget,1
		'	itemAlreadyExists = rsget("cnt")>0
		'rsget.Close

		'if itemAlreadyExists then
		'	response.write "<script>alert('온라인에서 출고할 수 없는 아이템이 있습니다. 작업이 취소되었습니다.');</script>"
		'	response.write "<script>location.replace('" + refer + "');</script>"
		'	dbget.close()	:	response.End
		'end if

		'1.온라인 입고 마스타
		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		rsget("socid") = targetid  '' - 입고브랜드
		rsget("chargeid") = reguser
		rsget("divcode") = "002"   '' - 가맹용입고
		rsget("vatcode") = "008"
		rsget("comment") = comment + VBCRLF + "가맹점 개별위탁 주문서 " + orgbaljucode + " 입고처리"
		rsget("ipchulflag") = "I"

		rsget.update
		iid = rsget("id")
		rsget.close

		baljucode = "ST" + Format00(6,Right(CStr(iid),6))

		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = rsget("socname_kor")
		end if
		rsget.close


		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " ,socname='" + targetname + "'" + VBCrlf
		sqlStr = sqlStr + " ,chargename='" + regname + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 디테일 입력
		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.buycash,"
		sqlStr = sqlStr + " d.realitemno, getdate(),getdate(),d.buycash,'W',d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d,"
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VBCrlf
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " and deldt is null"
		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 마스타 업데이트
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
	elseif (divcode="251") then
		''매입반품->오프재고
	elseif (divcode="261") then
		''오프재고->가맹점출고
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
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash,"
		sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),i.buycash,'W',d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
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
		''[온라인위탁재고->가맹점용매입] 인경우 온라인 내역에 출고로 잡히고 가맹점으로 매입입고됩니다. 입고 확정
		''xxxxx온라인내역에 출고로 잡히고 가맹점 매입입고. 변경
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
		rsget("socid") = "streetshop800"
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
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, i.sellcash, i.buycash,"
		sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),0,'W',d.itemgubun,d.itemname,d.itemoptionname,d.makerid" + VBCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d," + VBCrlf
		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VBCrlf
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and d.itemgubun='10'" + VBCrlf
		sqlStr = sqlStr + " and d.itemid=i.itemid" + VBCrlf
		sqlStr = sqlStr + " and d.deldt is null" + VBCrlf
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
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash,"
		sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),i.buycash,'M',d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
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
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	rsget.Open sqlStr, dbget, 1

elseif mode="modidetail" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " set baljuitemno = " + baljuitemno  + vbCrlf
	sqlStr = sqlStr + " ,realitemno = " + realitemno  + vbCrlf
	sqlStr = sqlStr + " ,comment = '" + comment + "'" + vbCrlf

	if sellcash<>"" then
		sqlStr = sqlStr + " ,sellcash = " + sellcash + "" + vbCrlf
	end if
	if suplycash<>"" then
		sqlStr = sqlStr + " ,suplycash = " + suplycash + "" + vbCrlf
	end if
	if buycash<>"" then
		sqlStr = sqlStr + " ,buycash = " + buycash + "" + vbCrlf
	end if

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
elseif mode="segumil" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set segumdate='" + datestr + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
elseif mode="ipkumil" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set ipkumdate='" + datestr + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
elseif mode="remijumun" then

	''미배송주문 내역 체크
	sqlStr = " select count(idx) as cnt  from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and baljuitemno-realitemno>0"
	sqlStr = sqlStr + " and (comment='3일내출고' or comment='5일내출고')"
	sqlStr = sqlStr + " and deldt is null"
'response.write sqlStr
	rsget.Open sqlStr, dbget, 1
		itemexists = (rsget("cnt")>0)
	rsget.Close

	sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	sqlStr = sqlStr + " and clinkcode  is not null"
	sqlStr = sqlStr + " and clinkcode<>''"
	rsget.Open sqlStr, dbget, 1
		itemAlreadyExists = (rsget("cnt")>0)
	rsget.Close

	if Not itemexists then
		response.write "<script>alert('재 주문할 내역이 없습니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	elseif itemAlreadyExists then
		response.write "<script>alert('재 주문서가 이미 작성되어 있습니다. 작성할 수 없습니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if


	''//미배송 주문서 작성
	sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
		targetid = rsget("targetid")
		targetname = rsget("targetname")
		baljuid = rsget("baljuid")
		baljuname = rsget("baljuname")
		reguser = session("ssBctId")
		regname = session("ssBctCname")
		divcode = rsget("divcode")
		vatinclude = rsget("vatinclude")
		targetname = rsget("targetname")
		obaljucode = rsget("baljucode")
		cwFlag     = rsget("cwFlag")
	rsget.Close


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
	rsget("scheduledate") = datestr
	rsget("statecd") = "0"
	rsget("comment") = obaljucode + " 미배송건 재작성"
    rsget("cwFlag")  = cwFlag

	rsget.update
	iid = rsget("idx")
	rsget.close

	baljucode = "RJ" + Format00(6,Right(CStr(iid),6))

	''디테일 저장
	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
	sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	sqlStr = sqlStr + " baljuitemno-realitemno,baljuitemno-realitemno,baljudiv" + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and baljuitemno-realitemno>0"
	sqlStr = sqlStr + " and comment='5일내출고'"
	sqlStr = sqlStr + " and deldt is null"

	rsget.Open sqlStr, dbget, 1


	''서머리 저장
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


	''브랜드 리스트
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

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " , obaljucode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1


	''원발주서에 링크코드 저장.
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set clinkcode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1

	response.write "<script>alert('재 주문서가 작성되어 있습니다.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif mode="returnjumun" then
	itemexists = true

	sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	sqlStr = sqlStr + " and clinkcode  is not null"
	sqlStr = sqlStr + " and clinkcode<>''"
	rsget.Open sqlStr, dbget, 1
		itemAlreadyExists = (rsget("cnt")>0)
	rsget.Close

	if Not itemexists then
		response.write "<script>alert('재 주문할 내역이 없습니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	elseif itemAlreadyExists then
		'response.write "<script>alert('재 주문서가 이미 작성되어 있습니다. 작성할 수 없습니다.');</script>"
		'response.write "<script>window.close();</script>"
		'dbget.close()	:	response.End
	end if


	''//미배송 주문서 작성
	sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
		targetid = rsget("targetid")
		targetname = rsget("targetname")
		baljuid = rsget("baljuid")
		baljuname = rsget("baljuname")
		reguser = session("ssBctId")
		regname = session("ssBctCname")
		divcode = rsget("divcode")
		vatinclude = rsget("vatinclude")
		targetname = rsget("targetname")
		obaljucode = rsget("baljucode")
	rsget.Close

	if baljuid<>"streetshop875" and baljuid <> "streetshop999" then
		response.write "<script>alert('streetshop875,streetshop999 만 작성 가능');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("targetid") = targetid
	rsget("targetname") = targetname

	''임시.
	if baljuid="streetshop875" then
		rsget("baljuid") = "streetshop875"
		rsget("baljuname") = "Singapore"
    else
        rsget("baljuid") = baljuid
		rsget("baljuname") = baljuname
	end if

	rsget("reguser") = reguser
	rsget("regname") = regname
	rsget("divcode") = divcode
	rsget("vatinclude") = vatinclude
	rsget("scheduledate") = datestr
	rsget("statecd") = " "
	rsget("comment") = obaljucode + " 반품 작성."

	rsget.update
	iid = rsget("idx")
	rsget.close

	baljucode = "RJ" + Format00(6,Right(CStr(iid),6))

	''디테일 저장
	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
	sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	sqlStr = sqlStr + " baljuitemno*-1,realitemno*-1,baljudiv" + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and deldt is null"

	rsget.Open sqlStr, dbget, 1


	''서머리 저장
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


	''브랜드 리스트
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

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " , obaljucode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1


	''원발주서에 링크코드 저장.
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set clinkcode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1

	response.write "<script>alert('반품 주문서가 작성되었습니다.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif mode="duplicatejumun" then

	''//미배송 주문서 작성
	sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
		targetid = rsget("targetid")
		targetname = rsget("targetname")
		baljuid = cpbaljuid
		baljuname = ""
		reguser = session("ssBctId")
		regname = session("ssBctCname")
		divcode = rsget("divcode")
		vatinclude = rsget("vatinclude")
		targetname = rsget("targetname")
		obaljucode = rsget("baljucode")
		ostatecd   = rsget("statecd")
		cwFlag     = rsget("cwFlag")
	rsget.Close


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
	rsget("scheduledate") = datestr
	rsget("statecd") = ostatecd
	rsget("comment") = obaljucode + " 복사 주문서 작성"
    rsget("cwFlag") = cwFlag
	rsget.update
	iid = rsget("idx")
	rsget.close

	baljucode = Left(obaljucode,2) + Format00(6,Right(CStr(iid),6))

	''디테일 저장
	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
	sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv" + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and deldt is null"

	rsget.Open sqlStr, dbget, 1


	''서머리 저장
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

    sqlStr = " update M"
    sqlStr = sqlStr + " set baljuname=c.socname_kor"
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master M" + vbCrlf
    sqlStr = sqlStr + "     Join [db_user].[dbo].tbl_user_c c"
    sqlStr = sqlStr + "     on M.baljuid=c.userid"+ vbCrlf
    sqlStr = sqlStr + " where M.idx=" + CStr(iid)

    dbget.Execute sqlStr

	''브랜드 리스트
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

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " , obaljucode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
	sqlStr = sqlStr + " where idx=" + CStr(iid)
	rsget.Open sqlStr, dbget, 1


	response.write "<script>alert('재 주문서가 작성되어 있습니다.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End

elseif mode="chulgoproc" then

	''합계 재작성
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



	''기본 master 정보
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set comment='" + comment + "'"  + vbCrlf
	sqlStr = sqlStr + " ,scheduledate='" + yyyymmdd + "'" + vbCrlf

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
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1

	''detail  마진 flag 설정
	''sqlStr = "update [db_storage].[dbo].tbl_ordersheet_detail "
	''sqlStr = sqlStr + " set defaultmaginflag=i.mwdiv"
	''sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
	''sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
	''sqlStr = sqlStr + " and d.itemgubun='10'"
	''sqlStr = sqlStr + " and d.itemid=i.itemid"
	''rsget.Open sqlStr, dbget, 1
	''defaultmaginflag
	''buymaginflag
	''suplymaginflag


	''출고 마스타에 입력. *-1
		sqlStr = "select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and d.realitemno<>0"

		rsget.Open sqlStr, dbget, 1
			itemexists = rsget("cnt")>0
		rsget.close

		if itemexists then
			'1.온라인 출고 마스타
			sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
			rsget.Open sqlStr,dbget,1,3
			rsget.AddNew
			rsget("code") = ""
			''출고처
			rsget("socid") = shopid
			rsget("socname") = baljuname
			rsget("chargeid") = reguser
			rsget("divcode") = "006"
			rsget("vatcode") = "008"
			rsget("comment") = orgbaljucode + " 주문 자동출고처리"
			rsget("chargename") = regname
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
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
			sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash,"
			sqlStr = sqlStr + " d.realitemno*-1, getdate(),getdate(),d.buycash,d.ipgoflag,d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d"
			''sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i"
			''sqlStr = sqlStr + " on d.itemgubun='10' and d.itemid=i.itemid"
			sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
			sqlStr = sqlStr + " and deldt is null"
			sqlStr = sqlStr + " and d.realitemno<>0"

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


			''출고된 내역 한정판매설정
			'발주서작성단계에서 한정수량을 줄인다. - 옵션별로 변경 요망.
			if (limitflag="true") then
				response.write "limitflag"

				sqlstr = " update [db_item].[dbo].tbl_item_option"
				sqlstr = sqlstr + " set optlimitsold=(case when (optlimitno - (optlimitsold - T.chulno)<0) then optlimitno else (optlimitsold - T.chulno) end)"
				sqlstr = sqlstr + " from "
				sqlstr = sqlstr + " ("
				sqlstr = sqlstr + " 	select d.itemid, d.itemoption, Min(d.itemno) as chulno"
				sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d,"
				sqlstr = sqlstr + " 	[db_item].[dbo].tbl_item_option i"
				sqlstr = sqlstr + " 	where d.mastercode = '" + CStr(baljucode) + "'"
				sqlStr = sqlStr + " 	and d.iitemgubun='10'"
				sqlstr = sqlstr + " 	and d.itemid=i.itemid"
				sqlstr = sqlstr + " 	and d.itemoption=i.itemoption"
				sqlstr = sqlstr + " 	and d.deldt is NULL"
				sqlstr = sqlstr + " 	and d.itemno<>0"
				sqlstr = sqlstr + " 	and i.optlimityn='Y'"
				sqlstr = sqlstr + " 	group by d.itemid, d.itemoption"
				sqlstr = sqlstr + " ) as T"
				sqlstr = sqlstr + " where [db_item].[dbo].tbl_item_option.itemid=T.itemid"
				sqlstr = sqlstr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"

				dbget.Execute sqlstr

				''상품한정수량
				sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
				sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
				sqlStr = sqlStr + " from (" + VBCrlf
				sqlStr = sqlStr + " 	select itemid, sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
				sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
				sqlStr = sqlStr + " 	where itemid in ("  + VBCrlf
				sqlstr = sqlstr + " 		select distinct d.itemid " + VBCrlf
				sqlstr = sqlstr + " 		from [db_storage].[dbo].tbl_acount_storage_detail d," + VBCrlf
				sqlstr = sqlstr + " 		[db_item].[dbo].tbl_item i" + VBCrlf
				sqlstr = sqlstr + " 		where d.mastercode = '" + CStr(baljucode) + "'" + VBCrlf
				sqlstr = sqlstr + " 		and d.itemid=i.itemid" + VBCrlf
				sqlStr = sqlStr + " 		and d.iitemgubun='10'" + VBCrlf
				sqlstr = sqlstr + " 		and d.deldt is NULL" + VBCrlf
				sqlstr = sqlstr + " 		and d.itemno<>0" + VBCrlf
				sqlstr = sqlstr + " 		and i.limityn='Y'" + VBCrlf
				sqlStr = sqlStr + " 	) "  + VBCrlf
				sqlStr = sqlStr + "		group by itemid" + VBCrlf
				sqlStr = sqlStr + " ) T" + VBCrlf
				sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.itemid" + VBCrlf
				sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.optioncnt>0"

				dbget.Execute sqlstr
			end if
		end if

		''오프라인용 입고 입력
		''####### 출고마스타 #######
		''sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
		''sqlStr = sqlStr + " (chargeid,shopid,divcode,totalsellcash,totalsuplycash,"
		''sqlStr = sqlStr + " vatcode,scheduledate,linkidx)"
		''sqlStr = sqlStr + " select '10x10',socid,divcode,totalsellcash*-1,totalsuplycash*-1,"
		''sqlStr = sqlStr + " vatcode,scheduledt,id"
		''sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master"
		''sqlStr = sqlStr + " where id = " + CStr(iid) + ""
		''rsget.Open sqlStr, dbget, 1

		''sqlStr = "select IDENT_CURRENT('[db_shop].[dbo].tbl_shop_ipchul_master') as idx"
		''rsget.Open sqlStr, dbget, 1
		''	regidx = rsget("idx")
		''rsget.Close

		''####### 출고디테일 #######
		''sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail"
		''sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption,designerid,sellcash,"
		''sqlStr = sqlStr + " suplycash,itemno,linkidx)"
		''sqlStr = sqlStr + " select " + CStr(regidx) + ",d.iitemgubun,d.itemid,d.itemoption,d.imakerid,"
		''''sqlStr = sqlStr + " d.sellcash,d.suplycash,d.itemno*-1,d.id"
		''sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail d"
		''sqlStr = sqlStr + " where d.mastercode='" + baljucode + "'"
		''sqlStr = sqlStr + " and d.deldt is NUll"
		''sqlStr = sqlStr + " and d.itemno<>0"
		''rsget.Open sqlStr, dbget, 1


		''상태변경
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
		sqlStr = sqlStr + " set statecd='7'" + vbCrlf
		sqlStr = sqlStr + " ,ipgodate='" + ipgodate + "'" + vbCrlf
		sqlStr = sqlStr + " ,alinkcode='" + baljucode + "'" + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(masteridx)
		rsget.Open sqlStr, dbget, 1

elseif mode="cpsheet" Then

	If (request("suplyer") <> "10x10") Then
		response.write "<script language='javascript'>"
		response.write "alert('공급자는 10x10 만 가능합니다.');"
		response.write "</script>"
		dbget.close()	:	response.End
	End If

	sqlStr = " select distinct s.makerid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail s "
	sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and s.makerid = d.makerid "
	sqlStr = sqlStr + " 		and d.shopid = '" & shopid & "' "
	sqlStr = sqlStr + " 		and d.comm_cd not in ('B012','B022') "
	sqlStr = sqlStr + " 		and d.comm_cd in ('B031','B011') "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and s.masteridx = " & idx
	sqlStr = sqlStr + " 	and d.shopid is NULL "
	sqlStr = sqlStr + " order by s.makerid "
	''response.Write sqlStr

	brandlist = ""
	rsget.Open sqlStr, dbget, 1
		do until rsget.eof
			brandlist = brandlist + rsget("makerid") + ","
			rsget.movenext
		loop
	rsget.Close

	if brandlist<>"" then
		brandlist = Left(brandlist,Len(brandlist)-1)
	end If

	If (brandlist <> "") Then
		response.write "<script language='javascript'>"
		response.write "alert('ERROR: 계약없는 브랜드 존재!!\n\n" & brandlist & "');"
		response.write "</script>"
		dbget.close()	:	response.End
	End If

elseif mode="delalinkipchul" then
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where code='" + alinkcode + "'"
	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master " + vbCrlf
	sqlStr = sqlStr + " set alinkcode=NULL " + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)
	rsget.Open sqlStr, dbget, 1
end if



if  (mode="addshopjumun") or (mode="chulgoproc") then
	refer = "/admin/fran/jumunlist.asp?menupos=497"
elseif  (mode="delmaster") then
	if ourl<>"" then
		refer = "/admin/fran/" + ourl + "?menupos=530&page=" + opage + "&shopid=" + oshopid + "&statecd=" + ostatecd + "&desinger=" + odesinger
	else
		refer = "/admin/fran/jumunlist.asp?menupos=497&page=" + opage + "&shopid=" + oshopid + "&statecd=" + ostatecd + "&desinger=" + odesinger
	end if
elseif ((mode="segumil") or (mode="ipkumil")) then
	response.write "<script language='javascript'>"
	response.write "alert('저장 되었습니다.');"
	response.write "window.close();"
	response.write "</script>"
	dbget.close()	:	response.End
elseif (mode = "cpsheet") Then
	refer = "/common/offshop/popshopjumunitem.asp?suplyer=10x10&shopid=" & shopid & "&idx=0&cwflag=" & request("cwFlag") & "&cp_idx=" & idx
	response.write "<script language='javascript'>"
	response.write "alert('상품 추가 페이지로 이동합니다.');"
	response.write "parent.location.replace('" & refer & "');"
	response.write "</script>"
	dbget.close()	:	response.End
end if
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

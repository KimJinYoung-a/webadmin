<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim mode,yyyymmdd,baljuid,targetid ,reguser, divcode, vatinclude ,masteridx, totalsellcash, tmpmakerid
dim beasongdate, songjangdiv, songjangname, songjangno ,alinkcode ,idx
dim comment, targetname, baljuname, regname, statecd ,opage, ourl
dim datestr, orgbaljucode ,ipgodate ,shopid ,waitflag, limitflag ,oshopid,ostatecd,odesinger
dim itemgubunarr, itemarr, itemoptionarr ,i,cnt,sqlStr
dim sellcasharr,suplycasharr,buycasharr,itemnoarr,designerarr,detailidxarr,baljuitemnoarr,realitemnoarr,commentarr
dim ipgoflagarr, defaultmaginflagarr, buymaginflagarr, suplymaginflagarr
dim sellcash, suplycash, buycash, baljuitemno ,itemgubun, itemid, itemoption ,realitemno
dim itemAlreadyExists ,brandlist ,iid,baljucode, IsForeignOrder, IsForeign_confirmed
dim itemexists, obaljucode , regidx ,AssignedRows
dim cpbaljuid , cwflag, foreign_statecd, brandcount
dim uniqregdate, errMSG, foreign_sellcasharr, foreign_suplycasharr
dim finishname, loginsite, currencyUnit, foreign_sellcash, foreign_suplycash , countryCd, exchangeRate
dim addshopid, newiid,newtargetid,newbaljuid, currentstatecd
dim ipchulflag
	foreign_statecd = request("foreign_statecd")
	cwflag = request("cwflag")
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
	cpbaljuid = request("cpbaljuid")

	uniqregdate = request("uniqregdate")

	''작성중인경우.
	waitflag = request("waitflag")
	limitflag = request("limitflag")

	foreign_sellcasharr = request("foreign_sellcasharr")
	foreign_suplycasharr = request("foreign_suplycasharr")
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

	itemgubun = replace(request("itemgubun"),"|","")
	itemid		= replace(request("itemid"),"|","")
	itemoption	= replace(request("itemoption"),"|","")
	sellcash	= replace(request("sellcash"),"|","")
	suplycash	= replace(request("suplycash"),"|","")
	buycash		= replace(request("buycash"),"|","")
	baljuitemno	= replace(request("baljuitemno"),"|","")
	realitemno  = replace(request("realitemno"),"|","")

	finishname = html2db(session("ssBctCname"))

	addshopid = request("addshopid")

IsForeignOrder = false		'/업체접수주문
IsForeign_confirmed = false		'/업체접수주문 컨펌완료여부

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="addshopjumun" then

	'// ========================================================================
	if (uniqregdate <> "") then
		'// 등록자 아이디 + 시간을 가지고 중복입력 체크
		sqlStr = "select top 1 idx from db_storage.dbo.tbl_ordersheet_master "
		sqlStr = sqlStr + " where regdate = '" + CStr(uniqregdate) + "' and reguser = '" + CStr(reguser) + "' "

		errMSG = ""
		rsget.Open sqlStr, dbget, 1
'		if Not rsget.Eof then
'		errMSG = "이미 저장내역이 저장되었습니다.(중복입력)"
	'	end if
		rsget.close

		if (errMSG <> "") then
			response.write "<script>alert('" + CStr(errMSG) + "');</script>"
			response.write errMSG
			dbget.close()	:	response.End
		end if
	end if

	sqlStr = "select top 1"
	sqlStr = sqlStr & " u.userid, u.shopname, isNULL(u.currencyUnit,'USD') as currencyUnit, isnull(u.countrylangcd,'EN') as countrylangcd"
	sqlStr = sqlStr & " , loginsite, isNULL(r.exchangeRate,1120) as exchangeRate"
	sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_user u"
	sqlStr = sqlStr & " join db_item.dbo.tbl_exchangeRate r"
	sqlStr = sqlStr & " 	on u.currencyUnit = r.currencyUnit"
	sqlStr = sqlStr & " 	and u.countrylangcd = r.countrylangcd"
	sqlStr = sqlStr & " 	and r.sitename='WSLWEB'"
	sqlStr = sqlStr & " where u.isusing = 'Y' and u.userid ='"& baljuid &"'"

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		currencyUnit = rsget("currencyUnit")
		countryCd = rsget("countrylangcd")
		exchangeRate = rsget("exchangeRate")
		loginsite = rsget("loginsite")
	end if
	rsget.close

	if targetid="10x10" then
		targetname = "텐바이텐"
	else
		sqlStr = " select top 1 socname_kor, socname from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			if loginsite="WSLWEB" then
				if countryCd="KR" then
					targetname = db2html(rsget("socname_kor"))
				else
					targetname = db2html(rsget("socname"))
				end if
			else
				targetname = db2html(rsget("socname_kor"))
			end if
		end if
		rsget.close
	end if

	if baljuname="" then
		sqlStr = " select top 1 socname_kor, socname from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + baljuid + "'"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			if loginsite="WSLWEB" then
				if countryCd="KR" then
					baljuname = db2html(rsget("socname_kor"))
				else
					baljuname = db2html(rsget("socname"))
				end if
			else
				baljuname = db2html(rsget("socname_kor"))
			end if
		end if
		rsget.close
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"

	'response.write sqlStr &"<Br>"
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

	if getcwflag(baljuid,"B013") = "1" then
		cwflag = cwflag
	else
		cwflag = "0"
	end if

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
	sqlStr = sqlStr + " ,targetname='" + html2db(targetname) + "'" + VbCrlf
	sqlStr = sqlStr + " ,baljuname='" + html2db(baljuname) + "'" + VbCrlf
	sqlStr = sqlStr + " ,cwflag='" + cwflag + "'" + VbCrlf

	if (uniqregdate <> "") then
		sqlStr = sqlStr + " ,regdate='" + CStr(uniqregdate) + "' " + VBCrlf
	end if

	if loginsite = "WSLWEB"	 then
		sqlStr = sqlStr + " ,currencyUnit='"+currencyUnit+"', foreign_statecd= 7" + VBCrlf
		sqlStr = sqlStr + " ,sitename='"& loginsite &"'" + VBCrlf
	end if

	sqlStr = sqlStr + " where idx=" + CStr(iid)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	buycasharr = Left(buycasharr,Len(buycasharr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	if request("foreign_sellcasharr") <> "" then
		foreign_sellcasharr = Left(foreign_sellcasharr,Len(foreign_sellcasharr)-1)
		foreign_sellcasharr = split(foreign_sellcasharr,"|")
	end if
	if request("foreign_suplycasharr") <> "" then
		foreign_suplycasharr = Left(foreign_suplycasharr,Len(foreign_suplycasharr)-1)
		foreign_suplycasharr = split(foreign_suplycasharr,"|")
	end if

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
		brandcount = 0

		if baljuid<>"" and designerarr(i)<>"" then
			brandcount = getcontractbranditemcount(baljuid, designerarr(i))

			if brandcount=0 then
				response.write "<script type='text/javascript'>"
				response.write "	alert('계약된 브랜드("& designerarr(i) &")가 아닙니다. 브랜드 계약을 먼저 하세요.')"
				response.write "</script>"
				dbget.close() : response.end
			end if
		end if

		foreign_sellcash = 0
		foreign_suplycash = 0

		sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
		sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv"+vbCrlf

		if request("foreign_sellcasharr") <> "" then
			sqlStr = sqlStr & " , foreign_sellcash" & VBCrlf
		end if
		if request("foreign_suplycasharr") <> "" then
			sqlStr = sqlStr & " , foreign_suplycash" & VBCrlf
		end if

		sqlStr = sqlStr + " )"  + vbCrlf
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
			sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
			'sqlStr = sqlStr + "0," + vbCrlf		'//본사쪽 어드민만 들어가있슴..이유가??
		end if

		sqlStr = sqlStr + "'0' "+ vbCrlf

		if request("foreign_sellcasharr") <> "" then
			sqlStr = sqlStr & " ,'" & foreign_sellcasharr(i) & "'"& vbCrlf
		end if
		if request("foreign_suplycasharr") <> "" then
			sqlStr = sqlStr & " ,'"& foreign_suplycasharr(i) & "'"& vbCrlf
		end if

		sqlStr = sqlStr + " )"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf		'/발주시 해외 소비자가
	sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf			'/발주시 해외 공급가
	sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf			'/확정 해외 소비자가
	sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf		'/확정 해외 공급가
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select sum(sellcash*baljuitemno) as totsell" + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*baljuitemno) as totsupp" + vbCrlf
	sqlStr = sqlStr + " 	,sum(buycash*baljuitemno) as totbuy" + vbCrlf
	sqlStr = sqlStr + " 	,sum(sellcash*realitemno) as realtotsell" + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*realitemno) as realtotsupp" + vbCrlf
	sqlStr = sqlStr + " 	,sum(buycash*realitemno) as realtotbuy" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_sellcash*baljuitemno) as totforeign_sellcash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash" + vbCrlf
	sqlStr = sqlStr + "  	from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + "  	where masteridx="  + CStr(iid) + vbCrlf
	sqlStr = sqlStr + " 	and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(iid)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off iid,targetid,baljuid

	'// 매장반품
	ShopReturnUpdateBySheetIdx(iid)

	addshopid = Split(addshopid, ",")
	for i = 0 to UBOund(addshopid)
		if (Trim(addshopid(i)) <> "") then
			sqlStr = " exec [db_storage].[dbo].[usp_Ten_OrderSheel_Cpoy] '" & baljucode & "', '" & Trim(addshopid(i)) & "' "
		    rsget.CursorLocation = adUseClient
		    rsget.Open sqlStr, dbget, adOpenForwardOnly
		    ''if Not rsget.Eof then
				newiid = rsget("masteridx")
				newtargetid = rsget("targetid")
				newbaljuid = rsget("baljuid")
		    ''end if
		    rsget.close

			if Not IsNull(newtargetid) then
				'주문서 기준 기주문 업데이트
				PreOrderUpdateBySheetIdx(newiid)

				'//기주문 업데이트
				PreOrderUpdateByBrand_off newiid,newtargetid,newbaljuid

				'// 매장반품
				ShopReturnUpdateBySheetIdx(newiid)
			end if
		end if
	next

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

			'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="modeshopjumunmasterdetail" then
	''edit master
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set comment='" + comment + "'"  + vbCrlf
	sqlStr = sqlStr + " ,scheduledate='" + yyyymmdd + "'" + vbCrlf

	If (statecd <> "false") then
		sqlStr = sqlStr + " ,statecd='" + statecd + "'" + vbCrlf
	End If

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

	''response.write sqlStr &"<Br>"
	''response.end
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
	foreign_sellcasharr = split(foreign_sellcasharr,"|")
	foreign_suplycasharr = split(foreign_suplycasharr,"|")

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
			if (request("foreign_sellcasharr") <> "") then
				sqlStr = sqlStr + " , foreign_sellcash=" + foreign_sellcasharr(i)
				sqlStr = sqlStr + " , foreign_suplycash=" + foreign_suplycasharr(i)
			end if
			sqlStr = sqlStr + " where idx=" + detailidxarr(i)

			'response.write sqlStr &"<Br>"
			rsget.Open sqlStr, dbget, 1
		end if
	next

	'// 우선은 강제로 상품명 자른다.(account 테이블 등 상품명이 저장되는 여러테이블에서 상품명 입력가능한지 체크 필요) : skyer9, 2012-08-02
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	'sqlStr = sqlStr + " set itemname=convert(varchar(64),[db_shop].[dbo].tbl_shop_item.shopitemname)"	' 이렇게 하면 안됨. 해외 상품명 입력해 논게 다 날라감. 2018.02.27 용만
	'sqlStr = sqlStr + " ,itemoptionname=[db_shop].[dbo].tbl_shop_item.shopitemoptionname"
	sqlStr = sqlStr + " set itemname=convert(varchar(64),itemname) where"
	'sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " masteridx=" + CStr(masteridx)
	'sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	'sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	'sqlStr = sqlStr + " and [db_storage].[dbo].tbl_ordersheet_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	if (request("foreign_sellcasharr") <> "") then
		sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf
	end if
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " 	sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " 	sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " 	sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " 	sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " 	sum(buycash*realitemno) as realtotbuy " + vbCrlf
	if (request("foreign_sellcasharr") <> "") then
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*baljuitemno) as totforeign_sellcash " + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash " + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash " + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash " + vbCrlf
	end if
	sqlStr = sqlStr + " 	from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " 	and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off masteridx,targetid,baljuid

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="delshopjumunarr" then

	if Right(detailidxarr,1)="," then
		detailidxarr = Left(detailidxarr,Len(detailidxarr)-1)
	end if

	if Trim(detailidxarr<>"") then
		sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " where idx in (" + detailidxarr + ")"

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
	end if

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off masteridx,targetid,baljuid

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="shopjumunitemadd" then

	sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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
		sqlStr = sqlStr + "0," + vbCrlf
		sqlStr = sqlStr + "'0'"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
		sqlStr = sqlStr + " where shopitemid=" + itemid
		sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

		'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="shopjumunitemaddarr" then

	if foreign_statecd<>"" then
		IsForeignOrder=true

		if foreign_statecd="7" then
			IsForeign_confirmed = true
		end if
	else
		IsForeign_confirmed = true
	end if

	itemgubunarr    = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr         = Left(itemarr,Len(itemarr)-1)
	itemoptionarr   = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr     = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr    = Left(suplycasharr,Len(suplycasharr)-1)
	buycasharr      = Left(buycasharr,Len(buycasharr)-1)
	itemnoarr       = Left(itemnoarr,Len(itemnoarr)-1)

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

		'response.write sqlStr &"<Br>"
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

			'response.write sqlStr &"<Br>"
			rsget.Open sqlStr, dbget, 1
		else
			sqlStr = " select makerid from db_shop.dbo.tbl_shop_item where" + vbCrlf
			sqlStr = sqlStr + " itemgubun='" + itemgubunarr(i) + "'" + vbCrlf
			sqlStr = sqlStr + " and shopitemid=" + itemarr(i) + vbCrlf
			sqlStr = sqlStr + " and itemoption='" + itemoptionarr(i) + "'"

			'response.write sqlStr &"<Br>"
			rsget.Open sqlStr, dbget, 1
				tmpmakerid = rsget("makerid")
			rsget.close

			if tmpmakerid<>"" then
				brandcount = 0

				if baljuid<>"" and tmpmakerid<>"" then
					brandcount = getcontractbranditemcount(baljuid, tmpmakerid)

					if brandcount=0 then
						response.write "<script type='text/javascript'>"
						response.write "	alert('계약된 브랜드("& tmpmakerid &")가 아닙니다. 브랜드 계약을 먼저 하세요.')"
						response.write "</script>"
						dbget.close() : response.end
					end if
				end if
			end if

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
			sqlStr = sqlStr + "" + itemnoarr(i) + "," + vbCrlf
			sqlStr = sqlStr + "'0'"
			sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
			sqlStr = sqlStr + " where shopitemid=" + itemarr(i)
			sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'"
			sqlStr = sqlStr + " and itemoption='" + itemoptionarr(i) + "'"

			'response.write sqlStr &"<Br>"
			rsget.Open sqlStr, dbget, 1

			if (IsForeignOrder) then
				sqlStr = "update d set" + vbCrlf
				sqlStr = sqlStr & " d.foreign_sellcash=isnull(mp.orgprice,0)" + vbCrlf

				'소수점두째자리 반올림
				sqlStr = sqlStr & " ,d.foreign_suplycash=round( (isnull(mp.orgprice,0)*(100-IsNULL(s.defaultsuplymargin,100))/100) ,1)" + vbCrlf

				'/0.25단위 반올림
				'sqlStr = sqlStr & " ,d.foreign_suplycash=floor(( (isnull(mp.orgprice,0)*(100-IsNULL(s.defaultsuplymargin,100))/100) *100+25)/50)*50*1.0/100" + vbCrlf
				sqlStr = sqlStr & " from [db_storage].[dbo].tbl_ordersheet_master m" + vbCrlf
				sqlStr = sqlStr & " join [db_storage].[dbo].tbl_ordersheet_detail d" + vbCrlf
				sqlStr = sqlStr & " 	on m.idx=d.masteridx" + vbCrlf
				sqlStr = sqlStr & " left join db_item.dbo.tbl_item_multiLang_price mp" + vbCrlf
				sqlStr = sqlStr & " 	on m.sitename=mp.sitename" + vbCrlf
				sqlStr = sqlStr & " 	and d.itemid=mp.itemid" + vbCrlf
				sqlStr = sqlStr & " 	and m.currencyUnit=mp.currencyUnit" + vbCrlf
				sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_designer s" + vbCrlf
				sqlStr = sqlStr & " 	on m.baljuid=s.shopid" + vbCrlf
				sqlStr = sqlStr & " 	and d.makerid=s.makerid" + vbCrlf
				sqlStr = sqlStr & " where m.idx="&masteridx&"" + vbCrlf
				sqlStr = sqlStr & " and d.itemgubun='"&itemgubunarr(i)&"'" + vbCrlf
				sqlStr = sqlStr & " and d.itemid="&itemarr(i)&"" + vbCrlf
				sqlStr = sqlStr & " and d.itemoption='"&itemoptionarr(i)&"'"

				'response.write sqlStr &"<Br>"
				dbget.execute sqlStr
			end if
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
		sqlStr = sqlStr + " 					and m.shopid = '" & shopid & "' "
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
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf

	if IsForeignOrder then
		sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf		'/발주시 해외 소비자가
		sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf			'/발주시 해외 공급가
		sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf			'/확정 해외 소비자가
		sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf		'/확정 해외 공급가
	end if

	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf

	if IsForeignOrder then
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*baljuitemno) as totforeign_sellcash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash" + vbCrlf
	end if

	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off masteridx,targetid,baljuid

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="modimaster" then
	sqlStr = "select isnull(alinkcode,'') as alinkcode, statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
	sqlStr = sqlStr + " where m.idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		baljucode = rsget("alinkcode")
		currentstatecd = rsget("statecd")
	rsget.close

	if C_ADMIN_AUTH then
		' 현재가 출고완료이고 출고 이전 상태로 바꿀려고 하는경우
		if currentstatecd="7" and currentstatecd<>statecd and (statecd<"7" or statecd=" ") then
			if baljucode<>"" and not(isnull(baljucode)) then
				' 관련 입출고 내역 삭제
				sqlStr = "update [db_storage].[dbo].tbl_acount_storage_master" & vbCrlf
				sqlStr = sqlStr & " set deldt=getdate() where" & vbCrlf
				sqlStr = sqlStr & " code='" & baljucode & "'" & vbCrlf

				'response.write sqlStr &"<Br>"
				rsget.Open sqlStr, dbget, 1

				sqlStr = "update [db_storage].[dbo].tbl_ordersheet_master" & vbCrlf
				sqlStr = sqlStr & " set alinkcode=NULL where" & vbCrlf
				sqlStr = sqlStr & " idx="& masteridx &"" & vbCrlf

				'response.write sqlStr &"<Br>"
				rsget.Open sqlStr, dbget, 1

				'주문서 기준 기주문 업데이트
				PreOrderUpdateBySheetIdx(masteridx)

				'// 매장반품
				ShopReturnUpdateBySheetIdx(masteridx)
			end if
		end if
	end if

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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		rsget("socid") = targetid  '' - 입고브랜드
		rsget("chargeid") = reguser
		rsget("divcode") = "001"   '' - 가맹용입고			'// 801 => 001 2016-01-07, skyer9
		rsget("vatcode") = "008"
		rsget("comment") = comment + VBCRLF + "가맹점 개별매입 주문서 " + orgbaljucode + " 입고처리"
		rsget("ipchulflag") = "I"

		rsget.update
			iid = rsget("id")
		rsget.close

		baljucode = "ST" + Format00(6,Right(CStr(iid),6))


		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + targetid + "'"

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

	    '개별매입입고 재고 반영 : 신규등록된 입고에 대해, 서머리정보를 업데이트한다.
		QuickUpdateNewIpgoDetailSummary baljucode, false

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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			targetname = rsget("socname_kor")
		end if
		rsget.close

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " ,socname='" + html2db(targetname) + "'" + VBCrlf
		sqlStr = sqlStr + " ,chargename='" + html2db(regname) + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

		'''2.온라인 입고 디테일 입력
'		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
'		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
'		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
'		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.buycash,"
'		sqlStr = sqlStr + " d.realitemno, getdate(),getdate(),d.buycash,'W',d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
'		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d,"
'		sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VBCrlf
'		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
'		sqlStr = sqlStr + " and d.itemid=i.itemid"
'		sqlStr = sqlStr + " and deldt is null"

		'''2.온라인 입고 디테일 입력
		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
		sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
		sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
		sqlStr = sqlStr + " select '" + baljucode + "',d.itemid, d.itemoption, d.sellcash, d.buycash,"
		sqlStr = sqlStr + " d.realitemno, getdate(),getdate(),d.buycash,'W',d.itemgubun,d.itemname,d.itemoptionname,d.makerid"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d"
		sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and deldt is null"

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

	    '개별위탁입고 재고 반영 : 신규등록된 입고에 대해, 서머리정보를 업데이트한다.
		QuickUpdateNewIpgoDetailSummary baljucode, false

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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

	elseif (divcode="131") then
		''[온라인위탁재고->가맹점용매입] 인경우 온라인 내역에 출고로 잡히고 가맹점으로 매입입고됩니다. 입고 확정
		''xxxxx온라인내역에 출고로 잡히고 가맹점 매입입고. 변경
		'1.온라인 출고 가능내역인지 확인 itemgubun start with 10

		sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and itemgubun<>'10'"
		sqlStr = sqlStr + " and deldt is null"

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
	elseif (divcode="201") then
		''[온라인매입재고->가맹점용매입] 인경우 온라인 내역에 출고로 잡히고 가맹점으로 매입입고됩니다. 입고 확정
		'1.온라인 출고 가능내역인지 확인 itemgubun start with 10

		sqlStr = " select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and itemgubun<>'10'"
		sqlStr = sqlStr + " and deldt is null"

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="chforeign_statecd" then

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set foreign_statecd="&foreign_statecd&"" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	dbget.execute sqlStr

elseif mode="delmaster" then

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'//기주문 업데이트
	PreOrderUpdateByBrand_off masteridx,targetid,baljuid

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

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

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="deldetail" then
	sqlStr = " delete from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" + itemid + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="segumil" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set segumdate='" + datestr + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

elseif mode="ipkumil" then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " set ipkumdate='" + datestr + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
elseif mode="remijumun" then

	''미배송주문 내역 체크
	sqlStr = " select count(idx) as cnt  from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and baljuitemno-realitemno>0"
	sqlStr = sqlStr + " and (comment='3일내출고' or comment='5일내출고')"
	sqlStr = sqlStr + " and deldt is null"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		itemexists = (rsget("cnt")>0)
	rsget.Close

	sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	sqlStr = sqlStr + " and clinkcode  is not null"
	sqlStr = sqlStr + " and clinkcode<>''"

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
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


	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	''브랜드 리스트
	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1


	''원발주서에 링크코드 저장.
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set clinkcode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(idx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(idx)

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

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
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

	if baljuid<>"streetshop011" then
		response.write "<script>alert('streetshop011 만 작성 가능');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("targetid") = targetid
	rsget("targetname") = targetname

	''임시.
	if baljuid="streetshop011" then
		rsget("baljuid") = "streetshop001"
		rsget("baljuname") = "대학로본점"
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

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	''브랜드 리스트
	brandlist = ""
	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(iid)

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	''원발주서에 링크코드 저장.
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlStr = sqlStr + " set clinkcode='" + obaljucode + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(idx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(idx)

	response.write "<script>alert('재 주문서가 작성되어 있습니다.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif mode="duplicatejumun" then

	''//미배송 주문서 작성
	sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
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
	rsget.Close

	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0"

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	''서머리 저장
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + "     select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + "     sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + "     sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + "     sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + "     sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + "     sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + "     from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + "     where masteridx="  + CStr(iid) + vbCrlf
	sqlStr = sqlStr + "     and deldt is null" + vbCrlf
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

	'response.write sqlStr &"<Br>"
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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(iid)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(iid)

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

	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1

	''기본 master 정보
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set comment='" + comment + "'"  + vbCrlf
	sqlStr = sqlStr + " ,scheduledate='" + yyyymmdd + "'" + vbCrlf
	sqlStr = sqlStr + " ,finishname='" + finishname + "'" + vbCrlf


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

	'response.write sqlStr &"<Br>"
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

	totalsellcash=0
	sqlStr = "select totalsellcash"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m"
	sqlStr = sqlStr + " where m.idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		totalsellcash = rsget("totalsellcash")
	rsget.close

	''출고 마스타에 입력. *-1
	sqlStr = "select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " where d.masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and d.deldt is null"
	sqlStr = sqlStr + " and d.realitemno<>0"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		itemexists = rsget("cnt")>0
	rsget.close

    divcode = "006"
    ipchulflag = "S"

    sqlStr = " select top 1 userdiv "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " [db_partner].[dbo].tbl_partner "
    sqlStr = sqlStr & " where id = '" & shopid & "' "
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
        if rsget("userdiv") = "900" then
            divcode = "999"
            ipchulflag = "E"
        end if
	end if
	rsget.close

	if itemexists then
		'1.온라인 출고 마스타
		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		''출고처
		rsget("socid") = shopid
		rsget("socname") = baljuname
		rsget("chargeid") = reguser
		rsget("finishname") = finishname
		rsget("divcode") = divcode
		rsget("vatcode") = "008"
		rsget("comment") = orgbaljucode + " 주문 자동출고처리"
		rsget("chargename") = regname
		rsget("ipchulflag") = ipchulflag

		rsget.update
			iid = rsget("id")
		rsget.close

		baljucode = "SO" + Format00(6,Right(CStr(iid),6))

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)

		'response.write sqlStr &"<Br>"
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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

		'// 평균매입가 => 출고내역매입가
		sqlStr = " exec [db_storage].[dbo].[usp_Ten_AvgIpgoPriceToAccoundStorageBuycash] '" & baljucode & "' "
		dbget.Execute sqlStr

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

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1

		''출고된 내역 한정판매설정
		'발주서작성단계에서 한정수량을 줄인다. - 옵션별로 변경 요망.
		if (limitflag="true") and (now()>#06/01/2016 23:00:00#) then '' 업체위탁 매입전환 관련
			response.write "limitflag"

			'' item
			sqlstr = " update [db_item].[dbo].tbl_item"
			sqlstr = sqlstr + " set limitsold=limitsold - T.chulno"
			sqlstr = sqlstr + " from "
			sqlstr = sqlstr + " ("
			sqlstr = sqlstr + " 	select d.itemid, sum(d.itemno) as chulno"
			sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d"
			sqlstr = sqlstr + " 	where d.mastercode = '" + CStr(baljucode) + "'"
			sqlstr = sqlstr + " 	and d.deldt is NULL"
			sqlstr = sqlstr + " 	and d.itemno<0"
			sqlstr = sqlstr + " 	and d.iitemgubun='10'"
			sqlstr = sqlstr + " 	group by d.itemid"
			sqlstr = sqlstr + " ) as T"
			sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"
			sqlstr = sqlstr + " and [db_item].[dbo].tbl_item.limityn='Y'"

			dbget.Execute(sqlStr)

			''옵션있는상품
			sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
			sqlStr = sqlStr + " set optlimitsold=optlimitsold - T.chulno" + vbCrlf
			sqlStr = sqlStr + " from " + vbCrlf
			sqlstr = sqlstr + " ("
			sqlstr = sqlstr + " 	select d.itemid, d.itemoption, sum(d.itemno) as chulno"
			sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d"
			sqlstr = sqlstr + " 	where d.mastercode = '" + CStr(baljucode) + "'"
			sqlstr = sqlstr + " 	and d.deldt is NULL"
			sqlstr = sqlstr + " 	and d.itemno<0"
			sqlstr = sqlstr + " 	and d.iitemgubun='10'"
			sqlstr = sqlstr + " 	group by d.itemid, d.itemoption"
			sqlstr = sqlstr + " ) as T"
			sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
			sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
			sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

			dbget.Execute(sqlStr)

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

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

    ''재고반영 ''쿼리 확인
    if (baljucode<>"") then  ''2016/05/31
        sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & baljucode & "','','',0,'',''"

        'response.write sqlStr &"<Br>"
    	dbget.Execute sqlStr, AssignedRows

		'// 매장재고 반영
		sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & shopid & "', '" & baljucode & "', 'N' "
		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr
    end if

	if (AssignedRows>0) then
	    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
	end if

    ''오프 재고 재작성.
    sqlStr = " exec db_summary.dbo.sp_Ten_RealtimeStock_offjupsuAll" + vbCrlf

    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)

elseif mode="delalinkipchul" then
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + vbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + vbCrlf
	sqlStr = sqlStr + " where code='" + alinkcode + "'"

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master " + vbCrlf
	sqlStr = sqlStr + " set alinkcode=NULL " + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(masteridx)

	'// 매장반품
	ShopReturnUpdateBySheetIdx(masteridx)
elseif mode="insforgnprice" Then
	sqlStr = " update d " + vbCrlf
	sqlStr = sqlStr + " set d.foreign_sellcash = T.foreign_sellcash, d.foreign_suplycash = T.foreign_suplycash " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail d " + vbCrlf
	sqlStr = sqlStr + " 	join ( " + vbCrlf
	sqlStr = sqlStr + " 		select T.idx, T.foreign_sellcash, round( (isnull(T.foreign_sellcash,0)*(100-IsNULL(sd.defaultsuplymargin,100))/100) ,1) as foreign_suplycash " + vbCrlf
	sqlStr = sqlStr + " 		from " + vbCrlf
	sqlStr = sqlStr + " 			( " + vbCrlf
	sqlStr = sqlStr + " 				select d.idx, d.makerid, m.baljuid as shopid, si.orgsellprice " + vbCrlf
	sqlStr = sqlStr + " 				, (case " + vbCrlf
	sqlStr = sqlStr + " 					when e.linkPriceType = '2' then CEILING((si.orgsellprice*e.multipleRate/e.exchangeRate) * 2) / 2 " + vbCrlf
	sqlStr = sqlStr + " 					when e.linkPriceType = '1' then CEILING((si.shopitemprice*e.multipleRate/e.exchangeRate) * 2) / 2 " + vbCrlf
	sqlStr = sqlStr + " 					else 0 end) as foreign_sellcash " + vbCrlf
	sqlStr = sqlStr + " 				from " + vbCrlf
	sqlStr = sqlStr + " 					[db_storage].[dbo].tbl_ordersheet_master m " + vbCrlf
	sqlStr = sqlStr + " 					join [db_shop].[dbo].tbl_shop_user s " + vbCrlf
	sqlStr = sqlStr + " 					on " + vbCrlf
	sqlStr = sqlStr + " 						m.baljuid = s.userid " + vbCrlf
	sqlStr = sqlStr + " 					join db_item.dbo.tbl_exchangeRate e " + vbCrlf
	sqlStr = sqlStr + " 					on " + vbCrlf
	sqlStr = sqlStr + " 						1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 						and s.currencyUnit = e.currencyUnit " + vbCrlf
	sqlStr = sqlStr + " 						and s.loginsite = e.sitename " + vbCrlf
	sqlStr = sqlStr + " 						and s.countrylangcd = e.countrylangcd " + vbCrlf
	sqlStr = sqlStr + " 					join [db_storage].[dbo].tbl_ordersheet_detail d " + vbCrlf
	sqlStr = sqlStr + " 					on " + vbCrlf
	sqlStr = sqlStr + " 						m.idx = d.masteridx " + vbCrlf
	sqlStr = sqlStr + " 					join db_shop.dbo.tbl_shop_item si " + vbCrlf
	sqlStr = sqlStr + " 					on " + vbCrlf
	sqlStr = sqlStr + " 						1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 						and d.itemgubun = si.itemgubun " + vbCrlf
	sqlStr = sqlStr + " 						and d.itemid = si.shopitemid " + vbCrlf
	sqlStr = sqlStr + " 						and d.itemoption = si.itemoption " + vbCrlf
	sqlStr = sqlStr + " 				where " + vbCrlf
	sqlStr = sqlStr + " 					1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 					and m.idx = " & masteridx & " " + vbCrlf
	sqlStr = sqlStr + " 					and d.itemgubun = '90' " + vbCrlf
	sqlStr = sqlStr + " 					and (d.foreign_sellcash = 0 or d.foreign_suplycash = 0)" + vbCrlf
	sqlStr = sqlStr + " 			) T " + vbCrlf
	sqlStr = sqlStr + " 			join db_shop.dbo.tbl_shop_designer sd " + vbCrlf
	sqlStr = sqlStr + " 			on " + vbCrlf
	sqlStr = sqlStr + " 				1 = 1 " + vbCrlf
	sqlStr = sqlStr + " 				and T.shopid = sd.shopid " + vbCrlf
	sqlStr = sqlStr + " 				and T.makerid = sd.makerid " + vbCrlf
	sqlStr = sqlStr + " 	) T " + vbCrlf
	sqlStr = sqlStr + " 	on " + vbCrlf
	sqlStr = sqlStr + " 		d.idx = T.idx " + vbCrlf
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set " + vbCrlf
	sqlStr = sqlStr + " jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf		'/발주시 해외 소비자가
	sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf			'/발주시 해외 공급가
	sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf			'/확정 해외 소비자가
	sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf		'/확정 해외 공급가
	sqlStr = sqlStr + " from (select " + vbCrlf
	sqlStr = sqlStr + " 	sum(foreign_sellcash*baljuitemno) as totforeign_sellcash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash" + vbCrlf
	sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash" + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	'response.write sqlStr &"<Br>"
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
end if

if mode="chulgoproc" then
	'/반품 주문일경우 한정수량 조절 팝업 띄움	'/2016.05.23 한용민 생성
	if totalsellcash < 0 and C_ADMIN_USER then
%>
		<script type="text/javascript">
			location.replace('<%= refer %>');
			alert('저장 되었습니다.\n\n반품 주문 입니다. 필요시 팝업창에서 한정 수량을 조절해 주세요.');
			var addreg = window.open('/admin/fran/poplimitcheckipgoNew.asp?alinkcode=<%= baljucode %>','addreg','width=1024,height=768,scrollbars=yes,resizable=yes');
			addreg.focus();
		</script>
	<% else %>
		<script type="text/javascript">
			alert('저장 되었습니다.');
			location.replace('<%= refer %>');
		</script>
<%
	end if
else
%>
	<script type="text/javascript">
		alert('저장 되었습니다.');
		location.replace('<%= refer %>');
	</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->

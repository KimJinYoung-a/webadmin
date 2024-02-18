<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 입출고 리스트
' History : 2009.04.07 서동석 생성
'			2011.12.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim mode,itemgubunarr,itemarr,menupos ,franitempricearr, fransuplycasharr, cmsitempricearr, cmssuplycasharr
dim itemoptionarr,itempricearr,chargeidarr,isusingarr,extbarcodearr, itemsuplyarr ,currState
dim designer,itemgubun,itemname,sellcash,suplycash ,idx ,songjangdiv,songjangno, isreq
dim detailidxarr, currjungsanidarr, shopitemnamearr ,shopbuyprice ,i,cnt,sqlStr ,extbarcodeAlreadyExists
dim sellcasharr,suplycasharr,itemnoarr,designerarr ,cksel ,extbarcodeAlreadyExistsString ,stockitemexists
dim discountsellpricearr, shopbuypricearr ,chargeid,shopid,divcode,vatcode
	menupos = requestCheckVar(request("menupos"),10)
	mode = requestCheckVar(request("mode"),32)
	itemgubunarr = request("itemgubunarr")
	itemarr = request("itemarr")
	itemoptionarr = request("itemoptionarr")
	itempricearr = request("itempricearr")
	itemsuplyarr = request("itemsuplyarr")
	chargeidarr = request("chargeidarr")
	isusingarr = request("isusingarr")
	extbarcodearr = request("extbarcodearr")
	detailidxarr  = request("detailidxarr")
	currjungsanidarr = request("currjungsanidarr")
	shopitemnamearr = (request("shopitemnamearr"))
	franitempricearr = request("franitempricearr")
	fransuplycasharr = request("fransuplycasharr")
	cmsitempricearr = request("cmsitempricearr")
	cmssuplycasharr = request("cmssuplycasharr")
	designer = requestCheckVar(request("designer"),32)
	itemgubun = requestCheckVar(request("itemgubun"),4)
	itemname = requestCheckVar(request("itemname"),124)
	sellcash = requestCheckVar(request("sellcash"),20)
	suplycash = requestCheckVar(request("suplycash"),20)
	shopbuyprice = requestCheckVar(request("shopbuyprice"),20)
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	itemnoarr  = request("itemnoarr")
	designerarr = request("designerarr")
	discountsellpricearr = request("discountsellpricearr")
	shopbuypricearr = request("shopbuypricearr")
	fransuplycasharr = request("fransuplycasharr")
	chargeid = requestCheckVar(request("chargeid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	divcode = requestCheckVar(request("divcode"),3)
	vatcode = requestCheckVar(request("vatcode"),3)
	idx = requestCheckVar(request("idx"),10)
	cksel = request("cksel")
	songjangdiv = requestCheckVar(request("songjangdiv"),2)
	songjangno = requestCheckVar((html2db(request("songjangno")),32)
	isreq      = requestCheckVar(request("isreq"),10)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="arrins" then
	sqlStr = " select top 1 statecd from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
	
	if not rsget.Eof then
	    currState = rsget("statecd")
		if currState>0 then
			response.write "<script type='text/javascript'>alert('현재 입고대기 상태가 아닙니다.');</script>"
			response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
			dbget.close()	:	response.End
		end if
	end if
	
	rsget.Close

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno, reqno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(shopbuypricearr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + + "," + vbCrlf
		
		if (currState=-2) then
            sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + vbCrlf
        else
            sqlStr = sqlStr + "0" + vbCrlf
        end if
        
        sqlStr = sqlStr + "" + ")"
        
        'response.write sqlStr &"<Br>"
		dbget.Execute sqlStr
	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + "		select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " 	,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " 	from " + vbCrlf
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " 	and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

    'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
	
elseif mode="addipchullist" then
	dim scheduledt
	scheduledt = request("scheduledt")
    
    ''isreq 입고요청. Flag , isbaljuExists 'Y'
    
	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno,reguserid,isbaljuExists)"
	sqlStr = sqlStr + " values('" + chargeid + "'"
	sqlStr = sqlStr + " ,'" + shopid + "'"
	sqlStr = sqlStr + " ,'" + divcode + "'"
	sqlStr = sqlStr + " ,'" + vatcode + "'"
	sqlStr = sqlStr + " ,'" + scheduledt + "'"
	if (isreq<>"") then
	    sqlStr = sqlStr + " ,-2"
	else
	    sqlStr = sqlStr + " ,0"
	end if
	sqlStr = sqlStr + " ,'" + songjangdiv + "'"
	sqlStr = sqlStr + " ,'" + songjangno + "'"
	sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"
	if (isreq<>"") then
	    sqlStr = sqlStr + " ,'Y'"
	else
	    sqlStr = sqlStr + " ,'N'"
	end if
	sqlStr = sqlStr + " )"

    'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "

    'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close

''  Join 사용.
''	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + VbCrlf
''	sqlStr = sqlStr + " set songjangname=IsNULL(T.divname,'')" + VbCrlf
''	sqlStr = sqlStr + " from [db_order].[dbo].tbl_songjang_div T" + VbCrlf
''	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)
''	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_master.songjangdiv=T.divcd"
''
''	dbget.Execute(sqlStr)

	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	sellcasharr = Left(sellcasharr,Len(sellcasharr)-1)
	suplycasharr = Left(suplycasharr,Len(suplycasharr)-1)
	shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
	itemnoarr = Left(itemnoarr,Len(itemnoarr)-1)
	designerarr = Left(designerarr,Len(designerarr)-1)

	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno,reqno)"  + vbCrlf
		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemgubunarr(i),2) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemarr(i),10) + "," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(itemoptionarr(i),4) + "'," + vbCrlf
		sqlStr = sqlStr + "'" + requestCheckVar(designerarr(i),32) + "'," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(sellcasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(suplycasharr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(shopbuypricearr(i),20) + "," + vbCrlf
		sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + "," + vbCrlf
		if (isreq<>"") then
		    sqlStr = sqlStr + "" + requestCheckVar(itemnoarr(i),10) + vbCrlf
		else
		    sqlStr = sqlStr + "0" + vbCrlf
		end if
		sqlStr = sqlStr + "" + ")"
        
	    'response.write sqlStr &"<Br>"
		dbget.Execute sqlStr
	next

	''상품명 옵션명 :: 사용안함?..
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"
	
	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

	'' Master Summary
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " ,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " ,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)
	
	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)
else
	response.write mode
	dbget.close()	:	response.End

end if

if ((mode ="offitemreg") or (mode="arrins")) then
	if (InStr(refer,"&react=true")<1) then
		refer = refer + "&react=true"
	end if

elseif mode="addipchullist" then
	refer = "/common/offshop/shop_ipchullist.asp?menupos="&menupos&""
end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
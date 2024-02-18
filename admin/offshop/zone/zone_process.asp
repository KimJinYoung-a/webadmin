<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
Dim ozone,idx , i , shopid,zonename,racktype,unit,regdate ,isusing ,mode , sql , designer
dim itemgubunarr,shopitemidarr,itemoptionarr ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,zoneidx ,parameter
dim itemname , itemid ,zonegroup , searchtype , zonegroup_name ,zoneidxarr ,menupos ,datefg
dim cdl,cdm ,cds ,sqlsearch2 ,sqlsearch ,sqlStr ,StartDay ,endDay
	idx = requestCheckVar(request("idx"),10)
	zonegroup = requestCheckVar(request("zonegroup"),10)
	shopid = requestCheckVar(request("shopid"),32)
	zonename = requestCheckVar(request("zonename"),128)
	racktype = requestCheckVar(request("racktype"),10)
	unit = requestCheckVar(request("unit"),20)
	regdate = requestCheckVar(request("regdate"),30)
	isusing = requestCheckVar(request("isusing"),1)
	mode = requestCheckVar(request("mode"),32)
	itemgubunarr = request("itemgubunarr")
	shopitemidarr = request("shopitemidarr")
	itemoptionarr = request("itemoptionarr")
	zoneidxarr = request("zoneidxarr")
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	zoneidx = requestCheckVar(request("zoneidx"),10)
	designer = requestCheckVar(request("designer"),32)
	itemname = requestCheckVar(request("itemname"),124)
	itemid = requestCheckVar(request("itemid"),10)
	searchtype = requestCheckVar(request("searchtype"),1)
	zonegroup_name = requestCheckVar(request("zonegroup_name"),32)
	menupos = requestCheckVar(request("menupos"),10)
	datefg = requestCheckVar(request("datefg"),10)
	cdl = requestCheckVar(request("cdl"),3)
	cdm = requestCheckVar(request("cdm"),3)
	cds = requestCheckVar(request("cds"),3)
	'response.write itemgubunarr

parameter = "isusing="&isusing&"&shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&menupos="&menupos&""
parameter = parameter & "&zonename="&zonename&"&designer="&designer&"&itemname="&itemname&"&itemid="&itemid&"&zonegroup="&zonegroup&"&racktype="&racktype&"&searchtype="&searchtype&""
parameter = parameter & "&datefg="&datefg&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&""

dim ref
	ref = request.ServerVariables("HTTP_REFERER")
	
'//샵별구역등록
if mode = "zonereg" then
	
	'신규등록
	if idx = "" then
		sql = "insert into db_shop.dbo.tbl_shop_zone" + vbcrlf
		sql = sql & " (shopid,zonegroup,racktype,zonename,unit,isusing) values (" + vbcrlf
		sql = sql & " '"&shopid&"','"&zonegroup&"',"&racktype&",'"&html2db(zonename)&"',"&unit&",'"&isusing&"'" + vbcrlf
		sql = sql & " )"
		
		'response.write sql &"<br>"
		dbget.execute sql
	
	'//수정모드	
	else
		sql = "update db_shop.dbo.tbl_shop_zone set" + vbcrlf
		sql = sql & " shopid = '"&shopid&"'" + vbcrlf
		sql = sql & " ,zonegroup = "&zonegroup&"" + vbcrlf		
		sql = sql & " ,racktype = "&racktype&"" + vbcrlf
		sql = sql & " ,zonename = '"&html2db(zonename)&"'" + vbcrlf
		sql = sql & " ,unit = "&unit&"" + vbcrlf		
		sql = sql & " ,isusing = '"&isusing&"'" + vbcrlf
		sql = sql & " where idx = "&idx&""

		'response.write sql &"<br>"
		dbget.execute sql	
	end if
	
	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='zone.asp?menupos="&menupos&"';"
	response.write "</script>"

'//삽별 상품 구역 지정
elseif mode = "zoneitemreg" then
	
	itemgubunarr = split(itemgubunarr,",")
	shopitemidarr = split(shopitemidarr,",")
	itemoptionarr = split(itemoptionarr,",")	
	zoneidxarr = split(zoneidxarr,",")
	
	'/트랜젝션
	dbget.beginTrans
	
	'//구역지정안함 선택시
	if zoneidx = "0" then

		for i = 0 to ubound(itemgubunarr)-1

		'//기존 상품 부터 로그테이블에 enddate를 엎어침
		sql = "update db_shop.dbo.tbl_shop_zoneitem_log set" + vbcrlf
		sql = sql & " enddate = getdate()" + vbcrlf
		sql = sql & " where shopid='"&shopid&"'" + vbcrlf					
		sql = sql & " and itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
		sql = sql & " and shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
		sql = sql & " and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf
		sql = sql & " and enddate is null" + vbcrlf

		'response.write sql &"<br>"
		dbget.execute sql
		
		sql = ""
		sql = "delete from db_shop.dbo.tbl_shop_zoneitem" + vbcrlf		
		sql = sql & " where shopid='"&shopid&"'" + vbcrlf
		sql = sql & " and itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
		sql = sql & " and shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
		sql = sql & " and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf

		'response.write sql &"<br>"
		dbget.execute sql
		
		next
	
	'//구역지정
	else
			
		for i = 0 to ubound(itemgubunarr)-1
			
			'//존재하는 상품 엎어침				
			if zoneidxarr(i) <> "" then
				
				'//유저가 삑사리 내서 똑같은 내역을 등록 할시 제낀다.
				if zoneidxarr(i) <> zoneidx then
				
					'//기존 상품 부터 로그테이블에 enddate를 엎어침
					sql = "update db_shop.dbo.tbl_shop_zoneitem_log set" + vbcrlf
					sql = sql & " enddate = getdate()" + vbcrlf
					sql = sql & " where shopid='"&shopid&"'" + vbcrlf					
					sql = sql & " and itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
					sql = sql & " and shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
					sql = sql & " and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf
					sql = sql & " and enddate is null" + vbcrlf
					
					'response.write sql &"<br>"
					dbget.execute sql
								
					'//상품 수정
					sql = ""
					sql = "update db_shop.dbo.tbl_shop_zoneitem set" + vbcrlf
					sql = sql & " zoneidx = "&zoneidx&" , regdate = getdate()" + vbcrlf
					sql = sql & " where shopid='"&shopid&"'" + vbcrlf
					sql = sql & " and itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
					sql = sql & " and shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
					sql = sql & " and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf
			
					'response.write sql &"<br>"
					dbget.execute sql
	
					'//로그 처리
					sql = ""
					sql = "insert into db_shop.dbo.tbl_shop_zoneitem_log" + vbcrlf 
					sql = sql & " (shopid,itemgubun,shopitemid,itemoption,zoneidx,startdate,isusing)" + vbcrlf
					sql = sql & " 		select shopid,itemgubun,shopitemid,itemoption,zoneidx,regdate,'Y'" + vbcrlf
					sql = sql & " 		from db_shop.dbo.tbl_shop_zoneitem" + vbcrlf
					sql = sql & " 		where itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
					sql = sql & " 		and shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
					sql = sql & " 		and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf
					sql = sql & " 		and shopid='"&shopid&"'" + vbcrlf
								
					'response.write sql &"<br>"
					dbget.execute sql
					
				end if	
							
			'//없는 상품 신규등록
			else
				
				'//상품 등록
				sql = "insert into db_shop.dbo.tbl_shop_zoneitem" + vbcrlf 
				sql = sql & " (shopid,itemgubun,shopitemid,itemoption,zoneidx)" + vbcrlf
				sql = sql & " 		select '"&shopid&"',i.itemgubun,i.shopitemid,i.itemoption,"&zoneidx&"" + vbcrlf
				sql = sql & " 		from db_shop.dbo.tbl_shop_item i" + vbcrlf
				sql = sql & " 		left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
				sql = sql & " 		on zi.shopid = '"&shopid&"'" + vbcrlf
				sql = sql & " 		and i.itemgubun = zi.itemgubun" + vbcrlf 
				sql = sql & " 		and i.shopitemid = zi.shopitemid" + vbcrlf
				sql = sql & " 		and i.itemoption = zi.itemoption" + vbcrlf
				sql = sql & " 		where zi.shopid is null" + vbcrlf		
				sql = sql & " 		and i.itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
				sql = sql & " 		and i.shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
				sql = sql & " 		and i.itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf
		
				'response.write sql &"<br>"
				dbget.execute sql				
				
				'//로그 처리
				sql = ""
				sql = "insert into db_shop.dbo.tbl_shop_zoneitem_log" + vbcrlf 
				sql = sql & " (shopid,itemgubun,shopitemid,itemoption,zoneidx,startdate,isusing)" + vbcrlf
				sql = sql & " 		select shopid,itemgubun,shopitemid,itemoption,zoneidx,regdate,'Y'" + vbcrlf
				sql = sql & " 		from db_shop.dbo.tbl_shop_zoneitem" + vbcrlf
				sql = sql & " 		where itemgubun='"& requestCheckVar(itemgubunarr(i),2) &"'" + vbcrlf
				sql = sql & " 		and shopitemid="& requestCheckVar(shopitemidarr(i),10) &"" + vbcrlf
				sql = sql & " 		and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'" + vbcrlf
				sql = sql & " 		and shopid='"&shopid&"'" + vbcrlf
							
				'response.write sql &"<br>"
				dbget.execute sql
			end if
			
		next

	end if

	If Err.Number = 0 Then
	    dbget.CommitTrans
	Else
	    dbget.RollBackTrans
	End If

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	location.href='zone_item.asp?"&parameter&"'"
	response.write "</script>"

'//삽별 상품 구역 지정 검색 결과 모두 저장
elseif mode = "zoneitemregall" then

	if (yyyy1="") then yyyy1 = Cstr(Year(now()))
	if (mm1="") then mm1 = Cstr(Month(now()))
	if (dd1="") then dd1 = Cstr(day(now()))
	if (yyyy2="") then yyyy2 = Cstr(Year(now()))
	if (mm2="") then mm2 = Cstr(Month(now()))
	if (dd2="") then dd2 = Cstr(day(now()))
				
	StartDay = DateSerial(yyyy1, mm1, dd1)
	EndDay = DateSerial(yyyy2, mm2, dd2+1)
	
	'/앞단에 검색결과 한큐에 몽땅 저장을 위한 검색쿼리
	if shopid <> "" then
		sqlsearch = sqlsearch & " and m.shopid = '"&shopid&"'"
	end if						
	if designer<>"" then
		sqlsearch2 = sqlsearch2 + " and i.makerid='" + CStr(designer) + "'"
	end if	
	if itemid<>"" then
		sqlsearch2 = sqlsearch2 + " and i.shopitemid=" + itemid + ""
	end if
	if itemname<>"" then
		sqlsearch2 = sqlsearch2 + " and i.shopitemname like '%" + itemname + "%'"
	end if
	if CDL<>"" then
		sqlsearch2 = sqlsearch2 + " and i.catecdl='" + CDL + "'"
	end if
	if CDM<>"" then
		sqlsearch2 = sqlsearch2 + " and i.catecdm='" + CDM + "'"
	end if
	if cds<>"" then
		sqlsearch2 = sqlsearch2 + " and i.catecdn='" + cds + "'"
	end if

	'//주문일 기준
	if datefg = "jumun" then
		if StartDay<>"" then
			sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(StartDay) + "'"
		end if
		if EndDay<>"" then
			sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(EndDay) + "'"
		end if
		
	'//매출일 기준
	elseif datefg = "maechul" then
		if StartDay<>"" then
			sqlsearch = sqlsearch + " and m.IXyyyymmdd>='" + CStr(StartDay) + "'"
		end if
		if EndDay<>"" then
			sqlsearch = sqlsearch + " and m.IXyyyymmdd<'" + CStr(EndDay) + "'"
		end if
	else
		if StartDay<>"" then
			sqlsearch = sqlsearch + " and m.shopregdate>='" + CStr(StartDay) + "'"
		end if
		if EndDay<>"" then
			sqlsearch = sqlsearch + " and m.shopregdate<'" + CStr(EndDay) + "'"
		end if		
	end if	

	if isusing = "Y" then
		sqlsearch2 = sqlsearch2 & " and zi.zoneidx is not null"
	elseif isusing = "N" then
		sqlsearch2 = sqlsearch2 & " and zi.zoneidx is null"
	end if
	if zonegroup <> "" then
		sqlsearch2 = sqlsearch2 & " and z.zonegroup = "&zonegroup&""
	end if
	if racktype <> "" then
		sqlsearch2 = sqlsearch2 & " and z.racktype = "&racktype&""
	end if
	if searchtype = "M" then
		sqlsearch2 = sqlsearch2 & " and t.shopid is not null"
	end if
		
	'/트랜젝션
	dbget.beginTrans
	
	'//구역지정안함 선택시
	if zoneidx = "0" then
		
		'//기존 상품 부터 로그테이블에 enddate를 엎어침
		sqlStr = "update l set" + vbcrlf
		sqlStr = sqlStr & " l.enddate = getdate()" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " Join db_shop.dbo.tbl_shop_zoneitem_log l" + vbcrlf
		sqlStr = sqlStr + " 	on i.itemgubun=l.itemgubun" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid=l.shopitemid" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption=l.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and l.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " 	and l.enddate is null" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on t.shopid = zi.shopid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
						
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = "delete zi from" + vbcrlf		
		sqlStr = sqlStr & " [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on i.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and i.shopitemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and zi.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
		
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
		
	'//구역지정
	else
			
		'//존재하는 상품 엎어침				
		'//기존 상품 부터 로그테이블에 enddate를 엎어침
		sqlStr = "update l set" + vbcrlf
		sqlStr = sqlStr & " l.enddate = getdate()" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " Join db_shop.dbo.tbl_shop_zoneitem_log l" + vbcrlf
		sqlStr = sqlStr + " 	on i.itemgubun=l.itemgubun" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid=l.shopitemid" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption=l.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and l.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " 	and enddate is null" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on t.shopid = zi.shopid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and t.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem tzi" + vbcrlf
		sqlStr = sqlStr & " 	on i.itemgubun = tzi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and i.shopitemid = tzi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemoption = tzi.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and tzi.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " 	and tzi.zoneidx = "&zoneidx&"" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
		sqlStr = sqlStr & " and tzi.shopid is null"		'이미같은 구역에 지정되어 있는 상품일 경우 제낌
						
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
					
		'//상품 수정
		sqlStr = "update zi set" + vbcrlf
		sqlStr = sqlStr & " zi.zoneidx = "&zoneidx&" , zi.regdate = getdate()" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on i.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and i.shopitemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and zi.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem tzi" + vbcrlf
		sqlStr = sqlStr & " 	on i.itemgubun = tzi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and i.shopitemid = tzi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemoption = tzi.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and tzi.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " 	and tzi.zoneidx = "&zoneidx&"" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
		sqlStr = sqlStr & " and tzi.shopid is null"		'이미같은 구역에 지정되어 있는 상품일 경우 제낌
		
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
						
		'//없는 상품 신규등록
		'//상품 등록
		sqlStr = "insert into db_shop.dbo.tbl_shop_zoneitem" + vbcrlf 
		sqlStr = sqlStr & " (shopid,itemgubun,shopitemid,itemoption,zoneidx)" + vbcrlf
		sqlStr = sqlStr & " select '"&shopid&"',i.itemgubun,i.shopitemid,i.itemoption,"&zoneidx&"" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on i.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and i.shopitemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and zi.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
		sqlStr = sqlStr & " and zi.shopid is null"		'이미 등록되어 있는 상품 제낌
		
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr				
		
		'//로그 처리
		sqlStr = "insert into db_shop.dbo.tbl_shop_zoneitem_log" + vbcrlf 
		sqlStr = sqlStr & " (shopid,itemgubun,shopitemid,itemoption,zoneidx,startdate,isusing)" + vbcrlf
		sqlStr = sqlStr & " select zi.shopid,i.itemgubun,i.shopitemid,i.itemoption,zi.zoneidx,zi.regdate,'Y'" + vbcrlf
		sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_zoneitem zi" + vbcrlf
		sqlStr = sqlStr & " 	on zi.shopid = '"&shopid&"'" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemgubun = zi.itemgubun" + vbcrlf
		sqlStr = sqlStr & " 	and i.shopitemid = zi.shopitemid" + vbcrlf
		sqlStr = sqlStr & " 	and i.itemoption = zi.itemoption" + vbcrlf
		sqlStr = sqlStr + " left Join ("
		sqlStr = sqlStr & " 	select" + vbcrlf
		sqlStr = sqlStr & " 	m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " 	Join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=d.orderno" + vbcrlf
		sqlStr = sqlStr & " 	and m.cancelyn='N' and d.cancelyn='N' " & sqlsearch
		sqlStr = sqlStr & " 	group by m.shopid ,d.itemgubun ,d.itemid ,d.itemoption,d.makerid" + vbcrlf
		sqlStr = sqlStr & " 	,d.itemname,d.itemoptionname" + vbcrlf
		sqlStr = sqlStr & " ) t" + vbcrlf		
		sqlStr = sqlStr + " 	on i.itemgubun=t.itemgubun"
		sqlStr = sqlStr + " 	and i.shopitemid=t.itemid"
		sqlStr = sqlStr + " 	and i.itemoption=t.itemoption"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone z" + vbcrlf
		sqlStr = sqlStr & " 	on zi.zoneidx = z.idx" + vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_zone_common c"
		sqlStr = sqlStr & " 	on z.zonegroup = c.zonegroup"
		sqlStr = sqlStr & " 	and c.isusing='Y' and c.zonegroup_type = 'GROUP'"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_Cate_large cl " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cl.code_large" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_mid cm " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cm.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cm.code_mid" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_cate_small cs " + vbcrlf
		sqlStr = sqlStr + " 	on i.catecdl=cs.code_large"
		sqlStr = sqlStr + "		and i.catecdm=cs.code_mid"
		sqlStr = sqlStr + "		and i.catecdn=cs.code_small" + vbcrlf
		sqlStr = sqlStr + " left Join db_shop.dbo.tbl_shop_zoneitem_log tl" + vbcrlf
		sqlStr = sqlStr + " 	on i.itemgubun=tl.itemgubun" + vbcrlf
		sqlStr = sqlStr + " 	and i.shopitemid=tl.shopitemid" + vbcrlf
		sqlStr = sqlStr + " 	and i.itemoption=tl.itemoption" + vbcrlf
		sqlStr = sqlStr + " 	and tl.shopid='"&shopid&"'" + vbcrlf
		sqlStr = sqlStr + " 	and tl.enddate is null" + vbcrlf
		sqlStr = sqlStr + " 	and tl.zoneidx = "&zoneidx&"" + vbcrlf
		sqlStr = sqlStr & " where i.isusing='Y' " & sqlsearch2
		sqlStr = sqlStr & " and tl.shopid is null"		'이미같은 구역에 지정되어 있는 상품일 경우 제낌
								
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
	end if

	If Err.Number = 0 Then
	     'dbget.RollBackTrans
	    dbget.CommitTrans
	Else
	    dbget.RollBackTrans
	End If

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	location.href='zone_item.asp?"&parameter&"'"
	response.write "</script>"
	
'//그룹등록
elseif mode = "zonecommonedit" then

	'신규등록
	if zonegroup = "" then
		sql = "insert into db_shop.dbo.tbl_shop_zone_common" + vbcrlf
		sql = sql & " (zonegroup_name,zonegroup_type,isusing) values (" + vbcrlf
		sql = sql & " '"&html2db(zonegroup_name)&"','GROUP','"&isusing&"'" + vbcrlf
		sql = sql & " )"
		
		'response.write sql &"<br>"
		dbget.execute sql
	
	'//수정모드	
	else
		sql = "update db_shop.dbo.tbl_shop_zone_common set" + vbcrlf
		sql = sql & " zonegroup_name = '"&html2db(zonegroup_name)&"'" + vbcrlf
		sql = sql & " ,isusing = '"&isusing&"'" + vbcrlf
		sql = sql & " where zonegroup = "&zonegroup&""

		'response.write sql &"<br>"
		dbget.execute sql	
	end if
	
	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='"& ref &"';"
	response.write "</script>"
		
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

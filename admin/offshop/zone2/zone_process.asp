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
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->
<%
Dim ozone,idx , i , shopid,zonename,racktype,unit,regdate ,isusing ,mode , sql , designer
dim itemgubunarr,shopitemidarr,itemoptionarr ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,zoneidx ,parameter
dim itemname , itemid ,zonegroup , searchtype , zonegroup_name ,zoneidxarr ,datefg,menupos
dim cdl,cdm ,cds ,sqlsearch2 ,sqlsearch ,sqlStr ,StartDay ,endDay, makeridarr, makerid ,chzoneidx
dim zoneisusing ,managershopyn ,searchgubun ,empno
	searchgubun = RequestCheckVar(request("searchgubun"),16)	
	idx = requestCheckVar(request("idx"),10)
	empno = request("empno")
	zoneisusing = requestCheckVar(request("zoneisusing"),1)
	zonegroup = requestCheckVar(request("zonegroup"),32)
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
	datefg = requestCheckVar(request("datefg"),16)
	cdl = requestCheckVar(request("cdl"),3)
	cdm = requestCheckVar(request("cdm"),3)
	cds = requestCheckVar(request("cds"),3)
	makeridarr = requestCheckVar(request("makerid"),32)
	chzoneidx = requestCheckVar(request("chzoneidx"),10)

managershopyn = "N"
	
parameter = "isusing="&isusing&"&shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&menupos="&menupos&""
parameter = parameter & "&zonename="&zonename&"&designer="&designer&"&itemname="&itemname&"&itemid="&itemid&"&zonegroup="&zonegroup&"&racktype="&racktype&"&searchtype="&searchtype&""
parameter = parameter & "&datefg="&datefg&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&zoneidx="&zoneidx&"&zoneisusing="&zoneisusing&"&searchgubun="&searchgubun

dim ref
	ref = request.ServerVariables("HTTP_REFERER")

'//샵별구역등록
'' 트랜잭션 제외 2015/07/21
if mode = "zonereg" then

	''dbget.beginTrans
	
	'신규등록
	if idx = "" then	

		sql = "INSERT INTO [db_shop].[dbo].[tbl_shop_brand_zone]" + vbcrlf
		sql = sql & " (shopid,zonename,unit,isusing,managershopyn) values (" + vbcrlf
		sql = sql & " '"&shopid&"','"&html2db(zonename)&"',"&unit&",'"&isusing&"','"&managershopyn&"'" + vbcrlf
		sql = sql & " )"
		
		'response.write sql &"<br>"
		dbget.execute sql

		sql ="select SCOPE_IDENTITY() "
		
		'response.write sql &"<br>"
		rsget.Open sql, dbget
		
		IF not (rsget.EOF or rsget.BOF) THEN
			idx = rsget(0)
		END IF
		rsget.close

		if empno <> "" then
			empno = split(empno,",")
			
			if isarray(empno) then
				for i = 0 to ubound(empno)
					sql = "INSERT INTO db_shop.dbo.tbl_shop_brand_zone_manager" + vbcrlf
					sql = sql & " (zoneidx ,parttype ,empno) values (" + vbcrlf
					sql = sql & " "&idx&",'SHOP','"& requestCheckVar(trim(empno(i)),32) &"'"
					sql = sql & " )"
					
					'response.write sql &"<br>"
					dbget.execute sql
				next
				
				managershopyn = "Y"
				
				sql = "update [db_shop].[dbo].[tbl_shop_brand_zone] set" + vbcrlf
				sql = sql & " managershopyn = '"&managershopyn&"'" + vbcrlf 
				sql = sql & " where idx = "&idx&""

				'response.write sql &"<br>"
				dbget.execute sql
			end if
		end if	
				
	'//수정모드	
	else
		if empno <> "" then
			empno = split(empno,",")
			
			if isarray(empno) then
				sql = "delete from db_shop.dbo.tbl_shop_brand_zone_manager" + vbcrlf
				sql = sql & " where zoneidx = "&idx&""

				'response.write sql &"<br>"
				dbget.execute sql
									
				for i = 0 to ubound(empno)
					sql = "INSERT INTO db_shop.dbo.tbl_shop_brand_zone_manager" + vbcrlf
					sql = sql & " (zoneidx ,parttype ,empno) values (" + vbcrlf
					sql = sql & " "&idx&",'SHOP','"& requestCheckVar(trim(empno(i)),32) &"'"
					sql = sql & " )"
					
					'response.write sql &"<br>"
					dbget.execute sql
				next
				
				managershopyn = "Y"
			end if
		end if

		sql = "UPDATE [db_shop].[dbo].[tbl_shop_brand_zone] SET" + vbcrlf
		sql = sql & " shopid = '"&shopid&"'" + vbcrlf
		sql = sql & " ,zonename = '"&html2db(zonename)&"'" + vbcrlf
		sql = sql & " ,unit = "&unit&"" + vbcrlf		
		sql = sql & " ,isusing = '"&isusing&"'" + vbcrlf
		sql = sql & " ,managershopyn = '"&managershopyn&"'" + vbcrlf
		sql = sql & " WHERE idx = "&idx&""

		'response.write sql &"<br>"
		dbget.execute sql	
	end if

	If Err.Number = 0 Then
	    ''dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('OK');"
		response.write "	opener.location.reload();"
		response.write "	self.close();"
		'response.write "	location.href='zone.asp?menupos="&menupos&"';"
		response.write "</script>"
	Else
	    ''dbget.RollBackTrans
	    
		response.write "<script type='text/javascript'>"
		response.write "	alert('에러발생 관리자 문의 하세요');"
		response.write "	opener.location.reload();"
		response.write "	self.close();"
		'response.write "	location.href='zone.asp?menupos="&menupos&"';"
		response.write "</script>"
	End If
	
'//삽별 상품 구역 지정
elseif mode = "zoneitemreg" then
	
	'/트랜젝션
	''dbget.beginTrans
	sql = ""
	
	'//구역지정안함 선택시
	if chzoneidx = "0" then
		makeridarr = split(request("makerid"),",")

		sql = ""
		for i = 0 to ubound(makeridarr)
			makerid = requestCheckVar(Trim(makeridarr(i)),32)
					
			sql = sql & "DELETE [db_shop].[dbo].[tbl_shop_brand_zone_detail]" + vbcrlf		
			sql = sql & " where shopid = '" & shopid & "' AND makerid = '" & makerid & "' " + vbcrlf
			
			'response.write sql &"<br>"
			dbget.execute sql

		next
		
	'//구역지정
	else
		makeridarr = split(request("makerid"),",")
		
		sql = ""
		for i = 0 to ubound(makeridarr)
			makerid = requestCheckVar(Trim(makeridarr(i)),32)

			sql = sql & "IF EXISTS(SELECT makerid FROM [db_shop].[dbo].[tbl_shop_brand_zone_detail] WHERE shopid = '" & shopid & "' AND makerid = '" & makerid & "') " & vbcrlf
			sql = sql & "	BEGIN " & vbcrlf
			sql = sql & "		UPDATE [db_shop].[dbo].[tbl_shop_brand_zone_detail] SET " & vbcrlf
			sql = sql & "		zoneidx = '" & chzoneidx & "' " & vbcrlf
			sql = sql & "		WHERE shopid = '" & shopid & "' AND makerid = '" & makerid & "' " & vbcrlf
			sql = sql & "	END " & vbcrlf
			sql = sql & "ELSE " & vbcrlf
			sql = sql & "	BEGIN " & vbcrlf
			sql = sql & "		INSERT INTO [db_shop].[dbo].[tbl_shop_brand_zone_detail](shopid, makerid, zoneidx) " & vbcrlf
			sql = sql & "		VALUES('" & shopid & "', '" & makerid & "', '" & chzoneidx & "') " & vbcrlf
			sql = sql & "	END " & vbcrlf

			'response.write sql &"<Br>"
			dbget.execute sql
		next
		
	end if

	If Err.Number = 0 Then
	    ''dbget.CommitTrans
	Else
	    ''dbget.RollBackTrans
	End If

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	'response.write "	opener.location.reload();"
	response.write "	location.href='zone_item.asp?"&parameter&"'"
	response.write "</script>"

end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 현금매출정산관리 
' History : 2013.10.24 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/payment_cls.asp"-->
	
<%
dim mode, shopid, yyyymmddarr, cnt100000wonarr, cnt50000wonarr, cnt10000wonarr, cnt5000wonarr, cnt1000wonarr, cnt500wonarr
dim cnt100wonarr, cnt50wonarr, cnt10wonarr, vaultcasharr, jungsanadminidarr, depositadminidarr, sqlstr, adminid, i
dim etctypearr, etcwonarr, masteridx, posidarr, bigoarr
	mode = requestCheckVar(Request("mode"),32)	
	shopid = requestCheckVar(Request("shopid"),32)
	posidarr = requestCheckVar(Request("posidarr"),10)
	yyyymmddarr = Request("yyyymmddarr")
	cnt100000wonarr = Request("cnt100000wonarr")
	cnt50000wonarr = Request("cnt50000wonarr")
	cnt10000wonarr = Request("cnt10000wonarr")
	cnt5000wonarr = Request("cnt5000wonarr")
	cnt1000wonarr = Request("cnt1000wonarr")
	cnt500wonarr = Request("cnt500wonarr")
	cnt100wonarr = Request("cnt100wonarr")
	cnt50wonarr = Request("cnt50wonarr")
	cnt10wonarr = Request("cnt10wonarr")
	vaultcasharr = Request("vaultcasharr")
	jungsanadminidarr = Request("jungsanadminidarr")
	depositadminidarr = Request("depositadminidarr")
	etctypearr = Request("etctypearr")
	etcwonarr = Request("etcwonarr")
	bigoarr = ReplaceRequestSpecialChar(Request("bigoarr"))
	
	adminid = session("ssBctId")
	
dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="cash_edit" then
	if yyyymmddarr="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('날짜가 없습니다');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	if shopid="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장이 없습니다');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	if posidarr="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('포스ID가 없습니다');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	
	if bigoarr<>"" then
		if checkNotValidHTML(bigoarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('코맨트에 유효하지 않은 글자가 포함되어 있습니다');"
			response.write "	location.href='"& refer &"';"
			response.write "</script>"
			dbget.close() : response.end			
		end if
	end if
	
	sqlstr = "if exists (" + vbcrlf
	sqlstr = sqlstr & " 	select top 1 *" + vbcrlf
	sqlstr = sqlstr & " 	from db_shop.dbo.tbl_shop_cash_management" + vbcrlf
	sqlstr = sqlstr & " 	where isusing='Y'" + vbcrlf
	sqlstr = sqlstr & " 	and yyyymmdd='" & trim(yyyymmddarr) & "'" + vbcrlf
	sqlstr = sqlstr & " 	and shopid='" & trim(shopid) & "'" + vbcrlf
	sqlstr = sqlstr & " 	and posid='" & trim(posidarr) & "'" + vbcrlf
	sqlstr = sqlstr & " )" + vbcrlf
	sqlstr = sqlstr & " 	update db_shop.dbo.tbl_shop_cash_management" + vbcrlf
	sqlstr = sqlstr & " 	set lastupdate=getdate()" + vbcrlf
	sqlstr = sqlstr & " 	,lastadminid='" & adminid & "'" + vbcrlf
	sqlstr = sqlstr & " 	,cnt100000won=" & trim(cnt100000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt50000won=" & trim(cnt50000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt10000won=" & trim(cnt10000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt5000won=" & trim(cnt5000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt1000won=" & trim(cnt1000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt500won=" & trim(cnt500wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt100won=" & trim(cnt100wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt50won=" & trim(cnt50wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, cnt10won=" & trim(cnt10wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, vaultcash=" & trim(vaultcasharr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, jungsanadminid='" & html2db(trim(jungsanadminidarr)) & "'" + vbcrlf
	sqlstr = sqlstr & " 	, depositadminid='" & html2db(trim(depositadminidarr)) & "'" + vbcrlf
	sqlstr = sqlstr & " 	, bigo='" & html2db(trim(bigoarr)) & "'" + vbcrlf
	sqlstr = sqlstr & " 	where isusing='Y'" + vbcrlf
	sqlstr = sqlstr & " 	and yyyymmdd='" & trim(yyyymmddarr) & "'" + vbcrlf
	sqlstr = sqlstr & " 	and shopid='" & trim(shopid) & "'" + vbcrlf
	sqlstr = sqlstr & " 	and posid='" & trim(posidarr) & "'" + vbcrlf
	sqlstr = sqlstr & " else" + vbcrlf
	sqlstr = sqlstr & " 	insert into db_shop.dbo.tbl_shop_cash_management(" + vbcrlf
	sqlstr = sqlstr & " 	yyyymmdd, shopid, posid, cnt100000won, cnt50000won, cnt10000won, cnt5000won, cnt1000won" + vbcrlf
	sqlstr = sqlstr & " 	, cnt500won, cnt100won, cnt50won, cnt10won, vaultcash, jungsanadminid, depositadminid" + vbcrlf
	sqlstr = sqlstr & " 	, isusing, regadminid, lastadminid, bigo" + vbcrlf
	sqlstr = sqlstr & " 	) values (" + vbcrlf
	sqlstr = sqlstr & " 	'" & trim(yyyymmddarr) & "', '" & trim(shopid) & "', '" & trim(posidarr) & "', " & trim(cnt100000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, " & trim(cnt50000wonarr) & "," & trim(cnt10000wonarr) & ", " & trim(cnt5000wonarr) & ", " & trim(cnt1000wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, " & trim(cnt500wonarr) & "," & trim(cnt100wonarr) & ", " & trim(cnt50wonarr) & ", " & trim(cnt10wonarr) & "" + vbcrlf
	sqlstr = sqlstr & " 	, " & trim(vaultcasharr) & ",'" & html2db(trim(jungsanadminidarr)) & "', '" & html2db(trim(depositadminidarr)) & "'" + vbcrlf
	sqlstr = sqlstr & " 	, 'Y', '" & html2db(adminid) & "', '" & html2db(adminid) & "', '" & html2db(bigoarr) & "'" + vbcrlf
	sqlstr = sqlstr & " 	)"
	
	'response.write sqlstr &"<br>"
	dbget.execute sqlstr
	
	sqlstr = "select top 1 idx"
	sqlstr = sqlstr & " from db_shop.dbo.tbl_shop_cash_management"
	sqlstr = sqlstr & " where isusing='Y'"
	sqlstr = sqlstr & " and yyyymmdd='" & trim(yyyymmddarr) & "'"
	sqlstr = sqlstr & " and shopid='" & trim(shopid) & "'"
	sqlstr = sqlstr & " and posid='" & trim(posidarr) & "'" + vbcrlf

	'response.write sqlstr & "<Br>"
	rsget.Open sqlstr,dbget,1
	if  not rsget.EOF  then
		masteridx = rsget("idx")	
	end if
	rsget.Close

	if masteridx="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('정상적으로 저장되지 않았습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if
	
	sqlstr = "delete from db_shop.dbo.tbl_shop_cash_management_etc where" + vbcrlf
	sqlstr = sqlstr & " masteridx="& trim(masteridx) &""

	'response.write sqlstr &"<br>"
	dbget.execute sqlstr

	etctypearr = split(etctypearr,",")
	etcwonarr = split(etcwonarr,",")

	if isarray(etctypearr) then
		for i = 0 to ubound(etctypearr) -1
			sqlstr = "insert into db_shop.dbo.tbl_shop_cash_management_etc(" + vbcrlf
			sqlstr = sqlstr & " masteridx, etctype, etcwon, isusing, regadminid, lastadminid" + vbcrlf
			sqlstr = sqlstr & " ) values (" + vbcrlf
			sqlstr = sqlstr & " "&masteridx&", '"& requestCheckVar(trim(etctypearr(i)),12) &"', '"& requestCheckVar(trim(etcwonarr(i)),20) &"', 'Y'" + vbcrlf
			sqlstr = sqlstr & " ,'" & html2db(adminid) & "', '" & html2db(adminid) & "'" + vbcrlf
			sqlstr = sqlstr & " )"

			'response.write sqlstr &"<br>"
			dbget.execute sqlstr			
		next
	end if

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='"& refer &"';"
	response.write "</script>"
	dbget.close() : response.end
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다');"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
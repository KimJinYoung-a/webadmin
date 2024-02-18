<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  목표매출
' History : 2013.03.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/targetmaechul/targetmaechul_cls.asp"-->

<%
dim menupos , mode ,solar_date ,yyyymm ,shopid ,gubuntype ,gubun ,targetmaechul ,i , sql
dim yyyy1 , mm1
	menupos = requestcheckvar(request("menupos"),10)
	mode = requestcheckvar(request("mode"),32)
	solar_date = request("solar_date")
	yyyymm = request("yyyymm")
	shopid = requestcheckvar(request("shopid"),32)
	gubuntype = requestcheckvar(request("gubuntype"),10)
	gubun = request("gubun")
	targetmaechul = request("targetmaechul")
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	mm1 = requestcheckvar(request("mm1"),2)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")
	
if mode = "tmreg" then
	
	if shopid = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('매장을 선택해 주세요');"
		response.write "	location.href='"& refer &"';"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	gubun = split(gubun,",")
	yyyymm = split(yyyymm,",")
	targetmaechul = split(targetmaechul,",")			

	for i = 0 to ubound(gubun)-1
	
		'/신규등록
		if trim(yyyymm(i)) = "" then
			sql = "insert into db_shop.dbo.tbl_targetmaechul_month_off" + vbcrlf
			sql = sql & " (yyyymm ,shopid ,gubuntype ,gubun ,targetmaechul ,regdate ,lastupdate ,lastadminid) values (" + vbcrlf
			sql = sql & " '"&trim(yyyy1)&"-"&trim(Format00(2,mm1))&"'" + vbcrlf
			sql = sql & " ,'"&shopid&"'" + vbcrlf
			sql = sql & " ,"&gubuntype&"" + vbcrlf
			sql = sql & " ,"&trim(gubun(i))&"" + vbcrlf
			sql = sql & " ,"&trim(targetmaechul(i))&"" + vbcrlf
			sql = sql & " ,getdate()" + vbcrlf
			sql = sql & " ,getdate()" + vbcrlf
			sql = sql & " ,'"&session("ssBctId")&"'"
			sql = sql & " )"
			
			'response.write sql
			dbget.execute sql
			
		'/기존목표매출수정	
		else
			sql = "update db_shop.dbo.tbl_targetmaechul_month_off set" + vbcrlf
			sql = sql & " targetmaechul = "&trim(targetmaechul(i))&"" + vbcrlf
			sql = sql & " ,lastupdate = getdate()" + vbcrlf
			sql = sql & " ,lastadminid = '"&session("ssBctId")&"'" + vbcrlf
			sql = sql & " where yyyymm = '"&trim(yyyy1)&"-"&Format00(2,trim(mm1))&"'" + vbcrlf
			sql = sql & " and shopid = '"&shopid&"'" + vbcrlf
			sql = sql & " and gubuntype = "&gubuntype&"" + vbcrlf
			sql = sql & " and gubun = "&trim(gubun(i))&""
			
			'response.write sql
			dbget.execute sql			
		end if
	next

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='"& refer &"';"
	response.write "</script>"
	dbget.close()	:	response.end
			
elseif mode = "targetreg" then
	
	if shopid = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('매장을 선택해 주세요');"
		response.write "	location.href='"& refer &"';"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	solar_date = split(solar_date,",")
	yyyymm = split(yyyymm,",")			
	targetmaechul = split(targetmaechul,",")			

	for i = 0 to ubound(solar_date)-1
	
		'/신규등록
		if trim(yyyymm(i)) = "" then
			sql = "insert into db_shop.dbo.tbl_targetmaechul_month_off" + vbcrlf
			sql = sql & " (yyyymm ,shopid ,gubuntype ,gubun ,targetmaechul ,regdate ,lastupdate ,lastadminid) values (" + vbcrlf
			sql = sql & " '"&trim(solar_date(i))&"'" + vbcrlf
			sql = sql & " ,'"&shopid&"'" + vbcrlf
			sql = sql & " ,"&gubuntype&"" + vbcrlf
			sql = sql & " ,"&gubun&"" + vbcrlf
			sql = sql & " ,"&trim(targetmaechul(i))&"" + vbcrlf
			sql = sql & " ,getdate()" + vbcrlf
			sql = sql & " ,getdate()" + vbcrlf
			sql = sql & " ,'"&session("ssBctId")&"'"
			sql = sql & " )"
			
			'response.write sql
			dbget.execute sql
			
		'/기존목표매출수정	
		else
			sql = "update db_shop.dbo.tbl_targetmaechul_month_off set" + vbcrlf
			sql = sql & " targetmaechul = "&trim(targetmaechul(i))&"" + vbcrlf
			sql = sql & " ,lastupdate = getdate()" + vbcrlf
			sql = sql & " ,lastadminid = '"&session("ssBctId")&"'" + vbcrlf
			sql = sql & " where yyyymm = '"&trim(solar_date(i))&"'" + vbcrlf
			sql = sql & " and shopid = '"&shopid&"'" + vbcrlf
			sql = sql & " and gubuntype = "&gubuntype&"" + vbcrlf
			sql = sql & " and gubun = "&gubun&""
			
			'response.write sql
			dbget.execute sql			
		end if
	next

	response.write "<script language='javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='"& refer &"';"
	response.write "</script>"
	dbget.close()	:	response.end

end if

%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
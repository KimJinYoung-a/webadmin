<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매장 타 사이트 매칭
' History : 2012.05.15 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim mode , i , menupos ,othershopid ,shopid ,siteseq , sql , userid
	menupos = requestCheckvar(request("menupos"),10)
	mode = requestCheckvar(request("mode"),32)
	siteseq = requestCheckvar(request("siteseq"),10)
	shopid = requestCheckvar(request("shopid"),32)
	othershopid = requestCheckvar(request("othershopid"),32)

userid = session("ssBctId")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode = "shopotherreg" then
	if siteseq = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('외부매장을 선택해 주세요');"
		response.write " 	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('텐바이텐 매장을 선택해 주세요');"
		response.write " 	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if		
	if othershopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('외부매장ID를 입력해 주세요');"
		response.write " 	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	
	siteseq = trim(siteseq)
	shopid = split(shopid,",")
	othershopid = split(othershopid,",")
	
	for i = 0 to ubound(shopid)-1

		sql = "if exists(" + vbcrlf
		sql = sql & "		select top 1 *" + vbcrlf
		sql = sql & "		from db_shop.dbo.tbl_shop_othersitelink" + vbcrlf
		sql = sql & "		where siteseq = "&siteseq&"" + vbcrlf
		sql = sql & "		and shopid = '"&trim(shopid(i))&"'" + vbcrlf
		sql = sql & " )" + vbcrlf
		sql = sql & "		update db_shop.dbo.tbl_shop_othersitelink set" + vbcrlf
		sql = sql & "		othershopid = '"&trim(othershopid(i))&"'" + vbcrlf
		sql = sql & "		,lastupdate = getdate()" + vbcrlf
		sql = sql & "		,lastadminuserid = '"&userid&"'" + vbcrlf
		sql = sql & "		where siteseq ="&siteseq&"" + vbcrlf
		sql = sql & "		and shopid = '"&trim(shopid(i))&"'" + vbcrlf
		sql = sql & " else" + vbcrlf
		sql = sql & "		insert into db_shop.dbo.tbl_shop_othersitelink" + vbcrlf
		sql = sql & "		(siteseq ,shopid ,othershopid ,regdate ,lastupdate ,lastadminuserid) values " + vbcrlf
		sql = sql & "		("&siteseq&",'"&trim(shopid(i))&"','"&trim(othershopid(i))&"',getdate(),getdate(),'"&userid&"')"
		
		'response.write sql &"<Br>"
		dbget.execute sql
	next

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write " 	location.href='"&refer&"';"
	response.write "</script>"
	dbget.close()	:	response.End
		
elseif mode = "shopotherdel" then
	if siteseq = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('외부매장을 선택해 주세요');"
		response.write " 	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	if shopid = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('텐바이텐 매장을 선택해 주세요');"
		response.write " 	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	
	sql = "delete from db_shop.dbo.tbl_shop_othersitelink where" + vbcrlf
	sql = sql & " siteseq ="&siteseq&"" + vbcrlf
	sql = sql & " and shopid = '"&shopid&"'" + vbcrlf
	
	'response.write sql &"<Br>"
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write " 	location.href='"&refer&"';"
	response.write "</script>"
	dbget.close()	:	response.End
	
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 정의되지 않았습니다');"
	response.write " 	history.go(-1);"
	response.write "</script>"
	dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
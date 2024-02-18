<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 모바일 keywordbanner
' History : 2013.12.16 한용민
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mobile/keywordbanner_cls.asp" -->

<%
dim idx, keywordtype, keyword, imagepath, linkpath, isusing, orderno, regdate, imgalt
dim lastdate, regadminid, lastadminid, YearUse, menupos, adminid, mode
Dim startdate , enddate

	idx = request("idx")
	keywordtype = request("keywordtype")
	keyword = html2db(request("keyword"))
	imagepath = request("imagepath")
	linkpath = html2db(request("linkpath"))
	isusing = request("isusing")
	orderno = request("orderno")						
	isusing = request("isusing")
	mode = request("mode")
	menupos = request("menupos")
	imgalt = request("imgalt")
	adminid=session("ssBctId")

	startdate			= Request("StartDate")& " " &Request("sTm")
	enddate			= Request("EndDate")& " " &Request("eTm")

dim sql

'/저장
if mode = "keywordbanneredit" then
	if keywordtype="" or isusing="" or orderno="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('필요한 조건이 없습니다.');"
		response.write "</script>"
		dbget.close()	:	response.end
	end if
	
	'/수정
	if idx<>"" then
		sql = "update db_sitemaster.dbo.tbl_mobile_main_keywordbanner" + vbcrlf
		sql = sql & " set keywordtype="&keywordtype&"" + vbcrlf
		sql = sql & " ,keyword='"&keyword&"'" + vbcrlf
		sql = sql & " ,imagepath='"&imagepath&"'" + vbcrlf
		sql = sql & " ,linkpath='"&linkpath&"'" + vbcrlf
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
		sql = sql & " ,orderno="&orderno&"" + vbcrlf
		sql = sql & " ,lastdate=getdate()" + vbcrlf	
		sql = sql & " ,lastadminid='"&adminid&"'" + vbcrlf
		sql = sql & " ,imgalt='"&imgalt&"'" + vbcrlf
		sql = sql & " ,startdate='"&startdate&"'" + vbcrlf
		sql = sql & " ,enddate='"&enddate&"'" + vbcrlf
		sql = sql & " where idx = "&idx&""
		
		'response.write sql
		dbget.execute sql
	
	'/신규등록
	else

		sql = "insert into db_sitemaster.dbo.tbl_mobile_main_keywordbanner" + vbcrlf
		sql = sql & " (keywordtype, keyword, imagepath, linkpath, isusing, orderno, regdate" + vbcrlf
		sql = sql & " , lastdate, regadminid, lastadminid, imgalt,startdate,enddate)" + vbcrlf
		sql = sql & " values ("  + vbcrlf
		sql = sql & " '"&keywordtype&"'" + vbcrlf
		sql = sql & " ,'"&keyword&"'"	 + vbcrlf
		sql = sql & " ,'"&imagepath&"'"	 + vbcrlf
		sql = sql & " ,'"&linkpath&"'" + vbcrlf
		sql = sql & " ,'"&isusing&"'"	 + vbcrlf
		sql = sql & " ,'"&orderno&"'"	 + vbcrlf
		sql = sql & " ,getdate()" + vbcrlf
		sql = sql & " ,getdate()"	 + vbcrlf
		sql = sql & " ,'"&adminid&"'"	 + vbcrlf
		sql = sql & " ,'"&adminid&"'" + vbcrlf
		sql = sql & " ,'"&imgalt&"'" + vbcrlf
		sql = sql & " ,'"&startdate&"'" + vbcrlf
		sql = sql & " ,'"&enddate&"'" + vbcrlf
		sql = sql & ")"
		
		'response.write sql
		dbget.execute sql
	end if

	response.write "<script type='text/javascript'>"
	response.write "	alert('저장되었습니다.');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.end
	
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

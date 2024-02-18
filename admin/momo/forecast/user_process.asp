<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성예보
' Hieditor : 2010.11.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim i ,yyyy,mm, mode , winner0 ,winner1 , winner2,contents0 , contents1 , contents2
	yyyy = request("yyyy")		
	mm = request("mm")
	mode = request("mode")
	winner0 = request("winner0")
	winner1 = request("winner1")
	winner2 = request("winner2")
	contents0 = request("contents0")
	contents1 = request("contents1")						
	contents2 = request("contents2")
			
dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")

'//신규 & 수정
if mode = "winneredit" then
	
	'// gubun 값 감성예보당첨자 0	
	sql = "delete from db_momo.dbo.tbl_winner where gubun=0 and yyyymm='"&yyyy&"-"&mm&"'"	
	
	'response.write sql &"<br>"
	dbget.execute sql

	sql = ""  '/1등
	sql = "insert into db_momo.dbo.tbl_winner (yyyymm ,gubun ,orderno ,userid ,contents)" + vbcrlf
	sql = sql & " values (" + vbcrlf		
	sql = sql & " '"&yyyy&"-"&mm&"'"		
	sql = sql & " ,0"	
	sql = sql & " ,0"
	sql = sql & " ,'"&db2html(winner0)&"'"
	sql = sql & " ,'"&db2html(contents0)&"'"
	sql = sql & " )"		

	'response.write sql &"<br>"
	dbget.execute sql

	sql = "" '/2등
	sql = "insert into db_momo.dbo.tbl_winner (yyyymm ,gubun ,orderno ,userid ,contents)" + vbcrlf
	sql = sql & " values (" + vbcrlf		
	sql = sql & " '"&yyyy&"-"&mm&"'"		
	sql = sql & " ,0"	
	sql = sql & " ,1"
	sql = sql & " ,'"&db2html(winner1)&"'"
	sql = sql & " ,'"&db2html(contents1)&"'"
	sql = sql & " )"		

	'response.write sql &"<br>"
	dbget.execute sql

	sql = "" '/3등
	sql = "insert into db_momo.dbo.tbl_winner (yyyymm ,gubun ,orderno ,userid ,contents)" + vbcrlf
	sql = sql & " values (" + vbcrlf		
	sql = sql & " '"&yyyy&"-"&mm&"'"		
	sql = sql & " ,0"	
	sql = sql & " ,2"
	sql = sql & " ,'"&db2html(winner2)&"'"
	sql = sql & " ,'"&db2html(contents2)&"'"
	sql = sql & " )"		

	'response.write sql &"<br>"
	dbget.execute sql			
	response.write "<script language='javascript'>"
	response.write "	alert('OK');"	
	response.write "	location.replace('" + referer + "');"		
	response.write "</script>"
end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

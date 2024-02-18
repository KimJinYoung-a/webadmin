<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 공용 Event 당첨자등록 수동 Y 로 전환
' History : 2009.04.14 한용민 생성 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim eCode , egKindCode ,strSql
	eCode = requestCheckVar(request("eCode"),10)	'//이벤트코드  컬쳐의 경우 무조건 4 입니다
	egKindCode = requestCheckVar(request("egKindCode"),10)	'//컬쳐스테이션의 이벤트코드

	'//컬쳐이벤트 당첨자발표 완료 처리 
	if eCode = 4 then
		
		strSql = "update db_culture_station.dbo.tbl_culturestation_event set"+vbcrlf
		strSql = strSql & " prizeyn = 'Y'"+vbcrlf
		strSql = strSql & " where evt_code = "&egKindCode&""+vbcrlf
		
		'response.write strSql&"<br>"
		dbget.execute strSql		
		
	'//일반이벤트 당첨자발표 완료 처리
	else
		strSql = "update db_event.dbo.tbl_event set"+vbcrlf
		strSql = strSql & " prizeyn = 'Y'"+vbcrlf
		strSql = strSql & " where evt_code = "&eCode&""+vbcrlf
	
		'response.write strSql&"<br>"			
		dbget.execute strSql
	end if
%>
	<script type='text/javascript'>
		opener.location.reload();
		alert('처리되었습니다');
		self.close();
	</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->


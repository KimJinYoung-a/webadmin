<%@ language=vbscript %>
<% option explicit %>
<% 
'###########################################################
' Description :   코멘트
' History : 2016.08.18 생성
'################################################################## 
%>
<!-- #include virtual="/partner/incSessionDesigner.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim sMode
dim  Comment, regid, commentidx,regtype,evtCode
dim strSql 
sMode		= requestCheckvar(Request("hidM"),2)
evtCode	= requestCheckvar(Request("eC"),10) 
Comment		= ReplaceRequestSpecialChar(Request("tCmt"))
regtype	= ReplaceRequestSpecialChar(Request("hidRT"))
regid		= session("ssBctId") 
commentidx = requestCheckvar(Request("iCidx"),10)
 
	
	
SELECT CASE sMode
	CASE "CI"	'코멘트 등록
	 strSql = "INSERT INTO db_event.dbo.tbl_partner_event_comment (evt_code, comment,regtype, regid  )"
	 strSql = strSql & " VALUES ("&evtCode&",   '"&comment&"','"&regtype&"', '"&regid&"') "
	 dbget.Execute strSql
	  
%>
	<script language="javascript">
		alert("등록되었습니다.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<% 
	 
CASE "CD"	'코멘트 삭제 
	strSql = "update   db_event.dbo.tbl_partner_event_comment  set isusing = 0 where comidx ="& commentidx &" and regid ='"&regid&"'"
	dbget.Execute strSql	 
%>
	<script language="javascript">
		alert("삭제되었습니다.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%	 
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
 
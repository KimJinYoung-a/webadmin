<%@ language=vbscript %>
<% option explicit %>
<% 
'###########################################################
' Description :   �ڸ�Ʈ
' History : 2016.08.18 ����
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
	CASE "CI"	'�ڸ�Ʈ ���
	 strSql = "INSERT INTO db_event.dbo.tbl_partner_event_comment (evt_code, comment,regtype, regid  )"
	 strSql = strSql & " VALUES ("&evtCode&",   '"&comment&"','"&regtype&"', '"&regid&"') "
	 dbget.Execute strSql
	  
%>
	<script language="javascript">
		alert("��ϵǾ����ϴ�.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<% 
	 
CASE "CD"	'�ڸ�Ʈ ���� 
	strSql = "update   db_event.dbo.tbl_partner_event_comment  set isusing = 0 where comidx ="& commentidx &" and regid ='"&regid&"'"
	dbget.Execute strSql	 
%>
	<script language="javascript">
		alert("�����Ǿ����ϴ�.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%	 
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
 
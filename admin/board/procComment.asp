<%@ language="VBScript" %>
<% option explicit %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim sMode
dim  Comment, regid, commentidx,regtype,iboard_idx
dim returnValue,objCmd
sMode		= requestCheckvar(Request("hidM"),2)
iboard_idx	= requestCheckvar(Request("ibidx"),10) 
Comment		= ReplaceRequestSpecialChar(Request("tCmt"))
regtype	= ReplaceRequestSpecialChar(Request("hidRT"))
regid		= session("ssBctId") 
commentidx = requestCheckvar(Request("iCidx"),10)

SELECT CASE sMode
	CASE "CI"	'�ڸ�Ʈ ���
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_board.[dbo].[sp_Ten_partnerA_notice_comment_Insert]( "&iboard_idx&",   '"&comment&"', '"&regid&"','"&regtype&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

	IF 	returnValue = 1 THEN
%>
	<script language="javascript">
		alert("��ϵǾ����ϴ�.");
		parent.jsGetCmt(); 
		self.location.href = "about:blank";
		</script>
<%
	ELSE
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
 
	response.end
CASE "CD"	'�ڸ�Ʈ ����
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_board.[dbo].[sp_Ten_partnerA_notice_comment_Delete]( "&commentidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
	IF 	returnValue = 1 THEN
%>
	<script language="javascript">
		alert("�����Ǿ����ϴ�.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%	ELSE
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
 
			response.end
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
 
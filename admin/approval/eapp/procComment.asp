<%@ language=vbscript %>
<% option explicit %> 
<%
'###########################################################
' Description : ���ڰ���
' Hieditor : ������ ����
'			 2022.07.11 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sMode
dim payrequestidx,reportIdx, Comment, adminID, commentidx
dim returnValue,objCmd
sMode		= requestCheckvar(Request("hidM"),2)
reportIdx	= requestCheckvar(Request("irIdx"),10)
payrequestidx=requestCheckvar(Request("iprIdx"),10) 
Comment		= ReplaceRequestSpecialChar(Request("tCmt"))
adminID		= session("ssBctId") 
commentidx = requestCheckvar(Request("iCidx"),10)

SELECT CASE sMode
	CASE "CI"	'�ڸ�Ʈ ���
		if Comment <> "" and not(isnull(Comment)) then
			Comment = ReplaceBracket(Comment)
		end If

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppComment_Insert]( "&reportidx&", "&payrequestidx&", '"&comment&"', '"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

	IF 	returnValue = 1 THEN
%>
	<script type='text/javascript'>
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
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppComment_Delete]( "&commentidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
	IF 	returnValue = 1 THEN
%>
	<script type='text/javascript'>
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
 
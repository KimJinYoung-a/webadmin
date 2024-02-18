	<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<%
dim sMode
dim payrequestidx,reportIdx, Comment, adminID, commentidx, sDoc_SMS
dim returnValue,objCmd
sMode		= requestCheckvar(Request("hidM"),2)
reportIdx	= requestCheckvar(Request("irIdx"),10)
payrequestidx=requestCheckvar(Request("iprIdx"),10)
Comment		= ReplaceRequestSpecialChar(Request("tCmt"))
adminID		= session("ssBctId")
commentidx = requestCheckvar(Request("iCidx"),10)
sDoc_SMS = requestCheckvar(Request("sms_send"),10)

Function fnGetMemberHpByReportIDX(reportIdx)
	Dim strSql

	strSql = " select isNull(usercell,'0') AS manager_hp "
	strSql = strSql + " from "
	strSql = strSql + " 	db_partner.dbo.tbl_eAppReport e "
	strSql = strSql + " 	join [db_partner].[dbo].tbl_user_tenbyten t "
	strSql = strSql + " 	on "
	strSql = strSql + " 		e.adminid = t.userid "
	strSql = strSql + " where reportIdx = " & reportIdx
	rsget.Open strSql,dbget,1
	'response.write strSql
	IF not rsget.EOF THEN
		If rsget("manager_hp") = "" Then
			fnGetMemberHpByReportIDX = "0"
		Else
			fnGetMemberHpByReportIDX = rsget("manager_hp")
		End If
	Else
		fnGetMemberHpByReportIDX = "0"
	END IF
	rsget.close
End Function

SELECT CASE sMode
	CASE "CI"	'코멘트 등록
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

		If sDoc_SMS = "o" Then
			Call SendNormalSMS_LINK(fnGetMemberHpByReportIDX(reportIdx),"",""&session("ssBctCname")&"님께서 전자결제에 코멘트를 남기셨습니다.(No." & reportidx & ")")
		End If

%>
	<script language="javascript">
		alert("등록되었습니다.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%
	ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF
	session.codePage = 949
	response.end
CASE "CD"	'코멘트 삭제
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
	<script language="javascript">
		alert("삭제되었습니다.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%	ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF
	session.codePage = 949
			response.end
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>

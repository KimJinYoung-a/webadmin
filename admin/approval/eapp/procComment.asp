<%@ language=vbscript %>
<% option explicit %> 
<%
'###########################################################
' Description : 전자결제
' Hieditor : 정윤정 생성
'			 2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
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
	CASE "CI"	'코멘트 등록
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
		alert("등록되었습니다.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%
	ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF 
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
	<script type='text/javascript'>
		alert("삭제되었습니다.");
		parent.jsGetCmt();
		self.location.href = "about:blank";
		</script>
<%	ELSE
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.1")
	END IF
 
			response.end
CASE ELSE
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
 
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, sellsite, jdate, ret, retval

sellsite = requestCheckVar(request("sellsite"), 32)
jdate = requestCheckVar(request("jdate"), 10)

If sellsite = "" OR jdate = "" Then
	Response.Write "X||Don't select sellsite OR Don't select extMeachulDate !!"
	dbget.close
	Response.End
End If


	sqlStr = ""
	sqlStr = sqlStr & "DECLARE @retval varchar(100) " & vbCrLf
	sqlStr = sqlStr & "DECLARE @RET int " & vbCrLf
	sqlStr = sqlStr & "exec @RET =[db_jungsan].[dbo].[sp_Ten_OUTAMLL_Jungsan_realOrder_modify_by_xSite_JungsanData]  '" & jdate & "','" & sellsite & "', @retval output " & vbCrLf
	sqlStr = sqlStr & "select @RET,@retval"
	
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

	if not rsget.Eof then
		ret = rsget(0)
		retval = rsget(1)
	end if
	
	Response.Write ret & "||" & retval
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
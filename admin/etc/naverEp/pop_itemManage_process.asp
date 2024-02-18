<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim mode, itemid, strSql, postfix, applyyn
Dim itemidArr, postfixArr, applyynArr, i
Dim tmpItemidArr, tmpPostfixArr, tmpApplyynArr
mode	= requestCheckVar(request("mode"), 3)
itemid	= requestCheckVar(request("itemid"), 32)
postfix	= requestCheckVar(request("postfix"), 32)
applyyn	= requestCheckVar(request("applyyn"), 1)

itemidArr = request("itemidArr")
postfixArr = request("postfixArr")
applyynArr = request("applyynArr")

If (mode = "del") Then
	strSql = ""
	strSql = strSql & "EXEC db_outmall.dbo.usp_EpShop_ItemPostfix_Set " & itemid & ",  '',  '', '"& session("ssBctID") &"' "
	dbCTget.execute strSql
	response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); top.location.reload(); " &_
					"</script>"
ElseIf (mode = "add") Then
	strSql = ""
	strSql = strSql & "EXEC db_outmall.dbo.usp_EpShop_ItemPostfix_Set " & itemid & ",  '"& postfix &"',  '"& applyyn &"', '"& session("ssBctID") &"' "
	dbCTget.execute strSql
	response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); top.location.reload(); " &_
					"</script>"
ElseIf (mode = "all") Then
	tmpItemidArr = Split(itemidArr, "*(^!")
	tmpPostfixArr = Split(postfixArr, "*(^!")
	tmpApplyynArr = Split(applyynArr, "*(^!")

	For i = 0 To Ubound(tmpItemidArr) - 1
		strSql = ""
		strSql = strSql & "EXEC db_outmall.dbo.usp_EpShop_ItemPostfix_Set " & tmpItemidArr(i) & ",  '"& tmpPostfixArr(i) &"',  '"& tmpApplyynArr(i) &"', '"& session("ssBctID") &"' "
		dbCTget.execute strSql
	Next
	response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); top.location.reload(); " &_
					"</script>"
Else
	response.write	"<script language='javascript'>" &_
					"	alert('잘못 된 접근입니다.'); top.window.close(); " &_
					"</script>"
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
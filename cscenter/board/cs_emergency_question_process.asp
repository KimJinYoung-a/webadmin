<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_emergencyQuestionCls.asp"-->
<%

dim mode
dim idx, orderserial, makerid, title, contents
dim categoryGubun, categoryName

mode 		= requestCheckVar(request("mode"), 32)
idx 		= requestCheckVar(request("idx"), 32)
orderserial = requestCheckVar(request("orderserial"), 32)
makerid 	= requestCheckVar(request("makerid"), 32)
title 		= requestCheckVar(request("title"), 64)
contents 	= requestCheckVar(request("contents"), 500)
categoryGubun	= requestCheckVar(request("categoryGubun"), 1)
categoryName	= CsEmergencyQuestionCategoryGubunToName(categoryGubun)


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim strSql

select case mode
	case "regEmergencyQuestion"
		strSql = ""
		strSql = strSql & " insert into db_cs.dbo.tbl_emergency_question_master"
		strSql = strSql & " (upcheGubun, upcheName, makerid, categoryGubun, categoryName, needReplyYN, title, contents, orderserial, buyName, itemids, deleteyn, currState, deadlineDate, regUserid) "
		strSql = strSql & "values"
		strSql = strSql & "('1', '텐바이텐 고객센터', '" & makerid & "', '" & categoryGubun & "', '" & categoryName & "', 'Y', '" & db2html(title) & "', '" & db2html(contents) & "', '" & orderserial & "', '', '', 'N', '1', '', '" & session("ssBctId") & "')"
		dbget.execute strSql

		response.write	"<script type='text/javascript'>" &_
						"	alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
						"</script>"
		dbget.close() : response.end
	case "modiEmergencyQuestion"
		strSql = ""
		strSql = strSql & " update db_cs.dbo.tbl_emergency_question_master"
		strSql = strSql & " set "
		strSql = strSql & " 	title = '" & db2html(title) & "'"
		strSql = strSql & " 	, contents = '" & db2html(contents) & "'"
		strSql = strSql & " 	, lastUpdate = getdate()"
		strSql = strSql & " where idx = " & idx & " and currState = '1' "
		dbget.execute strSql

		response.write	"<script type='text/javascript'>" &_
						"	alert('저장되었습니다.'); location.replace('" & refer & "') " &_
						"</script>"
		dbget.close() : response.end
	case "delEmergencyQuestion"
		strSql = ""
		strSql = strSql & " update db_cs.dbo.tbl_emergency_question_master"
		strSql = strSql & " set "
		strSql = strSql & " 	deleteyn = 'Y'"
		strSql = strSql & " 	, lastUpdate = getdate()"
		strSql = strSql & " where idx = " & idx & " and currState = '1' "
		dbget.execute strSql

		response.write	"<script type='text/javascript'>" &_
						"	alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
						"</script>"
		dbget.close() : response.end
	case else
		response.write "ERR"
		dbget.close() : response.end
end select

response.write	"<script type='text/javascript'>" &_
				"	alert('저장되었습니다.'); location.replace('/cscenter/board/boarduserlist.asp?menupos=" & menupos & "'); " &_
				"</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

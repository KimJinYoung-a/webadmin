<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/event/eventCls.asp"-->
<%
Dim evt_startdate, evt_enddate, gubun, lectureid, lecidx, company_name
Dim evt_name, isusing, menupos, isRegedEventUsing
Dim strSql, mode, idx, makerid, diycode

mode			= requestCheckVar(Request("mode"),1)
evt_startdate	= requestCheckVar(Request("evt_startdate"),10)
evt_enddate		= requestCheckVar(Request("evt_enddate"),10)
gubun			= requestCheckVar(request("gubun"), 1)
lectureid		= requestCheckVar(request("lectureid"), 32)
company_name	= requestCheckVar(request("company_name"), 32)
lecidx			= RequestCheckvar(request("lecidx"),10)
evt_name		= requestCheckVar(request("evt_name"), 100)
isusing			= RequestCheckvar(request("isusing"),1)
menupos			= RequestCheckvar(request("menupos"),10)
idx				= RequestCheckvar(request("idx"),10)
makerid			= requestCheckVar(request("actid"), 32)
diycode			= RequestCheckvar(request("diycode"),10)

Function getsRegedEventUseYN(actid, iidx)
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_academy_event] "
	sqlStr = sqlStr & " WHERE isusing = 'Y' "
	sqlStr = sqlStr & " and actid = '"&actid&"' "
	sqlStr = sqlStr & " and (evt_startdate < '"&evt_enddate&"' and evt_enddate > '"&evt_startdate&"') "
	If iidx <> "" Then
	sqlStr = sqlStr & " and idx <> '"&iidx&"' "	
	End If
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	If rsACADEMYget("cnt") > 0 Then
		getsRegedEventUseYN = "Y"
	Else
		getsRegedEventUseYN = "N"
	End If
	rsACADEMYget.Close
End Function

If gubun = "D" Then			'작가
	If mode = "I" Then
		If getsRegedEventUseYN(lectureid, idx) = "Y" Then
			response.write "<script>alert('이미 사용중인 이벤트가 있습니다.\n이벤트는 현재기간내에 하나만 사용가능합니다.');location.replace('/academy/events/event_regist.asp?menupos="&menupos&"');</script>"
			response.end
		Else
			strSql = ""
			strSql = strSql & " INSERT INTO [db_academy].[dbo].[tbl_academy_event] (gubun, actid, company_name, evt_startdate, evt_enddate, contentsCode, evt_name, isusing, regid, regdate) " & vbcrlf
			strSql = strSql & " VALUES ('D', '"&makerid&"', '"&company_name&"', '"&evt_startdate&"', '"&evt_enddate&"', '"&diycode&"', '"&evt_name&"', '"&isusing&"', '"&session("ssBctID")&"', getdate() ) " 
			dbACADEMYget.execute(strSql)
			response.write "<script>alert('저장 하였습니다');location.replace('/academy/events/eventlist.asp?menupos="&menupos&"');</script>"
			response.end
		End If
	ElseIf mode = "U" Then
		If getsRegedEventUseYN(lectureid, idx) = "Y" Then
			response.write "<script>alert('이미 사용중인 이벤트가 있습니다.\n이벤트는 현재기간내에 하나만 사용가능합니다.');location.replace('/academy/events/event_regist.asp?idx="&idx&"&menupos="&menupos&"');</script>"
			response.end
		Else
			strSql = ""
			strSql = strSql & " UPDATE [db_academy].[dbo].[tbl_academy_event] SET "
			strSql = strSql & " actid = '"&makerid&"'"
			strSql = strSql & " ,company_name = '"&company_name&"'"
			strSql = strSql & " ,evt_startdate = '"&evt_startdate&"'"
			strSql = strSql & " ,evt_enddate = '"&evt_enddate&"'"
			strSql = strSql & " ,contentsCode = '"&diycode&"'"
			strSql = strSql & " ,evt_name = '"&evt_name&"'"
			strSql = strSql & " ,isusing = '"&isusing&"'"
			strSql = strSql & " ,lastupdateid = '"&session("ssBctID")&"'"
			strSql = strSql & " ,lastupdate = getdate()"
			strSql = strSql & " WHERE idx = '"&idx&"' "
			dbACADEMYget.execute(strSql)
			response.write "<script>alert('수정 하였습니다');location.replace('/academy/events/eventlist.asp?menupos="&menupos&"');</script>"
			response.end
		End If
	End If
ElseIf gubun = "L" Then		'강사
	If mode = "I" Then
		If getsRegedEventUseYN(lectureid, idx) = "Y" Then
			response.write "<script>alert('이미 사용중인 이벤트가 있습니다.\n이벤트는 현재기간내에 하나만 사용가능합니다.');location.replace('/academy/events/event_regist.asp?menupos="&menupos&"');</script>"
			response.end
		Else
			strSql = ""
			strSql = strSql & " INSERT INTO [db_academy].[dbo].[tbl_academy_event] (gubun, actid, company_name, evt_startdate, evt_enddate, contentsCode, evt_name, isusing, regid, regdate) " & vbcrlf
			strSql = strSql & " VALUES ('L', '"&lectureid&"', '"&company_name&"', '"&evt_startdate&"', '"&evt_enddate&"', '"&lecidx&"', '"&evt_name&"', '"&isusing&"', '"&session("ssBctID")&"', getdate() ) " 
			dbACADEMYget.execute(strSql)
			response.write "<script>alert('저장 하였습니다');location.replace('/academy/events/eventlist.asp?menupos="&menupos&"');</script>"
			response.end
		End If
	ElseIf mode = "U" Then
		If getsRegedEventUseYN(lectureid, idx) = "Y" Then
			response.write "<script>alert('이미 사용중인 이벤트가 있습니다.\n이벤트는 현재기간내에 하나만 사용가능합니다.');location.replace('/academy/events/event_regist.asp?idx="&idx&"&menupos="&menupos&"');</script>"
			response.end
		Else
			strSql = ""
			strSql = strSql & " UPDATE [db_academy].[dbo].[tbl_academy_event] SET "
			strSql = strSql & " actid = '"&lectureid&"'"
			strSql = strSql & " ,company_name = '"&company_name&"'"
			strSql = strSql & " ,evt_startdate = '"&evt_startdate&"'"
			strSql = strSql & " ,evt_enddate = '"&evt_enddate&"'"
			strSql = strSql & " ,contentsCode = '"&lecidx&"'"
			strSql = strSql & " ,evt_name = '"&evt_name&"'"
			strSql = strSql & " ,isusing = '"&isusing&"'"
			strSql = strSql & " ,lastupdateid = '"&session("ssBctID")&"'"
			strSql = strSql & " ,lastupdate = getdate()"
			strSql = strSql & " WHERE idx = '"&idx&"' "
			dbACADEMYget.execute(strSql)
			response.write "<script>alert('수정 하였습니다');location.replace('/academy/events/eventlist.asp?menupos="&menupos&"');</script>"
			response.end
		End If
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
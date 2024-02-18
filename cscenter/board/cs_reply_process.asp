<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 [안내문구] 기본 카테고리
' History : 이상구 생성
'			2021.09.10 한용민 수정(이문재이사님요청 자사몰 필드추가, 소스표준화, 보안강화)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sqlStr, i, menupos, reguserid, sitename
dim mode, masteridx, detailidx, gubunCode, title, subtitle, contents, dispOrderNo, useYN
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	mode = requestcheckvar(request("mode"),32)
	masteridx = requestcheckvar(getNumeric(request("masteridx")),10)
	detailidx = requestcheckvar(getNumeric(request("detailidx")),10)
	gubunCode = requestcheckvar(request("gubunCode"),4)
	sitename = requestcheckvar(request("sitename"),32)

title = html2db(Trim(request("title")))
subtitle = html2db(Trim(request("subtitle")))
contents = html2db(Trim(request("contents")))

dispOrderNo = requestcheckvar(getNumeric(request("dispOrderNo")),10)
useYN = requestcheckvar(request("useYN"),1)

reguserid = session("ssBctid")

dim refer
refer = request.ServerVariables("HTTP_REFERER")

Select Case mode
	Case "insMaster"
		sqlStr = " insert into db_cs.dbo.tbl_reply_master(gubunCode, sitename, title, dispOrderNo, useYN, reguserid, modiuserid, regdate, lastupdate) values ("
		sqlStr = sqlStr & " '" + CStr(gubunCode) + "', N'"& sitename &"','" + CStr(title) + "', " + CStr(dispOrderNo) + ", '" + CStr(useYN) + "'"
		sqlStr = sqlStr & " , '" + CStr(reguserid) + "', '" + CStr(reguserid) + "', getdate(), getdate() )"

		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장되었습니다.');"
		response.write "	location.replace('/cscenter/board/cs_replymaster_list.asp?menupos=" + CStr(menupos) + "');"
		response.write "</script>"

	Case "modiMaster"
		sqlStr = " update db_cs.dbo.tbl_reply_master "
		sqlStr = sqlStr + " set gubunCode = '" + CStr(gubunCode) + "' "
		sqlStr = sqlStr + " , title = '" + CStr(title) + "' "
		sqlStr = sqlStr + " , dispOrderNo = " + CStr(dispOrderNo) + " "
		sqlStr = sqlStr + " , useYN = '" + CStr(useYN) + "' "
		sqlStr = sqlStr + " , modiuserid = '" + CStr(reguserid) + "' "
		sqlStr = sqlStr + " , lastupdate = getdate() "
		sqlStr = sqlStr + " , sitename = N'"& sitename &"'"
		sqlStr = sqlStr + " where idx = " + CStr(masteridx) + " "

		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장되었습니다.');"
		response.write "	location.replace('" + refer + "');"
		response.write "</script>"

	Case "insDetail"
		sqlStr = " insert into db_cs.dbo.tbl_reply_detail(masteridx, subtitle, contents, dispOrderNo, useYN, reguserid, modiuserid, regdate, lastupdate) "
		sqlStr = sqlStr + " values(" + CStr(masteridx) + ", '" + CStr(subtitle) + "', '" + CStr(contents) + "', " + CStr(dispOrderNo) + ", '" + CStr(useYN) + "', '" + CStr(reguserid) + "', '" + CStr(reguserid) + "', getdate(), getdate()) "

		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장되었습니다.');"
		response.write "	location.replace('/cscenter/board/cs_replydetail_list.asp?menupos=" + CStr(menupos) + "');"
		response.write "</script>"

	Case "modiDetail"
		sqlStr = " update db_cs.dbo.tbl_reply_detail "
		sqlStr = sqlStr + " set masteridx = " + CStr(masteridx) + " "
		sqlStr = sqlStr + " , subtitle = '" + CStr(subtitle) + "' "
		sqlStr = sqlStr + " , contents = '" + CStr(contents) + "' "
		sqlStr = sqlStr + " , dispOrderNo = " + CStr(dispOrderNo) + " "
		sqlStr = sqlStr + " , useYN = '" + CStr(useYN) + "' "
		sqlStr = sqlStr + " , modiuserid = '" + CStr(reguserid) + "' "
		sqlStr = sqlStr + " , lastupdate = getdate() "
		sqlStr = sqlStr + " where idx = " + CStr(detailidx) + " "

		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장되었습니다.');"
		response.write "	location.replace('" + refer + "');"
		response.write "</script>"

	Case Else
		response.write "ERROR"
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

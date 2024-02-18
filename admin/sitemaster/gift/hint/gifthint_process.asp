<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2015.01.26 한용민 생성
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim themeidx, themetype, title, executetime, isusing, orderno, regdate, lastadminid, lastupdate, mode, menupos, sqlStr
dim sIdx, sSortNo, sIsUsing, i, executedate
	themeidx = getNumeric(requestcheckvar(request("themeidx"),10))
	themetype = getNumeric(requestcheckvar(request("themetype"),10))
	title = requestcheckvar(request("title"),128)
	isusing = requestcheckvar(request("isusing"),1)
	orderno = getNumeric(requestcheckvar(request("orderno"),10))
	mode = requestcheckvar(request("mode"),32)
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	lastadminid = session("ssBctId")
	executetime = requestcheckvar(request("executetime"),8)
	executedate = requestcheckvar(request("executedate"),10)

if orderno="" then orderno=99
if isusing="" then isusing="Y"
if executetime="" then executetime="00:00:00"

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
if (InStr(referer,"10x10.co.kr")<1) then
	response.write "not valid Referer"
    'dbget.close() : response.end
end if

if mode="regtheme" then
	if title <> "" and not(isnull(title)) then
		title = ReplaceBracket(title)
	end If
	if checkNotValidHTML(title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.');"
		response.write "	location.replace('" + referer + "');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if themeidx<>"" then
	    sqlStr = "update db_board.dbo.tbl_gifthint" + VbCrlf
	    sqlStr = sqlStr + " set themetype=" + themetype + "" + VbCrlf
	    sqlStr = sqlStr + " ,title='" + db2html(title) + "'" + VbCrlf
	    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
	    sqlStr = sqlStr + " ,executetime='" + trim(executetime) + "'" + VbCrlf
	    sqlStr = sqlStr + " ,orderno=" + orderno + "" + VbCrlf
	    sqlStr = sqlStr + " ,regdate=getdate()" + VbCrlf
	    sqlStr = sqlStr + " ,lastadminid='" + lastadminid + "'" + VbCrlf
	    sqlStr = sqlStr + " ,lastupdate=getdate() where" + VbCrlf
	    sqlStr = sqlStr + " themeidx=" + CStr(themeidx)
	    
	    'response.write sqlStr & "<br>"
	    dbget.Execute sqlStr
	else
	    sqlStr = "insert into db_board.dbo.tbl_gifthint" + VbCrlf
	    sqlStr = sqlStr + " (themetype, title, isusing, executetime, orderno, lastadminid, lastupdate)"+ VbCrlf
	    sqlStr = sqlStr + " values("
	    sqlStr = sqlStr + " " + themetype + "" + VbCrlf
	    sqlStr = sqlStr + " ,'" + db2html(title) + "'" + VbCrlf
	    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
	    sqlStr = sqlStr + " ,'" + trim(executetime) + "'" + VbCrlf
	    sqlStr = sqlStr + " ," + orderno + "" + VbCrlf
	    sqlStr = sqlStr + " ,'" + lastadminid + "'" + VbCrlf
	    sqlStr = sqlStr + " ,getdate()" + VbCrlf
	    sqlStr = sqlStr + " )" + VbCrlf

	    'response.write sqlStr & "<br>"
	    dbget.Execute sqlStr
	end if

	response.write "<script type='text/javascript'>"
	response.write "	location.replace('/admin/sitemaster/gift/hint/gifthint.asp?menupos="& menupos &"');"
	response.write "</script>"
	dbget.close() : response.end

elseif mode="edittheme" then
	'@정렬번호 일괄저장
	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("use"&sIdx)
		if sSortNo="" then sSortNo="0"
		if sIsUsing="" then sIsUsing="N"

		sqlStr = sqlStr & "Update db_board.dbo.tbl_gifthint_item"
		sqlStr = sqlStr & " Set orderno='" & sSortNo & "'"
		sqlStr = sqlStr & " ,isusing='" & sIsUsing & "' Where"		'사이트 메인: 사용여부 > 선노출로 변경
		sqlStr = sqlStr & " itemidx='" & sIdx & "';" & vbCrLf
	next

    'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

	sqlStr = "Update db_board.dbo.tbl_gifthint"
	sqlStr = sqlStr & " Set lastupdate=getdate()"
	sqlStr = sqlStr & " ,lastadminid='" & lastadminid & "' Where"		'사이트 메인: 사용여부 > 선노출로 변경
	sqlStr = sqlStr & " themeidx='" & themeidx & "';" & vbCrLf

    'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	'response.write "	location.replace('/admin/sitemaster/gift/hint/gifthint_item.asp?themeidx="& themeidx &"&executedate="& executedate &"&menupos="& menupos &"');"
	response.write "	location.replace('"& referer &"');"
	response.write "</script>"
	dbget.close() : response.end

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('정상적인 경로가 아닙니다.');"
	response.write "</script>"
	dbget.close() : response.end
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
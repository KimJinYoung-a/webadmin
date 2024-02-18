<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
response.charSet = "utf-8"
'####################################################
' Description :  온라인 해외판매상품
' History : 2013.05.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
dim sqlstr, i, vitemid, siteisusing, multilangcnt, vRegUserID, sitename, mode
    sitename = requestCheckVar(Request("sitename"),32)
    mode = requestCheckVar(Request("mode"),16)
	vRegUserID = session("ssBctId")

multilangcnt=0

if mode="arrmodi" then
    for i=1 to request.form("check").count
        vitemid = getNumeric(requestCheckVar(request.form("check")(i),10))
        siteisusing = requestCheckVar(request.form("siteisusing_"& vitemid),1)

        '//상품의 언어팩의 갯수를 카운트 한다.
        sqlstr = "select count(ml.itemid) as multilangcnt"
        sqlstr = sqlstr & "	from [db_item].[dbo].[tbl_item_multiLang] ml with (nolock)"
        sqlstr = sqlstr & "	where ml.itemid="& vitemid &""

        'response.write sqlstr & "<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
            if not rsget.EOF  then
                multilangcnt = rsget("multilangcnt")
            End If
        rsget.Close

        '//사이트 등록
        sqlstr = " if exists(" + vbcrlf
        sqlstr = sqlstr & "		select top 1 *" + vbcrlf
        sqlstr = sqlstr & "		from db_item.dbo.tbl_item_multiSite_regItem with (nolock)" + vbcrlf
        sqlstr = sqlstr & "		where itemid = '" & vitemid & "' AND sitename = '" & sitename & "'" + vbcrlf
        sqlstr = sqlstr & "	)" + vbcrlf
        sqlstr = sqlstr & "		update db_item.dbo.tbl_item_multiSite_regItem set" + vbcrlf
        sqlstr = sqlstr & "		isusing=N'"& Siteisusing &"'" + vbcrlf
        sqlstr = sqlstr & "		,lastupdate = getdate()" + vbcrlf
        sqlstr = sqlstr & "		,lastuserid = N'"& vRegUserID &"'" + vbcrlf
        sqlstr = sqlstr & "		,multilangcnt="& multilangcnt &"" + vbcrlf
        sqlstr = sqlstr & "		WHERE itemid = '" & vitemid & "' AND sitename = '" & sitename & "' " + vbcrlf
        sqlstr = sqlstr & " else" + vbcrlf
        sqlstr = sqlstr & " 	insert into db_item.dbo.tbl_item_multiSite_regItem(" + vbcrlf
        sqlstr = sqlstr & " 	itemid, sitename, isusing, multilangcnt, regdate, reguserid, lastupdate, lastuserid" + vbcrlf
        sqlstr = sqlstr & "		) values (" + vbcrlf
        sqlstr = sqlstr & "		N'" & vitemid & "', N'" & sitename & "', N'" & Siteisusing & "', "& multilangcnt &", getdate(), N'" & vRegUserID & "'" + vbcrlf
        sqlstr = sqlstr & "		, getdate(), N'" & vRegUserID & "' " + vbcrlf
        sqlstr = sqlstr & "		)"

        'response.write sqlstr &"<Br>"
        dbget.execute sqlstr
    next

    response.write "<script type='text/javascript'>"
    response.write "	alert('저장되었습니다.');"
    session.codePage = 949
    response.write "	parent.document.location.reload();"
    response.write "</script>"
else
    response.write "<script type='text/javascript'>"
    response.write "	alert('정상적인 경로가 아닙니다.');"
    session.codePage = 949
    response.write "</script>"
end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>
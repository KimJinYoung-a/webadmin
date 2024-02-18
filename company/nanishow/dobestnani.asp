<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

dim idispnum, olditemid, iitemid, idisptext

idispnum = request("idispnum")
olditemid = request("olditemid")
iitemid = request("iitemid")
idisptext = request("idisptext")


'response.write idispnum + "<br>"
'response.write olditemid + "<br>"
'response.write iitemid + "<br>"
'response.write idisptext + "<br>"

dim sqlStr

sqlStr = "update tbl_etc_special"
sqlStr = sqlStr + " set dispnum='" + idispnum + "',"
sqlStr = sqlStr + " itemid=" + CStr(iitemid) + ","
sqlStr = sqlStr + " disptitle='" + Html2DB(idisptext) + "'"
sqlStr = sqlStr + " where sitename='nanistyle'"
sqlStr = sqlStr + " and itemid=" + CStr(olditemid)

rsget.Open sqlStr,dbget,1

dim referer

referer = request.ServerVariables("HTTP_REFERER")

response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
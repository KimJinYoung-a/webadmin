<%@ language=vbscript %>
<% option explicit %>
<% 
Response.AddHeader "Cache-Control","no-cache" 
Response.AddHeader "Expires","0" 
Response.AddHeader "Pragma","no-cache" 
%> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/noreplyboardcls.asp" -->
<%
dim referer

referer = request.ServerVariables("HTTP_REFERER")

dim mode,iid,checkflag

mode = request("mode")
iid = request("id")
checkflag = request("checkflag")


''필수 입력 체크.
if (mode="") or (iid="") then 
		dbget.close()	:	response.End
end if

dim oneboard,errmsg
set oneboard = new CNoReplyBoard
if mode="del" then
	errmsg = oneboard.delboard(iid)
elseif mode="check" then
	errmsg = oneboard.checkboard(iid,checkflag)
end if

set oneboard = Nothing

if errmsg<>"" then
	response.write "시스템 오류 : " + errmsg
else
	if mode="del" then
		response.write "<script>alert('삭제되었습니다.')</script>"
	else
		response.write "<script>alert('저장되었습니다.')</script>"
	end if

	response.write "<script>location.replace('" + referer + "')</script>"
	
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
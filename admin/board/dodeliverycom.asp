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

dim masterid
dim tx_com
dim writer

masterid = request("masterid")
tx_com = html2db(request("tx_com"))
writer = request("writer")


''�ʼ� �Է� üũ.
if (masterid="") or (tx_com="") or _
	(writer="") then 
		dbget.close()	:	response.End
end if

dim oneboard,errmsg
set oneboard = new CNoReplyBoard
errmsg = oneboard.writeCom(masterid,tx_com,writer)
set oneboard = Nothing

if errmsg<>"" then
	response.write "�ý��� ���� : " + errmsg
else
	response.write "<script>alert('����Ǿ����ϴ�.')</script>"
	response.write "<script>location.replace('" + referer + "')</script>"
	
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
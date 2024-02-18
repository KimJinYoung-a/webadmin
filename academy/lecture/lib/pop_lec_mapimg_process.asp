<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'#######################################################
'	History	:  2010.01.20 ÇÑ¿ë¹Î »ý¼º
'	Description : °í°´¼¾Å¸
'#######################################################
%>
<%
dim mode , map_idx , sql
	mode = RequestCheckvar(request("mode"),10)
	map_idx = RequestCheckvar(request("map_idx"),10)

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if mode = "del" then
	sql = "delete from dbo.tbl_map_Info where map_idx = "& map_idx &""
	
	'response.write sql &"<br>"
	dbACADEMYget.execute sql
	
end if

response.write "<script>alert('ok');</script>"
response.write "<script>location.href='"&referer&"'</script>"		
%>

<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
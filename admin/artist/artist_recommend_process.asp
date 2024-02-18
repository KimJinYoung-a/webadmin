<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.10 한용민 생성
'	Description : artist gallery
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->

<% 
dim mode , artist_idx
	mode = requestcheckvar(request("mode"),25)
	artist_idx = requestcheckvar(request("artist_idx"),10)

dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")	

''//삭제
if mode = "del" then
	
	if artist_idx = "" then
%>		
		<script language="javascript">
		alert('아티스트번호가 없습니다');
		history.go(-1);
		</script>
<%	
	dbget.close : response.end
	end if
	
	sql = "update db_contents.dbo.tbl_artist_recommend set" + vbcrlf
	sql = sql & " isusing = 'N'" + vbcrlf
	sql = sql & " where artist_idx ="&artist_idx&"" + vbcrlf
		
	'response.write sql &"<Br>"
	dbget.execute sql
end if	
%>	

<script language="javascript">
alert('저장되었습니다');
location.href='/admin/artist/artist_recommend.asp';
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
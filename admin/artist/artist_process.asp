<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.09 �ѿ�� ����
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
dim mode , itemid
	mode = requestcheckvar(request("mode"),25)
	itemid = requestcheckvar(request("itemid"),10)

dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")	


''//���ι�� ��ǰ ���
if mode = "mainbanneritem" then
	sql = "insert into db_contents.dbo.tbl_artist_banner (itemid ,gubun ,isusing)" + vbcrlf
	sql = sql & " values (" + vbcrlf
	sql = sql & " "&itemid&"" + vbcrlf	
	sql = sql & " ,0" + vbcrlf	
	sql = sql & " ,'Y'" + vbcrlf		
	sql = sql & " )" + vbcrlf
	
	'response.write sql &"<Br>"
	dbget.execute sql
end if	
%>	
<script language="javascript">
alert('����Ǿ����ϴ�');
location.href='<%=referer%>';
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
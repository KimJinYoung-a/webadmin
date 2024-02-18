<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station 
' Hieditor : 2009.04.02 ÇÑ¿ë¹Î »ý¼º
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<% 
dim idx
	idx = request("idx")

dim sql
	sql = "update db_culture_station.dbo.tbl_culturestation_editor_comment set isusing = 'N' where idx = "& idx &""
	
	response.write sql&"<br>"
	dbget.execute sql
%>

<script>
opener.location.reload();
self.close();
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


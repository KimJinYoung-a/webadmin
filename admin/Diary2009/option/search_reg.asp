<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%
dim idx , search_order , Diaryid_search
	idx = request("idx")
	search_order = request("search_order")
	Diaryid_search = request("Diaryid_search")
	
dim sql
	sql = "update db_diary2010.dbo.tbl_diary_info_search set search_order="&search_order&" where idx = "&idx&""
	response.write sql

dbget.execute sql
%>	

<script>
location.href="/admin/diary2009/option/pop_diary_info_reg.asp?mode=modify&diaryid=<%=Diaryid_search%>"
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
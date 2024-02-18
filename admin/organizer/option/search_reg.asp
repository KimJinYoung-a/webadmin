<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<%
dim idx , search_order , Diaryid_search
	idx = request("idx")
	search_order = request("search_order")
	Diaryid_search = request("Diaryid_search")
	
dim sql
	sql = "update db_diary2010.dbo.tbl_organizer_info_search set search_order="&search_order&" where idx = "&idx&""
	response.write sql

dbget.execute sql
%>	

<script>
location.href="/admin/organizer/option/pop_organizer_info_reg.asp?mode=modify&diaryid=<%=Diaryid_search%>"
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
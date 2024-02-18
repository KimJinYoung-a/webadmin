<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.28 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
dim mode , isusing ,idx,itemid
	mode = request("mode")
	itemid = request("itemid")
	isusing = request("isusing")
	idx = request("idx")

dim sql


'//신규
if mode = "new" then

	sql = ""
	sql = "insert into db_diary2009.dbo.tbl_organizer_alpha (itemid, isusing) values ("
	sql = sql & " '"& itemid & "','" & isusing & "')"

	response.write sql
	dbget.execute sql
%>
<script language="javascript">
	opener.location.href = '/admin/organizer/alpha_list.asp';
	window.close();
</script>
<%		 
 '//수정
elseif mode = "edit" then
	
	sql = ""
	sql = "update db_diary2009.dbo.tbl_organizer_alpha set"
	sql = sql & " itemid = '"& itemid &"'," 	
	sql = sql & " isusing = '"& isusing &"'" 	
	sql = sql & " where idx = "& idx &""
	
	response.write sql &"<br>"
	dbget.execute sql
%>
<script language="javascript">
	location.href = '/admin/organizer/alpha_list.asp';
	//window.close();
</script>

<%		
end if	

%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
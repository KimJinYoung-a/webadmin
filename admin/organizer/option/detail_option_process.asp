<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
dim mode , mode_type,keyword_option, diaryid
	mode = request("mode")
	mode_type = request("mode_type")
	keyword_option = request("keyword_option")
	diaryid = request("diaryid")

dim sql
	
'//키워드변경
if mode = "keyword" then

	'//추가
	if mode_type = "insert" then

		sql = ""
		sql = "insert into db_diary2010.dbo.tbl_organizer_keyword_master (organizerid , keyword_option) values ("
		sql = sql & " "& diaryid & ",'" & keyword_option & "')"

		response.write sql
		dbget.execute sql
	
	'//삭제
	elseif mode_type = "delete" then
	
		sql = ""
		sql = "delete from db_diary2010.dbo.tbl_organizer_keyword_master where organizerid = "& diaryid & " and keyword_option = '" & keyword_option & "'"

		response.write sql
		dbget.execute sql
	
	end if

end if
%>

	<script language="javascript">
		parent.location.href = '/admin/organizer/option/detail_option.asp?diaryid=<%=diaryid%>';
		window.close();
	</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
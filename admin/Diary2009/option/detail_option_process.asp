<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.10 한용민 생성
'	Description : 다이어리스토리
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/DiaryCls.asp"-->

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
		sql = "insert into db_diary2010.dbo.tbl_keyword_master (diaryid , keyword_option) values ("
		sql = sql & " "& diaryid & ",'" & keyword_option & "')"

		response.write sql
		dbget.execute sql
	
	'//삭제
	elseif mode_type = "delete" then
	
		sql = ""
		sql = "delete from db_diary2010.dbo.tbl_keyword_master where diaryid = "& diaryid & " and keyword_option = '" & keyword_option & "'"

		response.write sql
		dbget.execute sql
	
	end if

end if
%>

	<script language="javascript">
		parent.location.href = '/admin/diary2009/option/detail_option.asp?diaryid=<%=diaryid%>';
		window.close();
	</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
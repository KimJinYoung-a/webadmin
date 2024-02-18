<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
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
dim mode , option_value , option_order , isusing , mode_type , idx, stype
	mode = request("mode")
	option_value = request("option_value")
	option_order = request("option_order")
	isusing = request("isusing")
	mode_type = request("mode_type")
	idx = request("idx")
	stype = request("type")
	
	'response.write mode &"<br>"
	'response.write mode_type &"<br>"
dim sql

'// 컨텐츠	 
if mode_type = "contents" then	
	

	
'//키워드	
elseif mode_type = "keyword" then	
	
	'//신규
	if mode = "new" then

		sql = ""
		sql = "insert into db_diary2010.dbo.tbl_organizer_keyword_option (option_value , option_order,type, isusing) values ("
		sql = sql & " '"& option_value & "'," & option_order & ",'" & stype & "','" & isusing & "')"

		response.write sql
		dbget.execute sql
%>
	<script language="javascript">
		opener.location.href = '/admin/organizer/option/keyword_option.asp';
		window.close();
	</script>
<%		 
	 '//수정
	elseif mode = "edit" then
		
		sql = ""
		sql = "update db_diary2010.dbo.tbl_organizer_keyword_option set"
		sql = sql & " option_value = '"& option_value &"'," 	
		sql = sql & " option_order = "& option_order &"," 	
		sql = sql & " type = '"& stype &"'," 
		sql = sql & " isusing = '"& isusing &"'" 	
		sql = sql & " where idx = "& idx &""
		
		response.write sql &"<br>"
		dbget.execute sql
%>
	<script language="javascript">
		location.href = '/admin/organizer/option/keyword_option.asp';
		//window.close();
	</script>
	
	<%		
	end if	
end if
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
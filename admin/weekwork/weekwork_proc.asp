<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/weekwork/weekworkCls.asp"-->

<%
dim lastweek, thisweek, week_num, week_month, team
dim editsave
dim	mode, N
dim idx, userid, username
	idx = request("idx")
	mode = request("mode")
	team = request("team")
	week_num = request("Sweek_num")
	username = request("username")
	lastweek = request("lastweek")
	thisweek = request("thisweek")
	week_month = request("Sweek_month")
	
	userid = session("ssBctId")
	username = session("ssBctCname") 

'수정모드일때 리퀘스트로 값 받아온걸 업데이트 쿼리날림
	dim sqlstr, getdate
	if mode = "EDIT" then 
		sqlstr = " update db_temp.dbo.tbl_weekwork set " '수정모드일때 db업데이트
		sqlstr = sqlstr & " lastweek = '"& lastweek &"' "
		sqlstr = sqlstr & " ,thisweek = '"& thisweek &"' "
		sqlstr = sqlstr & " ,week_month = '"& week_month &"' "
		sqlstr = sqlstr & " ,week_num = '"& week_num &"' "		
		sqlstr = sqlstr & " ,rewrite_date = getdate() "
		sqlstr = sqlstr & " where idx = "& idx &" "
		dbget.execute sqlstr
	
	'신규입력 모드일때 리퀘스트로 받아온 값을 인설트 쿼리 날림
	elseif mode = "NEW" then
													
		sqlstr = "insert into db_temp.dbo.tbl_weekwork (team, userid, username, week_month, week_num, lastweek, thisweek, write_date, rewrite_date)"
		sqlstr = sqlstr & " values ('" & session("ssAdminPsn") & "','" & userid & "','" & username & "','" & week_month & "' , '" & week_num & "','" & lastweek & "'"
		sqlstr = sqlstr & " ,'" &thisweek & "',getdate(),getdate())"
		dbget.execute sqlstr
	end if
%>

<script language = "javascript">
	alert("저장되었습니다."); //저장되었습니다 라는 메시지띄움
	opener.location.reload(); //이창을 띄운 부모창을 리로드함
	self.close();			  //이창을 닫음
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
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

'��������϶� ������Ʈ�� �� �޾ƿ°� ������Ʈ ��������
	dim sqlstr, getdate
	if mode = "EDIT" then 
		sqlstr = " update db_temp.dbo.tbl_weekwork set " '��������϶� db������Ʈ
		sqlstr = sqlstr & " lastweek = '"& lastweek &"' "
		sqlstr = sqlstr & " ,thisweek = '"& thisweek &"' "
		sqlstr = sqlstr & " ,week_month = '"& week_month &"' "
		sqlstr = sqlstr & " ,week_num = '"& week_num &"' "		
		sqlstr = sqlstr & " ,rewrite_date = getdate() "
		sqlstr = sqlstr & " where idx = "& idx &" "
		dbget.execute sqlstr
	
	'�ű��Է� ����϶� ������Ʈ�� �޾ƿ� ���� �μ�Ʈ ���� ����
	elseif mode = "NEW" then
													
		sqlstr = "insert into db_temp.dbo.tbl_weekwork (team, userid, username, week_month, week_num, lastweek, thisweek, write_date, rewrite_date)"
		sqlstr = sqlstr & " values ('" & session("ssAdminPsn") & "','" & userid & "','" & username & "','" & week_month & "' , '" & week_num & "','" & lastweek & "'"
		sqlstr = sqlstr & " ,'" &thisweek & "',getdate(),getdate())"
		dbget.execute sqlstr
	end if
%>

<script language = "javascript">
	alert("����Ǿ����ϴ�."); //����Ǿ����ϴ� ��� �޽������
	opener.location.reload(); //��â�� ��� �θ�â�� ���ε���
	self.close();			  //��â�� ����
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
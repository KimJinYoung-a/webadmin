<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim eventidx,eventname,eventansCount,eventFile1,eventFile2,winnerOpenYn,mode,sql
eventname=request("eventname")
eventidx=request("eventidx")
mode=request("mode")

if mode="edit" and eventidx <>"" then

	sql = "select eventname, eventansCount ,eventFile1 ,eventFile2 ,winnerOpenYn from [db_cts].[dbo].[tbl_2007_diary_event_master] " &_
				" where eventidx =" & eventidx

	db2_rsget.open sql,db2_dbget,1


		if not db2_rsget.eof then
			eventname	=	db2_rsget("eventname")
			eventansCount= db2_rsget("eventansCount")
			eventFile1= db2_rsget("eventFile1")
			eventFile2= db2_rsget("eventFile2")
			winnerOpenYn= db2_rsget("winnerOpenYn")
		end if

end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<table border="1" cellpadding="0" cellspacing="0" class="a">
	<form name="eventfrm" method="post" action="http://testimgstatic.10x10.co.kr/linkweb/doDiary_quiz_event.asp" enctype="multipart/form-data">
	<input type="hidden" name="mode" value="<%= mode %>">
	<input type="hidden" name="eventidx" value="<%= eventidx %>">
	<tr>
		<td>이벤트명</td>
		<td><input type="text" name="eventname" value="<%= eventname %>" /></td>
	</tr>
	<tr>
		<td>정답글자수</td>
		<td><input type="text" name="eventansCount" size="2" maxlength="2" value="<%= eventansCount %>" /></td>
	</tr>
	<tr>
		<td>이미지1<br><font color="red">width:750(고정)</font></td>
		<td><input type="file" name="file1" size="32" /><br><%= eventFile1 %></td>
	</tr>
	<tr>
		<td>이미지2<br><font color="red">width:750(고정)</font></td>
		<td><input type="file" name="file2" size="32" /><br><%= eventFile2 %></td>
	</tr>
	<tr>
		<td>당첨자보여주기</td>
		<td>
			<input type="radio" value="Y" name="winnerOpenYn" <% if winnerOpenYn="Y" then response.write "checked" %> /> Y
			<input type="radio" value="N" name="winnerOpenYn" <% if winnerOpenYn="N" or trim(winnerOpenYn)="" then response.write "checked" %> /> N
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="확인" /></td>
	</tr>
	</form>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/db2close.asp" -->

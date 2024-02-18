<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/weekwork/weekworkCls.asp"-->

<%
	dim idx, week_num, username, lastweek, thisweek, i, week_month
	dim sqlstr, sqlsearch, arrlist
		lastweek = request("lastweek")
		thisweek = request("thisweek")
		idx = request("idx")
	
	dim opart
		set opart = new CWeekwork
			opart.getpartname()

	if idx <> "" then
		sqlsearch = sqlsearch & " and idx="& idx &""
	end if
		
		sqlstr = "select top 1"
		sqlstr = sqlstr & " idx, username, week_num, lastweek, thisweek, week_month"
		sqlstr = sqlstr & " from db_temp.dbo.tbl_weekwork"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by idx desc"
		
		rsget.Open sqlstr, dbget, 1
		
	if not rsget.EOF then
		arrlist = rsget.getrows()
	end if
		
		rsget.close
		
		idx = arrlist(0,0)
		username = arrlist(1,0)
		week_num = arrlist(2,0)
		lastweek = arrlist(3,0)
		thisweek = arrlist(4,0)
		week_month = arrlist(5,0)
		'write_date = arrlist(5,0)
		'team = arrlist(1,0)
		'userid = arrlist(2.0)
		'write_date = arrlist(4,0)
		'rewrite_date = arrlist(5,0)
		'lastweek = arrlist(7,0)
		'thisweek = arrlist(8,0)
%>


	<table border="1" width="100%">
	
		<tr>
			<td>번호</td>
			<td><%=idx%></td>
		</tr>
	
		<tr>
			<td>이름</td>
			<td><%=username%></td>
		</tr>
		
		<tr>
			<td>주차</td>
			<td>
				<%=week_month%>월
				<%=week_num%>주차 [지금은 <%=month(date)%>월 <%=weekselect%>주차 입니다]
			</td>
		</tr>
	
		<tr>
			<td colspan="2">지난주 한일</td>
		</tr>
		<tr>
			<td colspan="2" width="100%" height="150px">
			<textarea name="lastweek" class="textarea" style="width:100%; height:150px;"><%= lastweek %></textarea>
			</td>
		</tr>
		
		<tr>
			<td colspan="2">이번주 할일</td>
		</tr>
		<tr>
			<td colspan="2"  width="100%" height="150px">
				<textarea name="lastweek" class="textarea" style="width:100%; height:150px;"><%= thisweek %></textarea>
			</td>
		</tr>
		
		<tr align="center">
			<td colspan="2">
				<input type="button" name="editsave" value="닫기" onclick="self.close()">
			</td>
		</tr>
	</table>
	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sqlStr
dim ARrdate(), ARcnt(), AReventid()
dim ArdateOut(), ArOut()

dim ResultCount,ResultCount2
dim linaTotal, lina_only10Total

sqlStr = "select count(userid) as cnt, eventid"
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
sqlStr = sqlStr + " where eventid in ('lina','lina_only10')"
sqlStr = sqlStr + "group by eventid"
rsget.Open sqlStr,dbget,1
do until rsget.eof
	if rsget("eventid")="lina" then
		linaTotal = rsget("cnt")
	else
		lina_only10Total = rsget("cnt")
	end if
	rsget.movenext
loop
rsget.Close


sqlStr = "select top 30 Left(convert(varchar,regdate,21),10) as rdate, count(userid) as cnt, eventid"
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
sqlStr = sqlStr + " where eventid in ('lina','lina_only10')"
sqlStr = sqlStr + " group by  Left(convert(varchar,regdate,21),10), eventid"
sqlStr = sqlStr + " order by rdate desc"
rsget.Open sqlStr,dbget,1

ResultCount = rsget.RecordCount
redim ARrdate(ResultCount)
redim ARcnt(ResultCount)
redim AReventid(ResultCount)

dim i
i=0
do until rsget.Eof
	ARrdate(i) = rsget("rdate")
	ARcnt(i) = rsget("cnt")
	AReventid(i) = rsget("eventid")
	i=i+1
	rsget.MoveNext
loop

rsget.Close


sqlStr = "select top 30 Left(convert(varchar,regdate,21),10) as rdate, count(userid) as cnt"
sqlStr = sqlStr + " from [db_user].[dbo].tbl_deluser"
sqlStr = sqlStr + " where regdate >'2003-04-14'"
sqlStr = sqlStr + " group by  Left(convert(varchar,regdate,21),10)"
sqlStr = sqlStr + " order by rdate desc"
rsget.Open sqlStr,dbget,1

ResultCount2 = rsget.RecordCount

redim ArdateOut(ResultCount2)
redim ArOut(ResultCount2)

i=0
do until rsget.Eof
	ArdateOut(i) = rsget("rdate")
	ArOut(i) = rsget("cnt")
	i=i+1
	rsget.MoveNext
loop

rsget.Close

dim ps1
%>
<table width="700" border="0" cellspacing="0" cellpadding="0" class="a">
  <tr>
  	<td colspan="3" align="right">Lina : <%= FormatNumber(linaTotal,0) %> &nbsp;&nbsp;&nbsp;&nbsp;Lina_10 : <%= FormatNumber(lina_only10Total,0) %>&nbsp;&nbsp;&nbsp;&nbsp;Total : <%= FormatNumber(linaTotal+lina_only10Total,0) %></td>
  </tr>
  <tr>
  	<td width="120">³―Β₯</td>
  	<td>Graph</td>
  	<td width="120">Count</td>
  </tr>
  <% for i=0 to ResultCount-1 %>
  <%
  	ps1 = CLng(ARcnt(i)/(linaTotal+lina_only10Total)*500)
  %>
  <tr>
  	<td width="120"><%= ARrdate(i) %></td>
  	<td>
  		<% if AReventid(i)="lina" then %>
  		<img src="/images/dot1.gif" height="2" width="<%= ps1 %>" >
  		<% else %>
  		<img src="/images/dot2.gif" height="2" width="<%= ps1 %>" >
  		<% end if %>
  	</td>
  	<td width="120"><%= AReventid(i) %> : <%= FormatNumber(ARcnt(i),0) %></td>
  </tr>
  <% next %>
  <tr>
   <td colspan="3"><hr></td>
  </tr>
  <% for i=0 to ResultCount2-1 %>
  <%
  	ps1 = CLng(ArOut(i)/(linaTotal+lina_only10Total)*500)
  %>
  <tr>
  	<td width="120"><%= ArdateOut(i) %></td>
  	<td>
  		<img src="/images/dot2.gif" height="2" width="<%= ps1 %>" >
  	</td>
  	<td width="120">Ε»Επ : <%= FormatNumber(ArOut(i),0) %></td>
  </tr>
  <% next %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
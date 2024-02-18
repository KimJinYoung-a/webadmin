<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nextdateStr,searchnextdate

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)


dim sqlStr
dim totalCnt,totalCnt10,totalCnt10withID,totalCntLina
dim percentLina,percentUser

sqlStr = " select count(idx) as cnt"
sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
sqlStr = sqlStr + " where regdate>='" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
sqlStr = sqlStr + " and regdate<'" + searchnextdate + "'"
sqlStr = sqlStr + " and jumundiv<>9"
sqlStr = sqlStr + " and ipkumdiv>=4"
sqlStr = sqlStr + " and cancelyn='N'"

rsget.Open sqlStr,dbget,1
totalCnt = rsget("cnt")
rsget.close

sqlStr = sqlStr + " and sitename='10x10'"
rsget.Open sqlStr,dbget,1
totalCnt10 = rsget("cnt")
rsget.close

rsget.Open sqlStr + " and userid<>''",dbget,1
totalCnt10withID = rsget("cnt")
rsget.close

sqlStr = " select count(m.idx) as cnt"
sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m, [db_user].[dbo].tbl_user_n n"
sqlStr = sqlStr + " where m.regdate>='" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
sqlStr = sqlStr + " and m.regdate<'" + searchnextdate + "'"
sqlStr = sqlStr + " and m.jumundiv<>9"
sqlStr = sqlStr + " and m.ipkumdiv>=4"
sqlStr = sqlStr + " and m.cancelyn='N'"
sqlStr = sqlStr + " and m.sitename='10x10'"
sqlStr = sqlStr + " and m.userid<>''"
sqlStr = sqlStr + " and m.userid=n.userid"
sqlStr = sqlStr + " and n.eventid in ('lina','lina_only10')"

rsget.Open sqlStr ,dbget,1
totalCntLina = rsget("cnt")
rsget.close

if totalCnt10<>0 then
	percentUser = CLng(totalCnt10withID/totalCnt10*100)
end if

if totalCnt10withID<>0 then
	percentLina = CLng(totalCntLina/totalCnt10withID*100)
end if
%>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="760" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="150">총 구매건수</td>
	<td><%= totalCnt %></td>
</tr>
<tr>
	<td width="150">텐바이텐 구매건수</td>
	<td><%= totalCnt10 %></td>
</tr>
<tr>
	<td width="150">회원 구매건수</td>
	<td><%= totalCnt10withID %>(<%= percentUser%>%)</td>
</tr>
<tr>
	<td width="150">비회원 구매건수</td>
	<td><%= totalCnt10-totalCnt10withID %></td>
</tr>
<tr>
	<td width="150">회원 구매중 라이나건수</td>
	<td><%= totalCntLina %>(<%= percentLina %>%)</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
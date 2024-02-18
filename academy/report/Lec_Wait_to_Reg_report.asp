<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_reportcls.asp"-->

<%
dim report
dim fromdate
dim yyyy1,mm1,i
dim percent

yyyy1=RequestCheckvar(request("yyyy1"),4)
mm1=RequestCheckvar(request("mm1"),2)

if yyyy1="" then
	yyyy1=year(now)
end if
if mm1="" then
	mm1=month(dateadd("m","-2",now))
end if

fromdate=dateadd("d","-1",dateserial(yyyy1,mm1,"01"))

set report = new CJumunMaster
report.FRectFromDate=fromdate
report.GetWaitUserReportbyLecDate()

%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawYMBox yyyy1,mm1 %> ~ 현재 &nbsp;&nbsp;
		<br>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#BABABA" class="a">
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center">강좌월</td>
		<td></td>
		<td width="150">결제 건수/전체 대기자수</td>
	</tr>
	<% for i=0 to report.FResultCount-1 %>
	<% percent=Clng(report.FMasterItemList(i).Fregcount/report.FMasterItemList(i).FLectotcnt*100) %>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center"><%= report.FMasterItemList(i).FLecDate %></td>
		<td><img src="/images/dot4.gif" height="2" width="<%= percent/100*600 %>"></td>
		<td><%= report.FMasterItemList(i).Fregcount %> / <%= report.FMasterItemList(i).FLectotcnt %> (<%= percent %>%)</td>
	</tr>

	<% next %>

</table>

<%	set report = Nothing	%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
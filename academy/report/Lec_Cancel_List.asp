<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_reportcls.asp"-->
<%

dim research
dim yyyy1,mm1, i
dim fromdate, searchtype
dim TotalCnt,TotalCancelCnt

yyyy1= RequestCheckvar(request("yyyy1"),4)
mm1=RequestCheckvar(request("mm1"),2)
searchtype=RequestCheckvar(request("searchtype"),1)

if yyyy1="" then	yyyy1=year(now)
if mm1="" then mm1=month(now)-2

fromdate=dateadd("d","-1",dateserial(yyyy1,mm1,"01"))


dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate=fromdate
oreport.FRectSearchType = searchtype
oreport.GetCancelListbyLec_Date
%>


<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawYMBox yyyy1,mm1 %> ~ 현재 (강좌월)&nbsp;&nbsp;
		<input type="radio" name="searchtype" value="" <% if searchtype="" then response.write "checked" %>>전체
		<input type="radio" name="searchtype" value="2" <% if searchtype="2" then response.write "checked" %>>결제전취소
		<input type="radio" name="searchtype" value="4" <% if searchtype="4" then response.write "checked" %>>결제후취소
		<br>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00" class="a">
	<tr>
		<td width="100" align="center"><font color="#FFFFFf">강좌월</font></td>
		<td><font color="#FFFFFf">취소 횟수</font></td>
		<td width="100"><font color="#FFFFFf">취소 횟수</font></td>
		<td width="100"><font color="#FFFFFF">전체 건수</font></td>
		<td width="80"><font color="#FFFFFF">취소율</font></td>
	</tr>
	<% for i=0 to oreport.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center"><%= oreport.FMasterItemList(i).FLecDate %></td>
		<td>
			<img src="/images/dot2.gif" height="2" width="<%= Clng(oreport.FMasterItemList(i).FCancelCnt/600*100) %>">
			<br>
			<img src="/images/dot1.gif" height="2" width="<%= Clng(oreport.FMasterItemList(i).FLectotcnt/600*100) %>">
			</td>
		<td align="right"><%= FormatNumber(oreport.FMasterItemList(i).FCancelCnt,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FMasterItemList(i).FLectotcnt,0) %></td>
		<td align="right">
			<% if oreport.FMasterItemList(i).FLectotcnt<>0 then %>
			<%= CLng(oreport.FMasterItemList(i).FCancelCnt/oreport.FMasterItemList(i).FLectotcnt*100*100)/100 %> %
			<% end if %>
		</td>
	</tr>
	<% TotalCancelCnt=TotalCancelCnt+oreport.FMasterItemList(i).FCancelCnt %>
	<% TotalCnt=TotalCnt+oreport.FMasterItemList(i).FLectotcnt %>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td align="center">TotalCount</td>
		<td></td>
		<td align="right"><%= FormatNumber(TotalCancelCnt,0) %></td>
		<td align="right"><%= FormatNumber(TotalCnt,0) %></td>
		<td align="right">
			<% if TotalCnt<>0 then %>
			<%= CLng(TotalCancelCnt/TotalCnt*100*100)/100 %> %
			<% end if %>
		</td>
	</tr>
</table>

<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
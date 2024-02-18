<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim oreport
dim oldlist
dim stdate
dim yyyy1,mm1
Dim sort
Dim topcnt

yyyy1 = request("yyyy1")
mm1	  = request("mm1")
sort	  = request("sort")
topcnt	  = request("topcnt")
If topcnt="" Then topcnt=10

if yyyy1="" then
	stdate = CStr(Now)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if
oldlist = request("oldlist")

set oreport = new CJumunMaster
oreport.FRectOldJumun = oldlist
oreport.FRectFromDate = yyyy1 + "-" + mm1
oreport.FRectCnt = topcnt
oreport.GetLectureCountUserID

Dim i,p2
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a">
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역&nbsp;&nbsp;&nbsp;
			검색기간 : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;&nbsp;결과물갯수 : <input type="text" name="topcnt" value="<% = topcnt %>" size="5">
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">아이디</font></td>
          <td class="a" width="100"><font color="#FFFFFF">총횟수</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
				<td width="120" height="10" align="center">
					 <%= oreport.FMasterItemList(i).Fsitename %>
				</td>
				<td class="a" width="100" align="right">
					 <%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건
				</td>
        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
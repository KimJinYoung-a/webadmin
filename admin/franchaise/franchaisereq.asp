<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/offshop/franchaisereqcls.asp" -->

<%
dim page,gubun, onlymifinish
dim research

page = request("page")
gubun = request("gubun")
onlymifinish = request("onlymifinish")
research = request("research")

if research="" and onlymifinish="" then onlymifinish="on"


if (page = "") then
        page = "1"
end if

dim ofran
set ofran = new CFranChaiseReqList
ofran.FPageSize = 30
ofran.FCurrPage = page
ofran.FrectGubun = gubun
ofran.FRectOnlymifinish = onlymifinish
ofran.FRectNotDel = ""

ofran.GetReqList

dim i


%>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		<input type="radio" name="gubun" value="" <% if gubun="" then response.write "checked" %> >전체
		<input type="radio" name="gubun" value="1" <% if gubun="1" then response.write "checked" %> >투자상담
		<input type="radio" name="gubun" value="2" <% if gubun="2" then response.write "checked" %> >가맹점상담
		&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="onlymifinish" <% if onlymifinish="on" then response.write "checked" %> >처리안된목록

		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
  <tr bgcolor="#DDDDFF">
    <td width="120">신청자</td>
    <td width="80">예상투자비</td>
    <!-- <td width="80">매장개설자금</td> -->
    <td width="180">개설희망지역</td>
    <td width="110">신청일</td>
    <td width="40">처리구분</td>
  </tr>
<% for i = 0 to (ofran.FResultCount - 1) %>
  <tr bgcolor="#FFFFFF">
    <td width="120"><a href="franchaisereqdetail.asp?id=<%= ofran.FItemList(i).Fidx %>&gubun=<%= gubun %>&onlymifinish=<%=onlymifinish%>"> <%= ofran.FItemList(i).Fusername %></a></td>
    <td width="80"><%= ofran.FItemList(i).Finvest_money %></td>
   <!-- <td width="80"><%= ofran.FItemList(i).Fshop_mayfund %></td> -->
    <td width="180"><%= ofran.FItemList(i).Fshop_mayarea %></td>
    <td width="110"><%= ofran.FItemList(i).Fregdate %></td>
    <td width="40">
        <% if (ofran.FItemList(i).Ffinishflag = "0") then %>
      	<font color="red">미완료</font>
      	<% elseif (ofran.FItemList(i).Ffinishflag = "3") then %>
      	<font color="blue"><b>진행중</b></font>
      	<% elseif (ofran.FItemList(i).Ffinishflag = "7") then %>
      	완료
        <% else %>
        ???
        <% end if %>
    </td>
  </tr>
<% next %>
</table>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
  <tr bgcolor="#FFFFFF">
    <td align="center">
<% for i = 0 to (ofran.FTotalPage - 1) %>
<% if CStr(page)=CStr(i+1) then %>
	  <a href="?page=<%= (i+1) %>&gubun=<%= gubun %>&onlymifinish=<%=onlymifinish%>&research=<%=research %>">&nbsp;<b><%= (i+1) %></b>&nbsp;</a>
<% else %>
      <a href="?page=<%= (i+1) %>&gubun=<%= gubun %>&onlymifinish=<%=onlymifinish%>&research=<%=research %>">[<%= (i+1) %>]</a>
<% end if %>
<% next %>
    </td>
  </tr>
</table>
<%
set ofran = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
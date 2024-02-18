<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 마일리지
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopmileagecls.asp"-->
<%
dim page,shopid
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
shopid = requestCheckVar(request("shopid"),32)
page = requestCheckVar(request("page"),10)

if page="" then page=1
shopid = "cafe003"

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
yyyy2 = requestCheckVar(request("yyyy2"),4)
mm2 = requestCheckVar(request("mm2"),2)
dd2 = requestCheckVar(request("dd2"),2)

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-3)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

dim ooffmilde
dim i
set ooffmilde = new COffShopMileage
ooffmilde.FPageSize=100
ooffmilde.FCurrpage=page
ooffmilde.FRectStartDay = fromDate
ooffmilde.FRectEndDay = toDate
ooffmilde.FRectShopid=shopid
ooffmilde.COffShopMileageList
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="800" cellspacing="1" class="a" bgcolor=#FFFFFF>
<tr>
	<td align=right>총건수:<%= ooffmilde.FTotalCount%>, 페이지: <%= page %>/<%= ooffmilde.FTotalPage%></td>
</tr>
</table>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td width="100">회원번호</td>
	<td width="100">회원명</td>
	<td width="100">샾구분</td>
	<td width="80">마일리지</td>
	<td width="100">적요</td>
	<td width="80">저장일</td>
</tr>
<% for i=0 to ooffmilde.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffmilde.FItemList(i).Fpointuserno %></td>
	<td ><%= ooffmilde.FItemList(i).Fpointusername %></td>
	<td><%= ooffmilde.FItemList(i).Fshopid %></td>
	<td align="right"><%= ooffmilde.FItemList(i).Fpoint %></td>
	<td align="let"><%= ooffmilde.FItemList(i).Fjukyo %></td>
	<td><%= ooffmilde.FItemList(i).Fregdate %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height=20>
	<td colspan=6 align=center>
	<% if ooffmilde.HasPreScroll then %>
		<a href="?page=<%= ooffmilde.StarScrollPage-1 %>&menupos=<%= menupos %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ooffmilde.StarScrollPage to ooffmilde.FScrollCount + ooffmilde.StarScrollPage - 1 %>
		<% if i>ooffmilde.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ooffmilde.HasNextScroll then %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set ooffmilde = Nothing
%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
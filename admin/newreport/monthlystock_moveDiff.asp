<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1, tmpDate


page       		= requestCheckvar(request("page"),10)
research		= requestCheckvar(request("research"),10)
yyyy1       	= requestCheckvar(request("yyyy1"),10)
mm1         	= requestCheckvar(request("mm1"),10)

if (page="") then page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", -1, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
end if


'// ============================================================================
dim ojaego
set ojaego = new CMonthlyStock

ojaego.FPageSize = 200
ojaego.FCurrPage = page
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1

ojaego.GetMonthlyMoveDiffList

%>

<script language='javascript'>

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			&nbsp;
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 월 이동내역
		</td>

		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p>

* 최대 200개까지 표시됩니다.

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60">이동월</td>
		<td width="120">브랜드</td>
		<td width="30">구분</td>
		<td width=70>상품코드</td>
		<td width=40>옵션</td>
		<td width=40>수량</td>
		<td>비고</td>
	</tr>
<% if ojaego.FResultCount >0 then %>
	<% for i=0 to ojaego.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center><%= ojaego.FItemList(i).Fyyyymm %></td>
		<td align=center><%= ojaego.FItemList(i).Flastmakerid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemgubun %></td>
		<td align="right"><%= ojaego.FItemList(i).Fitemid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemoption %></td>
		<td align="right">
			<%= FormatNumber(ojaego.FItemList(i).FtotItemNo, 0) %>
		</td>
		<td>
	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height=50>
		<td align=center colspan="17">내역이 없습니다.</td>
	</tr>
<% end if %>
</table>

<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

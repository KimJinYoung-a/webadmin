<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/event/eventCls.asp"-->
<%
Dim i, page, searchKey, searchString
Dim oItem, makerid, research
research		= RequestCheckvar(request("research"),10)
page			= RequestCheckvar(request("page"),10)
searchKey		= requestcheckvar(Request("searchKey"),10)
searchString	= RequestCheckvar(Request("searchString"),128)
makerid			= requestcheckvar(Request("makerid"),32)

If page = "" Then page = 1

Set oItem = new CEvent
	oItem.FCurrPage			= page
	oItem.FPageSize			= 12
	oItem.FRectSearchKey	= searchKey
	oItem.FRectSearchString	= searchString
	oItem.FRectMakerid		= makerid
	oItem.getArtList
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function openerRegLecIdx(v){
	opener.$("#diycode").val(v);
	window.close();
}
</script>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="" method="POST">
<input type="hidden" name="page">
<tr height="80" bgcolor="FFFFFF">
	<td>
		<font size="4"><strong>판매중인 작품</strong></font>&nbsp;&nbsp;
		총 <%= oItem.FTotalCount %>건
	</td>
	<td align="right">
		<select name="searchKey" class="select">
			<option value="itemname" <%= chkiif(searchKey = "itemname", "selected", "") %>>작품명</option>
			<option value="itemid" <%= chkiif(searchKey = "itemid", "selected", "") %>>작품코드</option>
		</select>
		<input type="text" class="text" name="searchString" value="<%=searchString%>">
		<input type="button" class="button" value="검색" onclick="document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">작품 코드</td>
	<td>작품명</td>
	<td width="300">판매가격</td>
	<td width="100">등록일</td>
</tr>
<% For i = 0 to oItem.FResultCount - 1 %>
<tr height="30" bgcolor="FFFFFF" align="center" style="cursor:pointer;" onclick="openerRegLecIdx('<%= oItem.FItemList(i).FItemid %>')" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';>
	<td><%= oItem.FItemList(i).FItemid %></td>
	<td><%= oItem.FItemList(i).FItemname %></td>
	<td>
	<% if oItem.FItemList(i).IsSaleItem or oItem.FItemList(i).isCouponItem Then %>
		<% IF oItem.FItemList(i).IsSaleItem then %>
			<%= FormatNumber(oItem.FItemList(i).getRealPrice,0) %> <font color="red">[<%= oItem.FItemList(i).getSalePro %>]</font></span>
		<% End IF %>
		<% IF oItem.FItemList(i).IsCouponItem then %>
			<%= FormatNumber(oItem.FItemList(i).GetCouponAssignPrice,0) %> <font color="green">[<%= oItem.FItemList(i).GetCouponDiscountStr %>]<font color="green">
		<% End IF %>
	<% Else %>
		<% = FormatNumber(oItem.FItemList(i).getRealPrice,0) %></p>
	<% End if %>
	</td>
	<td><%= FormatDate(oItem.FItemList(i).FRegdate, "0000.00.00") %></td>	
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oItem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oItem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oItem.StartScrollPage to oItem.FScrollCount + oItem.StartScrollPage - 1 %>
    		<% if i>oItem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oItem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</body>
</html>
<% Set oItem = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
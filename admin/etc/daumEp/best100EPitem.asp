<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/daumEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim page, i, best100, catecode
page				= request("page")
catecode			= requestCheckvar(request("catecode"),3)

If page = "" Then page = 1
	
Set best100 = new epShop
	best100.FCurrPage		= page
	best100.FRectCateCode	= catecode
	best100.FPageSize		= 100
	best100.Best100EpItemList
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/daumEp/inc_daumHead.asp" -->
>> 베스트100EP리스트
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		카테고리: <%= fnDepth1CateSelectBox("catecode", catecode, "") %>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(best100.FTotalPage,0) %> 총건수: <%= FormatNumber(best100.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>순위</td>
	<td>카테고리</td>
	<td>이미지</td>
    <td>상품코드</td>
    <td>상품명</td>
    <td>브랜드ID</td>
    <td>품절여부</td>
	<td>상품등록일</td>
	<td>상품최종수정일</td>
	<td>판매가</td>
	<td>마진</td>
</tr>
<% For i=0 to best100.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20" align="center">
	<td><%= best100.FItemList(i).FRowNum %></td>
	<td><%= best100.FItemList(i).FCate1code %> / <%= best100.FItemList(i).FCatename %></td>
	<td><img src="<%= best100.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= best100.FItemList(i).FItemid %></td>
    <td><%= best100.FItemList(i).FItemname %></td>
    <td><%= best100.FItemList(i).FMakerid %></td>
    <td>
        <% if best100.FItemList(i).IsSoldOut then %>
            <% if best100.FItemList(i).FSellyn="N" then %>
            <font color="red">품절</font>
            <% else %>
            <font color="red">일시<br>품절</font>
            <% end if %>
        <% end if %>
    </td>
	<td><%= best100.FItemList(i).FRegdate %></td>
	<td><%= best100.FItemList(i).FLastupdate %></td>
	<td>
        <%= FormatNumber(best100.FItemList(i).FSellcash,0) %>
	</td>
	<td>
        <% if best100.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-best100.FItemList(i).Fbuycash/best100.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if best100.HasPreScroll then %>
		<a href="javascript:goPage('<%= best100.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + best100.StartScrollPage to best100.FScrollCount + best100.StartScrollPage - 1 %>
    		<% if i>best100.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if best100.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
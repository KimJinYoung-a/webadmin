<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/sitemaster/cardsale/cardsaleCls.asp"-->
<%
Dim page, i, oCardSale, research
Dim isusing
research	= request("research")
page    	= request("page")
isusing		= request("isusing")
If page = "" Then page = 1

If (research = "") Then
	isusing = "Y"
End If

Set oCardSale = new CCardSale
	oCardSale.FCurrPage		= page
	oCardSale.FPageSize		= 50
	oCardSale.FRectIsusing	= isusing
	oCardSale.getCardSaleItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function popManageCard(v){
	var pCM = window.open("/admin/sitemaster/cardsale/popManageCard.asp?idx="+v,"popManageCard","width=600,height=400,scrollbars=yes,resizable=yes");
	pCM.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		사용여부 :
		<select name="isusing" class="select">
			<option value="">-선택-</option>
			<option value="Y" <%= Chkiif(isusing="Y", "selected", "") %> >Y</option>
			<option value="N" <%= Chkiif(isusing="N", "selected", "") %> >N</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<input type="button" value="등록" class="button" onclick="popManageCard('');">
<br />
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%= FormatNumber(oCardSale.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCardSale.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30">idx</td>
	<td width="200">기간</td>
	<td width="140">사용 혜택</td>
	<td width="140">최소<br />구매 금액</td>
	<td width="140">최대<br />할인 금액</td>
	<td width="150">카드명</td>
	<td width="100">등록자</td>
	<td width="100">등록일</td>
	<td width="80">사용여부</td>
</tr>
<% For i=0 to oCardSale.FResultCount - 1 %>
<tr align="center" bgcolor= '#FFFFFF'" onclick="popManageCard('<%= oCardSale.FItemList(i).FIdx %>');" style="cursor:pointer;" onmouseover="this.style.background='orange'"; onmouseout="this.style.background='white';">
	<td align="center"><%= oCardSale.FItemList(i).FIdx %></td>
	<td align="center"><%= LEFT(oCardSale.FItemList(i).FStartdate,10) %> ~ <%= LEFT(oCardSale.FItemList(i).FEnddate, 10) %></td>
	<td align="center">
	<%
		Select Case oCardSale.FItemList(i).FSaleType
			Case "1"	response.write oCardSale.FItemList(i).FSalePrice &"won 할인"
			Case "2"	response.write oCardSale.FItemList(i).FSalePrice &"% 할인"
		End Select
	%>
	</td>
	<td align="center"><%= oCardSale.FItemList(i).FMinPrice %></td>
	<td align="center"><%= oCardSale.FItemList(i).FMaxPrice %></td>
	<td align="center"><%= oCardSale.FItemList(i).FCardName %></td>
	<td align="center"><%= oCardSale.FItemList(i).FReguserid %></td>
	<td align="center"><%= LEFT(oCardSale.FItemList(i).FRegdate, 10) %></td>
	<td align="center"><%= oCardSale.FItemList(i).FIsUsing %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="9" align="center" bgcolor="#FFFFFF">
        <% if oCardSale.HasPreScroll then %>
		<a href="javascript:goPage('<%= oCardSale.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oCardSale.StartScrollPage to oCardSale.FScrollCount + oCardSale.StartScrollPage - 1 %>
    		<% if i>oCardSale.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oCardSale.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<% SET oCardSale = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
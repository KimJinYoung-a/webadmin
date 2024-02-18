<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kaffa/itemsalecls.asp"-->
<%
Dim i
Dim regyn : regyn= requestCheckvar(request("regyn"),10)
Dim page  : page= requestCheckvar(request("page"),10)
Dim clsSale

if (page="") then page=1

Set clsSale = new CSale
clsSale.FCurrPage	= page
clsSale.FPageSize = 20
clsSale.FRectTenCodePreReg = regyn
clsSale.getTenSaleListWithKaffa
%>

<script language='javascript'>
function regByTenSale(tensalecode){
    if (confirm('TEN 세일코드'+tensalecode+'로 등록하시겠습니까?\n\n동일기간에 등록된 상품은 제외되며, 중국사이트 연동상품만 등록됩니다.')){
        document.frmSubmit.tensalecode.value=tensalecode;
        document.frmSubmit.mode.value="T";
        document.frmSubmit.submit();
    }
}
function goPage(p){
    document.frmSearch.page.value = p;
    document.frmSearch.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frmSearch" method="get"  >
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
		<td align="left">
		등록여부 :
		<select name="regyn">
            <option value="">전체
            <option value="Y" <%=CHKIIF(regyn="Y","selected","") %> >기등록
            <option value="N" <%=CHKIIF(regyn="N","selected","") %> >미등록
		</select>
		</td>
		<td  width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
		</td>
	</tr>
	</form>
</table>
<!---- /검색 ---->
<p>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="12">검색결과 : <b><%= FormatNumber(clsSale.FTotalCount,0) %></b>&nbsp;&nbsp;페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(clsSale.FTotalPage,0) %></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>TEN할인코드</td>
    	<td>TEN할인명</td>
    	<td>TEN할인율</td>
    	<td>TEN매입가구분</td>
    	<td>TEN시작일</td>
    	<td>TEN종료일</td>
    	<td>TEN상태</td>
    	<td>TEN등록일</td>
    	<td>가능수량</td>
    	<td>처리</td>
    </tr>
    <% For i = 0 To clsSale.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%=clsSale.FItemList(i).FTENsale_code %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_name %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_rate %>%</td>
    	<td><%=clsSale.FItemList(i).getTenSaleMarginGubun %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_startdate %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_enddate %></td>
    	<td><%=clsSale.FItemList(i).getTenSaleStateName %></td>
    	<td><%=clsSale.FItemList(i).FTENregdate %></td>
    	<td><%=clsSale.FItemList(i).FvalidCnt %></td>
    	<td>
    	<% if isNULL(clsSale.FItemList(i).FDiscountKey)  then %>
    	    <input type="button" value="등록" onClick="regByTenSale('<%=clsSale.FItemList(i).FTENsale_code %>')">
    	<% else %>
    	    <%=clsSale.FItemList(i).FDiscountKey%>
    	<% end if %>
    	</td>
    </tr>
    <% next %>
    <tr height="20">
		<td colspan="11" align="center" bgcolor="#FFFFFF">
		<% If clsSale.HasPreScroll Then %>
			<a href="javascript:goPage('<%= clsSale.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i=0 + clsSale.StartScrollPage To clsSale.FScrollCount + clsSale.StartScrollPage - 1 %>
			<% If i>clsSale.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If clsSale.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
		</td>
	</tr>
</table>

<%
Set clsSale = Nothing
%>
<form name="frmSubmit" method="post" action="saleitemProc.asp">
<input type="hidden" name="mode" value="T">
<input type="hidden" name="tensalecode" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
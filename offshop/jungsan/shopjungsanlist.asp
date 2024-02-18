<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 정산리스트
' History : 2009.04.07 서동석 생성
'			2012.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim page, shopid, divcode ,i,totalsum, totalsuply, totalerr
	page = request("page")
	divcode = request("divcode")

shopid = session("ssBctID")
	
divcode = "WS"

if page="" then page=1

dim ofranchulgojungsan
set ofranchulgojungsan = new CFranjungsan
	ofranchulgojungsan.FPageSize=50
	ofranchulgojungsan.FCurrpage = page
	ofranchulgojungsan.FRectshopid = shopid
	''ofranchulgojungsan.FRectdivcode = divcode
	ofranchulgojungsan.FRectStateUpcheView = "on"
	ofranchulgojungsan.getFranJungsanList

%>

<script language='javascript'>

function popMaster(iid){
	var popwin = window.open('popmeaipchulgo.asp?idx=' + iid,'franmeaipedit','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popSubmaster(iid){
	var popwin = window.open('franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

//공급받는자용
function popTaxPrint(taxNo, bizNo){
	var s_biz_no = "2118700620";	// 텐바이텐 사업자번호

	//	리얼서버	http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp
	<% if (application("Svr_Info")="Dev") then %>
	var popwinsub = window.open("http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+taxNo+"&cur_biz_no="+bizNo+"&s_biz_no="+s_biz_no+"&b_biz_no="+bizNo,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	<% else %>
	var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+taxNo+"&cur_biz_no="+bizNo+"&s_biz_no="+s_biz_no+"&b_biz_no="+bizNo,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
    <% end if %>
	popwinsub.focus();
}

function goView_Bill36524(tax_no, b_biz_no)
{
		window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><%=CTX_SEARCH%><br><%= CTX_conditional %></td>
	<td align="left">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="<%=CTX_SEARCH%>" onClick="document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" cellspacing="1" cellpadding=3 class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<%= CTX_search_result %> : <b><%= ofranchulgojungsan.FTotalCount %></b>
		&nbsp;
		<%= CTX_page %> : <b><%= page %> / <%= ofranchulgojungsan.FTotalpage %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td><%= CTX_number %></td>
	<td><%= CTX_SHOP %></td>
	<td><%= CTX_divide %></td>
	<td><%= CTX_title %></td>
	<td><%= CTX_Issue %></td>
	<td><%= CTX_real %>&nbsp;<%= CTX_Supply_price %></td>
	<td><%= CTX_tax_Bill_date %></td>
	<td><%= CTX_Deposit_Date %></td>
	<td><%= CTX_Status %></td>
	<td><%= CTX_DETAILVIEW %></td>
	<td><%= CTX_Bill %></td>
</tr>
<% if ofranchulgojungsan.FResultCount >0 then %>
<% for i=0 to ofranchulgojungsan.FResultCount-1 %>
<%
totalsum = totalsum + ofranchulgojungsan.FItemList(i).Ftotalsum
totalsuply  = totalsuply + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
totalerr = totalerr  + ofranchulgojungsan.FItemList(i).Ftotalsum -  ofranchulgojungsan.FItemList(i).Ftotalsuplycash
%>
<tr bgcolor="#FFFFFF" align="center">
	<td width=40><%= ofranchulgojungsan.FItemList(i).Fidx %></td>
	<td width=90>
		<a href="javascript:popMaster('<%= ofranchulgojungsan.FItemList(i).Fidx %>');">
		<%= ofranchulgojungsan.FItemList(i).Fshopid %></a>
	</td>
	<td width=60>
		<font color="<%= ofranchulgojungsan.FItemList(i).GetDivCodeColor %>">
		<%= ofranchulgojungsan.FItemList(i).GetDivCodeName %></font>
	</td>
	<td width=210>
		<a href="javascript:popSubmaster('<%= ofranchulgojungsan.FItemList(i).Fidx %>');">
		<%= ofranchulgojungsan.FItemList(i).Ftitle %></a>
	</td>
	<td align="right" width=76><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotalsum,0) %></td>
	<td align="right" width=76><%= formatNumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
	<td width=70><%= ofranchulgojungsan.FItemList(i).Ftaxdate %></td>
	<td width=70><%= ofranchulgojungsan.FItemList(i).Fipkumdate %></td>
	<td width=80>
		<font color="<%= ofranchulgojungsan.FItemList(i).GetStateColor %>">
		<%= ofranchulgojungsan.FItemList(i).GetStateName %></font>
	</td>
	<td width=60><a href="javascript:popSubmaster('<%= ofranchulgojungsan.FItemList(i).Fidx %>');"><%= CTX_DETAILVIEW %></a></td>
    <td width=60>
	    <% if Not IsNULL(ofranchulgojungsan.FItemList(i).FtaxNo) then %>
	    	<% if (LEft(ofranchulgojungsan.FItemList(i).FtaxNo,2)="TX") then %>
	    		<img src="/images/icon_print02.gif" width="16" onClick="goView_Bill36524('<%=ofranchulgojungsan.FItemList(i).FtaxNo%>','<%=ofranchulgojungsan.FItemList(i).FbizNo%>');" style="cursor:pointer">
	    	<% else %>
	   			<img src="/images/icon_print02.gif" width="16" onClick="popTaxPrint('<%=ofranchulgojungsan.FItemList(i).FtaxNo%>','<%=ofranchulgojungsan.FItemList(i).FbizNo%>');" style="cursor:pointer">
	    	<% end if %>
	    <% else %>
	    .
	    <% end if %>
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="#FFFFFF" align="center">
	<td><%= CTX_total %></td>
	<td></td>
	<td></td>
	<td></td>
	<td align="right"><%= formatNumber(totalsum,0) %></td>
	<td align="right"><%= formatNumber(totalsuply,0) %></td>
	<td align="right"><%= CTX_error %> : <%= formatNumber(totalerr,0) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<tr height="25" bgcolor="#FFFFFF" height=20>
	<td colspan=11 align=center>
	<% if ofranchulgojungsan.HasPreScroll then %>
		<a href="?page=<%= ofranchulgojungsan.StarScrollPage-1 %>&shopid=<%= shopid %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ofranchulgojungsan.StarScrollPage to ofranchulgojungsan.FScrollCount + ofranchulgojungsan.StarScrollPage - 1 %>
		<% if i>ofranchulgojungsan.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&shopid=<%= shopid %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ofranchulgojungsan.HasNextScroll then %>
		<a href="?page=<%= i %>&shopid=<%= shopid %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" >
	<td colspan=11 align=center>[<%= CTX_search_returns_no_results %>]</td>
</tr>
</table>
<% end if %>

<%
set ofranchulgojungsan = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/companyjungsancls.asp"-->

<%

dim page
dim ijungsan
dim yyyy1,mm1

page = request("page")
if (page="") then page=1
yyyy1 = request("yyyy1")
mm1 = request("mm1")

if yyyy1="" then
	yyyy1=Left(CStr(now()),4)
end if

if mm1="" then
	mm1=Mid(CStr(now()),6,2)
end if



set ijungsan = new CUpcheJungSan

ijungsan.FcurrPage = page
ijungsan.FPageSize=40
ijungsan.getDefaultInfo session("ssBctID")


if session("ssBctDiv")="999" then
	ijungsan.FRectRdSite = session("ssBctID")
else
	ijungsan.FRectSiteName = session("ssBctID")
end if

ijungsan.FRectYYYY = yyyy1
ijungsan.FRectMM = mm1
ijungsan.PartnerMiJungSanDeasangList

dim ix
dim bufsum, deasangsum, amountsum
bufsum =0
deasangsum =0
amountsum =0
%>
<script language='javascript'>

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="/company/viewordermaster.asp"
	frm.submit();

}
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
<table width="760" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawYMBox yyyy1,mm1 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="760" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="13">
	<table border="0" cellspacing="0" cellpadding="0" class="a">
	<tr>
		<td>* 커미션 : </td>
		<td><Font color="#3333FF"><%= CDbl(ijungsan.FCommission)*100 %> %</font></td>
	</tr>
	<tr>
		<td>* 총 건수 : </td>
		<td><Font color="#3333FF"><%= FormatNumber(ijungsan.FTotalCount,0) %></font></td>
	</tr>
	<tr>
		<td>* 정산대상 금액 : </td>
		<td > <% = FormatNumber((ijungsan.FTotalJungsan - ijungsan.FTotalBaesong),0)  %></td>
	</tr>
	<tr>
		<td>* 정산예정 금액 : </td>
		<td > <% =FormatNumber((ijungsan.FTotalJungsan - ijungsan.FTotalBaesong) * CDbl(ijungsan.FCommission),0) %></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td colspan="13" align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
</tr>
<tr >
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">구매자</td>
	<td width="72" align="center">결제금액</td>
	<td width="72" align="center">포장.배송료</td>
	<td width="90" align="center">정산대상금액</td>
	<td width="90" align="center">정산금액</td>
	<td width="100" align="center">주문일</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr>
	<td colspan="8" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ijungsan.FresultCount-1 %>
	<form name="frmOnerder_<%= ijungsan.FJungSanList(ix).FOrderSerial %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ijungsan.FJungSanList(ix).FOrderSerial %>">
	<input type="hidden" name="userid" value="<%= ijungsan.FJungSanList(ix).FUserID %>">
	<input type="hidden" name="buyname" value="<%= ijungsan.FJungSanList(ix).FBuyName %>">
	<input type="hidden" name="totalsum" value="<%= ijungsan.FJungSanList(ix).FSubTotalPrice %>">
	<input type="hidden" name="beasongpay" value="<%= ijungsan.FJungSanList(ix).FBeasongPay %>">
	<input type="hidden" name="deasangsum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay %>">
	<input type="hidden" name="jungsansum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay * CDbl(ijungsan.FCommission) %>">
	<tr class="a">
		<td align="center"><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ijungsan.FJungSanList(ix).FOrderSerial %>)" class="zzz"><%= ijungsan.FJungSanList(ix).FOrderSerial %></a></td>
		<% if ijungsan.FJungSanList(ix).FUserID<>"" then %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FUserID %></td>
		<% else %>
		<td align="center">&nbsp;</td>
		<% end if %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FBuyName %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="right"><%= FormatNumber(ijungsan.FJungSanList(ix).FDeasangPay,0) %></td>
		<%
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
		 %>
		<td align="right"><%= FormatNumber(bufsum* CDbl(ijungsan.FCommission),0) %></td>
		<td align="center"><%= Left(ijungsan.FJungSanList(ix).FRegDate,10) %></td>
	</tr>
	</form>
	<% next %>
	<tr>
		<td colspan="8" height="30" align="center">
		<% if ijungsan.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ijungsan.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + ijungsan.StarScrollPage to ijungsan.FScrollCount + ijungsan.StarScrollPage - 1 %>
			<% if ix>ijungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ijungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% end if %>
</table>
<form name="frmArrupdate" method="post" action="jungsanmaker.asp">
<input type="hidden" name="commission" value="<%= ijungsan.FCommission %>">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="userid" value="">
<input type="hidden" name="buyname" value="">
<input type="hidden" name="totalsum" value="">
<input type="hidden" name="deasangsum" value="">
<input type="hidden" name="beasongpay" value="">
<input type="hidden" name="jungsansum" value="">
</form>
<%
set ijungsan = nothing
%>

<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
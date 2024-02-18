<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/companyjungsancls.asp"-->

<%

dim page
dim ijungsan
Dim masterid
dim junsandate

masterid = request("masterid")
junsandate = request("junsandate")

page = request("page")
if (page="") then page=1

set ijungsan = new CUpcheJungSan

ijungsan.FcurrPage = page
ijungsan.FPageSize=20
ijungsan.getOldDefaultInfo masterid

ijungsan.FMasterid = masterid
ijungsan.FrectSiteName = session("ssBctID")
ijungsan.PartnerOldDetailJungSanDeasangList

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
		<td ><% = FormatNumber(ijungsan.FTotalJungsan,0)  %></td>
	</tr>
	<tr>
		<td>* 정산예정 금액 : </td>
		<td ><% = FormatNumber(ijungsan.FTotalJungsansum,0)  %></td>
	</tr>
	<tr>
		<td>* 기타사항 : </td>
		<td ><% = ijungsan.FEtcStr  %></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td colspan="7" align="right">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" class='a'>
		<tr>
			<td align="left"><a href="excel_jungsan_change.asp?masterid=<% = masterid %>&junsandate=<% = junsandate %>"><img src="/images/btn_excel.gif" border="0" width="75"></a></td>
			<td align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
		</tr>
		</table>
	</td>
</tr>
<tr >
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">구매자</td>
	<td width="72" align="center">결제금액</td>
	<td width="72" align="center">포장.배송료</td>
	<td width="90" align="center">정산대상금액</td>
	<td width="90" align="center">정산금액</td>
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
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FDeasangPay,0) %></td>
		<%
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
		 %>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).Fjungsansum,0) %></td>
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
<script language='javascript'>
	//deasangsum.innerText = '<%= FormatNumber(deasangsum,0) %>';
	//amountsum.innerText = '<%= FormatNumber(amountsum,0) %>';
</script>

<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="masterid" value="<% =masterid %>">
</form>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

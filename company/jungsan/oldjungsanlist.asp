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

page = request("page")
if (page="") then page=1

set ijungsan = new CUpcheJungSan

ijungsan.FcurrPage = page
ijungsan.FPageSize=20
ijungsan.getDefaultInfo session("ssBctID")

ijungsan.FrectSiteName = session("ssBctID")
ijungsan.PartnerOldJungSanDeasangList 

dim ix
dim bufsum, deasangsum, amountsum
bufsum =0
deasangsum =0
amountsum =0
%>
<script language='javascript'>
function ViewOrderDetail(frm){
    frm.target = 'orderdetail';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	</form>	

<table width="760" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="10">
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
		<td ><% = FormatNumber((ijungsan.FTotalJungsan * CDbl(ijungsan.FCommission)),0)  %></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td colspan="10" align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
</tr>
<tr >
	<td width="30" align="center">No</td>
	<td width="150" align="center">정산기간</td>
	<td width="40" align="center">건수</td>
	<td width="90" align="center">총금액</td>
	<td width="80" align="center">배송비</td>
	<td width="90" align="center">정산대상금액</td>
	<td width="90" align="center">정산금액</td>
	<td width="50" align="center">입금상태</td>
	<td align="center">입금일</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr>
	<td colspan="10" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ijungsan.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ijungsan.FJungSanList(ix).FOrderSerial %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ijungsan.FJungSanList(ix).FOrderSerial %>">
	<input type="hidden" name="userid" value="<%= ijungsan.FJungSanList(ix).FUserID %>">
	<input type="hidden" name="buyname" value="<%= ijungsan.FJungSanList(ix).FBuyName %>">
	<input type="hidden" name="totalsum" value="<%= ijungsan.FJungSanList(ix).FSubTotalPrice %>">
	<input type="hidden" name="beasongpay" value="<%= ijungsan.FJungSanList(ix).FBeasongPay %>">
	<input type="hidden" name="deasangsum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay %>">
	<input type="hidden" name="jungsansum" value="<%= ijungsan.FJungSanList(ix).FDeasangPay * CDbl(ijungsan.FCommission) %>">
	<tr class="a">
		<td align="center"><% =ix + 1  %></td>
		<td align="center"><a href="oldjungsanlist_detail_view.asp?masterid=<% =ijungsan.FJungSanList(ix).Fmasterid %>&junsandate=<% = ijungsan.FJungSanList(ix).FTotaldate %>"><%= ijungsan.FJungSanList(ix).FTotaldate %> ~ <%= ijungsan.FJungSanList(ix).FTotaldate2 %></a></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FTotalno %></td>
		<td align="center"><%= Formatnumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FTotaldeasang,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FTotaljungsansum,0) %></td>
		<% 
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
		 %>
		<td align="center">
		<% if ijungsan.FJungSanList(ix).FIpkumDiv = 0 then %>
		  <font color="blue">No</font>
		 <% else %>
          <font color="red">Yes</font>
		 <% end if %>
		</td>
		<td align="center">
		  <% if ijungsan.FJungSanList(ix).FIpkumDate<>"" then %>
		  <% =FormatDateTime(ijungsan.FJungSanList(ix).FIpkumDate,2)  %>
		 <% else %>
          &nbsp;
		 <% end if %>
		</td>
	</tr>
	</form>
	<% next %>
	<tr>
		<td colspan="13" height="30" align="center">
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
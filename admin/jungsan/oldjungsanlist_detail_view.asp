<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsancls.asp"-->

<%

dim page
dim ijungsan
Dim masterid,extsitename

extsitename = request("extsitename")

masterid = request("masterid")

page = request("page")
if (page="") then page=1

set ijungsan = new CUpcheJungSan

ijungsan.FcurrPage = page
ijungsan.FPageSize=3000
ijungsan.getOldDefaultInfo masterid

ijungsan.FMasterid = masterid
ijungsan.FrectSiteName = extsitename
ijungsan.PartnerOldDetailJungSanDeasangList

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

<table width="760" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="13">
	<table border="0" cellspacing="0" cellpadding="0" class="a">
	<tr>
		<td>* Ŀ�̼� : </td>
		<td><Font color="#3333FF"><%= CDbl(ijungsan.FCommission)*100 %> %</font></td>
	</tr>
	<tr>
		<td>* �� �Ǽ� : </td>
		<td><Font color="#3333FF"><%= FormatNumber(ijungsan.FTotalCount,0) %></font></td>
	</tr>
	<tr>
		<td>* ������ �ݾ� : </td>
		<td ><% = FormatNumber(ijungsan.FTotalJungsan,0)  %></td>
	</tr>
	<tr>
		<td>* ���꿹�� �ݾ� : </td>
		<td ><% = FormatNumber(ijungsan.FTotalJungsansum,0)  %></td>
	</tr>
	<tr>
		<td>* ��Ÿ���� : </td>
		<td ><% = ijungsan.FEtcStr  %></td>
	</tr>

	</table>
	</td>
</tr>
<tr>
	<td colspan="13" align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
</tr>
<tr >
	<td width="100" align="center">�ֹ���ȣ</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">������</td>
	<td width="72" align="center">�����ݾ�</td>
	<td width="72" align="center">����.��۷�</td>
	<td width="90" align="center">������ݾ�</td>
	<td width="90" align="center">����ݾ�</td>
	<td width="100" align="center">��ü�ֹ���ȣ</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr>
	<td colspan="8" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
	<% for ix=0 to ijungsan.FresultCount-1 %>
	<tr class="a">
		<td align="center"><%= ijungsan.FJungSanList(ix).FOrderSerial %></td>
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
		<td align="center"><%= ijungsan.FJungSanList(ix).Fauthcode & ijungsan.FJungSanList(ix).Fpaygatetid %></td>
	</tr>
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
<table cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" border="1" align="center" width="400">
<tr>
	<td class="a"><% =	ijungsan.Fetcstr %></td>
</tr>
</table>
<%
set ijungsan = nothing
%>

<form name="frm" method="get" >
<input type="hidden" name="extsitename" value="<%= extsitename %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="masterid" value="<% =masterid %>">
</form>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsancls.asp"-->

<%

dim page
dim extsitename
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate


extsitename = request("extsitename")

page = request("page")
if (page="") then page=1


nowdate = Left(CStr(now()),10)

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)


dim ijungsan
set ijungsan = new CUpcheJungSan
ijungsan.FCurrpage = page
ijungsan.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ijungsan.FRectRegEnd   = searchnextdate
ijungsan.FrectSiteName = extsitename
ijungsan.PartnerSiteJungSanDeasangList

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

function jungsanopen(num){
  window.open("jungsan_window.asp?masterid=" + num ,"window","status=no,toolbar=no,resizable=no,scrollbars=no, menubar=no,width=300,height=200");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
<table width="860" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr>
		<td class="a" width="500">
		���޻�:
		<% drawSelectBoxPartner "extsitename",extsitename %>
		<br>
		�˻��Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>&nbsp;(������ �Ⱓ ����)
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="860" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="11">
	<table border="0" cellspacing="0" cellpadding="0" class="a">
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
		<td ><% = FormatNumber((ijungsan.FTotalJungsan * CDbl(ijungsan.FCommission)),0)  %></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td colspan="11" align="right">page : <%= ijungsan.FCurrPage %>/<%=ijungsan.FTotalPage %></td>
</tr>
<tr >
	<td width="30" align="center">No</td>
	<td width="100" align="center">����Ʈ��</td>
	<td width="150" align="center">����Ⱓ</td>
	<td width="40" align="center">�Ǽ�</td>
	<td width="90" align="center">�ѱݾ�</td>
	<td width="80" align="center">��ۺ�</td>
	<td width="90" align="center">������ݾ�</td>
	<td width="90" align="center">����ݾ�</td>
	<td width="80" align="center">������</td>
	<td width="50" align="center">�Աݻ���</td>
	<td align="center">�Ա���</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr>
	<td colspan="11" align="center">[�˻������ �����ϴ�.]</td>
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
		<td align="center"><% =ijungsan.FJungSanList(ix).Fsitename %></td>
		<td align="center"><a href="oldjungsanlist_detail_view.asp?masterid=<% =ijungsan.FJungSanList(ix).Fmasterid %>&extsitename=<% =ijungsan.FJungSanList(ix).Fsitename %>"><%= ijungsan.FJungSanList(ix).FTotaldate %> ~ <%= ijungsan.FJungSanList(ix).FTotaldate2 %></a></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FTotalno %></td>
		<td align="center"><%= Formatnumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FTotaldeasang,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FTotaljungsansum,0) %></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FSegumil %></td>
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
		 <% if ijungsan.FJungSanList(ix).FIpkumDiv = 0 then %>
		   <input type="button" value="�����ϱ�" onclick="javascript:jungsanopen(<% =ijungsan.FJungSanList(ix).Fmasterid %>);">
		 <% else %>
           <% =FormatDateTime(ijungsan.FJungSanList(ix).FIpkumDate,2)  %>
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
<input type="hidden" name="extsitename" value="<%= extsitename %>">
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
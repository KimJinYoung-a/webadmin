<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<%
dim ojumun
set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = requestCheckvar(request("orderserial"),20)
ojumun.FRectSiteName=session("ssBctId")

if (ojumun.FRectOrderSerial<>"" and ojumun.FRectSiteName<>"") then
    ojumun.SearchJumunList
end if


if (ojumun.FResultCount<1) then
    dbget.close() : response.end
end if

dim ix
%>
<table border="1" cellspacing="0" cellpadding="0" class="a">
<tr>
  <td bgcolor="#22AAAA" width="100">�ֹ���ȣ</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FOrderSerial %></td>
  <td bgcolor="#22AAAA" width="100">����Ʈ</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FSitename %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">�������</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).JumunMethodName %></td>
  <td bgcolor="#22AAAA" width="100">�ֹ�����</td>
  <td bgcolor="#DDDDDD" width="200"><font color="<%= ojumun.FMasterItemList(0).IpkumDivColor %>"><%= ojumun.FMasterItemList(0).IpkumDivName %></font></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">������</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FDiscountRate %></td>
  <td bgcolor="#22AAAA" width="100">��ҿ���</td>
  <td bgcolor="#DDDDDD" width="200"><font color="<%= ojumun.FMasterItemList(0).CancelYnColor %>"><%= ojumun.FMasterItemList(0).CancelYnName %></font></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">�����ݾ�</td>
  <td bgcolor="#DDDDDD" width="200"><%= FormatNumber(ojumun.FMasterItemList(0).FSubTotalPrice,0) %></td>
  <td bgcolor="#22AAAA" width="100">�ֹ��ݾ�</td>
  <td bgcolor="#DDDDDD" width="200"><%= FormatNumber(ojumun.FMasterItemList(0).FTotalSum,0) %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">�ֹ���</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FRegDate %></td>
  <td bgcolor="#22AAAA" width="100">�Ա���</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FIpkumDate %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">������ID</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FUserID %></td>
  <td bgcolor="#22AAAA" width="100">������</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyName %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">��������ȭ</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyPhone %></td>
  <td bgcolor="#22AAAA" width="100">�������ڵ���</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FBuyHp %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">�������̸���</td>
  <td bgcolor="#DDDDDD" width="200"><a href="mailto:<%= ojumun.FMasterItemList(0).FBuyEmail %>" class="zzz"><%= ojumun.FMasterItemList(0).FBuyEmail %></a></td>
  <td bgcolor="#22AAAA" width="100">�Ա���</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FAccountName %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">������</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqName %></td>
  <td bgcolor="#22AAAA" width="100"></td>
  <td bgcolor="#DDDDDD" width="200"></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">��������ȭ</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqPhone %></td>
  <td bgcolor="#22AAAA" width="100">�������ڵ���</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FReqHp %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">�������ּ�</td>
  <td bgcolor="#DDDDDD" colspan="3">
  <input type="text"  value="<%= ojumun.FMasterItemList(0).FReqZipCode %>" size="7">
  <br>
  <input type="text" name="txzip1" value="<%= ojumun.FMasterItemList(0).FReqZipAddr %>" size="12">
  &nbsp;<input type="text" name="txzip1" value="<%= ojumun.FMasterItemList(0).FReqAddress %>" size="36">
  </td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">��Ÿ����</td>
  <td bgcolor="#DDDDDD" colspan="3">
  <%= ojumun.FMasterItemList(0).FComment %>
  </td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">��븶�ϸ���</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FMileTotalPrice %></td>
  <td bgcolor="#22AAAA" width="100">�����ȣ</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FDeliverno %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">ī����ι�ȣ</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FAuthcode %></td>
  <td bgcolor="#22AAAA" width="100">ī����</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FResultmsg %></td>
</tr>
<tr>
  <td bgcolor="#22AAAA" width="100">Inicis-ID</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).FPaygatetID %></td>
  <td bgcolor="#22AAAA" width="100">��������</td>
  <td bgcolor="#DDDDDD" width="200"><%= ojumun.FMasterItemList(0).Fjungsanflag %></td>
</tr>
</table>
<%
ojumun.SearchJumunDetail request("orderserial")
%>

<table border="1" cellspacing="0" cellpadding="0" class="a">
<tr>
	<td width="100">��ۿɼ�</td>
	<td width="200"><%= ojumun.FJumunDetail.BeasongOptionStr %></td>
</tr>
<tr>
	<td>��ۺ�</td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.BeasongPay,0) %></td>
</tr>
</table>

<table border="1" cellspacing="0" cellpadding="0" class="a">
<tr>
	<td width="50" align="center">��ǰID</td>
	<td width="50" align="center">�̹���</td>
	<td width="100" align="center">��ǰ��</td>
	<td width="50" align="center">����</td>
	<td width="70" align="center">�ɼ�Code</td>
	<td width="100" align="center">�ɼǸ�</td>
	<td width="70" align="center">Price</td>
	
</tr>
<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>
<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
<tr>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %></td>
	<td align="center"><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %>" target="_blank"><img src="<%= ojumun.FJumunDetail.FJumunDetailList(ix).FImageSmall %>" border="0"></a></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemOption %></td>
	<td align="center"><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %></td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %></td>
</tr>
<% end if %>
<% next %>
</table>
<% 
set ojumun = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
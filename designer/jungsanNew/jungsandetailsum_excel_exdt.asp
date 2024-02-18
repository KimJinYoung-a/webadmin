<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesignerNoCache.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id,gubun, itemvatyn, makerid, groupid
id      = requestCheckVar(request("id"),10)
gubun   = requestCheckVar(request("gubun"),20)
itemvatyn = requestCheckVar(request("itemvatyn"),10)

makerid = session("ssBctId")
groupid = getPartnerId2GroupID(makerid)

if (NOT chkAvailViewJungsanON(id,makerid,groupid)) then
    response.write "��ȸ ������ �����ϴ�"
    dbget.close()	:	response.End
end if

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectgubun = gubun
'ojungsan.FRectDesigner = makerid
'if (makerid<>"") then
'    ojungsan.JungsanMasterList
'end if
'ojungsan.FRectGroupID = groupid
'if (groupid<>"") then
'    ojungsan.JungsanMasterList
'end if
ojungsan.JungsanMasterList

if ojungsan.FresultCount <1 then
	dbget.close()	:	response.End
end if


dim ojungsanSubsmr
set ojungsanSubsmr = new CUpcheJungsan
ojungsanSubsmr.FRectId = id
ojungsanSubsmr.FRectdesigner = session("ssBctID")
ojungsanSubsmr.getJungsanSubSummary

Dim IsCommissionTax : IsCommissionTax=ojungsan.FItemList(0).IsCommissionTax
Dim IsCommissionETCTax : IsCommissionETCTax=ojungsan.FItemList(0).IsCommissionETCTax

dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2
sumttl1 = 0
sumttl2 = 0
%>
<!-- �������Ϸ� ���� ��� �κ� -->
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=" & "�¶��� " & ojungsan.FItemList(0).Ftitle & ".xls"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
/* ���� �ٿ�ε�� ����� ���ڷ� ǥ�õ� ��� ���� */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>



<%
dim TTLitemCNT, TTLSellcashSum, TTLCouponDiscountSum, TTLReducedpriceSum
dim TTLCommissionSum, TTLSuplycashSum
%>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="100">���걸��</td>
    	<td width="100">����</td>
    	<td width="50">����<br>����</td>
		<td width="50">�ѰǼ�</td>
		<% if (IsCommissionTax) then %>
		    <% if (IsCommissionETCTax) then %>
		    <td width="100">���θ�Ǻ��<br>(���»� �δ�)</td>
    		<td width="100">���޴���<br>(����Ȯ����)</td>    
		    <% else %>
    		<td width="90">�����Ѿ�</td>
    		<td width="80">�⺻�Ǹ�<br>������</td>
    		<td width="50">&nbsp;</td>
            <td width="80">�������ξ�<br>(�ٹ����ٺδ�)</td>
            <td width="80">�����ֹ���<br>(���»�����)</td>
    		<td width="100">������</td>
    		<td width="100">���޴���<br>(����Ȯ����)</td>
    		<% end if %>
		<% else %>
    		<td width="150">�ǸŰ��Ѿ�</td>
    		<td width="150">���ް��Ѿ�</td>
    		<td width="100">���޸�����</td>
		<% end if %>
	</tr>

    <% for i=0 to ojungsanSubsmr.FResultCount-1 %>
    <% IF (gubun=ojungsanSubsmr.FItemList(i).Fgubuncd and (Not IsCommissionTax or (IsCommissionTax and itemvatyn=ojungsanSubsmr.FItemList(i).FitemVatyn))) then %>
    <tr bgcolor="#FFFFFF">
        <td align="center"><%= ojungsanSubsmr.FItemList(i).getJSummaryGugunName %></td>
        <td><%= ojungsanSubsmr.FItemList(i).getJGubuncd2Name %></td>
        <td align="center"><%= ojungsanSubsmr.FItemList(i).getTaxTypeName %></td>
        <td align="center"><%= ojungsanSubsmr.FItemList(i).FitemCNT %></td>
        <% if (IsCommissionTax) then %>
            <% if (IsCommissionETCTax) then %>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCommissionSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSuplycashSum,0) %></td>
            <% else %>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSellcashSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCouponDiscountSum+ojungsanSubsmr.FItemList(i).getCommissionSum,0) %></td>
            <td align="center">
            <% if (ojungsanSubsmr.FItemList(i).getSellcashSum<>0) then %>
            <%= CLNG((ojungsanSubsmr.FItemList(i).getCouponDiscountSum+ojungsanSubsmr.FItemList(i).getCommissionSum)/ojungsanSubsmr.FItemList(i).getSellcashSum*100*100)/100 %> %
            <% end if %>
            </td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCouponDiscountSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getReducedpriceSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCommissionSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSuplycashSum,0) %></td>
            <% end if %>
        <% else %>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSellcashSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSuplycashSum,0) %></td>
            <td align="center">
                <% if ojungsanSubsmr.FItemList(i).getSellcashSum<>0 then %>
                <%= CLng((1-ojungsanSubsmr.FItemList(i).getSuplycashSum/ojungsanSubsmr.FItemList(i).getSellcashSum)*10000)/100 %> %
                <% end if %>
            </td>
        <% end if %>
    </tr>
    <% end if %>
    <% next %>

</table>

<p>

<%
set ojungsan = Nothing
set ojungsanSubsmr = Nothing


dim ojungsandetail
set ojungsandetail = new CUpcheJungsan
ojungsandetail.FRectId = id
ojungsandetail.FRectgubun = gubun
ojungsandetail.FRectdesigner = session("ssBctID")
ojungsandetail.FRectOrder = "orderserial"
ojungsandetail.FRectItemVatYn = itemvatyn
'' 1357 ���������� �������� �ٸ�(����������)
if (id>1357) and (gubun<>"")   then
    ojungsandetail.JungsanDetailList
end if
%>
<!-- �ֹ��Ǻ� ����Ʈ ����-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<% if (IsCommissionTax) then %>
	        <% if (IsCommissionETCTax) then %>
	        <td colspan="12" align="left">
	        <% else %>
            <td colspan="17" align="left">
            <% end if %>
        <% else %>
        <td colspan="13" align="left">
        <% end if %>
        
			<b>�ֹ�/���/�԰�Ǻ� �󼼸���Ʈ</b>
			&nbsp;&nbsp;
			<% if ojungsandetail.FRectgubun="maeip" then %>
			â���԰�Ȯ���� �������� ��ϵ˴ϴ�.
			<% else %>
			����� �������� ��ϵ˴ϴ�.
			<% end if %>

			<% if ojungsandetail.FResultCount>=10000 then %>
			(�ִ� <%= ojungsandetail.FResultCount %> �� ǥ��)
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="70">�����ڵ�</td>
      <td width="70">�Ǹ�ä��</td>
      <td width="50">������</td>
      <td width="50">������</td>
      <td width="60">��ǰ�ڵ�</td>
      <td>��ǰ��</td>
      <td>�ɼǸ�</td>
      <td width="35">����</td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
          <td width="90">���θ�Ǻ��<br>(���»� �δ�)</td>
          <td width="90">�����</td>
          <td width="90">�����հ�</td>  
        <% else %>
          <td width="60">�����Ѿ�</td>
          <td width="60">�⺻�Ǹ�<br>������</td>
          <td width="50">&nbsp;</td>
          <td width="70">�������ξ�<br>(�ٹ����ٺδ�)</td>
          <td width="80">�����ֹ���<br>(���»�����)</td>
          <td width="60">������</td>
          <td width="60">�����</td>
          <td width="65">�����հ�</td>
        <% end if %>
      <% else %>
      <td width="50">�ǸŰ�</td>
      <td width="50">���ް�</td>
      <td width="50">���޸�����</td>
      <td width="65">���ް���</td>
      <% end if %>
      <td width="65">�����</td>
    </tr>
<% if ojungsandetail.FResultCount>0 and ojungsandetail.FRectgubun<>"" then %>
    <% for i=0 to ojungsandetail.FResultCount-1 %>

    <%
	sumttl1 = sumttl1 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td><%= ojungsandetail.FItemList(i).Fmastercode %></td>
      <td><%= ojungsandetail.FItemList(i).Fsitename %></td>
      <td><%= ojungsandetail.FItemList(i).FBuyname %></td>
      <td><%= ojungsandetail.FItemList(i).FReqname %></td>
      <td align="center"><%= ojungsandetail.FItemList(i).Fitemid %></td>
      <td align="left"><%= ojungsandetail.FItemList(i).FItemName %></td>
      <td><%= ojungsandetail.FItemList(i).FItemOptionName %></td>
      <td><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo) %>"><%= ojungsandetail.FItemList(i).FItemNo %></font></td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).getCommission) %>"><%= FormatNumber(ojungsandetail.FItemList(i).getCommission,0) %></font></td>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
        <% else %>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash,0) %></font></td>
          <td align="right"><%= FormatNumber(ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission,0) %></td>
          <td align="center">
          <% if (ojungsandetail.FItemList(i).Fsellcash<>0) then %>
          <%= CLNG((ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)/ojungsandetail.FItemList(i).Fsellcash*100) %> %
          <% end if %>
          </td>
          <td align="right"><%= FormatNumber(ojungsandetail.FItemList(i).getCouponDiscount,0) %></td>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).getReducedprice) %>"><%= FormatNumber(ojungsandetail.FItemList(i).getReducedprice,0) %></font></td>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).getCommission) %>"><%= FormatNumber(ojungsandetail.FItemList(i).getCommission,0) %></font></td>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
         <% end if %>
      <% else %>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="center">
      <% if ojungsandetail.FItemList(i).Fsellcash<>0 then %>
      <%= 100-CLNG((ojungsandetail.FItemList(i).Fsuplycash)/ojungsandetail.FItemList(i).Fsellcash*100) %> %
      <% end if %>
      </td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <% end if %>
      <td align="center"><%=ojungsandetail.FItemList(i).FExecDate%></td>
    </tr>
     <% if (i mod 1000)=0 then response.flush %>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
      <td>�հ�</td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
            <td colspan="9"></td>
        <% else %>
            <td colspan="14"></td>
        <% end if %>
      <% else %>
      <td colspan="10"></td>
      <% end if %>
      <td align="right"><%=FormatNumber(sumttl2,0)%></td>
      <td ></td>
    </tr>
<% else %>

<% end if %>
</table>
<!-- �ֹ��Ǻ� ����Ʈ ��-->

<%
set ojungsandetail = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

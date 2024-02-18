<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������೻��
' History : ������ ����
'			2021.04.28 �ѿ�� ����(����Ǵ� �׸�� �����û. �繫��:������)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2, ItemNo, CpnNotAppliedPrice, CouponDiscountCommission, CouponDiscount, Reducedprice, Commission, suplycash
dim TTLitemCNT, TTLSellcashSum, TTLCouponDiscountSum, TTLReducedpriceSum
dim TTLCommissionSum, TTLSuplycashSum
dim id,gubun, itemvatyn, makerid, groupid
id      = requestCheckVar(request("id"),10)
gubun   = requestCheckVar(request("gubun"),20)
itemvatyn = requestCheckVar(request("itemvatyn"),10)

makerid = requestCheckVar(request("makerid"),32)
groupid = getPartnerId2GroupID(makerid)

'if (NOT chkAvailViewJungsanON(id,makerid,groupid)) then
'    response.write "��ȸ ������ �����ϴ�"
'    dbget.close()	:	response.End
'end if

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

Dim IsShowCpnNotAppliedPrice ''2018/07/02
IsShowCpnNotAppliedPrice = (ojungsan.FItemList(0).FYYYYMM>="2018-06") and (ojungsan.FItemList(0).FJGubun="CC") 
if (application("Svr_Info")	= "Dev") then IsShowCpnNotAppliedPrice = true

dim ojungsanSubsmr
set ojungsanSubsmr = new CUpcheJungsan
ojungsanSubsmr.FRectId = id
ojungsanSubsmr.FRectdesigner = session("ssBctID")
'ojungsanSubsmr.getJungsanSubSummary

Dim IsCommissionTax : IsCommissionTax=ojungsan.FItemList(0).IsCommissionTax
Dim IsCommissionETCTax : IsCommissionETCTax=ojungsan.FItemList(0).IsCommissionETCTax
sumttl1=0
sumttl2=0
ItemNo=0
CpnNotAppliedPrice=0
CouponDiscountCommission=0
CouponDiscount=0
Reducedprice=0
Commission=0
suplycash=0

' �������Ϸ� ���� ��� �κ�
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

<!--<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
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
		    <% if (IsShowCpnNotAppliedPrice) then %><td width="90">�Ǹ��Ѿ�</td><% end if %>
    		<td width="90">�����Ѿ�</td>
    		<td width="80">�⺻�Ǹ�<br>������</td>
    		<td width="50">&nbsp;</td>
            <td width="80">�������ξ�<br>(�ٹ����ٺδ�)</td>
            <td width="80">�����ֹ���<br>(���»�����)</td>
    		<td width="90">������</td>
    		<td width="90">�������������</td>
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
            <% if (IsShowCpnNotAppliedPrice) then %><td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).FCpnNotAppliedPriceSum,0) %></td><% end if %>
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
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getPGCommissionSum,0) %></td>
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

</table>-->

<%
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
        <td colspan="12"  align="left">
        <% else %>
        <td colspan="<%=CHKIIF(IsShowCpnNotAppliedPrice,"18","17")%>"  align="left">
        <% end if %>
        <% else %>
        <td colspan="12"  align="left">
        <% end if %>
      
			<b>�ֹ�/���/�԰�Ǻ� �󼼸���Ʈ</b>
			&nbsp;&nbsp;
			<% if ojungsandetail.FRectgubun="maeip" then %>
			���������԰�Ȯ���� �������� ��ϵ˴ϴ�.
			<% else %>
			��ۿϷ��� ����
			<% end if %>

			<% if ojungsandetail.FResultCount>=10000 then %>
			(�ִ� <%= ojungsandetail.FResultCount %> �� ǥ��)
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td>�����ڵ�</td>
      <td>�Ǹ�ä��</td>
      <td>������</td>
      <td>������</td>
      <td>��ǰ�ڵ�</td>
      <td>��ǰ��</td>
      <td>�ɼǸ�</td>
      <td>����</td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
          <td>���θ�Ǻ��(���»� �δ�)</td>
          <td>�����</td>
          <% '<td>�����հ�</td> %>
        <% else %>
          <% if (IsShowCpnNotAppliedPrice) then %><td>�Ǹ��Ѿ�</td><% end if %>
          <td>�����Ѿ�</td>
          <td>�⺻�Ǹż�����</td>
          <td>����������</td>
          <td>�������ξ�(�ٹ����ٺδ�)</td>
          <td>�����ֹ���(���»�����)</td>
          <td>��������</td>
          <% '<td>�������������</td> %>
          <td>�����</td>
          <% '<td>�����հ�</td> %>
        <% end if %>
      <% else %>
      <td>�ǸŴܰ�</td>
      <td>���޴ܰ�</td>
      <td>���޸�����</td>
      <td>���ް��հ�</td>
      <% end if %>
    </tr>
<% if ojungsandetail.FResultCount>0 and ojungsandetail.FRectgubun<>"" then %>
    <% for i=0 to ojungsandetail.FResultCount-1 %>

    <%
	sumttl1 = sumttl1 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash
	ItemNo = ItemNo + ojungsandetail.FItemList(i).FItemNo
	CpnNotAppliedPrice = CpnNotAppliedPrice + ojungsandetail.FItemList(i).FCpnNotAppliedPrice*ojungsandetail.FItemList(i).FItemNo
	CouponDiscountCommission = CouponDiscountCommission + (ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)*ojungsandetail.FItemList(i).FItemNo
	CouponDiscount = CouponDiscount + ojungsandetail.FItemList(i).getCouponDiscount*ojungsandetail.FItemList(i).FItemNo
	Reducedprice = Reducedprice + ojungsandetail.FItemList(i).getReducedprice*ojungsandetail.FItemList(i).FItemNo
	Commission = Commission + ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo
	suplycash = suplycash + ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td><%= ojungsandetail.FItemList(i).Fmastercode %></td>
      <td><%= ojungsandetail.FItemList(i).Fsitename %></td>
        <td>
            <% if C_CriticInfoUserLV1 then %>
                <%= ojungsandetail.FItemList(i).FBuyname %>
            <% else %>
                <%= AstarUserName(ojungsandetail.FItemList(i).FBuyname) %>
            <% end if %>
        </td>
        <td>
            <% if C_CriticInfoUserLV1 then %>
                <%= ojungsandetail.FItemList(i).FBuyname %>
            <% else %>
                <%= AstarUserName(ojungsandetail.FItemList(i).FReqname) %>
            <% end if %>
        </td>
      <td align="center"><%= ojungsandetail.FItemList(i).Fitemid %></td>
      <td align="left"><%= ojungsandetail.FItemList(i).FItemName %></td>
      <td><%= ojungsandetail.FItemList(i).FItemOptionName %></td>
      <td><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo) %>"><%= ojungsandetail.FItemList(i).FItemNo %></font></td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
			<% '���θ�Ǻ��(���»� �δ�)  %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<% '�����  %>
			<td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo,0) %></font></td>
			<!--<td align="right"><font color="<%'= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%'= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>-->
        <% else %>
			<%
			' �Ǹ��Ѿ�
			if (IsShowCpnNotAppliedPrice) then
			%>
				<td align="right">
					<font color="<%= MinusFont(ojungsandetail.FItemList(i).FCpnNotAppliedPrice*ojungsandetail.FItemList(i).FItemNo) %>">
					<%= FormatNumber(ojungsandetail.FItemList(i).FCpnNotAppliedPrice*ojungsandetail.FItemList(i).FItemNo,0) %></font>
				</td>
			<% end if %>
			<% ' �����Ѿ� %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<% ' �⺻�Ǹż����� %>
			<td align="right">
				<%= FormatNumber((ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)*ojungsandetail.FItemList(i).FItemNo,0) %>
			</td>
          <td align="center">
          <% if (ojungsandetail.FItemList(i).Fsellcash<>0) then %>
          <%= CLNG((ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)/ojungsandetail.FItemList(i).Fsellcash*100) %> %
          <% end if %>
          </td>
			<% ' �������ξ�(�ٹ����ٺδ�) %>
			<td align="right"><%= FormatNumber(ojungsandetail.FItemList(i).getCouponDiscount*ojungsandetail.FItemList(i).FItemNo,0) %></td>
			<% ' �����ֹ���(���»�����) %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).getReducedprice*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).getReducedprice*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<% ' ������ %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<!--<td align="right"><font color="<%'= MinusFont(ojungsandetail.FItemList(i).getPgCommission) %>"><%'= FormatNumber(ojungsandetail.FItemList(i).getPgCommission,0) %></font></td>-->
			<% ' ����� %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<!--<td align="right"><font color="<%'= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%'= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>-->
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

    </tr>
     <% if (i mod 1000)=0 then response.flush %>
    <% next %>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<% if (IsCommissionTax) then %>
			<% if (IsCommissionETCTax) then %>
			    <td><strong>�հ�</strong></td>
				<td colspan="7"></td>
				<td align="right"><strong><%= FormatNumber(Commission,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(suplycash,0) %></strong></td>
				<!--<td align="right"><%'=FormatNumber(sumttl2,0)%></td>-->
			<% else %>
			    <td colspan=7><strong>�հ�</strong></td>
				<td><strong><%= FormatNumber(ItemNo,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(CpnNotAppliedPrice,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(sumttl1,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(CouponDiscountCommission,0) %></strong></td>
				<td></td>
				<td align="right"><strong><%= FormatNumber(CouponDiscount,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(Reducedprice,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(Commission,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(suplycash,0) %></strong></td>
				<!--<td align="right"><%'=FormatNumber(sumttl2,0)%></td>-->
			<% end if %>
		<% else %>
			<td>�հ�</td>
			<td colspan="10"></td>
			<td align="right"><strong><%=FormatNumber(sumttl2,0)%></strong></td>
		<% end if %>
      
    </tr>
<% else %>

<% end if %>
</table>
<!-- �ֹ��Ǻ� ����Ʈ ��-->

<%
set ojungsan = Nothing
'set ojungsanSubsmr = Nothing
set ojungsandetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

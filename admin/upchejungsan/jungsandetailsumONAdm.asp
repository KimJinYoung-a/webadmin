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
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2, ItemNo, CpnNotAppliedPrice, CouponDiscountCommission, CouponDiscount, Reducedprice, Commission, suplycash
dim TTLitemCNT, TTLSellcashSum, TTLCouponDiscountSum, TTLReducedpriceSum
dim TTLCommissionSum, TTLSuplycashSum
dim TTLPgCommissionSum, TTLCpnNotAppliedPrice
dim id,gubun, makerid, groupid, itemvatyn
makerid = requestCheckVar(request("makerid"),32)
id      = requestCheckVar(request("id"),10)
gubun   = requestCheckVar(request("gubun"),20)
itemvatyn = requestCheckVar(request("itemvatyn"),10)


''groupid = requestCheckvar(request("groupid"),10) ''getPartnerId2GroupID(makerid)

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectgubun = gubun
ojungsan.FRectGroupID = groupid

''if (groupid<>"") then '' ������ �׷�ID �� ������� ��ȸ
    ojungsan.JungsanMasterList
''end if

if ojungsan.FresultCount <1 then
    set ojungsan = Nothing
	dbget.close()	:	response.End
end if

Dim IsShowCpnNotAppliedPrice ''2018/07/02
IsShowCpnNotAppliedPrice = (ojungsan.FItemList(0).FYYYYMM>="2018-06") and (ojungsan.FItemList(0).FJGubun="CC") 
if (application("Svr_Info")	= "Dev") then IsShowCpnNotAppliedPrice = true

dim ojungsanSubsmr
set ojungsanSubsmr = new CUpcheJungsan
ojungsanSubsmr.FRectId = id
ojungsanSubsmr.FRectdesigner = ojungsan.FItemList(0).Fdesignerid
ojungsanSubsmr.getJungsanSubSummary

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
%>

<script language='javascript'>
function JunsanDetailList(id,gubun,itemvatyn){
    location.href = '?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn + '&makerid=<%=makerid%>';
}
function ExcelJunsanDetailList(id,gubun,itemvatyn){
//alert('..');
//return;
    location.href = '/admin/upchejungsan/jungsandetailsumAdm_excel.asp?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn + '&makerid=<%=makerid%>';
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
        	<b>
        	�¶��� <%= ojungsan.FItemList(0).Ftitle %>&nbsp;[<%= ojungsan.FItemList(0).Fdesignerid %>]
        	&nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).Fdifferencekey %> ��
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).getJGuBunName %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).getTaxTypeName %>&nbsp;&nbsp;
            </b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

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
		    <% if (IsShowCpnNotAppliedPrice) then %><td width="90">�Ǹ��Ѿ�</td><% end if %>
    		<td width="90">�����Ѿ�<br>(��ǰ��������)</td>
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
		<td width="50">�󼼳���</td>
	</tr>

    <% for i=0 to ojungsanSubsmr.FResultCount-1 %>
    <%
        TTLitemCNT              = TTLitemCNT + ojungsanSubsmr.FItemList(i).FitemCNT
        TTLSellcashSum          = TTLSellcashSum + ojungsanSubsmr.FItemList(i).getSellcashSum
        TTLCouponDiscountSum    = TTLCouponDiscountSum + ojungsanSubsmr.FItemList(i).getCouponDiscountSum
        TTLReducedpriceSum      = TTLReducedpriceSum + ojungsanSubsmr.FItemList(i).getReducedpriceSum
        TTLCommissionSum        = TTLCommissionSum + ojungsanSubsmr.FItemList(i).getCommissionSum
        TTLPgCommissionSum      = TTLPgCommissionSum + ojungsanSubsmr.FItemList(i).getPgCommissionSum
        TTLSuplycashSum         = TTLSuplycashSum + ojungsanSubsmr.FItemList(i).getSuplycashSum
        
        TTLCpnNotAppliedPrice   = TTLCpnNotAppliedPrice + ojungsanSubsmr.FItemList(i).FCpnNotAppliedPriceSum

    %>
    <tr bgcolor='<%=CHKIIF(gubun=ojungsanSubsmr.FItemList(i).Fgubuncd and (Not IsCommissionTax or (IsCommissionTax and itemvatyn=ojungsanSubsmr.FItemList(i).FitemVatyn)),"#CCCCFF","#FFFFFF") %>' >
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

        <td align="center">
		  <a href="javascript:JunsanDetailList('<%= id %>','<%=ojungsanSubsmr.FItemList(i).Fgubuncd%>','<%=CHKIIF(IsCommissionTax,ojungsanSubsmr.FItemList(i).FitemVatyn,"")%>')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
		  <a href="#" onclick="ExcelJunsanDetailList('<%= id %>','<%=ojungsanSubsmr.FItemList(i).Fgubuncd%>','<%=CHKIIF(IsCommissionTax,ojungsanSubsmr.FItemList(i).FitemVatyn,"")%>'); return false;"><img src="/images/iexcel.gif" width="16" border="0"></a>
		</td>
    </tr>
    <% next %>

	<tr bgcolor="#FFFFFF">
		<td align=center>�Ѱ�</td>
		<td></td>
		<td></td>
		<td align=center><%= FormatNumber(TTLitemCNT,0) %></td>
		<% if (IsCommissionTax) then %>
		    <% if (IsCommissionETCTax) then %>
    		<td align=right><%= FormatNumber(TTLCommissionSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLSuplycashSum,0) %></td>
		    <% else %>
		    <% if (IsShowCpnNotAppliedPrice) then %><td align=right><%= FormatNumber(TTLCpnNotAppliedPrice,0) %></td><% end if %>
    		<td align=right><%= FormatNumber(TTLSellcashSum,0) %></td>
    		<td align="right"><%= FormatNumber(TTLCouponDiscountSum+TTLCommissionSum,0) %></td>
            <td align="center">
              <% if (TTLSellcashSum<>0) then %>
              <%= CLNG((TTLCouponDiscountSum+TTLCommissionSum)/TTLSellcashSum*100*100)/100 %> %
              <% end if %>
            </td>
    		<td align=right><%= FormatNumber(TTLCouponDiscountSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLReducedpriceSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLCommissionSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLPgCommissionSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLSuplycashSum,0) %></td>
    	    <% end if %>
		<% else %>
    		<td align=right><%= FormatNumber(TTLSellcashSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLSuplycashSum,0) %></td>
    		<td align=center>
		<% if TTLSellcashSum<>0 then %>
		    <%= CLng((1-TTLSuplycashSum/TTLSellcashSum)*10000)/100 %> %
		<% end if %>
		</td>
		<% end if %>
		<td></td>
	</tr>
</table>

<br>

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
		<td colspan="<%=CHKIIF(IsCommissionTax,20,13)%>" align="left">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
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
      <td>�ɼ��ڵ�</td>
      <td>��ǰ��</td>
      <td>�ɼǸ�</td>
      <td>����</td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
            <td>���θ�Ǻ��<br>(���»� �δ�)</td>  
            <td>�����</td>
            <% '<td>�����հ�<br>(����*�����)</td> %>
            <td>�ְ�������</td>
        <% else %>
            <% if (IsShowCpnNotAppliedPrice) then %><td>�Ǹ��Ѿ�</td><% end if %>
            <td>�����Ѿ�<br>(��ǰ��������)</td>
            <td>�⺻�Ǹ�<br>������</td>
            <td>����������</td>
            <td>�������ξ�<br>(�ٹ����ٺδ�)</td>
            <td>�����ֹ���<br>(���»�����)</td>
            <td>��������</td>
            <% '<td>��������<br>������</td> %>
            <td>�����</td>
            <% '<td>�����հ�<br>(����*�����)</td> %>
            <td>�ְ�������</td>
        <% end if %>
      <% else %>
      <td>�ǸŴܰ�</td>
      <td>���޴ܰ�</td>
      <td>���޸�����</td>
      <td>���ް��հ�<br>(����*���ް�)</td>
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
      <td align="center"><%= ojungsandetail.FItemList(i).FitemOption %></td>
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
            <td align="center"><%=ojungsandetail.FItemList(i).getpaymethodName %></td>  
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
            <td align="center"><%=ojungsandetail.FItemList(i).getpaymethodName %></td>
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
    <tr bgcolor="#FFFFFF" align="center">
        <% if (IsCommissionTax) then %>
            <% if (IsCommissionETCTax) then %>
                <td><strong>�հ�</strong></td>
				<td colspan="8"></td>
				<td align="right"><strong><%= FormatNumber(Commission,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(suplycash,0) %></strong></td>
                <!--<td align="right"><strong><%'=FormatNumber(sumttl2,0)%></strong></td>-->
            <% else %>
			    <td colspan=8><strong>�հ�</strong></td>
				<td><strong><%= FormatNumber(ItemNo,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(CpnNotAppliedPrice,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(sumttl1,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(CouponDiscountCommission,0) %></strong></td>
				<td></td>
				<td align="right"><strong><%= FormatNumber(CouponDiscount,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(Reducedprice,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(Commission,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(suplycash,0) %></strong></td>
                <!--<td align="right"><strong><%'=FormatNumber(sumttl2,0)%></strong></td>-->
            <% end if %>

            <td>&nbsp;</td>
        <% else %>
            <td><strong>�հ�</strong></td>
            <td colspan="11"></td>
            <td align="right"><strong><%=FormatNumber(sumttl2,0)%></strong></td>
        <% end if %>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="<%=CHKIIF(IsCommissionTax,20,13)%>" align="center"><img src="/images/icon_search.jpg" width="16" border="0" align="absbottom">&nbsp;�󼼳����� �����ϼ���.</td>
    </tr>
<% end if %>
</table>
<!-- �ֹ��Ǻ� ����Ʈ ��-->

<%
set ojungsan = Nothing
set ojungsanSubsmr = Nothing
set ojungsandetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
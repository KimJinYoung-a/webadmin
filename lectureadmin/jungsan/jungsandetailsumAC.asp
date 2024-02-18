<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id,gubun
dim makerid, groupid, itemvatyn
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

'ojungsan.FRectDesigner = makerid        ''�����귣���ϰ�� ������ ��. chkAvailViewJungsanON�� ��ü
'if (makerid<>"") then
'    ojungsan.JungsanMasterList
'end if
'ojungsan.FRectGroupID = groupid
'if (groupid<>"") then
'    ojungsan.JungsanMasterList
'end if

ojungsan.JungsanMasterList

if ojungsan.FresultCount <1 then
    set ojungsan = Nothing
	dbget.close()	:	response.End
end if

makerid = ojungsan.FItemList(0).Fdesignerid

dim ojungsanSubsmr
set ojungsanSubsmr = new CUpcheJungsan
ojungsanSubsmr.FRectId = id
ojungsanSubsmr.FRectdesigner = makerid
ojungsanSubsmr.getJungsanSubSummary

Dim IsCommissionTax : IsCommissionTax=ojungsan.FItemList(0).IsCommissionTax
Dim IsCommissionETCTax : IsCommissionETCTax=ojungsan.FItemList(0).IsCommissionETCTax

dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2

%>

<script language='javascript'>
function JunsanDetailList(id,gubun,itemvatyn){
    location.href = '?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn;
}
function ExcelJunsanDetailList(id,gubun,itemvatyn){
    <% if (LCASE(session("ssBctID"))="memorette") or (LCASE(session("ssBctID"))="1ppm") then %>
    location.href = 'jungsandetailsum_excel_exdt.asp?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn;    
    <% else %>
    location.href = 'jungsandetailsum_excel.asp?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn;
    <% end if %>
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
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

<%
dim TTLitemCNT, TTLSellcashSum, TTLCouponDiscountSum, TTLReducedpriceSum
dim TTLCommissionSum, TTLSuplycashSum
dim TTLPgCommissionSum
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
            <td width="80">�������ξ�<br>(���ΰŽ��δ�)</td>
            <td width="80">�����ֹ���<br>(���»�����)</td>
    		<td width="90">��ǰ�Ǹ�<br>������</td>
    		<td width="90">��������<br>������</td>
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
		  <a href="javascript:ExcelJunsanDetailList('<%= id %>','<%=ojungsanSubsmr.FItemList(i).Fgubuncd%>','<%=CHKIIF(IsCommissionTax,ojungsanSubsmr.FItemList(i).FitemVatyn,"")%>')"><img src="/images/iexcel.gif" width="16" border="0"></a>
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

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<p>

<%
set ojungsan = Nothing
set ojungsanSubsmr = Nothing


sumttl1 = 0
sumttl2 = 0

dim ojungsandetail
set ojungsandetail = new CUpcheJungsan
ojungsandetail.FRectId = id
ojungsandetail.FRectgubun = gubun
ojungsandetail.FRectdesigner = makerid
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
        <td colspan="12" align="left">
        <% end if %>
			<img src="/images/icon_arrow_down.gif" align="absbottom">
			<b>�ֹ�/���/�԰�Ǻ� �󼼸���Ʈ</b>
			&nbsp;&nbsp;
			<% if ojungsandetail.FRectgubun="maeip" then %>
			���������԰�Ȯ���� �������� ��ϵ˴ϴ�.
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
          <td width="90">�����հ�<br>(����*�����)</td>       
        <% else %>
          <td width="60">�����Ѿ�</td>
          <td width="60">�⺻�Ǹ�<br>������</td>
          <td width="50">&nbsp;</td>
          <td width="70">�������ξ�<br>(���ΰŽ��δ�)</td>
          <td width="80">�����ֹ���<br>(���»�����)</td>
          <td width="60">��ǰ�Ǹ�<br>������</td>
          <td width="60">��������<br>������</td>
          <td width="60">�����</td>
          <td width="80">�����հ�<br>(����*�����)</td>
        <% end if %>
      <% else %>
      <td width="50">�ǸŰ�</td>
      <td width="50">���ް�</td>
      <td width="60">���޸�����</td>
      <td width="80">���ް���<br>(����*���ް�)</td>
      <% end if %>
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
          <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).getPgCommission) %>"><%= FormatNumber(ojungsandetail.FItemList(i).getPgCommission,0) %></font></td>
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
    </tr>
    <% if (i mod 1000)=0 then response.flush %>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
      <td>�հ�</td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
        <td colspan="10"></td>    
        <% else %>
        <td colspan="15"></td>
        <% end if %>
      <% else %>
      <td colspan="10"></td>
      <% end if %>
      <td align="right"><%=FormatNumber(sumttl2,0)%></td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="<%=CHKIIF(IsCommissionTax,17,13)%>" align="center"><img src="/images/icon_search.jpg" width="16" border="0" align="absbottom">&nbsp;�󼼳����� �����ϼ���.</td>
    </tr>
<% end if %>
</table>
<!-- �ֹ��Ǻ� ����Ʈ ��-->

<%
set ojungsandetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
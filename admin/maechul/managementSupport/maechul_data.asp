<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->

<%
	dim bancancle,accountdiv,sitename,ipkumdatesucc, vPurchasetype, vatinclude
	dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
	dim i ,defaultdate,defaultdate1 , olddata
    dim mdgbn
	sitename = request("sitenamebox")
	accountdiv = request("accountdiv")
	vPurchasetype = request("purchasetype")
	bancancle = NullFillWith(request("bancancle"), "1")
	vatinclude = request("vatinclude")

	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 60�������� �˻�
	yyyy1 = NullFillWith(request("yyyy1"), left(defaultdate1,4))
	mm1 = NullFillWith(request("mm1"), mid(defaultdate1,6,2))
	dd1 = NullFillWith(request("dd1"), right(defaultdate1,2))
	yyyy2 = NullFillWith(request("yyyy2"), year(now))
	mm2 = NullFillWith(request("mm2"), month(now))
	mm2 = TwoNumber(mm2)
	dd2 = NullFillWith(request("dd2"), day(now))
	dd2 = TwoNumber(dd2)
    mdgbn = NullFillWith(request("mdgbn"),"m")

	dim Omaechul_list
	set Omaechul_list = new cManagementSupportMaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc
	Omaechul_list.frectpurchasetype = vPurchasetype
	Omaechul_list.frectvatinclude = vatinclude
	Omaechul_list.frectGroupByMonth=mdgbn

	Omaechul_list.fmaechul_list()


	Dim vSum_TotItemNo, vSum_TotReducedPrice, vSum_TotBuycash, vSum_TotBuycashCouponNotApplied
	Dim vSum_TotOrgitemcost, vSum_TotOrgitemcostDLV, vSum_TotItemcostCouponNotApplied, vSum_TotItemcostCouponNotAppliedDLV, vSum_TotItemcost, vSum_TotItemcostDLV
	Dim vSum_TotReducePrice, vSum_TotReducePriceDLV, vSum_SpendCouponSum, vSum_SpendCouponSumDLV, vSum_MaechulItem, vSum_MaechulItemDLV
	Dim vSum_SpendMileSum, vSum_SpendMileSumDLV
%>
<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">��ǰ����� / ��¥ <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;&nbsp;&nbsp;
			<input type="radio" name="mdgbn" value="m" <%= CHKIIF(mdgbn="m","checked","") %> >����
			<input type="radio" name="mdgbn" value="d" <%= CHKIIF(mdgbn="d","checked","") %> >�Ϻ�
			</td>
		</tr>
    	<tr>
    		<td height="25">
	    	<input type=radio name="bancancle" value="1" <% if bancancle="1" then  response.write "checked" %>>��ǰ����
	    	<input type=radio name="bancancle" value="2" <% if bancancle="2" then  response.write "checked" %>>��ǰ�Ǹ�
	    	<input type=radio name="bancancle" value="3" <% if bancancle="3" then  response.write "checked" %>>��ǰ����
	    	/ �������� <select name="accountdiv">
	    		<option value="" <% if accountdiv = "" then response.write "selected" %>>��ü</option>
	    		<option value="7" <% if accountdiv = "7" then response.write "selected" %>>������</option>
				<option value="14" <% if accountdiv = "14" then response.write "selected" %>>����������</option>
	    		<option value="20" <% if accountdiv = "20" then response.write "selected" %>>�ǽð�</option>
	    		<option value="50" <% if accountdiv = "50" then response.write "selected" %>>�ܺθ�</option>
	    		<option value="80" <% if accountdiv = "80" then response.write "selected" %>>�ÿ�</option>
	    		<option value="100" <% if accountdiv = "100" then response.write "selected" %>>�ſ�ī��</option>
	    	</select>
	    	&nbsp;&nbsp;&nbsp;
	    	/ �������� <select name="vatinclude">
	    	    <option value="" <% if vatinclude = "" then response.write "selected" %>>��ü</option>
	    		<option value="Y" <% if vatinclude = "Y" then response.write "selected" %>>����</option>
	    		<option value="N" <% if vatinclude = "N" then response.write "selected" %>>�鼼</option>
	    	</select>
	    	&nbsp;&nbsp;&nbsp;
	    	����Ʈ���� : <% Drawsitename "sitenamebox",sitename %>
	    	&nbsp;&nbsp;&nbsp;
	    	�⺻������ : <% drawPartnerCommCodeBox true,"selljungsantype","purchasetype",vPurchasetype,"" %>
	    	</td>
	    </tr>
	    </table>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="70" rowspan="2">���<%=CHKIIF(mdgbn="m","��","��") %></td>
    <td align="center" width="50" rowspan="2">�ѻ�ǰ<br>����</td>
	<% if (C_InspectorUser = False) then %>
    <td align="center" colspan="2">�Һ��ڰ�<br>A</td>
    <td align="center" colspan="2">���αݾ�<br>B</td>
    <td align="center" colspan="2">�ǸŰ�(���ΰ�)<br>C=A-B</td>
    <td align="center" colspan="2">��ǰ��������<br>D</td>
    <td align="center" colspan="2">�����Ѿ�<br>E=C-D</td>
    <td align="center" colspan="4">���ʽ�����<br>��������(F)=E-ȯ�Ҿ�(reducePrice)<br>��������(G)</td>
	<% end if %>
    <td align="center" colspan="2">����<br>��ǰ(H)=E-F-G</td>
    <td align="center" width="50" rowspan="2">���</td>
    <td align="center" width="10" rowspan="2"></td>
    <td align="center" colspan="2">���ϸ���<br>���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
	<td>��������</td>
	<td>��ۺ�����</td>
	<td>��������<br>�Ⱥ�(��ǰ)</td>
	<td>��������<br>�Ⱥ�(��ۺ�)</td>
	<% end if %>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>���ϸ���<br>�Ⱥ�(��ǰ)</td>
	<td>���ϸ���<br>�Ⱥ�(��ۺ�)</td>
</tr>
<%
Dim vYear, vMonth, vDay
For i = 0 To Omaechul_list.ftotalcount -1
	vYear	= Year(Omaechul_list.flist(i).fbaesongdate)
	vMonth	= TwoNumber(Month(Omaechul_list.flist(i).fbaesongdate))
	vDay	= TwoNumber(Day(Omaechul_list.flist(i).fbaesongdate))
%>
<tr align="center" bgcolor="#FFFFFF">
    <td align="center">
    <% IF(mdgbn="m") then %>
        <%= Omaechul_list.flist(i).fbaesongdate %>
    <% else %>
    	<% if right(FormatDateTime(Omaechul_list.flist(i).fbaesongdate,1),3) = "�����" then %>
    		<font color="blue"><%= Omaechul_list.flist(i).fbaesongdate %></font>
    	<% elseif right(FormatDateTime(Omaechul_list.flist(i).fbaesongdate,1),3) = "�Ͽ���" then %>
    		<font color="red"><%= Omaechul_list.flist(i).fbaesongdate %></font>
    	<% else %>
    		<%= Omaechul_list.flist(i).fbaesongdate %>
    	<% end if %>
    <% end if %>
	</td>
    <td align="center"><%= Replace(Omaechul_list.flist(i).ftot_itemno,"-","") %></td>
	<% if (C_InspectorUser = False) then %>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost_d) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost - Omaechul_list.flist(i).ftot_itemcostCouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost_d - Omaechul_list.flist(i).ftot_itemcostCouponNotApplied_d) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied_d) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied - Omaechul_list.flist(i).ftot_itemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied_d - Omaechul_list.flist(i).ftot_itemcost_d) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost_d) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost_d-Omaechul_list.flist(i).ftot_reducedPrice_d) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendCouponSum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendCouponSum_d) %></td>
	<% end if %>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost-(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice)-Omaechul_list.flist(i).ftot_DivSpendCouponSum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost_d-(Omaechul_list.flist(i).ftot_itemcost_d-Omaechul_list.flist(i).ftot_reducedPrice_d)-Omaechul_list.flist(i).ftot_DivSpendCouponSum_d) %></td>
	<td align="center" >[<a href="/admin/upchejungsan/upcheselllist.asp?datetype=chulgoil&yyyy1=<%=vYear%>&mm1=<%=vMonth%>&dd1=<%=vDay%>&yyyy2=<%=vYear%>&mm2=<%=vMonth%>&dd2=<%=vDay%>&delivertype=all" target="_blank">��</a>]</td>
	<td align="center" ></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendMileSum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendMileSum_d) %></td>
</tr>
<%
	vSum_TotItemNo 						= vSum_TotItemNo + Omaechul_list.flist(i).ftot_itemno
	vSum_TotReducedPrice 				= vSum_TotReducedPrice + Omaechul_list.flist(i).ftot_reducedPrice
	vSum_TotBuycash 					= vSum_TotBuycash + Omaechul_list.flist(i).ftot_buycash
	vSum_TotBuycashCouponNotApplied 	= vSum_TotBuycashCouponNotApplied + Omaechul_list.flist(i).ftot_buycashCouponNotApplied
	vSum_TotOrgitemcost 				= vSum_TotOrgitemcost + Omaechul_list.flist(i).ftot_orgitemcost
	vSum_TotOrgitemcostDLV 				= vSum_TotOrgitemcostDLV + Omaechul_list.flist(i).ftot_orgitemcost_d
	vSum_TotItemcostCouponNotApplied 	= vSum_TotItemcostCouponNotApplied + Omaechul_list.flist(i).ftot_itemcostCouponNotApplied
	vSum_TotItemcostCouponNotAppliedDLV = vSum_TotItemcostCouponNotAppliedDLV + Omaechul_list.flist(i).ftot_itemcostCouponNotApplied_d
	vSum_TotItemcost 					= vSum_TotItemcost + Omaechul_list.flist(i).ftot_itemcost
	vSum_TotItemcostDLV 				= vSum_TotItemcostDLV + Omaechul_list.flist(i).ftot_itemcost_d
	vSum_TotReducePrice					= vSum_TotReducePrice + Omaechul_list.flist(i).ftot_reducedPrice
	vSum_TotReducePriceDLV				= vSum_TotReducePriceDLV + Omaechul_list.flist(i).ftot_reducedPrice_d
	vSum_SpendCouponSum					= vSum_SpendCouponSum + Omaechul_list.flist(i).ftot_DivSpendCouponSum
	vSum_SpendCouponSumDLV				= vSum_SpendCouponSumDLV + Omaechul_list.flist(i).ftot_DivSpendCouponSum_d
	vSum_MaechulItem					= vSum_MaechulItem + (Omaechul_list.flist(i).ftot_itemcost-(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice)-Omaechul_list.flist(i).ftot_DivSpendCouponSum)
	vSum_MaechulItemDLV					= vSum_MaechulItemDLV + (Omaechul_list.flist(i).ftot_itemcost_d-(Omaechul_list.flist(i).ftot_itemcost_d-Omaechul_list.flist(i).ftot_reducedPrice_d)-Omaechul_list.flist(i).ftot_DivSpendCouponSum_d)

	vSum_SpendMileSum					= vSum_SpendMileSum + Omaechul_list.flist(i).ftot_DivSpendMileSum
	vSum_SpendMileSumDLV				= vSum_SpendMileSumDLV + Omaechul_list.flist(i).ftot_DivSpendMileSum_d

Next
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" rowspan="2">
	�Ѱ�
	</td>
	<td align="center"  rowspan="2"><%= Replace(vSum_TotItemNo,"-","") %></td>
	<% if (C_InspectorUser = False) then %>
	<td align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcost) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcostDLV) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcost - vSum_TotItemcostCouponNotApplied) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcostDLV - vSum_TotItemcostCouponNotAppliedDLV) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotAppliedDLV) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied - vSum_TotItemcost) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotAppliedDLV - vSum_TotItemcostDLV) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcost) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcostDLV) %></td>

	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcost-vSum_TotReducePrice) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcostDLV-vSum_TotReducePriceDLV) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_SpendCouponSum) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_SpendCouponSumDLV) %></td>
	<% end if %>
	<td align="right"><%= NullOrCurrFormat(vSum_MaechulItem) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_MaechulItemDLV) %></td>
	<td></td>
	<td></td>
	<td align="right"><%= NullOrCurrFormat(vSum_SpendMileSum) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_SpendMileSumDLV) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if (C_InspectorUser = False) then %>
    <td colspan="2"><%= NullOrCurrFormat(vSum_TotOrgitemcost + vSum_TotOrgitemcostDLV) %></td>
    <td colspan="2"><%= NullOrCurrFormat((vSum_TotOrgitemcost - vSum_TotItemcostCouponNotApplied) + (vSum_TotOrgitemcostDLV - vSum_TotItemcostCouponNotAppliedDLV)) %></td>
    <td colspan="2"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied + vSum_TotItemcostCouponNotAppliedDLV) %></td>
    <td colspan="2"><%= NullOrCurrFormat((vSum_TotItemcostCouponNotApplied - vSum_TotItemcost) + (vSum_TotItemcostCouponNotAppliedDLV - vSum_TotItemcostDLV)) %></td>
    <td colspan="2"><%= NullOrCurrFormat(vSum_TotItemcost + vSum_TotItemcostDLV) %></td>
    <td colspan="4"><%= NullOrCurrFormat((vSum_TotItemcost-vSum_TotReducePrice) + (vSum_TotItemcostDLV-vSum_TotReducePriceDLV) + vSum_SpendCouponSum+ vSum_SpendCouponSumDLV) %></td>
	<% end if %>
    <td colspan="2"><%= NullOrCurrFormat(vSum_MaechulItem + vSum_MaechulItemDLV) %></td>
    <td></td>
    <td></td>
    <td colspan="2"><%= NullOrCurrFormat(vSum_SpendMileSum + vSum_SpendMileSumDLV) %></td>
</tr>
<% if (C_InspectorUser = False) then %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" rowspan="2">
	������
	</td>
	<td align="center" rowspan="2"></td>
	<td align="right" colspan="2" rowspan="2">�Һ񰡴��=&gt</td>
	<td align="center">
	<% if vSum_TotOrgitemcost<>0 then %>
	    <%= CLNG((vSum_TotOrgitemcost-vSum_TotItemcostCouponNotApplied)/vSum_TotOrgitemcost*100*100)/100 %> %
	<% end if %>
	</td>
	<td align="center">
	<% if vSum_TotOrgitemcostDLV<>0 then %>
	    <%= CLNG((vSum_TotOrgitemcostDLV-(vSum_TotOrgitemcostDLV-vSum_TotItemcostCouponNotAppliedDLV))/vSum_TotOrgitemcostDLV*100*100)/100 %> %
	<% end if %>
	</td>
	<td align="right" colspan="2" rowspan="2">�ǸŰ����=&gt</td>
	<td align="center">
	<% if vSum_TotItemcostCouponNotApplied<>0 then %>
	    <%= CLNG((vSum_TotItemcostCouponNotApplied-vSum_TotItemcost)/vSum_TotItemcostCouponNotApplied*100*100)/100 %> %
	<% end if %>
	</td>
	<td align="center">
	<% if vSum_TotItemcostCouponNotAppliedDLV<>0 then %>
	    <%= CLNG((vSum_TotItemcostCouponNotAppliedDLV-vSum_TotItemcostDLV)/vSum_TotItemcostCouponNotAppliedDLV*100*100)/100 %> %
	<% end if %>
	</td>
	<td align="right" colspan="2" rowspan="2"></td>

	<td align="right" colspan="4" rowspan="2"></td>
	<td align="right" colspan="2" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" colspan="2" rowspan="2"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="2">
    <% if (vSum_TotOrgitemcost+vSum_TotOrgitemcostDLV)<>0 then %>
        <%= CLNG(((vSum_TotOrgitemcost+vSum_TotOrgitemcostDLV)-(vSum_TotItemcostCouponNotApplied+(vSum_TotOrgitemcostDLV-vSum_TotItemcostCouponNotAppliedDLV)))/(vSum_TotOrgitemcost+vSum_TotOrgitemcostDLV)*100*100)/100 %> %
    <% end if %>
    </td>
    <td colspan="2">
    <% if (vSum_TotItemcostCouponNotApplied+vSum_TotItemcostCouponNotAppliedDLV)<>0 then %>
        <%= CLNG(((vSum_TotItemcostCouponNotApplied+vSum_TotItemcostCouponNotAppliedDLV)-(vSum_TotItemcost+vSum_TotItemcostDLV))/(vSum_TotItemcostCouponNotApplied+vSum_TotItemcostCouponNotAppliedDLV)*100*100)/100 %> %
    <% end if %>
    </td>
</tr>
<% end if %>
</table>

<% set Omaechul_list = nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

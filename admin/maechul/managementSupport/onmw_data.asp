<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
const jErrShow = true
	dim bancancle,accountdiv,sitename,ipkumdatesucc, vPurchasetype, vatinclude
	dim yyyy1,yyyy2,mm1,mm2
	dim i ,defaultdate,defaultdate1 , olddata
    dim mdgbn, targetGbn, dlvdiv, sitegrp, vbizsec
    dim supptype

	sitename = request("sitenamebox")
	accountdiv = request("accountdiv")
	vPurchasetype = request("purchasetype")
	bancancle = NullFillWith(request("bancancle"), "1")
	vatinclude = request("vatinclude")
    supptype   = request("supptype")
	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 60�������� �˻�
	yyyy1 = NullFillWith(request("yyyy1"), left(defaultdate1,4))
	mm1 = NullFillWith(request("mm1"), mid(defaultdate1,6,2))
	yyyy2 = NullFillWith(request("yyyy2"), year(now))
	mm2 = NullFillWith(request("mm2"), month(now))
	mm2 = TwoNumber(mm2)
    mdgbn = NullFillWith(request("mdgbn"),"m")
    targetGbn = NullFillWith(request("targetGbn"),"")
    dlvdiv = NullFillWith(request("dlvdiv"),"")
    sitegrp = NullFillWith(request("sitegrp"),"")
    vbizsec = NullFillWith(request("bizsec"),"")

	dim Omaechul_list
	set Omaechul_list = new cManagementSupportMaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-01"
	Omaechul_list.FRectEndDate = CStr(DateAdd("d",-1,DateAdd("m",1,yyyy2 & "-" & mm2 & "-01")))
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc
	Omaechul_list.frectpurchasetype = vPurchasetype
	Omaechul_list.frectvatinclude = vatinclude

	Omaechul_list.FRectOnOff = targetGbn
	Omaechul_list.FRectDLVdiv = dlvdiv
	Omaechul_list.frectGroupByMwDiv="on"
	Omaechul_list.frectGroupByMonth=mdgbn
	Omaechul_list.frectGroupBySitename=sitegrp
	Omaechul_list.FRectBizSectionCd=vbizsec
    Omaechul_list.FRectSupptype = supptype
	Omaechul_list.fmaechul_listByGbn()


	Dim vSum_TotItemNo, vSum_TotReducedPrice, vSum_TotBuycash, vSum_TotBuycashCouponNotApplied
	Dim vSum_TotOrgitemcost, vSum_TotItemcostCouponNotApplied,  vSum_TotItemcost
	Dim vSum_TotReducePrice, vSum_SpendCouponSum, vSum_MaechulItem
	Dim vSum_SpendMileSum
	Dim vSum_jPrice,vSum_jPriceEtc,vSum_jPriceEtcChulgo
    Dim vSum_HanDlePrice , vSum_CalcuMeachul, vSum_CalcuMeachulNoVat, vSum_ErrJungsan

%>
<h3>������(��ۺ�)</h3>
<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">��ǰ����� / ��¥ <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
			&nbsp;&nbsp;&nbsp;
			<!--
			<input type="radio" name="mdgbn" value="m" <%= CHKIIF(mdgbn="m","checked","") %> >����
			<input type="radio" name="mdgbn" value="d" <%= CHKIIF(mdgbn="d","checked","") %> disabled >�Ϻ�
			-->
			* �⺻ ����μ� :
			<% Call DrawBizSectionGain("O,T,C","bizsec", vbizsec,"") %>

			&nbsp;&nbsp;&nbsp;
			��ǰ�ͼ�
			<select name="targetGbn">
			<option value="" <%=CHKIIF(targetGbn="","selected","") %> >��ü
			<option value="ON" <%=CHKIIF(targetGbn="ON","selected","") %> >�¶���
			<option value="IT" <%=CHKIIF(targetGbn="IT","selected","") %> >���̶��_�¶���
			<option value="AC" <%=CHKIIF(targetGbn="AC","selected","") %> >��ī����
			<option value="NOAC" <%=CHKIIF(targetGbn="NOAC","selected","") %> >�¶���,���̶��_�¶���
			<!--
			<option value="OF" <%=CHKIIF(targetGbn="OF","selected","") %> >��������
			<option value="OT" <%=CHKIIF(targetGbn="OT","selected","") %> >���̶��_��������
			-->
			</select>

			&nbsp;&nbsp;&nbsp;
			���Ա���
			<select name="dlvdiv">
			<option value="" <%=CHKIIF(dlvdiv="","selected","") %> >��ü
			<option value="s" <%=CHKIIF(dlvdiv="s","selected","") %> >��ǰ(����+Ư��+��ü)
			<option value="d" <%=CHKIIF(dlvdiv="d","selected","") %> >��ۺ�(����+�ٹ�)
			<option value="M" <%=CHKIIF(dlvdiv="M","selected","") %> >����
			<option value="W" <%=CHKIIF(dlvdiv="W","selected","") %> >Ư��
			<option value="U" <%=CHKIIF(dlvdiv="U","selected","") %> >��ü
			<option value="Y" <%=CHKIIF(dlvdiv="Y","selected","") %> >����
			<option value="Z" <%=CHKIIF(dlvdiv="Z","selected","") %> >�ٹ�
			</select>
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
	    	���ó���� : <% Drawsitename "sitenamebox",sitename %>
	    	<input type="radio" name="sitegrp" value="" <%= CHKIIF(sitegrp="","checked","") %> >�հ�
			<input type="radio" name="sitegrp" value="g" <%= CHKIIF(sitegrp="g","checked","") %>  >���ó��
	    	&nbsp;&nbsp;&nbsp;
	    	�������� : <% drawPartnerCommCodeBox true,"selljungsantype","purchasetype",vPurchasetype,"" %>

	    	&nbsp;&nbsp;&nbsp;
	    	/   <input type="radio" name="supptype" value="S" <%= CHKIIF(supptype="S","checked","") %> > ���ް���
	    	    <input type="radio" name="supptype" value="" <%= CHKIIF(supptype="","checked","") %> > �հ�ݾ�
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
<table width="100%" class="a">
<tr bgcolor="#FFFFFF">
    <td>
        ����(M) : �¶��� ����,�ٹ� = ��޾�, �¶��� Ư��,��ü,���� = ��޾�-���԰�(������), ���̶�� = ��޾�,
    </td>
</tr>
</table>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2" align="center" width="70" >���<%=CHKIIF(mdgbn="m","��","��") %></td>
	<% if (sitegrp<>"") then %>
	<td rowspan="2" align="center" width="60" >���ó</td>
	<td rowspan="2" align="center" width="60" >��������</td>
	<% end if %>
	<td rowspan="2" align="center" width="60" >����<br>�μ�</td>
	<td rowspan="2" align="center" width="40" >��ǰ<br>�ͼ�</td>
	<td rowspan="2" align="center" width="40" >����<br>����</td>
    <td rowspan="2" align="center" width="50" >�ѻ�ǰ<br>����</td>
	<% if (C_InspectorUser = False) then %>
    <td rowspan="2" align="center" >�Һ��ڰ�<br>A</td>
    <td rowspan="2" align="center" >���αݾ�<br>B</td>
    <td rowspan="2" align="center" >�ǸŰ�(���ΰ�)<br>C=A-B</td>
    <td rowspan="2" align="center" >��ǰ��������<br>D</td>
    <td rowspan="2" align="center" >�����Ѿ�<br>E=C-D</td>
    <td align="center" colspan="2">���ʽ�����<br>��������(F)=E-ȯ�Ҿ�(reducePrice)<br>��������(G)</td>
    <td rowspan="2" align="center" >��޾�<br>(H)=E-F-G</td>
    <td rowspan="2" align="center" width="5" ></td>
    <!--<td rowspan="2" align="center" >���ϸ���<br>���Ⱥ�</td>-->
    <td rowspan="2" align="center" >��޾׿���<br>(�ֹ��ø��԰�)<br>(S)</td>
    <td rowspan="2" align="center" >��޾�<br>������(%)<br>S/H</td>
	<% end if %>
    <td rowspan="2" align="center" >����(M)</td>
    <td rowspan="2" align="center" >����<br>(vat����)</td>
    <td rowspan="2" align="center" width="5" ></td>
    <td rowspan="2" align="center" >�����<br>(J1)</td>
    <td rowspan="2" align="center" >��Ÿ����<br>(��ǰ��ۺ��)</td>
    <td rowspan="2" align="center" >��Ÿ�������<br>(����,�ν���)</td>
    <% if (jErrShow) then %>
    <td rowspan="2" align="center" >�������<br>(S-J1)</td>
    <% end if %>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td>��������(F)</td>
	<td>��������(G)<br>�Ⱥ�</td>
	<% end if %>
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
	<% if (sitegrp<>"") then %>
	<td align="center"><%= Omaechul_list.flist(i).fsitename %></td>
	<td align="center"><%= Omaechul_list.flist(i).fsellTypeName %></td>
	<% end if %>
	<td align="center"><%= Omaechul_list.flist(i).fsellBizCdName %></td>
	<td align="center"><%= Omaechul_list.flist(i).getItemGubunName %></td>
	<td align="center"><%= Omaechul_list.flist(i).getMwGubunName %></td>
    <td align="center"><%= Replace(Omaechul_list.flist(i).ftot_itemno,"-","") %></td>
	<% if (C_InspectorUser = False) then %>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost - Omaechul_list.flist(i).ftot_itemcostCouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied) %></td>
    <td align="right">
        <% if (Omaechul_list.flist(i).ftot_itemcostCouponNotApplied - Omaechul_list.flist(i).ftot_itemcost<0) then %>
        <font color=red><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied - Omaechul_list.flist(i).ftot_itemcost) %></font>
        <% else %>
        <%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied - Omaechul_list.flist(i).ftot_itemcost) %>
        <% end if %>
    </td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendCouponSum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getHanDlePrice) %></td>
	<td align="center" >
	<!--
	[<a href="/admin/upchejungsan/upcheselllist.asp?datetype=chulgoil&yyyy1=<%=vYear%>&mm1=<%=vMonth%>&dd1=<%=vDay%>&yyyy2=<%=vYear%>&mm2=<%=vMonth%>&dd2=<%=vDay%>&delivertype=all" target="_blank">��</a>]
	-->
	</td>
	<!--<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendMileSum) %></td>-->
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_buycash) %></td>
	<td align="center">
	<% if (Omaechul_list.flist(i).getHanDlePrice<>0) then %>
	<%= CLNG(Omaechul_list.flist(i).ftot_buycash/Omaechul_list.flist(i).getHanDlePrice*100*100)/100 %>
	<% else %>
	-
	<% end if %>
	</td>
	<% end if %>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getCalcuMeachul) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getCalcuMeachulNoVat) %></td>
	<td align="right"></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fjPrice) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fjPriceEtc) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fjPriceEtcChulgo) %></td>
	<% if (jErrShow) then %>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getErrJungsan) %></td>
	<% end if %>
</tr>
<%
	vSum_TotItemNo 						= vSum_TotItemNo + Omaechul_list.flist(i).ftot_itemno
	vSum_TotReducedPrice 				= vSum_TotReducedPrice + Omaechul_list.flist(i).ftot_reducedPrice
	vSum_TotBuycash 					= vSum_TotBuycash + Omaechul_list.flist(i).ftot_buycash
	vSum_TotBuycashCouponNotApplied 	= vSum_TotBuycashCouponNotApplied + Omaechul_list.flist(i).ftot_buycashCouponNotApplied
	vSum_TotOrgitemcost 				= vSum_TotOrgitemcost + Omaechul_list.flist(i).ftot_orgitemcost
	vSum_TotItemcostCouponNotApplied 	= vSum_TotItemcostCouponNotApplied + Omaechul_list.flist(i).ftot_itemcostCouponNotApplied
	vSum_TotItemcost 					= vSum_TotItemcost + Omaechul_list.flist(i).ftot_itemcost
	vSum_TotReducePrice					= vSum_TotReducePrice + Omaechul_list.flist(i).ftot_reducedPrice
	vSum_SpendCouponSum					= vSum_SpendCouponSum + Omaechul_list.flist(i).ftot_DivSpendCouponSum
	vSum_MaechulItem					= vSum_MaechulItem + (Omaechul_list.flist(i).ftot_itemcost-(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice)-Omaechul_list.flist(i).ftot_DivSpendCouponSum)

	vSum_SpendMileSum					= vSum_SpendMileSum + Omaechul_list.flist(i).ftot_DivSpendMileSum

	vSum_jPrice                         = vSum_jPrice + Omaechul_list.flist(i).fjPrice
	vSum_jPriceEtc                      = vSum_jPriceEtc + Omaechul_list.flist(i).fjPriceEtc
	vSum_jPriceEtcChulgo                = vSum_jPriceEtcChulgo + Omaechul_list.flist(i).fjPriceEtcChulgo

	vSum_HanDlePrice                    = vSum_HanDlePrice + Omaechul_list.flist(i).getHanDlePrice
	vSum_CalcuMeachul                   = vSum_CalcuMeachul + Omaechul_list.flist(i).getCalcuMeachul
	vSum_CalcuMeachulNoVat              = vSum_CalcuMeachulNoVat + Omaechul_list.flist(i).getCalcuMeachulNoVat
	vSum_ErrJungsan                     = vSum_ErrJungsan + Omaechul_list.flist(i).getErrJungsan
Next
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" rowspan="2">
	�Ѱ�
	</td>
	<% if (sitegrp<>"") then %>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<% end if %>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ><%= Replace(vSum_TotItemNo,"-","") %></td>
	<% if (C_InspectorUser = False) then %>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcost) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcost - vSum_TotItemcostCouponNotApplied) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied - vSum_TotItemcost) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotItemcost) %></td>

	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcost-vSum_TotReducePrice) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_SpendCouponSum) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_HanDlePrice) %></td>
	<td rowspan="2" ></td>
	<!--<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_SpendMileSum) %></td>-->
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotBuycash) %></td>
	<td rowspan="2" align="center">
	<% if (vSum_HanDlePrice<>0) then %>
	<%= CLNG(vSum_TotBuycash/vSum_HanDlePrice*100*100)/100 %>
	<% else %>
	-
	<% end if %>
	</td>
	<% end if %>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_CalcuMeachul) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_CalcuMeachulNoVat) %></td>
	<td rowspan="2" align="right"></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_jPrice) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_jPriceEtc) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_jPriceEtcChulgo) %></td>
	<% if (jErrShow) then %>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_ErrJungsan) %></td>
	<% end if %>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if (C_InspectorUser = False) then %>
    <td colspan="2"><%= NullOrCurrFormat(vSum_TotItemcost-vSum_TotReducePrice+vSum_SpendCouponSum)  %></td>
	<% end if %>
</tr>
<% if (C_InspectorUser = False) then %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" rowspan="2">
	������
	</td>
	<% if (sitegrp<>"") then %>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<% end if %>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="right" rowspan="2">�Һ񰡴��=&gt</td>
	<td align="center">
	<% if vSum_TotOrgitemcost<>0 then %>
	    <%= CLNG((vSum_TotOrgitemcost-vSum_TotItemcostCouponNotApplied)/vSum_TotOrgitemcost*100*100)/100 %> %
	<% end if %>
	</td>

	<td align="right" rowspan="2">�ǸŰ����=&gt</td>
	<td align="center">
	<% if vSum_TotItemcostCouponNotApplied<>0 then %>
	    <%= CLNG((vSum_TotItemcostCouponNotApplied-vSum_TotItemcost)/vSum_TotItemcostCouponNotApplied*100*100)/100 %> %
	<% end if %>
	</td>

	<td align="right" rowspan="2"></td>

	<td align="right" colspan="2" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<!--<td align="right" rowspan="2"></td>-->
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<% if (jErrShow) then %>
	<td align="right" rowspan="2"></td>
	<% end if %>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td >
    <% if (vSum_TotOrgitemcost)<>0 then %>
        <%= CLNG(((vSum_TotOrgitemcost)-(vSum_TotItemcostCouponNotApplied))/(vSum_TotOrgitemcost)*100*100)/100 %> %
    <% end if %>
    </td>
    <td >
    <% if (vSum_TotItemcostCouponNotApplied)<>0 then %>
        <%= CLNG(((vSum_TotItemcostCouponNotApplied)-(vSum_TotItemcost))/(vSum_TotItemcostCouponNotApplied)*100*100)/100 %> %
    <% end if %>
    </td>
</tr>
<% end if %>
</table>

<% set Omaechul_list = nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

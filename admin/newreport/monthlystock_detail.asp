<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<!--
<H1>������</H1>
<p>BSC���>>����ڻ�  �޴� �����</p>
-->
<%
''dbget.close():response.end

Const isShowIpgoPrice = FALSE
Const isOnlySys = FALSE
Const isShowOffreturn = FALSE
Dim isViewUser : isViewUser = (session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal,mwgubun,makerid,newitem,itemgubun,vatyn
dim research,offrt2on
dim minusinc, bPriceGbn,buseo
dim purchasetype, ordTp, swSppPrc

yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
isusing     = requestCheckvar(request("isusing"),10)
sysorreal   = requestCheckvar(request("sysorreal"),10)
mwgubun     = requestCheckvar(request("mwgubun"),10)
makerid     = requestCheckvar(request("makerid"),32)
newitem     = requestCheckvar(request("newitem"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
offrt2on    = requestCheckvar(request("offrt2on"),10)
research    = requestCheckvar(request("research"),10)
minusinc    = requestCheckvar(request("minusinc"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
buseo       = requestCheckvar(request("buseo"),10)
purchasetype   = requestCheckvar(request("purchasetype"),10)
ordTp       = requestCheckvar(request("ordTp"),10)
swSppPrc	= requestCheckvar(request("swSppPrc"),32)

if sysorreal="" then sysorreal="sys" ''real
if (research="") or (not isShowOffreturn) then offrt2on="on"
if (isViewUser="") then sysorreal="sys"
if (isViewUser="") then bPriceGbn="P"
if (isViewUser="") then isusing=""

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
ojaego.FRectIsUsing  = isusing
ojaego.FRectGubun    = sysorreal
ojaego.FRectMakerid  = makerid
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectNewItem  = newitem
ojaego.FRectVatYn    = vatyn
ojaego.FRectItemGubun = itemgubun
ojaego.FRectMinusInclude = minusinc
ojaego.FRectTargetGbn = buseo
ojaego.FRectPurchaseType = purchasetype
ojaego.FRectOrdTp = ordTp
ojaego.FRectShopSuplyPrice    = swSppPrc

if (buseo="IT") then
    ojaego.FRectITSOnlyOrNot = "O"
else
    ojaego.FRectITSOnlyOrNot = "N"
end if

if (bPriceGbn="P") then
    ojaego.FRectIsFix = "on"
end if

''ojaego.FRectOFFReturn2OnStock = offrt2on

if makerid<>"" then
    ojaego.GetMonthlyRealJeagoDetailByMakerWithPreMonth ''GetMonthlyRealJeagoDetailByMakerNew
else
	ojaego.GetMonthlyRealJeagoDetailWithPreMonth

end if

dim i
dim totno, totbuy, subTotno, subTotbuy
dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy

dim iURL
%>
<script language='javascript'>
function TnPopItemStockWithGubun(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function TnPopItemStockModifyMW(itemgubun,itemid,itemoption) {
	var popwin = window.open("pop_item_stock_edit.asp?yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"TnPopItemStockModifyMW","width=600 height=300 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
			&nbsp;&nbsp;
			<font color="#CC3333">�귣�� :</font> <% drawSelectBoxDesignerwithName "makerid",makerid %>
			<!--
			&nbsp;&nbsp;
			<input type="radio" name="newitem" value="" <% if newitem="" then response.write "checked" %> >��ü��ǰ
        	<input type="radio" name="newitem" value="new" <% if newitem="new" then response.write "checked" %> >�Ż�ǰ
        	-->
        	&nbsp;&nbsp;|&nbsp;&nbsp;
	        	��������
	        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >��ü
	        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >����
	        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >�鼼
	        	&nbsp;&nbsp;
			    <input type="checkbox" name="swSppPrc" value="Y" <%= CHKIIF(swSppPrc="Y","checked","") %> >���ް��� ǥ��
		</td>

		<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <% IF not (isOnlySys) THEN %>
		    <% if (Not isViewUser) then %>
			<font color="#CC3333">�����:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >�ý������
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >�ǻ����
        	&nbsp;&nbsp;
        	<% end if %>
        	<% end if %>

        	<% if (Not isViewUser) then %>
        	<font color="#CC3333">��ǰ��뱸��:</font>
        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >��ü
        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >�����
        	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >������
        	&nbsp;&nbsp;
        	<% end if %>

        	<font color="#CC3333">���Ա���:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >����
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >��Ź
        	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >��ü -->
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >������

        	<% if (isShowOffreturn) then %>
            <br><input type="checkbox" name="offrt2on" <%= CHKIIF(offrt2on="on","checked","") %> >�����ǰOn����
            <% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<font color="#CC3333">���̳ʽ�����:</font>
		<input type="radio" name="minusinc" value="" <%= CHKIIF(minusinc="","checked","") %> >���̳ʽ���� ����(��ü)
		<input type="radio" name="minusinc" value="P" <%= CHKIIF(minusinc="P","checked","") %> >(+)���
	    <input type="radio" name="minusinc" value="M" <%= CHKIIF(minusinc="M","checked","") %> >���̳ʽ���� ��
	    &nbsp;&nbsp;
	    <% if (Not isViewUser) then %>
	    <font color="#CC3333">���԰�����:</font>
	    <input type="radio" name="bPriceGbn" value="" <%= CHKIIF(bPriceGbn="","checked","") %> >������԰�
	    <input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %> >�ۼ��ø��԰�
	    <input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %> disabled >��ո��԰�
	    <% end if %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	    <font color="#CC3333">�μ�����:</font>
	        <select name="buseo">
        	<option value="ON" <%= CHKIIF(buseo="ON","selected" ,"") %> >�¶���
        	<option value="OF" <%= CHKIIF(buseo="OF","selected" ,"") %> >��������
        	<option value="IT" <%= CHKIIF(buseo="IT","selected" ,"") %> >���̶��(��)
        	<option value="ET" <%= CHKIIF(buseo="ET","selected" ,"") %> >���
        	<option value="EG" <%= CHKIIF(buseo="EG","selected" ,"") %> >���׷���
        	</select>
	    <font color="#CC3333">��ǰ����:</font>
        	<select name="itemgubun">
        	<option value="10" <%= CHKIIF(itemgubun="10","selected" ,"") %> >�Ϲ�(10)
        	<option value="55" <%= CHKIIF(itemgubun="55","selected" ,"") %> >CS��Ÿ����(55)
        	<option value="70" <%= CHKIIF(itemgubun="70","selected" ,"") %> >����ǰ(70)
        	<option value="75" <%= CHKIIF(itemgubun="75","selected" ,"") %> >������(75)
        	<option value="85" <%= CHKIIF(itemgubun="85","selected" ,"") %> >����ǰ(85)
        	<option value="80" <%= CHKIIF(itemgubun="80","selected" ,"") %> >����ǰ(80)
        	<option value="90" <%= CHKIIF(itemgubun="90","selected" ,"") %> >��������(90)
        	</select>

        	&nbsp;
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	        <font color="#CC3333">���ı���:</font>
	        <% if makerid<>"" then %>
	        <input type="radio" name="ordTp" value="" <%= CHKIIF(ordTp="","checked","") %> >��ǰ�ڵ�
	        <input type="radio" name="ordTp" value="S" <%= CHKIIF(ordTp="S","checked","") %> >������(�⸻)
	        <% else %>
	        <input type="radio" name="ordTp" value="" <%= CHKIIF(ordTp="","checked","") %> >�⺻
	        <% end if %>
	    </td>
	</tr>

	</form>
</table>
<!-- �˻� �� -->

<p>

* ������� : <font color="red">������� ����</font> ǥ�þȵ�.

<p>

<% if makerid<>"" then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="8">��ǰ����</td>
        <td colspan="2">�������(��������)<br>A</td>
        <td colspan="2">�������(��)<br>B</td>
        <td colspan="2">�⸻���(��������)<br>C</td>
        <td colspan="2">�Ѹ������<br>D=A+B-C</td>
        <td width="1" bgcolor="#FFFFFF"></td>
        <td colspan="2">���LOSS<br>E</td>
        <td colspan="2">��ǰ�������<br>F=A+B+E-C</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="65">�����</td>
    	<td width="40"></td>
    	<td>����</td>
    	<td width="50">��ǰ�ڵ�</td>
    	<td width="40">�ɼ�<br>�ڵ�</td>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td >����<br>����</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td  bgcolor="#FFFFFF"></td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    </tr>

    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum

    totPreno = totPreno + ojaego.FItemList(i).FTotPreCount
    totPrebuy= totPrebuy + ojaego.FItemList(i).FTotPreBuySum

    totIpno  = totIpno + ojaego.FItemList(i).FTotIpCount
    totIpBuy = totIpBuy + ojaego.FItemList(i).FTotIpBuySum

    totLossno  = totLossno + ojaego.FItemList(i).FTotLossCount
    totLossBuy = totLossBuy + ojaego.FItemList(i).FTotLossBuySum

    subTotno    = subTotno + ojaego.FItemList(i).FTotCount
    subTotbuy   = subTotbuy + ojaego.FItemList(i).FTotBuySum

    subPreno    = subPreno + ojaego.FItemList(i).FTotPreCount
    subPrebuy   = subPrebuy + ojaego.FItemList(i).FTotPreBuySum
    subIpno     = subIpno + ojaego.FItemList(i).FTotIpCount
    subIpBuy    = subIpBuy + ojaego.FItemList(i).FTotIpBuySum
    subLossno   = subLossno + ojaego.FItemList(i).FTotLossCount
    subLossBuy  = subLossBuy + ojaego.FItemList(i).FTotLossBuySum


    %>
    <% if ((ojaego.FItemList(i).FIsUsing="N") or (ojaego.FItemList(i).FOptionUsing="N")) then %>
    <tr align="center" bgcolor="<%= adminColor("dgray") %>">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
    	<td><%= left(ojaego.FItemList(i).Fregdate,10) %></td>
    	<td>
    		<% if (datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD)) <= 3 then %>
    		<font color="red"><%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>����</font>
    		<% else %>
    		<%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>����
    		<% end if %>
    	</td>
    	<td><%= ojaego.FItemList(i).FItemGubun %></td>
    	<% if ( isViewUser) then %>
    	<td><%= ojaego.FItemList(i).FItemID %></td>
    	<% else %>
    	<td><a href="javascript:TnPopItemStockWithGubun('<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>');"><%= ojaego.FItemList(i).FItemID %></a></td>
    	<% end if %>
    	<td><%= ojaego.FItemList(i).FItemOption %></td>
    	<td align="left"><%= ojaego.FItemList(i).FItemName %></td>
    	<td><%= ojaego.FItemList(i).FItemOptionName %></td>
    	<td>
    		<% if (ojaego.FItemList(i).FMaeIpGubun <> "Z") then %>
    			<%= ojaego.FItemList(i).getMaeipGubunName %>
    		<% else %>
    		    <% if ( isViewUser) then %>
    		    -
    		    <% else %>
    			<a href="javascript:TnPopItemStockModifyMW('<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>')">-</a>
    			<% end if %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotPreCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotPreBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotIpCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotIpBuySum,0) %></td>
        <td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getWongaCnt,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getWongaSum,0) %></td>
    	<td ></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaCnt,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaSum,0) %></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totPreno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totIpno,0) %></td>
    	<td align="right" ><%= FormatNumber(totIpBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totPreno+totIpno-totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy+totIpBuy-totbuy,0) %></td>
    	<td ></td>
    	<td align="right" ><%= FormatNumber(totLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(totLossBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totPreno+totIpno-totno+totLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy+totIpBuy-totbuy+totLossBuy,0) %></td>
    </tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td >�귣��</td>
        <td colspan="2">�������(��������)<br>A</td>
        <td colspan="2">�������(��)<br>B</td>
        <td colspan="2">�⸻���(��������)<br>C</td>
        <td colspan="2">�Ѹ������<br>D=A+B-C</td>
        <td width="1" bgcolor="#FFFFFF"></td>
        <td colspan="2">���LOSS<br>E</td>
        <td colspan="2">��ǰ�������<br>F=A+B+E-C</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td >�귣��ID</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td  bgcolor="#FFFFFF"></td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    	<td >����</td>
    	<td >�ݾ�(���԰�)</td>
    </tr>

    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum

    totPreno = totPreno + ojaego.FItemList(i).FTotPreCount
    totPrebuy= totPrebuy + ojaego.FItemList(i).FTotPreBuySum

    totIpno  = totIpno + ojaego.FItemList(i).FTotIpCount
    totIpBuy = totIpBuy + ojaego.FItemList(i).FTotIpBuySum

    totLossno  = totLossno + ojaego.FItemList(i).FTotLossCount
    totLossBuy = totLossBuy + ojaego.FItemList(i).FTotLossBuySum

    subTotno    = subTotno + ojaego.FItemList(i).FTotCount
    subTotbuy   = subTotbuy + ojaego.FItemList(i).FTotBuySum

    subPreno    = subPreno + ojaego.FItemList(i).FTotPreCount
    subPrebuy   = subPrebuy + ojaego.FItemList(i).FTotPreBuySum
    subIpno     = subIpno + ojaego.FItemList(i).FTotIpCount
    subIpBuy    = subIpBuy + ojaego.FItemList(i).FTotIpBuySum
    subLossno   = subLossno + ojaego.FItemList(i).FTotLossCount
    subLossBuy  = subLossBuy + ojaego.FItemList(i).FTotLossBuySum


    iURL = "monthlystock_detail.asp?menupos="& menupos &"&mwgubun="& mwgubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&makerid="& ojaego.FItemList(i).FMakerid &"&newitem="& newitem &"&itemgubun="& itemgubun&"&vatyn="&vatyn
    iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&buseo&"&swSppPrc="&swSppPrc
    if Not(isOnlySys) THEN iURL=iURL&"&sysorreal="& sysorreal
    %>
    <% if (ojaego.FItemList(i).FMakerUsing="Y") then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
    <% end if %>
    	<td align="left"><a href="<%= iURL %>" ><%= ojaego.FItemList(i).FMakerid %></a></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotPreCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotPreBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotIpCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotIpBuySum,0) %></td>
        <td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getWongaCnt,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getWongaSum,0) %></td>
    	<td ></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaCnt,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaSum,0) %></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td align="right" ><%= FormatNumber(totPreno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totIpno,0) %></td>
    	<td align="right" ><%= FormatNumber(totIpBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totPreno+totIpno-totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy+totIpBuy-totbuy,0) %></td>
    	<td ></td>
    	<td align="right" ><%= FormatNumber(totLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(totLossBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totPreno+totIpno-totno+totLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy+totIpBuy-totbuy+totLossBuy,0) %></td>
    </tr>
    </tr>
</table>

<% end if %>

<%
set ojaego = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

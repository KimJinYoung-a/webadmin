<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
Const isShowIpgoPrice = FALSE
Dim isShowSysWithReal : isShowSysWithReal = FALSE
Dim isViewUser : isViewUser = (session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal,mwgubun,makerid,newitem,shopid,showminus, itemmwdiv, buseo
dim vatyn, showSupplyPrice, ordTp
dim showminusOnly

yyyy1       = RequestCheckVar(request("yyyy1"),10)
mm1         = RequestCheckVar(request("mm1"),10)
isusing     = RequestCheckVar(request("isusing"),10)
sysorreal   = RequestCheckVar(request("sysorreal"),10)
mwgubun     = RequestCheckVar(request("mwgubun"),10)
makerid     = RequestCheckVar(request("makerid"),32)
newitem     = RequestCheckVar(request("newitem"),10)
shopid      = RequestCheckVar(request("shopid"),32)
showminus   = RequestCheckVar(request("showminus"),32)
itemmwdiv   = RequestCheckVar(request("itemmwdiv"),32)
vatyn       = requestCheckvar(request("vatyn"),10)
showSupplyPrice 	= requestCheckvar(request("showSupplyPrice"),10)
buseo   			= RequestCheckVar(request("buseo"),32)
ordTp   			= RequestCheckVar(request("ordTp"),10)
showminusOnly       = requestCheckvar(request("showminusOnly"),10)

if (isViewUser) then showminus=""
if (isViewUser) then showminusOnly=""
if sysorreal="" then sysorreal="sys"
if (isViewUser) then sysorreal="sys"
if (isViewUser) then isusing=""

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FPageSize = 3000
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
ojaego.FRectIsUsing  = isusing
ojaego.FRectGubun    = sysorreal
ojaego.FRectMakerid  = makerid
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectNewItem  = newitem
ojaego.FRectShopID = shopid
ojaego.FRectShowMinus = showminus
ojaego.FRectShowMinusOnly = showminusOnly
ojaego.FRectItemMwDiv = itemmwdiv
ojaego.FRectVatYn    = vatyn
ojaego.FRectShopSuplyPrice    = showSupplyPrice
ojaego.FRectTargetGbn = buseo
ojaego.FRectOrdTp = ordTp

if (makerid<>"") and (shopid<>"") then
    IF (isShowSysWithReal) then
        ojaego.FRectGubun = "sys"
        ojaego.GetShopMonthlyRealJeagoDetailByMakerSysWithReal
    else
	    ojaego.GetShopMonthlyRealJeagoDetailByMakerNew
    end if
elseif (makerid="") and (shopid<>"") then
    IF (isShowSysWithReal) then
        ojaego.FRectGubun = "sys"
        ojaego.GetShopMonthlyRealJeagoDetailSysWithReal
    else
	    ojaego.GetShopMonthlyRealJeagoDetailNew
	end if
else
    IF (isShowSysWithReal) then
        ojaego.FRectGubun = "sys"
        ojaego.GetShopMonthlyRealJeagoDetailByShopidSysWithReal
    else
	    ojaego.GetShopMonthlyRealJeagoDetailByShopidNew
	end if
end if

dim i
dim totno, totbuy, totShopbuy, totsell, totavgIpgoPrice
dim totRealno, totRealbuy, totRealsell
dim iURI

dim CLDiv : CLDiv = "L"
if Left(Now(), 7) = (yyyy1 + "-" + mm1) then
	CLDiv = "C"
end if

%>
<script language='javascript'>
<% if (Not isViewUser) then %>
function TnPopOffItemStock(shopid,itemgubun,itemid,itemoption){
	//var popwin = window.open("/admin/stock/itemcurrentstock_shop.asp?menupos=709&shopid="+shopid+"&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	var popwin = window.open("/common/offshop/shop_itemcurrentstock.asp?shopid="+shopid+"&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"PopOffItemStock","width=1280 height=960 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popAssignCommCD(imakerid,iyyyymm,ishopid){
    var iURL = "popAssignMonthlyCommCd.asp?makerid=" + imakerid+"&yyyymm="+iyyyymm+"&shopid="+ishopid
    var popwin = window.open(iURL,'popAssignMonthlyCommCd','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}
<% end if %>
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
			&nbsp;&nbsp;
			���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<font color="#CC3333">�μ�����:</font>
	        <select name="buseo">
	        <option value="" <%= CHKIIF(buseo="","selected" ,"") %> >��ü
        	<option value="OF" <%= CHKIIF(buseo="OF","selected" ,"") %> >��������
        	<option value="IT" <%= CHKIIF(buseo="IT","selected" ,"") %> >���̶��(��)
        	<option value="ET" <%= CHKIIF(buseo="ET","selected" ,"") %> >���
        	<option value="EG" <%= CHKIIF(buseo="EG","selected" ,"") %> >���׷���
        	</select>
        	&nbsp;&nbsp;
			<font color="#CC3333">�귣�� :</font> <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;&nbsp;
			<input type="radio" name="newitem" value="" <% if newitem="" then response.write "checked" %> >��ü��ǰ
        	<input type="radio" name="newitem" value="new" <% if newitem="new" then response.write "checked" %> >�Ż�ǰ
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <% IF Not (isShowSysWithReal) THEN %>
		    <% if (Not isViewUser) then %>
			<font color="#CC3333">�����:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >�ý������
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >�ǻ����
        	&nbsp;&nbsp;
        	<% end if %>
        	<% END IF %>

			<% if (Not isViewUser) then %>
        	<font color="#CC3333">��ǰ��뱸��:</font>
        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >��ü
        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >�����
        	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >������
        	&nbsp;&nbsp;
        	<font color="#CC3333">��ǰ���Ա���:</font>
        	<input type="radio" name="itemmwdiv" value="" <% if itemmwdiv="" then response.write "checked" %> >��ü
        	<input type="radio" name="itemmwdiv" value="M" <% if itemmwdiv="M" then response.write "checked" %> <% if (makerid="") or (shopid = "") then %>disabled<% end if %> >����
        	<input type="radio" name="itemmwdiv" value="W" <% if itemmwdiv="W" then response.write "checked" %> <% if (makerid="") or (shopid = "") then %>disabled<% end if %> >��Ź
			<input type="radio" name="itemmwdiv" value="Z" <% if itemmwdiv="Z" then response.write "checked" %> <% if (makerid="") or (shopid = "") then %>disabled<% end if %> >������
        	&nbsp;&nbsp;
        	<font color="#CC3333">��������</font>
        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >��ü
        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >����
        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >�鼼
        	&nbsp;&nbsp;
			<input type="checkbox" name="showSupplyPrice" value="Y" <%= CHKIIF(showSupplyPrice="Y","checked","") %> >���ް��� ǥ��
			<br>
        	<% END IF %>
        	<font color="#CC3333">��౸��:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
        	<input type="radio" name="mwgubun" value="B011" <% if mwgubun="B011" then response.write "checked" %> >��Ź�Ǹ�
        	<input type="radio" name="mwgubun" value="B012" <% if mwgubun="B012" then response.write "checked" %> >��ü��Ź
        	<input type="radio" name="mwgubun" value="B013" <% if mwgubun="B013" then response.write "checked" %> >�����Ź
        	<input type="radio" name="mwgubun" value="B022" <% if mwgubun="B022" then response.write "checked" %> >�������
        	<input type="radio" name="mwgubun" value="B031" <% if mwgubun="B031" then response.write "checked" %> >������
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >������
        	<% if (Not isViewUser) then %>
        	<br>
        	<input type="checkbox" name="showminus" <%= CHKIIF(showminus="on","checked","") %> >���̳ʽ���� ����
			<input type="checkbox" name="showminusOnly" <%= CHKIIF(showminusOnly="on","checked","") %> >���̳ʽ����
        	<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	        <font color="#CC3333">���ı���:</font>
	        <% if makerid<>"" then %>
	        <input type="radio" name="ordTp" value="" <%= CHKIIF(ordTp="","checked","") %> >��ǰ�ڵ�
	        <input type="radio" name="ordTp" value="S" <%= CHKIIF(ordTp="S","checked","") %> >������
	        <% else %>
	        <input type="radio" name="ordTp" value="" <%= CHKIIF(ordTp="","checked","") %> >�⺻
	        <% end if %>
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p>
<% if (makerid<>"") and (shopid<>"") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
    <td>�ִ� <%= ojaego.FPageSize %>�� ���� ǥ�õ˴ϴ�.</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <% if (Not isViewUser) then %>
    	<td width="65" rowspan="2">�����</td>
    	<td width="40" rowspan="2"></td>
    	<% end if %>
    	<% if (FALSE) then %>
        	<td width="50" rowspan="2">����</td>
        	<td width="50" rowspan="2">��ǰ�ڵ�</td>
        	<td width="50" rowspan="2">�ɼ��ڵ�</td>
    	<% else %>
    	    <td width="100">�����ڵ�</td>
    	<% end if %>
    	<td width="100">�����ڵ�</td>
    	<td rowspan="2">��ǰ��</td>
    	<td rowspan="2">�ɼǸ�</td>
    	<td width="35" rowspan="2">��ǰ<br>�Ӽ�</td>
		<td width="50" rowspan="2">���<br>����</td>
    	<td width="50" rowspan="2">(��)�ǸŴܰ�</td>
    	<td colspan="4">�ý������</td>
    	<td width="50" rowspan="2">����</td>
    	<td colspan="3">�ǻ����</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">������</td>
    	<td width="80">�ǸŰ�</td>
    	<td width="80">���԰�(����)</td>
    	<td width="50">���Ը���</td>
    	<td width="50">������</td>
    	<td width="80">�ǸŰ�</td>
    	<td width="80">���԰�(����)</td>
    </tr>
    <% else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <% if (Not isViewUser) then %>
    	<td width="65">�����</td>
    	<td width="40"></td>
    	<% end if %>
    	<% if (FALSE) then %>
        	<td width="50">����</td>
        	<td width="50">��ǰ�ڵ�</td>
        	<td width="50">�ɼ��ڵ�</td>
    	<% else %>
    	    <td width="100">�����ڵ�</td>
    	<% end if %>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td width="35">��ǰ<br>�Ӽ�</td>
    	<td width="50">���<br>����</td>
    	<td width="50">������</td>
    	<td width="50">(��)�ǸŴܰ�</td>
    	<td width="80">�Һ��ڰ�</td>
    	<td width="80">���԰�(����)</td>
    	<td width="50">���Ը���</td>
    	<td width="80">���ް�(����)</td>
    	<td width="50">���޸���</td>
    	<% IF(isShowIpgoPrice)THEN %><td width="90">���԰�<br>(���Խñ���)</td><% end if %>
    </tr>
    <% end if %>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totShopbuy = totShopbuy + ojaego.FItemList(i).FTotShopBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum

    if (isShowSysWithReal) then
        totRealno   = totRealno + ojaego.FItemList(i).FTotRealCount
        totRealbuy  = totRealbuy + ojaego.FItemList(i).FTotRealBuySum
        totRealsell = totRealsell + ojaego.FItemList(i).FTotRealSellSum
    end if

    if not IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then
        totavgIpgoPrice = totavgIpgoPrice + ojaego.FItemList(i).FavgIpgoPriceSum
    end if
    %>
    <% if (isShowSysWithReal) then %>
    <% if ((ojaego.FItemList(i).FIsUsing="N") or (ojaego.FItemList(i).FOptionUsing="N")) and (Not isViewUser) then %>
    <tr align="center" bgcolor="<%= adminColor("dgray") %>">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
        <% if (Not isViewUser) then %>
    	<td><%= left(ojaego.FItemList(i).Fregdate,10) %></td>
    	<td>
    		<% if (datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD)) <= 3 then %>
    		<font color="red"><%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>����</font>
    		<% else %>
    		<%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>����
    		<% end if %>
    	</td>
    	<% end if %>
    	<% if (FALSE) then %>
        	<td><%= ojaego.FItemList(i).FItemGubun %></td>
        	<% if (isViewUser) then %>
        	<td><%= ojaego.FItemList(i).FItemID %></td>
        	<% else %>
        	<td><a href="javascript:TnPopOffItemStock('<%= shopid %>','<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>');"><%= ojaego.FItemList(i).FItemID %></a></td>
        	<% end if %>
        	<td><%= ojaego.FItemList(i).FItemOption %></td>
    	<% else %>
    	    <% if (isViewUser) then %>
    	    <td><%= ojaego.FItemList(i).getLogisticsCode %></td>
    	    <% else %>
    	    <td><a href="javascript:TnPopOffItemStock('<%= shopid %>','<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>');"><%= ojaego.FItemList(i).getLogisticsCode %></a></td>
    	    <% end if %>
    	<% end if %>
    	<td align="left"><%= ojaego.FItemList(i).FItemName %></td>
    	<td><%= ojaego.FItemList(i).FItemOptionName %></td>
    	<td>111</td>
		<td><%= ojaego.FItemList(i).getMaeipGubunName %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FCurrshopitemprice,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="center" ><%= FormatNumber(ojaego.FItemList(i).FTotRealCount-ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotRealCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotRealSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotRealBuySum,0) %></td>
    </tr>
    <% else %>
    <% if ((ojaego.FItemList(i).FIsUsing="N") or (ojaego.FItemList(i).FOptionUsing="N")) and (Not isViewUser)  then %>
    <tr align="center" bgcolor="<%= adminColor("dgray") %>">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
        <% if (Not isViewUser) then %>
    	<td><%= left(ojaego.FItemList(i).Fregdate,10) %></td>
    	<td>
    		<% if (datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD)) <= 3 then %>
    		<font color="red"><%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>����</font>
    		<% else %>
    		<%= datediff("m",ojaego.FItemList(i).Fregdate , ojaego.FRectYYYYMMDD) %>����
    		<% end if %>
    	</td>
    	<% end if %>
    	<% if (FALSE) then %>
        	<td><%= ojaego.FItemList(i).FItemGubun %></td>
        	<% if (isViewUser) then %>
        	<td><%= ojaego.FItemList(i).FItemID %></td>
        	<% else %>
        	<td><a href="javascript:TnPopOffItemStock('<%= shopid %>','<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>');"><%= ojaego.FItemList(i).FItemID %></a></td>
        	<% end if %>
        	<td><%= ojaego.FItemList(i).FItemOption %></td>
    	<% else %>
    	    <% if (isViewUser) then %>
    	    <td><%= ojaego.FItemList(i).getLogisticsCode %></td>
    	    <% else %>
    	    <td><a href="javascript:TnPopOffItemStock('<%= shopid %>','<%= ojaego.FItemList(i).FItemGubun %>','<%= ojaego.FItemList(i).FItemID %>','<%= ojaego.FItemList(i).FItemOption %>');"><%= ojaego.FItemList(i).getLogisticsCode %></a></td>
    	    <% end if %>
    	<% end if %>
    	<td align="left"><%= ojaego.FItemList(i).FItemName %></td>
    	<td><%= ojaego.FItemList(i).FItemOptionName %></td>
		<td><%= ojaego.FItemList(i).getITemMaeipGubunName %></td>
    	<td><%= ojaego.FItemList(i).getMaeipGubunName %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FCurrshopitemprice,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right">
			<% if Not IsNull(ojaego.FItemList(i).FTotBuySum) then  %>
				<%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %>
			<% end if %>
		</td>
    	<td>
    		<% if Not IsNull(ojaego.FItemList(i).FTotBuySum) and ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right" >
			<% if Not IsNull(ojaego.FItemList(i).FTotShopBuySum) then  %>
				<%= FormatNumber(ojaego.FItemList(i).FTotShopBuySum,0) %>
			<% end if %>
		</td>
    	<td>
    		<% if Not IsNull(ojaego.FItemList(i).FTotShopBuySum) and ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotShopBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<% IF(isShowIpgoPrice)THEN %>
    	<td align="right">
    	<% if IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then %>
    	-
    	<% else %>
    	<%= FormatNumber(ojaego.FItemList(i).FavgIpgoPriceSum,0) %>
    	<% end if %>
    	</td>
    	<% end if %>
    </tr>
    <% end if %>
    <% next %>
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="#FFFFFF">
        <% if (isViewUser) then %>
        <td colspan="5">�Ѱ�</td>
        <% else %>
    	<td colspan="7">�Ѱ�</td>
    	<% end if %>
    	<td align="center" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td >
    	    <% if totsell<>0 then %>
    		<%= clng((1-(totbuy)/(totsell))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="center" ><%= FormatNumber(totRealno-totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealbuy,0) %></td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
        <% if ( isViewUser) then %>
        <td colspan="5">�Ѱ�</td>
        <% else %>
    	<td colspan="7">�Ѱ�</td>
    	<% end if %>
    	<td align="center" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" >
			<% if Not IsNull(totbuy) then  %>
				<%= FormatNumber(totbuy,0) %>
			<% end if %>
		</td>
    	<td >
    	    <% if Not IsNull(totbuy) and totsell<>0 then %>
    		<%= clng((1-(totbuy)/(totsell))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right" >
			<% if Not IsNull(totShopbuy) then  %>
				<%= FormatNumber(totShopbuy,0) %>
			<% end if %>
		</td>
    	<td >
    	    <% if Not IsNull(totShopbuy) and totsell<>0 then %>
    		<%= clng((1-(totShopbuy)/(totsell))*100)/100 %>
    		<% end if %>
    	</td>
    	<% IF(isShowIpgoPrice)THEN %><td align="right" ><%= FormatNumber(totavgIpgoPrice,0) %></td><% end if %>
    </tr>
    <% end if %>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2">
    		<% if (makerid="") and (shopid<>"") then %>
	    		�귣��ID
	    	<% else %>
	    		��ID
	    	<% end if %>
    	</td>
    	<td rowspan="2" <% if (makerid="") and (shopid<>"") then %>width=50<% end if %> >
    		<% if (makerid="") and (shopid<>"") then %>
	    		��౸��
	    	<% else %>
	    		����
	    	<% end if %>
    	</td>
    	<td colspan="4">�ý������</td>
    	<td width="80" rowspan="2">����</td>
    	<td colspan="3">�ǻ����</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80">������</td>
    	<td width="100">�Һ��ڰ�<br>(������)</td>
    	<td width="100">���԰�<br>(����)</td>
    	<td width="80">���Ը���</td>
    	<td width="80">������</td>
    	<td width="100">�Һ��ڰ�<br>(������)</td>
    	<td width="100">���԰�<br>(����)</td>
    </tr>
    <% else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>
    		<% if (makerid="") and (shopid<>"") then %>
	    		�귣��ID
	    	<% else %>
	    		��ID
	    	<% end if %>
    	</td>
    	<td <% if (makerid="") and (shopid<>"") then %>width=50<% end if %> >
    		<% if (makerid="") and (shopid<>"") then %>
	    		��౸��
	    	<% else %>
	    		��౸��
	    	<% end if %>
    	</td>
    	<td width="80">��������</td>
    	<td width="100">�Һ��ڰ�<br>(������)</td>
    	<td width="100">���԰�<br>(����)</td>
    	<td width="80">���Ը���</td>
    	<td width="100">���ް�<br>(����)</td>
    	<td width="80">���޸���</td>
    	<td width="50"></td>
    	<% IF(isShowIpgoPrice)THEN %><td width="100">���԰�<br>(���Խñ���)</td><% end if %>
    </tr>
    <% end if %>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totShopbuy  = totShopbuy + ojaego.FItemList(i).FTotShopBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum

    if not IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then
        totavgIpgoPrice = totavgIpgoPrice + ojaego.FItemList(i).FavgIpgoPriceSum
    end if

    if (isShowSysWithReal) then
        totRealno   = totRealno + ojaego.FItemList(i).FTotRealCount
        totRealbuy  = totRealbuy + ojaego.FItemList(i).FTotRealBuySum
        totRealsell = totRealsell + ojaego.FItemList(i).FTotRealSellSum
    end if

 	if (shopid <> "") then
 		iURI = "monthlystockShop_detail.asp?menupos="& menupos &"&mwgubun="& mwgubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&makerid="& ojaego.FItemList(i).FMakerid &"&newitem="& newitem &"&shopid="& shopid &"&showminus="&showminus&"&showminusOnly="&showminusOnly&"&buseo="&buseo&"&vatyn="&vatyn&"&showSupplyPrice="&showSupplyPrice
 	else
 		iURI = "monthlystockShop_detail.asp?menupos="& menupos &"&mwgubun="& mwgubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&makerid="& makerid &"&newitem="& newitem &"&shopid="& ojaego.FItemList(i).Fshopid &"&showminus="&showminus&"&showminusOnly="&showminusOnly&"&buseo="&buseo&"&vatyn="&vatyn&"&showSupplyPrice="&showSupplyPrice
 	end if
    if Not(isShowSysWithReal) THEN iURI=iURI&"&sysorreal="& sysorreal
    %>
    <% if (isShowSysWithReal) then %>
    <% if (ojaego.FItemList(i).FMakerUsing="Y") or (isViewUser) then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
    <% end if %>
    	<td align="left">
   			<% if (makerid="") and (shopid<>"") then %>
	    		<a href="<%= iURI %>" ><%= ojaego.FItemList(i).FMakerid %></a>
	    	<% else %>
	    		<a href="<%= iURI %>" ><%= ojaego.FItemList(i).Fshopid %></a>
	    	<% end if %>
    	</td>
    	<td align="left">
   			<% if (makerid="") and (shopid<>"") then %>
				<%= ojaego.FItemList(i).getMaeipGubunName %>
	    	<% else %>
	    		<%= ojaego.FItemList(i).Fshopname %>
	    	<% end if %>
    	</td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotRealCount-ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotRealCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotRealSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotRealBuySum,0) %></td>

    </tr>
    <% else %>
    <% if (ojaego.FItemList(i).FMakerUsing="Y") or (isViewUser) then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
    <% end if %>
    	<td align="left">
   			<% if (makerid="") and (shopid<>"") then %>
				<a href="<%= iURI %>" ><%= ojaego.FItemList(i).FMakerid %></a>
	    	<% else %>
	    		<a href="<%= iURI %>" ><%= ojaego.FItemList(i).Fshopid %> (<%= ojaego.FItemList(i).Fshopname %>)</a>
	    	<% end if %>
    	</td>
    	<td align="left">
   			<% if (makerid="") and (shopid<>"") then %>
				<%= ojaego.FItemList(i).getMaeipGubunName %>
	    	<% else %>
	    		<%= ojaego.FItemList(i).getMaeipGubunName %>
	    	<% end if %>
    	</td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right">
    	<% if not isNULL(ojaego.FItemList(i).FTotBuySum) then %>
    	<%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %>
    	<% end if %>
    	</td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotShopBuySum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotShopBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td>
    	    <% if isNULL(ojaego.FItemList(i).FMaeIpGubun) then %>
    	    <% if (makerid<>"") and (shopid="") then %>
    	    <img src="/images/icon_arrow_link.gif" onClick="popAssignCommCD('<%= makerid %>','<%= yyyy1 + "-" + mm1 %>','<%= ojaego.FItemList(i).Fshopid %>')" style="cursor:pointer">
    	    <% elseif (makerid="") and (shopid="") then %>

    	    <% else %>
    		<img src="/images/icon_arrow_link.gif" onClick="popAssignCommCD('<%= ojaego.FItemList(i).FMakerid %>','<%= yyyy1 + "-" + mm1 %>','<%= shopid %>')" style="cursor:pointer">
    		<% end if %>
    		<% end if %>

    		<a target="moon1" href="/admin/offshop/stock/OutItemListByBrand.asp?shopid=<%= shopid %>&makerid=<%= ojaego.FItemList(i).FMakerid %>&research=on&cType=L&CLDiv=<%= CLDiv %>&LstYYYYMM=<%=yyyy1%>-<%=mm1%>&errExist=&ipchulcode=">
				<% if (ojaego.FItemList(i).FErrItemCnt <> 0) then %>����:<%= ojaego.FItemList(i).FErrItemCnt %><% else %>=<% end if %>
			</a>
    	</td>
    	<% IF(isShowIpgoPrice)THEN %>
    	<td align="right">
    	<% if IsNULL(ojaego.FItemList(i).FavgIpgoPriceSum) then %>
    	-
        <% else %>
    	<%= FormatNumber(ojaego.FItemList(i).FavgIpgoPriceSum,0) %>
    	<% end if %>
    	</td>
    	<% end if %>
    </tr>
    <% end if %>
    <% next %>
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td></td>
    	<td align="center" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td>
    	    <% if totsell<>0 then %>
    		<%= clng((1-(totbuy)/(totsell))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="center" ><%= FormatNumber(totRealno-totno,0) %></td>
    	<td align="center" ><%= FormatNumber(totRealno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealbuy,0) %></td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td></td>
    	<td align="center" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td>
    	    <% if totsell<>0 then %>
    		<%= clng((1-(totbuy)/(totsell))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(totShopbuy,0) %></td>
    	<td>
    	    <% if totsell<>0 then %>
    		<%= clng((1-(totShopbuy)/(totsell))*100)/100 %>
    		<% end if %>
    	</td>
    	<td></td>
    	<% IF(isShowIpgoPrice)THEN %><td align="right" ><%= FormatNumber(totavgIpgoPrice,0) %></td><% end if %>
    </tr>
    <% end if %>
</table>

<% end if %>

<%
set ojaego = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

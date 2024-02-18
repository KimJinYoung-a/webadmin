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
Dim isShowSysWithReal : isShowSysWithReal = FALSE    '''�ý������/�ǻ���� ����ǥ��
Dim isViewUser : isViewUser = FALSE ''(session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal, research, shopid, showminus
dim mwgubun, vatyn, showSupplyPrice, buseo
dim showminusOnly
dim etcjungsantype

yyyy1     = RequestCheckVar(request("yyyy1"),10)
mm1       = RequestCheckVar(request("mm1"),10)
isusing   = RequestCheckVar(request("isusing"),10)
sysorreal = RequestCheckVar(request("sysorreal"),10)
research  = RequestCheckVar(request("research"),10)
shopid    = RequestCheckVar(request("shopid"),32)
mwgubun   = RequestCheckVar(request("mwgubun"),10)
showminus   		= RequestCheckVar(request("showminus"),32)
vatyn       		= requestCheckvar(request("vatyn"),10)
showSupplyPrice 	= requestCheckvar(request("showSupplyPrice"),10)
buseo       		= requestCheckvar(request("buseo"),10)
showminusOnly       = requestCheckvar(request("showminusOnly"),10)
etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)

if (sysorreal="") then sysorreal="sys" ''real
if (isViewUser) then showminus=""
if (isViewUser) then showminusOnly=""
if (isViewUser) then sysorreal="sys"
if (isViewUser) then isusing=""

if (research="") and (showminus="") then showminus="on"
if (research="") and (mwgubun="") then mwgubun="M"
if (research="") and (etcjungsantype="") then etcjungsantype="41" ''����+�Ǹź�
if (research="") and (buseo="") then buseo="3X" ''3pl����
if (research="") and (showSupplyPrice="") then showSupplyPrice="Y"

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

dim oshopjaego
set oshopjaego = new CMonthlyStock
oshopjaego.FRectYYYYMM   = yyyy1 + "-" + mm1
oshopjaego.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
oshopjaego.FRectIsUsing = isusing
oshopjaego.FRectGubun = sysorreal
oshopjaego.FRectShopid = shopid
oshopjaego.FRectMwDiv    = mwgubun
oshopjaego.FRectShowMinus = showminus
oshopjaego.FRectShowMinusOnly = showminusOnly
oshopjaego.FRectVatYn    = vatyn
oshopjaego.FRectShopSuplyPrice    = showSupplyPrice
oshopjaego.FRectTargetGbn = buseo
oshopjaego.FRectetcjungsantype = etcjungsantype

IF (isShowSysWithReal) then
    oshopjaego.FRectGubun = "sys"
    oshopjaego.GetShopMonthlyJeagoSumSysWithReal
ELSE
    oshopjaego.GetShopMonthlyJeagoSumNew
END IF

dim i
dim totno, totbuy, totsell, totavgBuy, offtotavgBuy
dim offtotno, offtotbuy, totshopBuy, offtotsell
dim totRealno, totRealbuy, totRealsell

dim iURL

%>
<script type='text/javascript'>

function reActdailySummary1(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('�Ϻ������ ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopdailystock1&yyyymm=' + yyyymm,'reActMonthSummary1','width=600,height=600');
	popwin.focus();
}
function reActMonthSummary(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthlystock&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}

function reActMonthSummary10(){
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly10&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_1() {
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ����(��������Ӹ�)�� ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly101&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_2() {
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ����(���зθ���)�� ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly102&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_3() {
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ����(���з� �� ����)�� ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly103&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_4() {
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ����(���+�����Ź)�� ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly104&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_5() {
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ����(���� ��)�� ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly105&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary11(){
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly11&yyyymm=' + yyyymm,'reActMonthSummary11','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary20(){
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly20&yyyymm=' + yyyymm,'reActMonthSummary20','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary21(){
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly21&yyyymm=' + yyyymm,'reActMonthSummary21','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary30(){
    //alert('������..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly30&yyyymm=' + yyyymm,'reActMonthSummary30','width=100,height=100');
	popwin.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
			&nbsp;&nbsp;
			���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
			&nbsp;&nbsp;
			<font color="#CC3333">�μ�����:</font>
	        <% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
        	&nbsp;&nbsp;
        	<font color="#CC3333">������:</font>
        	<% 'drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype",etcjungsantype,"" %>
        	<select class="select" name="etcjungsantype"  >
            <option value="">-����-</option>
            <option value="1" <%=CHKIIF(etcjungsantype="1","selected","")%> >�Ǹź�����</option>
            <option value="2" <%=CHKIIF(etcjungsantype="2","selected","")%> >��������</option>
            <option value="3" <%=CHKIIF(etcjungsantype="3","selected","")%> >����������</option>
            <option value="4" <%=CHKIIF(etcjungsantype="4","selected","")%> >����������</option>
            <option value="41" <%=CHKIIF(etcjungsantype="41","selected","")%> >������+�Ǹź�����</option>
            </select>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <% IF NOt (isShowSysWithReal) THEN %>
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

        	<font color="#CC3333">��������</font>
        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >��ü
        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >����
        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >�鼼
        	&nbsp;&nbsp;
			<input type="checkbox" name="showSupplyPrice" value="Y" <%= CHKIIF(showSupplyPrice="Y","checked","") %> >���ް��� ǥ��
        	<br>
        	<font color="#CC3333">���Ա���:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >����(�������+������+ITS�����Ź)
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >��Ź(��Ź�Ǹ�+��ü��Ź+�����Ź)
        	<input type="radio" name="mwgubun" value="C" <% if mwgubun="C" then response.write "checked" %> >�����Ź
        	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >��ü -->
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >������
        	<% if (Not isViewUser) then %>
        	<br>
        	<input type="checkbox" name="showminus" <%= CHKIIF(showminus="on","checked","") %> >���̳ʽ���� ����
			<input type="checkbox" name="showminusOnly" <%= CHKIIF(showminusOnly="on","checked","") %> >���̳ʽ����
        	<% end if %>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<br>

* <font color="red">�������� ��ǰ����</font>�� ���� ��� ǥ�õ��� �ʽ��ϴ�.<br />
* ���� <font color="red">����ڻ� ���Ա���</font>�� ���� ���곻���� �ۼ��� �� �����˴ϴ�.

<!-- �׼� ���� -->
<% ''if C_ADMIN_AUTH or (session("ssBctId") = "faxy") then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�Ϻ���������ۼ�" onclick="reActdailySummary1();" >
			<input type="button" class="button" value="����ڻ����ۼ�" onclick="reActMonthSummary();" >
			&nbsp;&nbsp;
			<!--
			<input type="button" class="button" value="���ۼ� 1�ܰ�" onclick="reActMonthSummary10();">
			&nbsp;
			-->
			<input type="button" class="button" value="���ۼ� 1-1�ܰ�" onclick="reActMonthSummary10_1();">
			<input type="button" class="button" value="���ۼ� 1-2�ܰ�" onclick="reActMonthSummary10_2();">
			<input type="button" class="button" value="���ۼ� 1-3�ܰ�" onclick="reActMonthSummary10_3();">
			<input type="button" class="button" value="���ۼ� 1-4�ܰ�" onclick="reActMonthSummary10_4();">
			<input type="button" class="button" value="���ۼ� 1-5�ܰ�" onclick="reActMonthSummary10_5();">
			<input type="button" class="button" value="���ۼ� 2�ܰ�" onclick="reActMonthSummary11();">
			<input type="button" class="button" value="���ۼ� 3-1�ܰ�" onclick="reActMonthSummary20();">
			<input type="button" class="button" value="���ۼ� 3-2�ܰ�" onclick="reActMonthSummary21();">
			<input type="button" class="button" value="���ۼ� 4�ܰ�" onclick="reActMonthSummary30();">
		</td>
		<td align="right">
		</td>
	</tr>
</table>

<p>
<% ''end if %>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td rowspan="2">����</td>
        <td rowspan="2">������</td>
    	<td width="110" rowspan="2">���Ա���</td>
    	<td colspan="3">�ý������</td>
    	<td width="39" rowspan="2">����</td>
    	<td colspan="3">�ǻ����</td>
    	<td  width="90" rowspan="2">�귣�庰<br>����ڻ�</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="110">������</td>
    	<td width="110">�Һ��ڰ�*����</td>
    	<td width="110">���԰�*����<br>(���� ���԰�)</td>
    	<td width="110">������</td>
    	<td width="110">�Һ��ڰ�*����</td>
    	<td width="110">���԰�*����<br>(���� ���԰�)</td>
    </tr>
	<% else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="110">�μ�</td>
	    <td width="110">����ID</td>
	    <td >����</td>
	    <td >������</td>
    	<td  width="110">���Ա���</td>
    	<td width="110">��������</td>
    	<td width="110">�Һ��ڰ�*����</td>
    	<!-- td width="110">��ո���</td -->
    	<td width="110">���԰�*����<br>(���� ���԰�)</td>
    	<td width="110">���ް�*����<br>(���� ���԰�)</td>
    	<% IF(isShowIpgoPrice)THEN %><td width="110">�Ǹ��԰�*����</td><% END IF %>
    	<td  width="90">�귣�庰<br>����ڻ�</td>
    	<!-- td  width="90">�귣�庰<br>���ȸ����</td -->
    </tr>
    <% end if %>

    <% for i=0 to oShopJaego.FResultCount-1 %>
    <% if (oShopJaego.FItemList(i).FMaeIpGubun="Z") and (oShopJaego.FItemList(i).FTotCount=0) then %>

    <% else %>
    <% if (TRUE) or oShopJaego.FItemList(i).FMaeIpGubun<>"Z" then %>
    <%
    totno   = totno + oShopJaego.FItemList(i).FTotCount
    totbuy  = totbuy + CCur(oShopJaego.FItemList(i).FTotBuySum)
    totshopBuy  = totshopBuy + CCur(oShopJaego.FItemList(i).FTotShopBuySum)
    totsell = totsell + CCur(oShopJaego.FItemList(i).FTotSellSum)

    if Not IsNULL(oShopJaego.FItemList(i).FavgIpgoPriceSum) THEN
       totavgBuy = totavgBuy + oShopJaego.FItemList(i).FavgIpgoPriceSum
    end if

    if (isShowSysWithReal) then
        totRealno   = totRealno + oShopJaego.FItemList(i).FTotRealCount
        totRealbuy  = totRealbuy + oShopJaego.FItemList(i).FTotRealBuySum
        totRealsell = totRealsell + oShopJaego.FItemList(i).FTotRealSellSum
    end if

    iURL = "monthlystockShop_detail.asp?menupos="& menupos &"&mwgubun="& oShopJaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1&"&mm1="& mm1 &"&isusing="& isusing &"&shopid="&oShopJaego.FItemList(i).FShopID&"&showminus="&showminus&"&showminusOnly="&showminusOnly&"&buseo="&oShopJaego.FItemList(i).FtargetGbn&"&vatyn="&vatyn&"&showSupplyPrice="&showSupplyPrice
    if Not(isShowSysWithReal) THEN iURL=iURL&"&sysorreal="& sysorreal
    %>
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getBusiName %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).FShopName %></a></td>
    	<td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td><%= oShopJaego.FItemList(i).getEtcJungsanTypeName %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="center"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealCount-oShopJaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealBuySum,0) %></td>
    	<td align="center"><a href="<%= iURL %>" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getBusiName %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).FShopID %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).FShopName %></a></td>
        <td><%= oShopJaego.FItemList(i).getEtcJungsanTypeName %></td>
    	<td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotShopBuySum,0) %></td>
    	<% IF(isShowIpgoPrice)THEN %>
    	<td align="right">
        	<% If IsNULL(oShopJaego.FItemList(i).FavgIpgoPriceSum) then %>
        	-
        	<% else %>
        	<%= FormatNumber(oShopJaego.FItemList(i).FavgIpgoPriceSum,0) %>
        	<% end if %>
    	</td><% END IF %>
    	<td align="center"><a href="<%= iURL %>" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
    	<!--td align="center"><a href="javascript:alert('�غ���.');" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td -->
    </tr>
    <% end if %>
    <% end if %>
    <% end if %>
    <% next %>
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td >�Ѱ�</td>
    	<td></td>
    	<td></td>
        <td></td>
        <td></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="center" ><%= FormatNumber(totRealno-totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealbuy,0) %></td>
    	<td align="center"></td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td >�Ѱ�</td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<!-- td></td -->
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totshopBuy,0) %></td>
        <% IF(isShowIpgoPrice)THEN %><td align="right"><%= FormatNumber(totavgBuy,0) %></td><% END IF %>
    	<td align="center"></td>
    	<!--td align="center"><a href="avascript:alert('�غ���.');" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td-->
    </tr>
    <% end if %>
</table>



<%
set oShopJaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

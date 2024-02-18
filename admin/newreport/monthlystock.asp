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
Const isOnlySys = FALSE
Dim isViewUser : isViewUser = FALSE ''(session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn
dim mwgubun, buseo, itemgubun
dim purchasetype, swSppPrc

yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
isusing     = requestCheckvar(request("isusing"),10)
sysorreal   = requestCheckvar(request("sysorreal"),10)
research    = requestCheckvar(request("research"),10)
newitem     = requestCheckvar(request("newitem"),10)
mwgubun     = requestCheckvar(request("mwgubun"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
minusinc   = requestCheckvar(request("minusinc"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
buseo       = requestCheckvar(request("buseo"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
purchasetype   = requestCheckvar(request("purchasetype"),10)
swSppPrc	= requestCheckvar(request("swSppPrc"),32)

if (sysorreal="") then sysorreal="sys"  ''real
if (isViewUser="") then sysorreal="sys"
if (isViewUser="") then bPriceGbn="P"
if (isViewUser="") then isusing=""

if (research="") then
	buseo = "3X"
    bPriceGbn="P"
	swSppPrc = "Y"
	mwgubun = "M"
end if

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
ojaego.FRectIsUsing = isusing
ojaego.FRectGubun = sysorreal
ojaego.FRectNewItem = newitem
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectVatYn    = vatyn
ojaego.FRectItemGubun = itemgubun
ojaego.FRectMinusInclude = minusinc
ojaego.FRectPurchaseType = purchasetype
ojaego.FRectTargetGbn = buseo
ojaego.FRectShopSuplyPrice    = swSppPrc

if (buseo="IT") then
    ojaego.FRectITSOnlyOrNot = "O"
else
    ojaego.FRectITSOnlyOrNot = "N"
end if

if (bPriceGbn="P") then
    ojaego.FRectIsFix = "on"
end if
ojaego.GetMonthlyJeagoSumWithPreMonth '' GetMonthlyJeagoSumNew ''


dim i
dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy

dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy


dim iURL
dim nBusiName
%>
<script type='text/javascript'>

function reActMonthSummary(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' ���� ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=monthlystock&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}
function reActdailySummary1(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('�Ϻ������ STEP1 ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=dailystock1&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}
function reActdailySummary2(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('�Ϻ������ STEP2 ��� ������ ���ۼ� �Ͻðڽ��ϱ�?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=dailystock2&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
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

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<% IF Not (isOnlySys) THEN %>
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
        <% end if %>
        	<font color="#CC3333">���Ա���:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >����
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >��Ź
        	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >��ü -->
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >������

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
	    <input type="radio" name="bPriceGbn" value="" <%= CHKIIF(bPriceGbn="","checked","") %>  >������԰�
	    <input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >�ۼ��ø��԰�
	    <input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %> disabled >��ո��԰�
	    <% end if %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	    	<font color="#CC3333">�μ�����:</font>
	        <% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
			&nbsp;
	    	<font color="#CC3333">��ǰ����:</font>
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;
			�������� : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

	<!--* <font color="red">���� �ۼ��� �������</font>�� ���ܵ˴ϴ�.(�߰��� �ϰ� �ݿ���, ��ǰ�������Ȳ���� ������� ��ü ���ΰ�ħ�ϸ� �ݿ���)<br>-->
	* <font color="red">����Ȯ�� ��</font> ������� �ۼ��� ��� ����ڻ꿡 �ݿ��ȵ�(���̻�Կ��� ��û�ؾ� �ݿ���-mwgubun)

<p>

<!-- �׼� ���� -->
<% ''if C_ADMIN_AUTH or (session("ssBctId") = "faxy") then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�Ϻ���������ۼ�STEP1" onclick="reActdailySummary1();" <% if (Left(DateAdd("m", -1, Now()), 7) > (yyyy1 + "-" + mm1)) then %>disabled<% end if %> ><!-- ���� ��û���� disable -->
			<input type="button" class="button" value="�Ϻ���������ۼ�STEP2" onclick="reActdailySummary2();" <% if (Left(DateAdd("m", -1, Now()), 7) > (yyyy1 + "-" + mm1)) then %>disabled<% end if %> ><!-- ���� ��û���� disable -->
			<input type="button" class="button" value="����ڻ����ۼ�" onclick="reActMonthSummary();" <% if (Left(DateAdd("m", -1, Now()), 7) > (yyyy1 + "-" + mm1)) then %>disabled<% end if %> ><!-- ���� ��û���� disable -->
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
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="4">��ǰ����</td>
        <td colspan="2">�������(��������)<br>A</td>
        <td colspan="2">�������(��)<br>B</td>
        <td colspan="2">�⸻���(��������)<br>C</td>
        <td colspan="2">�Ѹ������<br>D=A+B-C</td>
        <td width="1" bgcolor="#FFFFFF"></td>
        <td colspan="2">���LOSS<br>E</td>
        <td colspan="2">��ǰ�������<br>F=A+B+E-C</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td >�μ�</td>
	    <td >����</td>
	    <td >�ڵ屸��</td>
    	<td >���Ա���</td>
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
    if i<ojaego.FResultCount-1 then
        nBusiName= ojaego.FItemList(i+1).getBusiName
    else
        nBusiName=""
    end if

    if (ojaego.FItemList(i).getBusiName=nBusiName) then nBusiName=""
    if (i=ojaego.FResultCount-1) then nBusiName="L"

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


    iURL = "monthlystock_detail.asp?menupos="& menupos &"&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
    iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype&"&swSppPrc="&swSppPrc
    if Not(isOnlySys) THEN iURL=iURL&"&sysorreal="& sysorreal
    %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= ojaego.FItemList(i).getBusiName %></td>
        <td><a href="<%= iURL %>" target="_blank"><%= GetItemGubunName(ojaego.FItemList(i).Fitemgubun) %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
    	<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
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

    <% if (nBusiName<>"") then %>
    <tr align="center" bgcolor="#EEFFEE">

    	<td></td>
    	<td>�Ұ�</td>
    	<td></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(subPreno,0) %></td>
    	<td align="right" ><%= FormatNumber(subPrebuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subIpno,0) %></td>
    	<td align="right" ><%= FormatNumber(subIpBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subTotno,0) %></td>
    	<td align="right" ><%= FormatNumber(subTotbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subPreno+subIpno-subtotno,0) %></td>
    	<td align="right" ><%= FormatNumber(subPrebuy+subIpBuy-subtotbuy,0) %></td>
    	<td ></td>
    	<td align="right" ><%= FormatNumber(subLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(subLossBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subPreno+subIpno-subtotno+subLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(subPrebuy+subIpBuy-subtotbuy+subLossBuy,0) %></td>
    </tr>
    <tr  bgcolor="#FFFFFF">
    	<td colspan="17"></td>
    </tr>
    <%
        subTotno=0
        subTotbuy=0

        subPreno   =0
        subPrebuy  =0
        subIpno    =0
        subIpBuy    =0
        subLossno   =0
        subLossBuy  =0

    %>
    <% end if %>
    <% next %>



    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
    	<td>�Ѱ�</td>
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



<%
set ojaego = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

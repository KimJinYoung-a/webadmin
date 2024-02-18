<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyInventoryCls.asp"-->
<%

dim research, i
dim yyyy1,mm1, yyyymm1, makerid, showsupply
dim stockPlace, shopid, stockGubun, showShopid
dim targetGbn, itemgubun, mwdiv
dim ArrList

research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1       	= requestCheckvar(request("mm1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
stockGubun  = requestCheckvar(request("stockGubun"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsupply   = requestCheckvar(request("showsupply"),10)
shopid    	= requestCheckvar(request("shopid"),32)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
mwdiv   = requestCheckvar(request("mwdiv"),10)
showShopid   = requestCheckvar(request("showShopid"),10)


dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


yyyymm1 = yyyy1 + "-" + mm1


dim oCMonthlyInventory
set oCMonthlyInventory = new CMonthlyInventory

oCMonthlyInventory.FRectYYYYMM = yyyymm1
oCMonthlyInventory.FRectStockPlace = stockPlace
oCMonthlyInventory.FRectStockGubun = stockGubun
oCMonthlyInventory.FRectMakerid = makerid
oCMonthlyInventory.FRectBySupplyPrice = showsupply
oCMonthlyInventory.FRectShopid = shopid
oCMonthlyInventory.FRectItemgubun = itemgubun
oCMonthlyInventory.FRectTargetGbn = targetGbn
oCMonthlyInventory.FRectMwdiv = mwdiv
oCMonthlyInventory.FRectShowShopid = showShopid

oCMonthlyInventory.GetMonthlyInventorySUM

if oCMonthlyInventory.FTotalCount>0 then
	ArrList = oCMonthlyInventory.farrlist
end if

dim oitem

%>
<script src="/js/jquery-1.7.1.min.js"></script>
<script>
function jsRewrite(yyyymm) {


    // �������
    realCall(yyyymm, 'makeStockBeginStock');
    realCall(yyyymm, 'makeStockIpgo');
    realCall(yyyymm, 'makeStockMove');
    realCall(yyyymm, 'makeStockSell');
    realCall(yyyymm, 'makeStockSellOnGift');
    realCall(yyyymm, 'makeStockSellUpcheWitak');
    realCall(yyyymm, 'makeStockShopLoss');
    realCall(yyyymm, 'makeStockCsChulgo');
    realCall(yyyymm, 'makeStockEndStock');
}

function realCall(yyyymm, mode) {
    var url;
    var host = window.location.protocol + "//" + window.location.host + '/admin/newreport/monthlyInventorySum_process.asp?yyyymm=' + yyyymm + '&silent=Y';

    url = host + '&mode=' + mode;
    var data = '{}';

    $.ajax({
        type : 'POST',
        url : url,
        data : data,
        async : false,
        dataType: 'html',
        contentType: 'application/x-www-form-urlencoded; charset=euc-kr',
        error:function(request, status, error) {
            alert("code:"+request.status+"\n"+"message:"+request.responseText+"\n"+"error:"+error);
        },
        success : function(data) {
            if (data.indexOf('{') > 0) {
                data = data.substring(data.indexOf('{'));
            }
            // alert(data);

            var obj = JSON.parse(data);
            if (obj.code == '000') {
                alert(obj.message);
            } else {
                alert(obj.message);
            }
        }
    });
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" target="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">* ��/�� :</font> <% DrawYMBox yyyy1,mm1 %>
			&nbsp;&nbsp;
			<font color="#CC3333">* �귣�� :</font> <%	drawSelectBoxDesignerWithName "makerid", makerid %>
            &nbsp;&nbsp;
            <font color="#CC3333">* ���� :</font> <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showsupply" value="Y" <%= CHKIIF(showsupply="Y","checked","") %> disabled> ���ް��� ǥ��
            &nbsp;&nbsp;
			<input type="checkbox" name="showShopid" value="Y" <%= CHKIIF(showShopid="Y", "checked", "") %> > ����ǥ��
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.target='';document.frm.action='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
            <font color="#CC3333">* ���Ա��� :</font>
		    <select name="mwdiv" class="select">
		        <option value="" <%= CHKIIF(mwdiv="","selected" ,"") %> >��ü</option>
        		<option value="M" <%= CHKIIF(mwdiv="M","selected" ,"") %> >����</option>
        		<option value="W" <%= CHKIIF(mwdiv="W","selected" ,"") %> >��Ź</option>
        	</select>
        	&nbsp;&nbsp;
		    <font color="#CC3333">* �����ġ :</font>
		    <select name="stockPlace" class="select">
		        <option value="" <%= CHKIIF(stockPlace="","selected" ,"") %> >��ü</option>
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >����</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >����</option>
				<option value="E" <%= CHKIIF(stockPlace="E","selected" ,"") %> >��Ÿ</option>
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">* ����� :</font>
        	<select name="stockGubun" class="select">
        	<option value="">��ü
        	<option value="M" <%= CHKIIF(stockGubun="M","selected" ,"") %> >����
        	<option value="W" <%= CHKIIF(stockGubun="W","selected" ,"") %> >��Ź
        	<option value="T" <%= CHKIIF(stockGubun="T","selected" ,"") %> >3PL
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">* �ڵ屸�� :</font>
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;&nbsp;
			<font color="#CC3333">* ���԰����� :</font>
			<input type="radio" name="priceGubun" value="V" checked> ��ո��԰�
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

* 3�� CS���� ����� : 10-2867096-0012

<p />

<div style="float: right; margin-bottom: 5px;">
    <input type="button" value="���ۼ�(<%= yyyy1 & "-" & mm1 %>)" onclick="jsRewrite('<%= yyyy1 & "-" & mm1 %>')" />
</div>

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="5">��ǰ����</td>
        <td colspan="2">�������(��������)</td>
        <td colspan="2">�������(��)</td>
        <td colspan="2">����̵�(��)</td>
        <td colspan="2">����Ǹ�(��)</td>
        <td colspan="2">������1(��)</td>
        <td colspan="2">������2(��)</td>
        <td colspan="2">�����Ÿ���(��)</td>
        <td colspan="2">���CS���(��)</td>
        <td colspan="2">����(��)</td>
		<td colspan="2"><b>�⸻���(��)</b></td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>����</td>
	    <td>�ڵ�<br />����</td>
	    <td>����<br />����</td>
	    <td>���<br>��ġ</td>
		<td>�����̵�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    	<td>����</td>
    	<td>�ݾ�</td>
    </tr>
    <% if isarray(arrlist) then %>
    <% for i=0 to ubound(arrlist,2) %>
    <%
    set oitem = new CMonthlyInventoryItem
    Call oitem.SetValueByArray(arrlist, i)
    %>
    <tr align="right" bgcolor="#FFFFFF" >
        <td align="center"><%= oitem.GetShopDivBasic() %></td>
        <td align="center"><%= oitem.Fitemgubun %></td>
        <td align="center"><%= oitem.GetMwdivName %></td>
        <td align="center"><%= oitem.GetStockPlaceName %></td>
        <td align="center"><%= oitem.Fshopid %></td>
        <td><%= FormatNumber(oitem.FBeginingNo, 0) %></td>
        <td><%= FormatNumber(oitem.FBeginingSum, 0) %></td>
        <td><%= FormatNumber(oitem.FMaeipNo, 0) %></td>
        <td><%= FormatNumber(oitem.FMaeipSum, 0) %></td>
        <td><%= FormatNumber(oitem.FMoveNo, 0) %></td>
        <td><%= FormatNumber(oitem.FMoveSum, 0) %></td>
        <td><%= FormatNumber(oitem.FSellNo, 0) %></td>
        <td><%= FormatNumber(oitem.FSellSum, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoOneNo, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgSOneum, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoTwoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoTwoSum, 0) %></td>
        <td><%= FormatNumber(oitem.FEtcChulgoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FEtcChulgoSum, 0) %></td>
        <td><%= FormatNumber(oitem.FCsChulgoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FCsChulgoSum, 0) %></td>
        <td><%= FormatNumber(oitem.getDiffNo, 0) %></td>
        <td><%= FormatNumber(oitem.getDiffSum, 0) %></td>
        <td><%= FormatNumber(oitem.FEndingNo, 0) %></td>
        <td><%= FormatNumber(oitem.FEndingNo, 0) %></td>
    </tr>
    <% next %>
    <% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

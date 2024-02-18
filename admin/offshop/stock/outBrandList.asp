<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������θ��� ������� ��ǰ ���� �귣��
' History : 2011.08
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shopstockClearCls.asp"-->

<%

'response.write "������"
'response.end

dim shopid, makerid, research, yyyy1, mm1
dim dispDiv

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
research     = RequestCheckVar(request("research"),32)

dim usingyn, centermwdiv ,NoZeroStock, comm_cd
usingyn      = RequestCheckVar(request("usingyn"),32)
centermwdiv  = RequestCheckVar(request("centermwdiv"),32)
NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
comm_cd      = RequestCheckVar(request("comm_cd"),32)

yyyy1         = RequestCheckVar(request("yyyy1"),4)
mm1           = RequestCheckVar(request("mm1"),2)

dispDiv           = RequestCheckVar(request("dispDiv"),2)


Dim PreMonth : PreMonth = DateAdd("m",-1,Now())
if (yyyy1="") then
    yyyy1 = Left(CStr(PreMonth),4)
    mm1   = Mid(CStr(PreMonth),6,2)
end if


''����
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if
if (research="") then NoZeroStock="on"

dim oOffOutItem
set oOffOutItem = new CShopStockClear
oOffOutItem.FRectShopID		= shopid
oOffOutItem.FRectMakerID	= makerid
oOffOutItem.FRectCommCD		= comm_cd
oOffOutItem.FRectDispDiv	= dispDiv
''oOffOutItem.FRectYYYYMM	= yyyy1 + "-" + mm1


if (shopid<>"") or (makerid<>"") then
    oOffOutItem.GetShopStockClearBrandList
end if

dim i
%>
<script language='javascript'>

function popOffOutItemList(makerid, shopid, cType){
    var popUrl = "/admin/offshop/stock/OutItemListByBrand.asp?makerid="+makerid+"&shopid="+shopid+"&cType="+cType;
    var popwin = window.open(popUrl,'OutItemListByBrand','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popMonthlyStock(makerid, shopid, yyyymm){
    var popUrl = "/admin/newreport/monthlystockShop_detail.asp?menupos=1346&showminus=on&sysorreal=sys&makerid="+makerid+"&shopid="+shopid+"&yyyymm="+yyyymm;
    var popwin = window.open(popUrl,'popMonthlyStock','width=1100,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    <!-- ����� �������� ����
		    ��� ��/��
		    <input type="text" name="yyyy1" value="<%= yyyy1 %>" size="4" readOnly class="text_ro">
		    <input type="text" name="mm1" value="<%= mm1 %>" size="2" readOnly class="text_ro">
		    &nbsp;
		    -->
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    ���� : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- ���� ��ü -->
    		    ���� : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        ���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp; <!-- drawSelectBoxOffShop -->
		    <% end if %>

		    <% if (C_IS_Maker_Upche) then %>
		        <input type="hidden" name="makerid" value="<%= makerid %>">
		    <% else %>
    			�귣�� :
    			<% drawSelectBoxDesignerwithName "makerid", makerid %> &nbsp;&nbsp;
			<% end if %>
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<!-- ��ǰ ��뱸�� : <% drawSelectBoxUsingYN "usingyn", usingyn %> &nbsp;&nbsp; -->

		    ���Ա��� :
		    <% drawSelectBoxOFFJungsanCommCD "comm_cd",comm_cd %>

            &nbsp;&nbsp;
			ǥ�ñ��� :
			<select class="select" name="dispDiv">
				<option value="SY" <%if (dispDiv = "SY") then %>selected<% end if %> >SYS���</option>
				<option value="ER" <%if (dispDiv = "ER") then %>selected<% end if %> >�������</option>
				<option value="SM" <%if (dispDiv = "SM") then %>selected<% end if %> >�������</option>
			</select>
			�ִ� �귣�常
            <!--
             <input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > ���0�� �귣�� �˻� ����.
             -->
		</td>
	</tr>

	</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
    <tr height="60">
        <td>
        * ���̳ʽ� ���� ��� 0 ���� ������.<br>
        * ���>0 �̰� �Ǹŷ� <1 �� ��ǰ
        </td>
    </tr>
	<tr height="30">
		<td align="left">
			�˻���� �� <%= oOffOutItem.FTotalCount %> ��
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="150">�귣��ID</td>
    	<td width="90">(��)���Ա���</td>
    	<td width="70">�ý������<br>�����ǰ</td>
    	<td width="70">�ý������</td>
		<td width="70">����</td>
    	<td width="70"><b>�ǻ����</b></td>
		<td width="70">����</td>
		<td width="80"><b>��ȿ���<br>(�ǻ�+����)</b></td>
    	<td width="70">2���� <br>�Ǹż���</td>
    	<td width="70">2���� <br>�����</td>
    	<td width="70">���� <br>�԰���</td>
    	<td width="70">���� <br>�԰���</td>
    	<!--
    	<td width="90">�������<br>��ǰ��</td>
    	<td width="90">�������<br>��ǰ���</td>
    	-->
        <td >���</td>
    </tr>
    <% if (shopid="") then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <td colspan="13">[���� <Strong>����</Strong> �� �����ϼ���.]</td>
    </tr>
    <% else %>
    <% for i=0 to oOffOutItem.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= oOffOutItem.FItemList(i).Fmakerid %></td>
        <td><%= oOffOutItem.FItemList(i).Fcomm_name %></td>
        <td><a href="javascript:popMonthlyStock('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','<%= yyyy1 %>','<%= mm1 %>');"><%= oOffOutItem.FItemList(i).FItemCnt %></a></td>
        <td><%= oOffOutItem.FItemList(i).Ftotsysstockno %></td>
        <td><%= oOffOutItem.FItemList(i).Ftoterrrealcheckno %></td>
		<td><b><%= oOffOutItem.FItemList(i).Ftotrealstockno %></b></td>
		<td><%= oOffOutItem.FItemList(i).Ftoterrsampleitemno %></td>
		<td><b><%= (oOffOutItem.FItemList(i).Ftotrealstockno + oOffOutItem.FItemList(i).Ftoterrsampleitemno) %></b></td>
        <td><%= oOffOutItem.FItemList(i).FtotSellNo %></td>
        <td><%= FormatNumber(oOffOutItem.FItemList(i).FtotRealSellPrice,0) %></td>
        <td>
            <% if IsRecentIpchul(oOffOutItem.FItemList(i).Ffirstipgodate) then %>
            <b><%= oOffOutItem.FItemList(i).Ffirstipgodate %></b>
            <% else %>
            <%= oOffOutItem.FItemList(i).Ffirstipgodate %>
            <% end if %>
        </td>
        <td>
            <% if IsRecentIpchul(oOffOutItem.FItemList(i).Flastipgodate) then %>
            <b><%= oOffOutItem.FItemList(i).Flastipgodate %></b>
            <% else %>
            <%= oOffOutItem.FItemList(i).Flastipgodate %>
            <% end if %>
        </td>
        <td>
            <!--
            <input type="button" value="�������(�ӽ�)" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','C')">
            &nbsp;

            <input type="button" value="��������" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','M')">
            &nbsp;
            -->
            <input type="button" value="���� �ν�ó��" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','L')">
			<input type="button" value="���� ���" class="button" onClick="popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','S')">
        </td>
    </tr>
    <% next %>
    <% end if %>
</table>
<%
set oOffOutItem = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

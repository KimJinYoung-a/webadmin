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
<!-- #include virtual="/lib/classes/datamart/brandStockTurnOverCls.asp"-->


<%
rw "�������"
response.end

dim shopid, makerid, research, yyyy1, mm1

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
set oOffOutItem = new CBrandStockTurnOver
oOffOutItem.FRectShopID       = shopid
oOffOutItem.FRectMakerID      = makerid
oOffOutItem.FRectComm_cd      = comm_cd
oOffOutItem.FRectYYYYMM       = yyyy1 + "-" + mm1


''oOffOutItem.FRectNoZeroStock  = NoZeroStock
if (shopid<>"") or (makerid<>"") then
    if ((shopid<>"") and (makerid<>"")) then oOffOutItem.FRectYYYYMM=""
        
    oOffOutItem.getOutItemBrandList
end if

dim i
%>
<script language='javascript'>
function popBrandStock(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTaking(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock_taking.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTakingInput(stIdx){
    var popUrl = "/common/offshop/shop_brandcurrentstock_byjobkey.asp?idx="+stIdx+"&sType=stTaking";
    var popwin = window.open(popUrl,'popBrandStockInput','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popOffOutItemList(makerid, shopid, yyyy1, mm1){
    var popUrl = "/common/offshop/stock/OutItemList.asp?makerid="+makerid+"&shopid="+shopid+"&yyyy1="+yyyy1+"&mm1="+mm1;
    var popwin = window.open(popUrl,'popItemStockTurnOver','scrollbars=yes,resizable=yes');
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
		    ��� ��/�� <% CALL DrawYMBox(yyyy1,mm1) %>
		    &nbsp;
		    <% if (C_IS_SHOP) then %>
    		    <input type="hidden" name="shopid" value="<%= shopid %>">
    		    ���� : <%= shopid %>
		    <% elseif (C_IS_Maker_Upche) then %>
    		    <!-- ���� ��ü -->
    		    ���� : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		    <% else %>
		        ���� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
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
			
		    ������Ա��� : 
		    <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
			
             &nbsp;&nbsp;
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
        <td width="60">��/��</td>
        <td width="110">�귣��ID</td>
    	<td width="100">���Ա���</td>
    	<td width="70">��� ���� <br>��ǰ ����</td>
    	<td width="70"><%= mm1 %>�� <br>�Ǹż���</td>
    	<td width="70"><%= mm1 %>�� <br>�����</td>
    	
    	<td width="90">�������<br>��ǰ��</td>
    	<td width="90">�������<br>��ǰ���</td>
        <td >���</td>
    </tr>
    <% if (shopid="") then %>
    <tr align="center" bgcolor="#FFFFFF" height="30">
        <td colspan="14">[���� <Strong>����</Strong> �� �����ϼ���.]</td>
    </tr>
    <% else %>
    <% for i=0 to oOffOutItem.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= oOffOutItem.FItemList(i).Fyyyymm %></td>
        <td><a href="javascript:popOffOutItemList('<%= oOffOutItem.FItemList(i).Fmakerid %>','<%= shopid %>','<%= Left(oOffOutItem.FItemList(i).FYYYYMM,4) %>','<%= Right(oOffOutItem.FItemList(i).FYYYYMM,2) %>');"><%= oOffOutItem.FItemList(i).Fmakerid %></a></td>
        <td><%= oOffOutItem.FItemList(i).Fcomm_name %></td>
        <td><%= oOffOutItem.FItemList(i).FItemCnt %></td>
        <td><%= oOffOutItem.FItemList(i).FtotSellno %></td>
        <td><%= FormatNumber(oOffOutItem.FItemList(i).FtotRealSellPrice,0) %></td>
        <td><%= FormatNumber(oOffOutItem.FItemList(i).FitemTaragetCnt,0) %></td>
        <td><%= FormatNumber(oOffOutItem.FItemList(i).FstockTaragetCnt,0) %></td>
        <td></td>
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

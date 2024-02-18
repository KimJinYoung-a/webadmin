<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� ȯ�� ����
' History : 2010.08.07 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->

<%
dim shopid : shopid = requestCheckvar(request("shopid"),32)
dim mode  : mode = requestCheckvar(request("mode"),32)
dim exchangeRate  : exchangeRate = requestCheckvar(request("exchangeRate"),32)
dim multipleRate  : multipleRate = requestCheckvar(request("multipleRate"),32)
dim decimalPointLen  : decimalPointLen = requestCheckvar(request("decimalPointLen"),32)
dim decimalPointCut  : decimalPointCut = requestCheckvar(request("decimalPointCut"),32)
dim currencyUnit_Pos : currencyUnit_Pos = requestCheckvar(request("currencyUnit_Pos"),32)
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

Dim sqlStr
if (mode="edit") then
 
    sqlStr = " update db_shop.dbo.tbl_shop_user"&VbCRLF
    sqlStr = sqlStr & " set exchangeRate="&exchangeRate&VbCRLF
    sqlStr = sqlStr & " ,multipleRate="&multipleRate&VbCRLF
    sqlStr = sqlStr & " ,decimalPointLen="&decimalPointLen&VbCRLF
    sqlStr = sqlStr & " ,decimalPointCut="&decimalPointCut&VbCRLF
    sqlStr = sqlStr & " ,currencyUnit_Pos='"&currencyUnit_Pos&"'"&VbCRLF 
    sqlStr = sqlStr & " where userid='"&shopid&"'"
    dbget.Execute sqlStr
    
    response.write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
    dbget.close() : response.end
end if

dim ochargeuser
set ochargeuser = new COffShopChargeUser
	ochargeuser.FRectShopID = shopid
	ochargeuser.GetOffShopList

Dim IsForeignShop : IsForeignShop=ochargeuser.FItemList(0).IsForeignShop

if Not(IsForeignShop) then
    response.write "<script>alert('�ؿ� �������� ���� �Ǿ� ���� �ʽ��ϴ�.');window.close();</script>"
    dbget.Close() : response.end
end if
%>

<script language='javascript'>

function SavecurrencyUnit(frm){
    if (frm.currencyUnit_Pos.value.length<1){
        alert('ȭ������� �Է��ϼ���.');
        frm.currencyUnit_Pos.focus();
        return;
    }
    
    if (frm.exchangeRate.value.length<1){
        alert('ȯ���� �Է��ϼ���');
        frm.exchangeRate.focus();
        return;
    }
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }    
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmcurrencyUnit" method="post" action="" >
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="mode" value="edit">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">ȭ�����</td>
    <td>
    	<% DrawexchangeRate "currencyUnit_Pos",ochargeuser.FItemList(0).fcurrencyUnit_Pos,"" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">ȯ��</td>
    <td>
        <input type="text" class="text" name="exchangeRate" value="<%= ochargeuser.FItemList(0).FexchangeRate %>" size=9 maxlength=12>
        ex) 1300
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���</td>
    <td>
        <input type="text" class="text" name="multipleRate" value="<%= ochargeuser.FItemList(0).FmultipleRate %>" size=3 maxlength=9>
        ex) 1.5
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">ȭ��Ҽ���</td>
    <td>
         ǥ�� <input type="text" class="text" name="decimalPointLen" value="<%= ochargeuser.FItemList(0).FdecimalPointLen %>" size=2 maxlength=2> �ڸ�
		 ���� <input type="text" class="text" name="decimalPointCut" value="<%= ochargeuser.FItemList(0).FdecimalPointCut %>" size=2 maxlength=2> �ڸ�
    </td>
</tr>
		
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SavecurrencyUnit(frmcurrencyUnit);" class="button"></td>
</tr>
</form>
</table>

<%
set ochargeuser = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ��������
' History : �̻󱸻���
'			2018.09.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<%
dim coupontype, couponidx, i, buf
    coupontype = requestCheckvar(request("coupontype"),32)
    couponidx = requestCheckvar(request("couponidx"),32)

if ((coupontype = "") or (couponidx = "")) then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

'==============================================================================
'��ǰ����
'dim oitemcoupon
'set oitemcoupon = new CUserItemCoupon
'oitemcoupon.FRectUserID = userid
'oitemcoupon.FRectAvailableYN = "Y"
'oitemcoupon.FRectDeleteYN = "Y"
'oitemcoupon.FPageSize = 200
'oitemcoupon.FCurrPage = 1
'oitemcoupon.GetCouponList

'==============================================================================
'���ʽ�����
dim ocscoupon
set ocscoupon = New CCSCenterCoupon
    ocscoupon.FRectBonusCouponIdx = couponidx
    ocscoupon.GetOneCSCenterCoupon
%>
<script type="text/javascript">

function SubmitForm(){
	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		document.frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>���� �������</b>
	</td>
</tr>
</table>

<form name="frm" method="post" action="domodifycoupon.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="copy">
<input type="hidden" name="coupontype" value="<%= coupontype %>">
<input type="hidden" name="couponidx" value="<%= couponidx %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
    <td height="30" width="20%" bgcolor="#f1f1f1">���̵� :</td>
    <td bgcolor="#FFFFFF" width="25%" >
        <b><%= ocscoupon.FOneItem.Fuserid %></b>
    </td>
    <td height="30" width="20%" bgcolor="#f1f1f1">������ :</td>
    <td bgcolor="#FFFFFF"  >
        <%= ocscoupon.FOneItem.Fcouponname %>
    </td>
</tr>
<tr align="left">
    <td height="30" bgcolor="#f1f1f1">���ΰ� :</td>
    <td bgcolor="#FFFFFF" >
        <%= ocscoupon.FOneItem.Fcouponvalue %><%= ocscoupon.FOneItem.GetCouponTypeUnit %>
    </td>
    <td height="30" bgcolor="#f1f1f1">�ּұ��űݾ� :</td>
    <td bgcolor="#FFFFFF" ><%= ocscoupon.FOneItem.Fminbuyprice %> </td>
</tr>
<tr align="left">
    <td height="30" bgcolor="#f1f1f1">��뿩�� :</td>
    <td bgcolor="#FFFFFF" >
        <%= ocscoupon.FOneItem.Fisusing %>
    </td>
    <td height="30" bgcolor="#f1f1f1">�����ֹ���ȣ :</td>
    <td bgcolor="#FFFFFF" ><%= ocscoupon.FOneItem.Forderserial %></td>
</tr>
<tr align="left">
    <td height="30" bgcolor="#f1f1f1">��ȿ�Ⱓ :</td>
    <td bgcolor="#FFFFFF" colspan=3>
        <acronym title="<%= ocscoupon.FOneItem.Fstartdate %>"><%= Left(ocscoupon.FOneItem.Fstartdate,10) %></acronym> ~ <acronym title="<%= ocscoupon.FOneItem.Fexpiredate %>"><%= Left(ocscoupon.FOneItem.Fexpiredate,10) %></acronym>
    </td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF" colspan=4>
        <% if (ocscoupon.FOneItem.Fisusing = "Y") and (ocscoupon.FOneItem.Fdeleteyn <> "Y") then %>
            <input type="button" value="�������" onClick="SubmitForm();" class="button">
        <% end if %>
        <input type="button" value=" â �� �� " onClick="self.close()" class="button">
	</td>
</tr>
</table>
</form>

<%
'set OUserInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
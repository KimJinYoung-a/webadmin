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
    response.write "<script type='text/javascript'>alert('�߸��� �����Դϴ�.'); history.back();</script>"
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

'==============================================================================
dim totay, expireday, baseday, daybeforeonemonth
dim basedayadd1, basedayadd5, basedayadd7, basedayadd14, basedayadd30
totay = Left(now(), 10)
expireday = Left(ocscoupon.FOneItem.Fexpiredate,10)
baseday = expireday

'baseday = totay
'if (expireday > totay) then
'	baseday = expireday
'end if

basedayadd1 = Left(DateAdd("d", 1, baseday), 10)
basedayadd5 = Left(DateAdd("d", 5, baseday), 10)
basedayadd7 = Left(DateAdd("d", 7, baseday), 10)
basedayadd14 = Left(DateAdd("d", 14, baseday), 10)
basedayadd30 = Left(DateAdd("d", 30, baseday), 10)
daybeforeonemonth = Left(DateAdd("d", -300, totay), 10)
%>
<script type="text/javascript">

function SubmitForm(){
	var ischecked = false;
	for(var i=0; i<document.frm.extendday.length; i++){
		if(document.frm.extendday[i].checked) { ischecked = true; }
	}

	if (ischecked == false) {
		alert("������ ���ڸ� �����ϼ���.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		document.frm.submit();
	}
}

function ChangeDate(d){
	var div = document.getElementById('datechanged');

	switch (d) {
	case 1  :
		div.innerHTML = "[������ : <%= basedayadd1 %>]";
		break;
	case 5  :
		div.innerHTML = "[������ : <%= basedayadd5 %>]";
		break;
	case 7  :
		div.innerHTML = "[������ : <%= basedayadd7 %>]";
		break;
	case 14 :
		div.innerHTML = "[������ : <%= basedayadd14 %>]";
		break;
	case 30 :
		div.innerHTML = "[������ : <%= basedayadd30 %>]";
		break;
	default    :
		div.innerHTML = "";
		break;
	}

	div.innerHTML = "<font color=red>" + div.innerHTML + "</font>";
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>���� �Ⱓ����</b>
	</td>
</tr>
</table>

<form name="frm" method="post" action="domodifycoupon.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="expiredate">
<input type="hidden" name="coupontype" value="<%= coupontype %>">
<input type="hidden" name="couponidx" value="<%= couponidx %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="margin:0px;">
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
		<table width=100% class="a">
		<tr>
			<td width=35%><acronym title="<%= ocscoupon.FOneItem.Fstartdate %>"><%= Left(ocscoupon.FOneItem.Fstartdate,10) %></acronym> ~ <acronym title="<%= ocscoupon.FOneItem.Fexpiredate %>"><%= Left(ocscoupon.FOneItem.Fexpiredate,10) %></acronym></td>
			<td><div id="datechanged"></div></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" bgcolor="#f1f1f1">�Ⱓ���� :</td>
	<td bgcolor="#FFFFFF" colspan=3>
		[������ : <%= baseday %>]
		<input type=radio name=extendday value=1 onclick="ChangeDate(1)"> 1��
		<input type=radio name=extendday value=5 onclick="ChangeDate(5)"> 5��
		<input type=radio name=extendday value=7 onclick="ChangeDate(7)"> 7��
		<input type=radio name=extendday value=14 onclick="ChangeDate(14)"> 14��
		<input type=radio name=extendday value=30 onclick="ChangeDate(30)"> 30��
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF" colspan=4>
		<% if (ocscoupon.FOneItem.Fisusing <> "Y") and (ocscoupon.FOneItem.Fdeleteyn <> "Y") then %>
			<input type="button" value="�����ϱ�" onClick="SubmitForm();" class="button">
		<% end if %>
		<input type="button" value=" â�ݱ� " onClick="self.close()" class="button">
	</td>
</tr>
</table>
</form>

<%
'set OUserInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
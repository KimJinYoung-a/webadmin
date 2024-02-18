<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ���ڵ� ��ǰ�˻�
' History : 2009.04.07 ������ ����
'			2013.02.13 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
if C_IS_OWN_SHOP or C_IS_SHOP then
	IS_HIDE_BUYCASH = True
end if

dim page, chargeid, shopid, ErrStr, barcode, digitflag, research, itemgubun, itemid,itemoption
dim sqlStr, menupos
	shopid = requestCheckVar(request("shopid"),32)
	barcode = requestCheckVar(request("barcode"),32)
	digitflag = requestCheckVar(request("digitflag"),32)
	research = requestCheckVar(request("research"),2)
	menupos = requestCheckVar(request("menupos"),10)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if

	chargeid = request("chargeid")
else
	'/��ü
	if (C_IS_Maker_Upche) then
		chargeid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if

		chargeid = requestCheckVar(request("chargeid"),32)
	end if
end if

if trim(barcode)<>"" then

	'//���ڵ尡 �������, ������ڵ�� �ʼ��� �˻�
	sqlStr = "select top 1"
	sqlStr = sqlStr + " itemgubun,shopitemid,itemoption"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where extbarcode='" + trim(barcode) + "'"

	'response.write sqlStr & "<Br>"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("shopitemid")
		itemoption = rsget("itemoption")
	end if
	rsget.Close
end if

'if itemid = "" then
'	itemgubun = Left(barcode,2)
'	itemid = Mid(barcode,3,6)	'CLng(Mid(barcode,3,6))
'	itemoption = Right(barcode,4)
'end if

if itemid = "" then		'//��ǰ�ڵ尡 ���°͸� �ؿ� Ÿ����		'//12�ڸ� 14�ڸ� ���ڵ� �г�	'/2014.02.11 �ѿ�� �߰�
	if BF_IsMaybeTenBarcode(barcode) then
	    itemgubun 	= BF_GetItemGubun(barcode)
		itemid 		= BF_GetItemId(barcode)
		itemoption 	= BF_GetItemOption(barcode)
	end if
end if

dim ioffitem
set ioffitem = new COffShopItem
	ioffitem.FRectShopid = shopid
	ioffitem.FRectDesigner = chargeid
	ioffitem.FRectItemgubun	= itemgubun
	ioffitem.FRectItemId	= itemid
	ioffitem.FRectItemOption= itemoption

	if (itemgubun<>"") and (CStr(itemid)<>"") and (itemoption<>"") then
		if not(isNumeric(itemid)) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('���ڵ尡 �߸��Ǿ����ϴ�');"
			response.write "	location.href='"&refer&"'"
			response.write "</script>"
			dbget.close()	:	response.end
		end if

		ioffitem.GetOffLineJumunByOneItemCode
	end if
%>

<script type='text/javascript'>

function search(frm){
	frm.submit();
}

function AddArr(upfrm){
    //���԰� ���ް� ��������
	opener.ReActItems(upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.buycasharr.value,upfrm.suplycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value);

}


function GetOnLoad(){
	document.frm.barcode.focus();
	document.frm.barcode.select();

}
window.onload = GetOnLoad;

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="shopid"value="<%= shopid %>">
<input type="hidden" name="chargeid"value="<%= chargeid %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="digitflag" value="<%= digitflag %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<% if digitflag="P" then %>
			<%= chargeid %> ���
		<% elseif digitflag="M" then %>
			<%= chargeid %> ��ǰ
		<% elseif digitflag="itemadd" then %>
			<%= chargeid %> ��ǰ�߰�
		<% end if %>

		<p align="right">
		���ڵ� :
		<input type="text" name="barcode" value="" size="16" maxlength="20" AUTOCOMPLETE="off">
		</p>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="search(frm);"></td>
</tr>
</form>
</table>

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<% if ioffitem.FResultCount>0 then %>
	<% if IsNULL(ioffitem.FOneItem.Fdefaultsuplymargin) then %>
		<%
		ErrStr = "[��� �ȵ� �귣�� �Դϴ�. ���� ������ ��� ���.]"
		%>
	<% elseif ioffitem.FOneItem.Fchargediv<>"6" and ioffitem.FOneItem.Fchargediv<>"8" then %>
		<%
		ErrStr = "[��ü��Ź�̳� ��ü���Ը� ��밡���մϴ�." +  ioffitem.FOneItem.Fchargediv + "]"
		%>
	<% else %>
	<form name="upfrm" >
		<input type="hidden" name="itemgubunarr" value="<%= ioffitem.FOneItem.FItemgubun %>|">
		<input type="hidden" name="itemarr" value="<%= ioffitem.FOneItem.Fshopitemid %>|">
		<input type="hidden" name="itemoptionarr" value="<%= ioffitem.FOneItem.Fitemoption %>|">
		<input type="hidden" name="sellcasharr" value="<%= ioffitem.FOneItem.Fshopitemprice %>|">
		<input type="hidden" name="suplycasharr" value="<%= ioffitem.FOneItem.GetOfflineSuplycash %>|"> <!-- ���ް� -->
		<input type="hidden" name="buycasharr" value="<%= ioffitem.FOneItem.GetOfflineBuycash %>|"> <!-- ���԰� -->

		<% if digitflag<>"P" and digitflag<>"itemadd" then %>
			<input type="hidden" name="itemnoarr" value="-1|">
		<% else %>
			<input type="hidden" name="itemnoarr" value="1|">
		<% end if %>

		<input type="hidden" name="itemnamearr" value='<%= replace(ioffitem.FOneItem.Fshopitemname,"'","") %>|'>
		<input type="hidden" name="itemoptionnamearr" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>|">
		<input type="hidden" name="designerarr" value="<%= ioffitem.FOneItem.Fmakerid %>|">
	</form>

	<script type='text/javascript'>
		AddArr(upfrm);
	</script>

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ioffitem.FResultCount %></b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<font color="blue">
			[<%= ioffitem.FOneItem.Fmakerid %>] <%= ioffitem.FOneItem.Fshopitemname %> <%= ioffitem.FOneItem.Fshopitemoptionname %>
			</font> �߰� �Ϸ�
		</td>
	</tr>
	</table>
	<% end if %>
<% elseif research="on" then %>
	<%
	ErrStr = "[�˻������ �����ϴ�.]"
	%>
<% end if %>

<% if ErrStr<>"" then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ioffitem.FResultCount %></b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<font color="red"><%= ErrStr %></font>
		</td>
	</tr>
	</table>

	<script type='text/javascript'>
		alert('<%= ErrStr %>');
	</script>
<% end if %>

<%
set ioffitem = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

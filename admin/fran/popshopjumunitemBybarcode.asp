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

dim suplyer, shopid, barcode, idx, ErrStr, research, digitflag, menupos
dim itemgubun, itemid,itemoption, sqlStr, ioffitem, isusing, shopsuplycash, buycash, foreign_suplycash
	isusing = request("isusing")
	suplyer = request("suplyer")
	shopid = request("shopid")
	barcode = request("barcode")
	idx = request("idx")
	research = request("research")
	digitflag = request("digitflag")
	menupos = request("menupos")

if digitflag="" then digitflag="P"

if C_ADMIN_USER then

'/����
elseif (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if

	isusing="Y"
else
	'/��ü
	if (C_IS_Maker_Upche) then
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if

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

if itemid = "" then
	itemgubun 	= BF_GetItemGubun(barcode)
	itemid 		= BF_GetItemId(barcode)
	itemoption 	= BF_GetItemOption(barcode)
end if

set ioffitem = new COffShopItem
	ioffitem.FRectShopid = shopid
	ioffitem.FRectItemgubun	= itemgubun
	ioffitem.FRectItemId	= itemid
	ioffitem.FRectItemOption= itemoption
	ioffitem.frectisusing = isusing

	''rw shopid & " " & itemgubun & " " & itemid & " " & itemoption & " " & BF_GetItemId(barcode)
	if (itemgubun<>"") and (CStr(itemid)<>"") and (itemoption<>"") then
		ioffitem.GetOffLineJumunByOneItemCode
	end if

%>

<script language='javascript'>

function search(frm){
	frm.submit();
}

function AddArr(upfrm){
	<% if (digitflag = "MV") then %>
	opener.ReActItems(upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value);
	<% else %>
	opener.ReActItems(
		'<%= idx %>'
		, upfrm.itemgubunarr.value
		, upfrm.itemarr.value
		, upfrm.itemoptionarr.value
		, upfrm.sellcasharr.value
		, upfrm.suplycasharr.value
		, upfrm.buycasharr.value
		, upfrm.itemnoarr.value
		, upfrm.itemnamearr.value
		, upfrm.itemoptionnamearr.value
		, upfrm.designerarr.value
		, upfrm.foreign_sellcasharr.value
		, upfrm.foreign_suplycasharr.value);
	<% end if %>
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
<input type="hidden" name="suplyer"value="<%= suplyer %>">
<input type="hidden" name="shopid"value="<%= shopid %>">
<input type="hidden" name="idx"value="<%= idx %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="digitflag" value="<%= digitflag %>">

<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<% if digitflag="P" then %>
			���
		<% elseif digitflag="M" then %>
			��ǰ
		<% elseif digitflag="MV" then %>
			�̵�
		<% end if %>

		<p align="right">
		���ڵ� :
		<input type="text" name="barcode" value="<%= barcode %>" size="13" maxlength="20" AUTOCOMPLETE="off">
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
	<% if IsNULL(ioffitem.FOneItem.Fdefaultsuplymargin) or IsNULL(ioffitem.FOneItem.FShopMargin) then %>
		<%
		ErrStr = "[��� �ȵ� �귣�� �Դϴ�. ���� ������ ��� ���.]"
		%>
	<% elseif ioffitem.FOneItem.Fchargediv<>"2" and ioffitem.FOneItem.Fchargediv<>"4" and ioffitem.FOneItem.Fchargediv<>"5" then %>
		<%
		ErrStr = "[��ü��Ź�̳� ��ü������ ����� �� �����ϴ�." +  ioffitem.FOneItem.Fchargediv + "]"
		%>
	<% else %>
		<%
	    ''���� Ʈ���̵�.. ���� 10%..
	    if shopid="streetshop881" then
	        ioffitem.FOneItem.Fshopbuyprice = 0
	    end if
		shopsuplycash = ioffitem.FOneItem.GetFranchiseSuplycash
		buycash		  = ioffitem.FOneItem.GetFranchiseBuycash

		if ioffitem.Floginsite="WSLWEB" then
			'/ �ؿ� ���. �����ܿ��� ��ǰ���̺� ����� ������� �����ؼ� ó�� ���Ѱ� �־���
			if ioffitem.FOneItem.Fforeign_suplyprice="" or isnull(ioffitem.FOneItem.Fforeign_suplyprice) or ioffitem.FOneItem.Fforeign_suplyprice=0 then
				foreign_suplycash = shopsuplycash
			else
				foreign_suplycash = ioffitem.FOneItem.Fforeign_suplyprice
			end if
		end if
		%>

		<form name="upfrm" >
		<input type="hidden" name="itemgubunarr" value="<%= ioffitem.FOneItem.FItemgubun %>|">
		<input type="hidden" name="itemarr" value="<%= ioffitem.FOneItem.Fshopitemid %>|">
		<input type="hidden" name="itemoptionarr" value="<%= ioffitem.FOneItem.Fitemoption %>|">
		<input type="hidden" name="sellcasharr" value="<%= ioffitem.FOneItem.Fshopitemprice %>|">
		<input type="hidden" name="suplycasharr" value="<%= shopsuplycash %>|">
		<% if IS_HIDE_BUYCASH = True then %>
		<input type="hidden" name="buycasharr" value="-1|"> <!-- ���԰� -->
		<% else %>
		<input type="hidden" name="buycasharr" value="<%= buycash %>|"> <!-- ���԰� -->
		<% end if %>
		<input type="hidden" name="foreign_sellcasharr" value="<%= getdisp_price(ioffitem.FOneItem.Fforeign_sellprice, ioffitem.fcurrencyChar) %>|">
		<input type="hidden" name="foreign_suplycasharr" value="<%= getdisp_price(foreign_suplycash, ioffitem.fcurrencyChar) %>|">

		<% if digitflag<>"P" and (digitflag <> "MV") then %>
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

	<script language='javascript'>
		alert('<%= ErrStr %>');
	</script>
<% end if %>

<%
set ioffitem = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

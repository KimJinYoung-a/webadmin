<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ����� ����Ʈ ��ǰ�߰�
' Hieditor : 2009.04.07 ������ ����
'			 2010.08.04 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim PriceEditEnable
dim chargeid,page,idx,research,react,shopid
dim isusing, itemname, imageon

PriceEditEnable = false

''��ü�ΰ��, ������ �������ΰ��
if (C_IS_Maker_Upche) then
	chargeid = session("ssBctID")
else
	chargeid = requestCheckVar(request("chargeid"),32)
end if

if Not (C_IS_SHOP) and Not (C_IS_Maker_Upche) then PriceEditEnable = true

react = requestCheckVar(request("react"),32)
page  = requestCheckVar(request("page"),10)
idx = requestCheckVar(request("idx"),10)
shopid = requestCheckVar(request("shopid"),32)
isusing = requestCheckVar(request("isusing"),1)
itemname = requestCheckVar(request("itemname"),124)
imageon = requestCheckVar(request("imageon"),2)
research = requestCheckVar(request("research"),2)

if page="" then page=1
if research="" then imageon="on"
if research="" then isusing="on"

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = chargeid
	ioffitem.FRectShopid = shopid
	ioffitem.FRectIpChulId = idx
	ioffitem.FRectOnlyUsing = isusing
	ioffitem.FRectItemName = Html2Db(itemname)

	if chargeid<>"" then
		ioffitem.GetNotIpChulList
	end if

dim i
%>
<script type='text/javascript'>
function ReSearch(page){
	frm.page.value= page;
	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function RefreshParent(){
	opener.ReAct();
}

function AnSearch(frm){
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.shopbuypricearr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (!IsDigit(frm.sellcash.value)){
					alert('�ǸŰ��� ���ڸ� �����մϴ�.');
					frm.sellcash.focus();
					return;
				}

				if (!IsDigit(frm.suplycash.value)){
					alert('���ް��� ���ڸ� �����մϴ�.');
					frm.suplycash.focus();
					return;
				}

				if (!IsInteger(frm.itemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('������ �Է��ϼ���.');
					frm.itemno.focus();
					return;
				}

<% if (C_IS_SHOP) then %>
			// �����ΰ�� ������ -�� �ƴѰ�� ������ ����
			if (frm.itemno.value*-1<0){
				if (!confirm(frm.itemgubun.value + '-' + frm.itemid.value + '-' + frm.itemoption.value + ' : ' + '����' + frm.itemno.value + '\n��ǰ�ΰ�� ���̳ʽ��� �Է��ϼž� �մϴ�. ��� ���� �Ͻðڽ��ϱ�?')){
					frm.itemno.focus();
					return ;
				}
			}
<% end if %>

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.designer.value + "|";
			}
		}
	}

	var ret = confirm('���� ��ǰ�� �԰������� �߰� �Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.submit();
	}

}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="chargeid" value="<%= chargeid %>">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="idx" value="<%= idx %>">
	<tr>
		<td class="a" >
		<input type="checkbox" name="isusing" value="on" <% if isusing="on" then response.write "checked" %> > ��� ��ǰ���˻�
		&nbsp;
		<input type="checkbox" name="imageon" value="on" <% if imageon="on" then response.write "checked" %> > �̹���ǥ��
		&nbsp;
		��ǰ�� : <input type="text" name="itemname" value="<%= itemname %>" size="10" maxlength="32">
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<% if ioffitem.FresultCount>0 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="right">�ѰǼ�: <%= ioffitem.FTotalCount %> &nbsp; <%= Page %>/<%= ioffitem.FTotalPage %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" >* �԰����� �߰��� ��ǰ�� �̰��� ǥ�õ��� �ʽ��ϴ�.</td>
		<td colspan="3" align="right"><input type="button" value="���� ��ǰ �߰�" onclick="AddArr()"></td>
	</tr>
	<% end if %>
	<tr bgcolor="#DDDDFF">
		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
		<% if imageon="on" then %>
		<td width="50">�̹���</td>
		<% end if %>
		<td width="80">BarCode</td>
		<td width="80">��ǰ��</td>
		<td width="80">�ɼǸ�</td>
		<td width="80">�ǸŰ�</td>
		<% if (C_IS_SHOP) then %>
			<td width="80">����<br>���ް�</td>
		<% elseif (C_IS_Maker_Upche) then %>
			<td width="80">�ٹ�����<br>���ް�</td>
		<% else %>
			<td width="80">����<br>���ް�</td>
			<td width="80">�ٹ�����<br>���ް�</td>
		<% end if %>
		<td width="50">����<br>����</td>
		<td width="40">����</td>
	</tr>
	<% for i=0 to ioffitem.FResultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
	<input type="hidden" name="designer" value="<%= ioffitem.FItemList(i).Fmakerid %>">
	<% if Not (PriceEditEnable) then %>
	<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
	<input type="hidden" name="suplycash" value="<%= ioffitem.FItemList(i).GetOfflineBuycash %>">
	<input type="hidden" name="shopbuyprice" value="<%= ioffitem.FItemList(i).GetOfflineSuplycash %>">
	<% end if %>

	<tr bgcolor="#FFFFFF">
		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<% if imageon="on" then %>
		<td ><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<% end if %>
		<td ><%= ioffitem.FItemList(i).GetBarCode %></td>
		<td ><%= ioffitem.FItemList(i).FShopItemName %></td>
		<td ><%= ioffitem.FItemList(i).FShopItemOptionName %></td>
		<% if Not (PriceEditEnable) then %>
			<td align="right"><%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %></td>
			<% if (C_IS_SHOP) then %>
			<td align="right"><%= FormatNumber(ioffitem.FItemList(i).GetOfflineSuplycash,0) %></td>
			<% elseif (C_IS_Maker_Upche) then %>
			<td align="right"><%= FormatNumber(ioffitem.FItemList(i).GetOfflineBuycash,0) %></td>
			<% else %>
			<td align="right"><%= FormatNumber(ioffitem.FItemList(i).GetOfflineSuplycash,0) %></td>
			<td align="right"><%= FormatNumber(ioffitem.FItemList(i).GetOfflineBuycash,0) %></td>
			<% end if %>
		<% else %>
			<td ><input type="text" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>" size="7" maxlength="9"></td>
			<td ><input type="text" name="suplycash" value="<%= ioffitem.FItemList(i).GetOfflineBuycash %>" size="7" maxlength="9"></td>
			<td ><input type="text" name="shopbuyprice" value="<%= ioffitem.FItemList(i).GetOfflineSuplycash %>" size="7" maxlength="9"></td>
		<% end if %>
		<td align="center">
		<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
			<% if (C_IS_SHOP) then %>
			<%= 100-Clng(ioffitem.FItemList(i).GetOfflineSuplycash/ioffitem.FItemList(i).Fshopitemprice*100*100)/100 %> %
			<% elseif (C_IS_Maker_Upche) then %>
			<%= 100-Clng(ioffitem.FItemList(i).GetOfflineBuycash/ioffitem.FItemList(i).Fshopitemprice*100*100)/100 %> %
			<% else %>
			<%= 100-Clng(ioffitem.FItemList(i).GetOfflineSuplycash/ioffitem.FItemList(i).Fshopitemprice*100*100)/100 %> %
			<%= 100-Clng(ioffitem.FItemList(i).GetOfflineBuycash/ioffitem.FItemList(i).Fshopitemprice*100*100)/100 %> %
			<% end if %>
		<% end if %>
		</td>
		<td ><input type="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:ReSearch('<%= ioffitem.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
			<% if i>ioffitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:ReSearch('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ioffitem.HasNextScroll then %>
			<a href="javascript:ReSearch('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<form name="frmArrupdate" method="post" action="do_shopipchulitem.asp">
<input type="hidden" name="mode" value="arrins">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="shopbuypricearr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="designerarr" value="">
</form>

<% if react="true" then %>
	<script type='text/javascript'>
		RefreshParent();
	</script>
<% end if %>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
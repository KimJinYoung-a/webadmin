<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ����� ����Ʈ ��ǰ�߰�
' Hieditor : 2009.04.07 ������ ����
'			 2010.08.04 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%

'' =============================================================================
'' �Ʒ� 3���� �޴� �˻������� �⺻������ �����ؾ� �Ѵ�.
'' (���� ������� ������ǰ, �ֹ�����(����), �ֹ�����(�� ü))
'' =============================================================================
'' /common/offshop/stock/shortagestock_shop.asp
'' /common/offshop/popshopitem2.asp
'' /common/offshop/popshopjumunitem.asp
'' =============================================================================

dim PriceEditEnable ,page, chargeid, shopid ,isusing, itemname, imageon, research , i
dim mode , onlyActive
dim ipgo, sell7days, includepreorder, shortagetype

PriceEditEnable = false

''��ü�ΰ��, ������ �������ΰ��
if (C_IS_Maker_Upche) then
	chargeid = session("ssBctID")
else
	chargeid = request("chargeid")
end if

if Not (C_IS_SHOP) and Not (C_IS_Maker_Upche) then PriceEditEnable = true

	onlyActive = RequestCheckVar(request("onlyActive"),32)
	mode = RequestCheckVar(request("mode"),32)
	page = request("page")
	shopid  = request("shopid")
	isusing = request("isusing")
	itemname = request("itemname")
	imageon = request("imageon")
	research = request("research")

	ipgo = RequestCheckVar(request("ipgo"),32)
	sell7days = RequestCheckVar(request("sell7days"),32)
	includepreorder = RequestCheckVar(request("includepreorder"),32)
	shortagetype = RequestCheckVar(request("shortagetype"),32)


if page="" then page=1
if research="" then imageon="on"
if research="" and isusing="" then isusing="Y"
if mode="" then mode="bybrand"

if research = "" then
	ipgo = "on"
	includepreorder = "on"
	shortagetype = "7"
	mode = "all"
end if

if (C_IS_SHOP) then
    isusing="Y"
end if

dim ioffitem
	set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = chargeid
	ioffitem.FRectshopid = shopid
	ioffitem.FRectOnlyUsing = isusing
	ioffitem.FRectItemName = Html2Db(itemname)
	ioffitem.FRectOnlyActive = onlyActive
	ioffitem.FRectOrder = mode

	ioffitem.FRectIpGoOnly = ipgo
	ioffitem.FRectSell7days = sell7days
	ioffitem.FRectIncludePreOrder = includepreorder
	ioffitem.FRectShortageType = shortagetype

	if chargeid<>"" then
		if (shopid<>"") then
	        ioffitem.GetOffShopItemList
	    else
	        response.write "<script>alert('������ ���� ���� �ʾҽ��ϴ�. ');</script>"
	    end if
	end if
%>

<script language='javascript'>

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
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
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

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
			}
		}
	}

	opener.ReActItems(upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.shopbuypricearr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value);

	//�ʱ�ȭ
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frm.cksel.checked = false;
				frm.itemno.value="0"
			}
		}
	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="chargeid" value="<%= chargeid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="30" maxlength="32">
		<% if not(C_IS_SHOP) then %>
			&nbsp;&nbsp;
			* ��ǰ��뿩�� : <% drawSelectBoxUsingYN "isusing", isusing %>
		<% end if %>

		&nbsp;&nbsp;
		<input type="checkbox" name="imageon" value="on" <% if imageon="on" then response.write "checked" %> > �̹���ǥ��
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
    	<input type=checkbox name="ipgo" <% if ipgo = "on" then response.write " checked" %>>�԰��Ȱ͸�
        <input type=checkbox name="sell7days" <% if sell7days = "on" then response.write " checked" %>>�ֱ�7���Ǹų����ִ°͸�
        <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write " checked" %>>���ֹ����Ժ�����&nbsp;
        ������� : <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write " checked" %>>��ü&nbsp;
        <input type="radio" name="shortagetype" value="3" <% if shortagetype="3" then response.write " checked" %>>3����&nbsp;
        <input type="radio" name="shortagetype" value="7" <% if shortagetype="7" then response.write " checked" %>>7����&nbsp;
        <input type="radio" name="shortagetype" value="14" <% if shortagetype="14" then response.write " checked" %>>14����&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" name="mode" value="all" <% if mode="all" then response.write "checked" %> >��ü
		<input type="radio" name="mode" value="by7sell" <% if mode="by7sell" then response.write "checked" %> >����7�� �ǸŻ�ǰ
		<input type="radio" name="mode" value="byevent" <% if mode="byevent" then response.write "checked" %> disabled ><font color=gray>�ٹ����� ��ȹ��ǰ[�غ���]</font>
		<input type="radio" name="mode" value="byrecent" <% if mode="byrecent" then response.write "checked" %> >�Ż�ǰ
		<input type="radio" name="mode" value="byshopfav" <% if mode="byshopfav" then response.write "checked" %> disabled ><font color=gray>���ɻ�ǰ[�غ���]</font>
		<input type="radio" name="mode" value="byetc" <% if mode="byetc" then response.write "checked" %> >��Ÿ�Ҹ�ǰ <!-- 70 -->
		<br>
		<input type="radio" name="mode" value="bybrand" <% if mode="bybrand" then response.write "checked" %> >�귣�庰
		<input type="radio" name="mode" value="byonbest" <% if mode="byonbest" then response.write "checked" %> >�¶��� ����Ʈ
		<!-- <input type="radio" name="mode" value="byonfav" <% if mode="byonfav" then response.write "checked" %> >�¶��� �α��ǰ -->
		<input type="radio" name="mode" value="byoffbest" <% if mode="byoffbest" then response.write "checked" %> >�������� ����Ʈ
		<input type="radio" name="mode" value="byoffbestAll" <% if mode="byoffbestAll" then response.write "checked" %> >�������� ����Ʈ(����)
		&nbsp;&nbsp;
		<input type="checkbox" name="onlyActive" <% if onlyActive="on" then response.write "checked" %>> �¶��� �� ���� ������� ��ǰ������
	</td>
</tr>
</form>
</table>
<p>
<!-- �׼� ���� -->
<% if ioffitem.FresultCount>0 then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="���� ��ǰ �߰�" onclick="AddArr()">
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->
<p>
<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ioffitem.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ioffitem.FTotalCount %></b>
		&nbsp;
		������ : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
	</td>
</tr>
<% end if %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<% if imageon="on" then %>
	<td width="50">�̹���</td>
	<% end if %>
	<td width="80">BarCode</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="60">�ǸŰ�</td>
	<% if (C_IS_SHOP) then %>
		<td width="60">����<br>���ް�</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td width="60">�ٹ�����<br>���ް�</td>
	<% else %>

		<td width="60">�ٹ�����<br>���ް�</td>
		<td width="60">����<br>���ް�</td>
	<% end if %>
	<td width="50">����<br>����</td>
	<td width="40">����</td>
</tr>
<% for i=0 to ioffitem.FResultCount -1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
<% if Not (PriceEditEnable) then %>
<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="suplycash" value="<%= ioffitem.FItemList(i).GetOfflineBuycash %>">
<input type="hidden" name="shopbuyprice" value="<%= ioffitem.FItemList(i).GetOfflineSuplycash %>">
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<% if imageon="on" then %>
	<td ><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<% end if %>
	<td ><%= ioffitem.FItemList(i).GetBarCode %></td>
	<td align="left"><%= ioffitem.FItemList(i).FShopItemName %></td>
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
	<td ><input type="text" class="text" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>" size="7" maxlength="9"></td>
	<td ><input type="text" class="text" name="suplycash" value="<%= ioffitem.FItemList(i).GetOfflineBuycash %>" size="7" maxlength="9"></td>
	<td ><input type="text" class="text" name="shopbuyprice" value="<%= ioffitem.FItemList(i).GetOfflineSuplycash %>" size="7" maxlength="9"></td>
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
	<td ><input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="#FFFFFF">
	<td colspan="11" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:ReSearch('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
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

<form name="frmArrupdate" method="post" action="">
	<input type="hidden" name="mode" value="arrins">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="itemnamearr" value="">
	<input type="hidden" name="itemoptionnamearr" value="">
	<input type="hidden" name="designerarr" value="">
</form>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
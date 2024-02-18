<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : pos ��ǰ����
' Hieditor : 2011.01.13 ������ ����
'			 2011.03.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopLocaleItemcls.asp"-->

<%
dim designer,page,usingyn ,shopid, offgubun ,onexpire ,posshopid
dim research,pricediff,imageview ,itemgubun, itemid, itemname ,i, PriceDiffExists
dim oexchangerate, IsCommaValid
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	pricediff   = RequestCheckVar(request("pricediff"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	onexpire    = RequestCheckVar(request("onexpire"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),2)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID

    ''�������� ��� �� ���̵� �ݵ�� �����Ǿ���.
    if (shopid="") then
        response.write "���� ���̵� �������� �ʾҽ��ϴ�. ������ ���ǿ��."
        dbget.close() : response.end
    end if
end if

IF (designer="") then
    ''��ǥ �귣��(POS��ǰ)
    designer = getDefaultPosBrand(shopid)
end if

if page="" then page=1
if research<>"on" then usingyn="Y"

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectOnlineExpiredItem = onexpire

	'' gubun 00 is Maybe Pos Item
	ioffitem.FRectItemgubun="00"

	if (designer<>"") then
	    ioffitem.GetOffNOnLineShopItemList
	end if

if (shopid="") and (designer<>"") then
    posshopid = getShopIDbyPosBrand(designer)
end if

if (posshopid="") and (shopid<>"") then posshopid=shopid

IsCommaValid = false

set oexchangerate = new COffShopLocale
	oexchangerate.frectuserid = posshopid
    if posshopid <> "" then
    	oexchangerate.fexchangeratecheck()
    	IsCommaValid = oexchangerate.foneitem.fcurrencyUnit<>"WON" and oexchangerate.foneitem.fcurrencyUnit<>"KRW" and oexchangerate.foneitem.fcurrencyUnit<>""
    end if
set oexchangerate = Nothing
%>

<script language='javascript'>

function NotUsingCheckAll(){
    var frm;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
		    if (frm.isusing[0].checked==true){
		        frm.isusing[1].checked = true;
		        frm.cksel.checked = true;
		        AnCheckClick(frm.cksel);
		    }
		}
	}
}

function popOffItemEdit(ibarcode,itemgubun,itemid,itemoption,makerid, shopid){

	var popwin = window.open('/common/offshop/popoffitemreg_Etc.asp?barcode=' + ibarcode +'&itemgubun=' + itemgubun+'&itemid=' + itemid+'&itemoption=' + itemoption+'&makerid=' + makerid + '&shopid=' +shopid,'offitemedit','width=500,height=350,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('/admin/offshop/popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function OffEtcItemReg(makerid, shopid){

	var subwin;

	if (confirm('�������� POS���� ��ǰ�� ����ϴ� �۾��Դϴ�.����Ͻðڽ��ϱ�?')){
		subwin = window.open('popoffitemreg_Etc.asp?makerid=' + makerid + '&shopid=' +shopid,'window_reg','width=500,height=350,scrollbars=yes,resizable=yes,status=no');
		subwin.focus();
	}
}

function ReSearch(page){
	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function ModiArr(){
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
	upfrm.itempricearr.value = "";
	upfrm.itemsuplyarr.value = "";

	upfrm.extbarcodearr.value = "";

	upfrm.shopbuypricearr.value = "";
    upfrm.shopid.value="<%= posshopid %>";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

                if (!<%= CHKIIF(IsCommaValid,"IsDouble","IsDigit") %>(frm.tx_orgsellprice.value)){
					alert('�Һ��ڰ����� ���ڸ� �����մϴ�.');
					frm.tx_orgsellprice.focus();
					return;
				}

				if (!<%= CHKIIF(IsCommaValid,"IsDouble","IsDigit") %>(frm.tx_sellcash.value)){
					alert('�ǸŰ��� ���ڸ� �����մϴ�.');
					frm.tx_sellcash.focus();
					return;
				}

				if (frm.tx_sellcash.value<1){
					if (!confirm('�ǸŰ��� 1���� �۽��ϴ�. ��� �����Ͻðڽ��ϱ�?')){
					    frm.tx_sellcash.focus();
					    return;
					}
				}

                if (frm.tx_orgsellprice.value*1<frm.tx_sellcash.value*1){
					alert('�Һ��ڰ��� �ǸŰ����� Ŀ���մϴ�..');
					frm.tx_orgsellprice.focus();
					return;
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";

				upfrm.orgsellpricearr.value = upfrm.orgsellpricearr.value + frm.tx_orgsellprice.value + "|";
				upfrm.itempricearr.value = upfrm.itempricearr.value + frm.tx_sellcash.value + "|";
				upfrm.itemsuplyarr.value = upfrm.itemsuplyarr.value + frm.tx_suplycash.value + "|";

				upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.tx_shopbuyprice.value + "|";

				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";

				if (frm.isusing[0].checked){
					upfrm.isusingarr.value = upfrm.isusingarr.value + "Y" + "|";
				}else{
					upfrm.isusingarr.value = upfrm.isusingarr.value + "N" + "|";
				}
			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.mode.value = "arrmodi";
		upfrm.submit();
	}
}

function samePrice(frm){
    frm.tx_orgsellprice.value=frm.oldonlineorgprice.value*1 + frm.oldonlineOptAddprice.value*1;  //�Һ��ڰ�
	frm.tx_sellcash.value=frm.oldonlineprice.value*1 + frm.oldonlineOptAddprice.value*1;         //�ǸŰ�

	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function getOnLoad(){
    alert('��ǥ ID�� �������� �ʾҽ��ϴ�. - ��ǥ ID ���� �� ��밡��. <%= shopid %>');
}

<% if (designer="") then %>
	window.onload = getOnLoad;
<% end if %>

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    <% if (Not C_IS_SHOP) then %>
			��ǥ �귣�� : <% FnDrawOptPosBrand shopid,"designer",designer  %>
		<% else %>
			<input type="hidden" name="designer" value="<%= designer %>">
		<% end if %>
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
		&nbsp;
		��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
     	�������:<% drawSelectBoxUsingYN "usingyn", usingyn %>
     	&nbsp;
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	<% if (designer<>"") then %>
	    <input type="button" class="button" value="POS ��ǰ���" onclick="OffEtcItemReg('<%= designer %>','<%= posshopid %>')">
	<% end if %>
	</td>
	<td align="right">
	    <% if ioffitem.FresultCount>0 then %>
		<input type="button" class="button" value="���û�ǰ �ϰ�����" onclick="ModiArr()">
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">

		�˻���� : <b><%= ioffitem.FTotalcount %></b>
		&nbsp;
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>

		<b><%= page %> / <%= ioffitem.FTotalpage %></b>

		<% if (ioffitem.FTotalpage - ioffitem.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<% if (imageview<>"") then %>
	<td width="50">�̹���</td>
	<% end if %>
	<td width="70">�귣��ID</td>
	<td width="90">��ǰ�ڵ�</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="20"></td>
	<td width="50">�Һ��ڰ�</td>
	<td width="50">�ǸŰ�</td>
	<td width="40">������<br>(%)</td>
	<td width="50">���԰�</td>
	<td width="30">����<br>����</td>
	<td width="60">������ڵ�</td>
	<td width="70">��� ����<br><input class="button" type="button" value="������" onClick="NotUsingCheckAll();"></td>
</tr>
<% for i=0 to ioffitem.FresultCount -1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
<input type="hidden" name="oldonlineprice" value="<%= ioffitem.FItemlist(i).FOnLineItemprice %>">
<input type="hidden" name="oldonlineorgprice" value="<%= ioffitem.FItemlist(i).FOnLineItemOrgprice %>">
<input type="hidden" name="oldonlineOptAddprice" value="<%= ioffitem.FItemlist(i).FOnlineOptaddprice %>">
<input type="hidden" name="tx_shopbuyprice" value="0">
<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<% if (imageview<>"") then %>
		<td width="50">
			<a href="javascript:popOffImageEdit('<%= ioffitem.FItemlist(i).GetBarCode %>')">
			<img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></a>
		</td>
	<% end if %>
	<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
	<td>
		<a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>','<%= ioffitem.FItemlist(i).Fitemgubun %>','<%= ioffitem.FItemlist(i).Fshopitemid %>','<%= ioffitem.FItemlist(i).Fitemoption %>','<%= designer %>','<%= posshopid %>');">
		<%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></a>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopItemName %>
		<% if ioffitem.FItemlist(i).Fitemoption<>"0000" then %>
			<font color="blue">[<%= ioffitem.FItemlist(i).FShopitemOptionname %>]</font>
		<% end if %>

		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <br>�ɼ��߰��ݾ�: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
    <% PriceDiffExists = false %>
	<td align="center" >
	    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
	    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice) or (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice) then %>
	    <input type="button" class="button" value=">" onclick="samePrice(frmBuyPrc_<%= i %>);">
	    <% PriceDiffExists = true %>
	    <% end if %>
	    <% end if %>
	</td>
    <td align="right" >
        <input type="text" class="text" name="tx_orgsellprice" value="<%= ioffitem.FItemlist(i).FShopItemOrgprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
        <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
        <% if (ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemOrgprice)  then %>
            <font color="red"><strong><%= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
        <% else %>
            <% if (PriceDiffExists) then %>
            <%= ioffitem.FItemlist(i).FOnlineitemorgprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
            <% end if %>
        <% end if %>
        <% end if %>
    </td>
	<td align="right" >
	    <input type="text" class="text" name="tx_sellcash" value="<%= ioffitem.FItemlist(i).FShopItemprice %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
	    <% if (ioffitem.FItemlist(i).FItemGubun="10") then %>
        <% if (ioffitem.FItemlist(i).FOnLineItemprice+ ioffitem.FItemlist(i).FOnlineOptaddprice<>ioffitem.FItemlist(i).FShopItemprice)  then %>
	        <font color="red"><strong><%= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %></strong></font>
	    <% else %>
	        <% if (PriceDiffExists) then %>
	        <%= ioffitem.FItemlist(i).FOnLineItemprice + ioffitem.FItemlist(i).FOnlineOptaddprice %>
	        <% end if %>
        <% end if %>
        <% end if %>
	</td>
	<td align="center" >
        <% if (ioffitem.FItemlist(i).FShopItemOrgprice<>0) then %>
            <% if ioffitem.FItemlist(i).FShopItemOrgprice<>ioffitem.FItemlist(i).FShopItemprice then %>
            OFF:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FShopItemOrgprice-ioffitem.FItemlist(i).FShopItemprice)/ioffitem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
            <% end if %>
	    <% end if %>

	    <% if (ioffitem.FItemlist(i).FOnlineitemorgprice<>0) then %>
	        <% if ioffitem.FItemlist(i).FOnlineitemorgprice<>ioffitem.FItemlist(i).FOnLineItemprice then %>
            ON:<font color="#FF3333"><%= CLng((ioffitem.FItemlist(i).FOnlineitemorgprice-ioffitem.FItemlist(i).FOnLineItemprice)/ioffitem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
            <% end if %>
	    <% end if %>
	</td>
	<td align="right" >
		<input type="text" name="tx_suplycash" value="<%= ioffitem.FItemlist(i).Fshopsuplycash %>" size="6" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)">
	</td>
	<td align="right" >
	<% if (ioffitem.FItemlist(i).FShopItemprice<>0) and (ioffitem.FItemlist(i).Fshopsuplycash<>0) then %>
		<font color="blue"><%= CLng((ioffitem.FItemlist(i).FShopItemprice-ioffitem.FItemlist(i).Fshopsuplycash)/ioffitem.FItemlist(i).FShopItemprice*100) %>%</font>
	<% end if %>
	</td>
	<td align="right" ><input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="12" maxlength="20" style="border:1px #999999 solid; " onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
	<td align="left" >
	<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
	<input type="radio" name="isusing" value="Y" checked onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
	<input type="radio" name="isusing" value="N" onclick="CheckThis(frmBuyPrc_<%= i %>)">N
	<% else %>
	<input type="radio" name="isusing" value="Y" onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
	<input type="radio" name="isusing" value="N" checked onclick="CheckThis(frmBuyPrc_<%= i %>)"><font color="red">N</font>
	<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="18" align="center">
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
		<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ioffitem.HasNextScroll then %>
		<a href="javascript:ReSearch('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>

<form name="frmArrupdate" method="post" action="popoffitemreg_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="orgsellpricearr" value="">
	<input type="hidden" name="itempricearr" value="">
	<input type="hidden" name="itemsuplyarr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="extbarcodearr" value="">
	<input type=hidden name=shopid value="">
	<input type=hidden name=isforeignshop value="<%= chkIIF(IsCommaValid,"on","") %>">
</form>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
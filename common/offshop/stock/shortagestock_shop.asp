<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ����������������ǰ
' Hieditor : 2011.07.13 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
if C_IS_OWN_SHOP or C_IS_SHOP then
	IS_HIDE_BUYCASH = True
end if

'' =============================================================================
'' �Ʒ� 3���� �޴� �˻������� �⺻������ �����ؾ� �Ѵ�.
'' (���� ������� ������ǰ, �ֹ�����(����), �ֹ�����(�� ü))
'' =============================================================================
'' /common/offshop/stock/shortagestock_shop.asp
'' /common/offshop/popshopitem2.asp
'' /common/offshop/popshopjumunitem.asp
'' =============================================================================

dim page , shopid , isusing , makerid , itemid , itemname , generalbarcode , i , sell7days
dim cdl , cdm , cds , shortagetype , comm_cd ,includepreorder ,research , parameter , ipgo
dim onlinemwdiv, ordby, sellyn
dim forcemakerid
    page = requestCheckVar(getNumeric(request("page")),10)
    research = requestCheckVar(request("research"),2)
    isusing = requestCheckVar(request("isusing"),1)
    makerid = requestCheckVar(request("makerid"),32)
    itemid = requestCheckVar(request("itemid"),10)
    itemname = requestCheckVar(request("itemname"),64)
    generalbarcode = requestCheckVar(request("generalbarcode"),20)
    comm_cd = requestCheckVar(request("comm_cd"),16)
    cdl = requestCheckVar(getNumeric(request("cdl")),3)
    cdm = requestCheckVar(getNumeric(request("cdm")),3)
    cds = requestCheckVar(getNumeric(request("cds")),3)
    shortagetype = requestCheckVar(request("shortagetype"),10)
    includepreorder = requestCheckVar(request("includepreorder"),2)
    sell7days = requestCheckVar(request("sell7days"),2)
    ipgo = requestCheckVar(request("ipgo"),2)
	shopid = requestCheckVar(request("shopid"),32)
	onlinemwdiv = requestCheckVar(request("onlinemwdiv"),1)
	ordby = requestCheckVar(request("ordby"),32)
	sellyn      = requestCheckvar(request("sellyn"),2)
    forcemakerid = RequestCheckVar(request("forcemakerid"),32)

if page="" then page=1
if (research<>"on") and (sellyn="") then
    sellyn = "YS"
end if
if (research<>"on") and (includepreorder="") then
    includepreorder = "on"
end if
if (research<>"on") and (ipgo="") then
    ipgo = "on"
end if
if (research<>"on") and (shortagetype="") then
    shortagetype = 7
end if
if (research<>"on") and (isusing="") then
    isusing = "Y"
end if

if (research<>"on") and (ordby="") then
    ordby = "BI"
end if

if C_ADMIN_USER then

'/�����ϰ�� ���� ���常 ��밡��
elseif (C_IS_SHOP) then
	'/���α��� ���� �̸�
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	if (C_IS_Maker_Upche) then
		shopid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

'if shopid = "" then shopid = "streetshop011"

parameter = "page="&page&"&research="&research&"&shopid="&shopid&"&isusing="&isusing&"&makerid="&makerid&"&itemid="&itemid&"&itemname="&itemname&"&sell7days="&sell7days&""
parameter = parameter & "&generalbarcode="&generalbarcode&"&comm_cd="&comm_cd&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&shortagetype="&shortagetype&"&includepreorder="&includepreorder
parameter = parameter & "&ipgo="&ipgo&"&sellyn="&sellyn&""

dim oshortage
set oshortage  = new cshortagestock_list
    oshortage.FPageSize = 100
    oshortage.FCurrPage = page
    oshortage.frectcdl = cdl
    oshortage.frectcdm = cdm
    oshortage.frectcds = cds
    oshortage.frectincludepreorder = includepreorder
    oshortage.frectsell7days = sell7days
    oshortage.Frectshopid = shopid
    oshortage.Frectisusing = isusing
    if forcemakerid = "" then
        oshortage.Frectmakerid = makerid
    else
        oshortage.Frectmakerid = forcemakerid
    end if
    oshortage.Frectitemid = itemid
    oshortage.Frectitemname = itemname
    oshortage.Frectcomm_cd = comm_cd
    oshortage.Frectgeneralbarcode = generalbarcode
    oshortage.Frectshortagetype = shortagetype
    oshortage.Frectipgo = ipgo
	oshortage.FRectOnlineMWdiv = onlinemwdiv
	oshortage.FRectOrder = ordby
	oshortage.FRectSellYN       = sellyn

    if shopid <> "" then
        oshortage.fshortagestock_list
    else
        response.write "<script language='javascript'>"
        response.write "    alert('������ ������ �ּ���');"
        response.write "</script>"
    end if

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if
%>

<script type="text/javascript">

//�����ٿ�ε�
function exceldownload(){
    var exceldownload = window.open('/common/offshop/stock/shortagestock_shop_excel.asp?<%=parameter%>','exceldownload','width=1024,height=768,scrollbars=yes,resizable=yes');
    exceldownload.focus();
}

//�ʿ����Ŭ����
function inputiteno(shortitemno,formi){
    formi.itemno.value=shortitemno;

    formi.cksel.checked=true;
    AnCheckClick(formi.cksel);
}

//�귣��Ŭ����
function searchmakerid(makerid){
    frm.makerid.value=makerid;
    frm.submit();
}

function CheckThis(frm){
    frm.cksel.checked=true;
    AnCheckClick(frm.cksel);
}

//�˻���ư
function reg(page){

    if(frm.itemid.value!=''){
        if (!IsDouble(frm.itemid.value)){
            alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
            frm.itemid.focus();
            return;
        }
    }

    frm.page.value=page;
    frm.submit();
}

//�ǻ�����Է�
function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
        alert('������ �����ϴ�. - ��üƯ�� ��ǰ�� ��� ���� ����.');
        return; //��üƯ�� ��ǰ�� ���?
    <% else %>
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
        popwin.focus();
    <% end if %>
}

var jumunitem = null;

//�ٹ����� �ֹ��� �ۼ�
function jumundirect(){
    var upfrm = document.frmArrupdate;
    var frm; var tmpshopid = '';
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

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.buycasharr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('���� Ʋ�� ������ ���õǾ� �ֽ��ϴ� \n����(�ֹ���)�� ���� �ؾ� �մϴ�.');
	                    return;
                	}
                }

                if (frm.comm_cd.value=="B012" || frm.comm_cd.value=="B022"){
                    alert('��üƯ���̳� ��ü������ �ֹ� �ϽǼ� �����ϴ�.');
                    frm.itemno.focus();
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

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";

                //[db_storage].[dbo].tbl_ordersheet_master�� ���� ������ ��� ���͸��԰��� , ������԰��� ���ٷ���
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopbuyprice.value + "|";
                upfrm.buycasharr2.value = upfrm.buycasharr2.value + frm.shopsuplycash.value + "|";
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
            }
        }
    }

	//�� ���� ����... ������... ���� ��ǰ ���� ����..
    var jumunitem = window.open('','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');

	<% if C_ADMIN_USER then %>
    	upfrm.action='/admin/fran/jumuninput.asp';
    <% else %>
    	upfrm.action='/common/offshop/shop_jumuninput.asp';
    <% end if %>

    upfrm.target='jumunitem';
    upfrm.shopid.value=tmpshopid;
    upfrm.submit();
    jumunitem.focus();

    //���� �˾��� ��� �ִ� ���
	//if(jumunitem != null){
    //}

    //���� �˾��� ���°�� �˾� ����
    //else {
    //	jumunitem = window.open('/admin/fran/jumuninput.asp?suplyer=10x10&shopid=<%=shopid%>','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');
	//}

	//�˾� �ε��� 0.1�� �ڿ�.. ��ǰ ��������... �̷��� �ؾ� ���� ��ǰ ���� ����..
	//window.setTimeout("jumunitem.ReActItems('0',frmArrupdate.itemgubunarr2.value,frmArrupdate.itemidadd2.value,frmArrupdate.itemoptionarr2.value,frmArrupdate.sellcasharr2.value,frmArrupdate.suplycasharr2.value,frmArrupdate.buycasharr2.value,frmArrupdate.itemnoarr2.value,frmArrupdate.itemnamearr2.value,frmArrupdate.itemoptionnamearr2.value,frmArrupdate.designerarr2.value);",500)
	//jumunitem.focus();

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                frm.cksel.checked = false;
                frm.itemno.value="0"
                dL(frm.cksel);
            }
        }
    }
}

//��ü �ֹ��� �ۼ�
function jumundirect_upche(){
    var upfrm = document.frmArrupdate;
    var frm; var tmpshopid = ''; var tmpmakerid = '';
    var pass = false;
    var searchfrm = document.frm;

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            pass = ((pass)||(frm.cksel.checked));
        }
    }

    var ret;

    //if (searchfrm.makerid.value == ''){
    //    alert('�귣��(����ó)�� ������ �ּ���.');
    //    return;
    //}

    if (!pass) {
        alert('���� �������� �����ϴ�.');
        return;
    }

    upfrm.itemgubunarr2.value = "";
    upfrm.itemidadd2.value = "";
    upfrm.itemoptionarr2.value = "";
    upfrm.sellcasharr2.value = "";
    upfrm.suplycasharr2.value = "";
    upfrm.shopbuypricearr2.value = "";
    upfrm.itemnoarr2.value = "";
    upfrm.itemnamearr2.value = "";
    upfrm.itemoptionnamearr2.value = "";
    upfrm.designerarr2.value = "";

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){

                if (tmpshopid==''){
                	tmpshopid = frm.shopid.value;
                } else {
                	if (tmpshopid != frm.shopid.value){
	                    alert('���� Ʋ�� ������ ���õǾ� �ֽ��ϴ� \n����(�ֹ���)�� ���� �ؾ� �մϴ�.');
	                    return;
                	}
                }

                if (tmpmakerid==''){
                	tmpmakerid = frm.makerid.value;
                } else {
                	if (tmpmakerid != frm.makerid.value){
	                    alert('���� Ʋ�� �귣��(����ó)�� ���õǾ� �ֽ��ϴ� \n��ü�ֹ��� ��� �귣��(����ó)�� �����ؾ� �մϴ�');
	                    return;
                	}
                }

                if (frm.comm_cd.value=="B011" || frm.comm_cd.value=="B031"){
                    alert('�ٹ�����Ư���̳� �������� �ֹ� �ϽǼ� �����ϴ�.');
                    frm.itemno.focus();
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

                upfrm.itemgubunarr2.value = upfrm.itemgubunarr2.value + frm.itemgubun.value + "|";
                upfrm.itemidadd2.value = upfrm.itemidadd2.value + frm.itemid.value + "|";
                upfrm.itemoptionarr2.value = upfrm.itemoptionarr2.value + frm.itemoption.value + "|";
                upfrm.sellcasharr2.value = upfrm.sellcasharr2.value + frm.shopitemprice.value + "|";
                upfrm.suplycasharr2.value = upfrm.suplycasharr2.value + frm.shopsuplycash.value + "|";
                upfrm.shopbuypricearr2.value = upfrm.shopbuypricearr2.value + frm.shopbuyprice.value + "|";
                upfrm.itemnoarr2.value = upfrm.itemnoarr2.value + frm.itemno.value + "|";
                upfrm.itemnamearr2.value = upfrm.itemnamearr2.value + frm.itemname.value + "|";
                upfrm.itemoptionnamearr2.value = upfrm.itemoptionnamearr2.value + frm.itemoptionname.value + "|";
                upfrm.designerarr2.value = upfrm.designerarr2.value + frm.makerid.value + "|";
            }
        }
    }

	//�� ���� ����... ������... ���� ��ǰ ���� ����..
    var jumunitem = window.open('','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');

	<% if C_ADMIN_USER then %>
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
    <% else %>
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
    <% end if %>

    upfrm.target='jumunitem';
    upfrm.shopid.value=tmpshopid;
    upfrm.chargeid.value=tmpmakerid;
    upfrm.submit();
    jumunitem.focus();

    //���� �˾��� ��� �ִ� ���
	//if(jumunitem != null){
    //}

    //���� �˾��� ���°�� �˾� ����
    //else {
    //	jumunitem = window.open('/common/offshop/shop_ipchulinput.asp?chargeid=<%=makerid%>&shopid=<%=shopid%>&isreq=Y','jumunitem','width=1024,height=768,scrollbars=yes,resizable=yes');
	//}

	//�˾� �ε��� 0.1�� �ڿ�.. ��ǰ ��������... �̷��� �ؾ� ���� ��ǰ ���� ����..
	//window.setTimeout("jumunitem.ReActItems(frmArrupdate.itemgubunarr2.value,frmArrupdate.itemidadd2.value,frmArrupdate.itemoptionarr2.value,frmArrupdate.sellcasharr2.value,frmArrupdate.suplycasharr2.value,frmArrupdate.shopbuypricearr2.value,frmArrupdate.itemnoarr2.value,frmArrupdate.itemnamearr2.value,frmArrupdate.itemoptionnamearr2.value,frmArrupdate.designerarr2.value);",500)
	//jumunitem.focus();

    for (var i=0;i<document.forms.length;i++){
        frm = document.forms[i];
        if (frm.name.substr(0,9)=="frmBuyPrc") {
            if (frm.cksel.checked){
                frm.cksel.checked = false;
                frm.itemno.value="0"
                dL(frm.cksel);
            }
        }
    }
}

</script>

- �ٹ����� �������� �ֹ�
<Br>&nbsp; &nbsp; ���걸���� �ٹ�����Ư��&������ �̰�, ����(�ֹ���)�� ���� �ؾ� �ֹ��� �ۼ� ���� �մϴ�.
<br>- ��ü �ֹ�
<Br>&nbsp; &nbsp; ���걸���� ��üƯ��&��ü���� �̰� , �귣��(����ó)�� ����(�ֹ���)�� ���� �ؾ� �ֹ��� �ۼ� ���� �մϴ�.
<br>�ʿ����(3��) = (3���Ǹź� x 1) - (��ȿ��� + ���ֹ���)
<br>�ʿ����(7��) = (7���Ǹź� x 1) - (��ȿ��� + ���ֹ���)
<br>�ʿ����(14��) = (7���Ǹź� x 2) - (��ȿ��� + ���ֹ���)
<!--<br>&nbsp; &nbsp; �ʿ����(28��) = (7���Ǹź� x 4) - (��ȿ��� + ���ֹ���)-->

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        ���� :
        <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
        	<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
        <% elseif (C_IS_SHOP) then %>
    		<%= shopid %>
    	<% else %>
        	<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
        <% end if %>
        ��뿩��:<% drawSelectBoxUsingYN "isusing", isusing %>
        �¶����Ǹſ���:<% drawSelectBoxSellYN "sellyn", sellyn %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			&nbsp;
			ON���Ա���:
			<select class="select" name="onlinemwdiv">
				<option></option>
				<option value="M" <% if (onlinemwdiv = "M") then %>selected<% end if %> >����</option>
				<option value="W" <% if (onlinemwdiv = "W") then %>selected<% end if %> >Ư��</option>
				<option value="U" <% if (onlinemwdiv = "U") then %>selected<% end if %> >��ü</option>
			</select>
		<% end if %>

		&nbsp;
        <!-- #include virtual="/common/module/categoryselectbox.asp"-->
    </td>

    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="javascript:reg('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        <% if C_ADMIN_AUTH then %>
        * �귣��[������] <input type="text" class="text" name="forcemakerid" value="<%= forcemakerid %>" size="16" maxlength="32">
        &nbsp;&nbsp;
        <% end if %>
        �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        &nbsp;
        ��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg('');">
        &nbsp;
        ��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
        ������ڵ� :
        <input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
        ������� : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
    	<input type=checkbox name="ipgo" <% if ipgo = "on" then response.write " checked" %> onclick="reg('');">�԰�Ȱ͸�(����)
        <input type=checkbox name="sell7days" <% if sell7days = "on" then response.write " checked" %> onclick="reg('');">�ֱ�7���Ǹų����ִ°͸�
        <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write " checked" %> onclick="reg('');">���ֹ����Ժ�����&nbsp;
        ������ : <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write " checked" %> onclick="reg('');">��ü&nbsp;
        <input type="radio" name="shortagetype" value="3" <% if shortagetype="3" then response.write " checked" %> onclick="reg('');">3����&nbsp;
        <input type="radio" name="shortagetype" value="7" <% if shortagetype="7" then response.write " checked" %> onclick="reg('');">7����&nbsp;
        <input type="radio" name="shortagetype" value="14" <% if shortagetype="14" then response.write " checked" %> onclick="reg('');">14����&nbsp;
        <!--<input type="radio" name="shortagetype" value="28" <%' if shortagetype="28" then response.write " checked" %> onclick="reg('');">28����-->
		&nbsp;
		���ļ��� :
		<select class="select" name="ordby">
			<option value="BI" <% if (ordby = "BI") then %>selected<% end if %> >�귣��</option>
			<option value="I" <% if (ordby = "I") then %>selected<% end if %> >��ǰ�ڵ� ����</option>
		</select>
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">
		<input type="button" class="button" value="�����ٿ�ε�" onclick="exceldownload()">
    </td>
    <td align="right">
    	<input type="button" class="button" value="��ٱ��ϴ��" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="��ٱ��Ϻ���" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
        <% if oshortage.FresultCount>0 then %>
            <input type="button" class="button" value="���ùٷ��ֹ��ۼ�(�ٹ����ٹ���)" onclick="jumundirect()">
        <% end if %>
        <% if oshortage.FresultCount>0 then %>
        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
            	<input type="button" class="button" value="���ùٷ��ֹ��ۼ�(��ü)" onclick="jumundirect_upche()">
            <%' end if %>
        <% end if %>
    </td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="30">
        �˻���� : <b><%= oshortage.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oshortage.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td width=80>����</td>
    <td>��౸��</td>
	<td width=50>�̹���</td>
    <td>�귣��</td>
    <td width=80>��ǰ�ڵ�</td>
    <td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width=50>�ǸŰ�</td>
	<td width=50>���</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
    	<td width=50>���԰�</td>
    <% end if %>

    <td width=30>����<br>�԰�</td>
    <td width=30>��ü<br>�԰�</td>
    <td width=30>�Ǹ�</td>
    <td width=40>�ý���<br>���</td>
    <td width=30>����</td>
	<td width=30>�ǻ�<br>���</td>
    <td width=30>����</td>
    <td width=30>��ȿ<br>���</td>

	<td width=30>OFF<br>3��</td>
	<td width=30>OFF<br>7��</td>
	<td width=30>3��</td>
	<td width=30>7��</td>
	<td width=30>14��</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td width=30>ON<br>����<br>����</td>
		<td width=30>����<br>����<br>����</td>
	<% end if %>

    <td width=40>�ֹ�<br>����</td>
    <td width=90>���</td>
</tr>
<% if oshortage.FresultCount > 0 then %>
<% for i=0 to oshortage.FresultCount -1 %>
<form method="get" action="" name="frmBuyPrc<%=i%>">

<% if oshortage.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
<input type="hidden" name="comm_cd" value="<%= oshortage.FItemlist(i).fcomm_cd %>">
<input type="hidden" name="itemgubun" value="<%= oshortage.FItemlist(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= oshortage.FItemlist(i).fitemid %>">
<input type="hidden" name="itemoption" value="<%= oshortage.FItemlist(i).fitemoption %>">
<input type="hidden" name="shopitemprice" value="<%= oshortage.FItemlist(i).fshopitemprice %>">
<% if IS_HIDE_BUYCASH = True then %>
<input type="hidden" name="shopsuplycash" value="-1">
<% else %>
<input type="hidden" name="shopsuplycash" value="<%= oshortage.FItemlist(i).fshopsuplycash %>">
<% end if %>
<input type="hidden" name="shopbuyprice" value="<%= oshortage.FItemlist(i).fshopbuyprice %>">
<input type="hidden" name="itemname" value="<%= oshortage.FItemlist(i).fshopitemname %>">
<input type="hidden" name="itemoptionname" value="<%= oshortage.FItemlist(i).fshopitemoptionname %>">
<input type="hidden" name="makerid" value="<%= oshortage.FItemlist(i).fmakerid %>">
<input type="hidden" name="shopid" value="<%= oshortage.FItemlist(i).fshopid %>">
    <td >
        <input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
    </td>
    <td >
        <%= oshortage.FItemlist(i).fshopid %>
    </td>
    <td>
        <%= GetJungsanGubunName(oshortage.FItemlist(i).fcomm_cd) %>
    </td>
    <td>
        <img src="<%= oshortage.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0>
    </td>
    <td>
        <a href="javascript:searchmakerid('<%= oshortage.FItemlist(i).fmakerid %>');" onfocus="this.blur()"><%= oshortage.FItemlist(i).fmakerid %></a>
    </td>
    <td>
		<!--
        <a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= oshortage.FItemList(i).Fitemgubun %>&itemid=<%= oshortage.FItemList(i).Fitemid %>&itemoption=<%= oshortage.FItemList(i).Fitemoption %>" target=_blank ><%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %></a>
		-->
		<a href="/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=<%= oshortage.FItemlist(i).fshopid %>&barcode=<%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %>" target=_blank ><%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %></a>
    </td>
    <td align="left">
        <%= oshortage.FItemlist(i).fshopitemname %><Br>
    </td>
    <td align="left">
        <% if oshortage.FItemlist(i).fshopitemoptionname <> "" then %>
            <%=oshortage.FItemlist(i).fshopitemoptionname%>
        <% end if %>
    </td>
    <td align="right">
        <%= FormatNumber(oshortage.FItemlist(i).fshopitemprice,0) %>
    </td>
    <td align="right">
        <%= FormatNumber(oshortage.FItemlist(i).fshopbuyprice,0) %>
		<% if oshortage.FItemList(i).Fshopitemprice<>0 then %>
		<br>(<%= 100-Clng(oshortage.FItemList(i).fshopbuyprice/oshortage.FItemList(i).Fshopitemprice*100*100)/100 %> %)
		<% end if %>
    </td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td align="right">
	        <%= FormatNumber(oshortage.FItemlist(i).fshopsuplycash,0) %>
	    </td>
	<% end if %>

    <td align="right">
        <%= oshortage.FItemlist(i).flogicsipgono + oshortage.FItemlist(i).flogicsreipgono %>    <!--�����԰��ǰ-->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).fbrandipgono + oshortage.FItemlist(i).fbrandreipgono %>		<!--�귣���԰��ǰ-->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).fsellno+oshortage.FItemlist(i).fresellno %>       <!--���Ǹ���Ȳ -->
    </td>
    <td bgcolor="#EEEEFF" align="right">
        <b><%= oshortage.FItemlist(i).fsysstockno %></b>       <!--�ý������-->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).Ferrrealcheckno %>       <!--����-->
    </td>
    <td bgcolor="#EEEEFF" align="right">
		<b><%= oshortage.FItemlist(i).frealstockno %></b>          <!-- �ǻ���� -->
    </td>
    <td align="right">
        <%= oshortage.FItemlist(i).ferrsampleitemno %>      <!--����-->
    </td>
    <td bgcolor="#EEEEFF" align="right">
        <b><%= oshortage.FItemlist(i).getAvailStock %></b>     <!--��ȿ���-->
    </td>
	<td align="right"><%= oshortage.FItemlist(i).fsell3days %></td>		<!--�Ǹż���-->
	<td align="right"><%= oshortage.FItemlist(i).fsell7days %></td>
	<td align="right">													<!-- ���ʿ���� -->
        <% if oshortage.FItemlist(i).frequire3daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire3daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire3daystock*-1 %></font>
            </a>
        <% else %>
            0
        <% end if %>
	</td>
	<td align="right">
        <% if oshortage.FItemlist(i).frequire7daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire7daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire7daystock*-1 %></font>
            </a>
        <% else %>
            0
        <% end if %>
	</td>
	<td align="right">
        <% if oshortage.FItemlist(i).frequire14daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire14daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire14daystock*-1 %></font>
            </a>
        <% else %>
            0
        <% end if %>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
	    <td><%= oshortage.FItemlist(i).Fonlinemwdiv %></td>
		<td><%= oshortage.FItemlist(i).fcentermwdiv %></td>
	<% end if %>

    <td>
        <input type="text" class="text" name="itemno" value="0" size="2" maxlength="4" onKeyDown="CheckThis(frmBuyPrc<%= i %>);">
    </td>
    <td>
        <% if oshortage.FItemList(i).Fpreorderno>0 or oshortage.FItemList(i).Fpreorderno<0 then %>
        	���ֹ�:
            <% if oshortage.FItemList(i).Fpreorderno<>oshortage.FItemList(i).Fpreordernofix then response.write CStr(oshortage.FItemList(i).Fpreorderno) + " -> " %>
        	<%= oshortage.FItemList(i).Fpreordernofix %><br>
        <% end if %>
        <% if oshortage.FItemList(i).IsSoldOut then %>
			<font color="red">ON:ǰ��</font><Br>
        <% end if %>
		<% if oshortage.FItemList(i).Flimityn="Y" then %>
			<font color="blue">ON:����(<%= oshortage.FItemList(i).getLimitNo %>)</font><Br>

		<% end if %>
        <input type="button" class="button" value="�ǻ��Է�" onclick="popOffErrInput('<%= shopid %>','<%= oshortage.FItemList(i).Fitemgubun %>','<%= oshortage.FItemList(i).Fitemid %>','<%= oshortage.FItemList(i).Fitemoption %>');">     <!--�ǻ�����Է�-->
    </td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
    <td colspan="30" align="center">
        <% if oshortage.HasPreScroll then %>
            <span class="list_link"><a href="javascript:reg(<%=oshortage.StartScrollPage-1%>)">[pre]</a></span>
        <% else %>
        [pre]
        <% end if %>
        <% for i = 0 + oshortage.StartScrollPage to oshortage.StartScrollPage + oshortage.FScrollCount - 1 %>
            <% if (i > oshortage.FTotalpage) then Exit for %>
            <% if CStr(i) = CStr(oshortage.FCurrPage) then %>
            <span class="page_link"><font color="red"><b><%= i %></b></font></span>
            <% else %>
            <a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
            <% end if %>
        <% next %>
        <% if oshortage.HasNextScroll then %>
            <span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
        <% else %>
        [next]
        <% end if %>
    </td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
    <td colspan="30" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
<form name="frmArrupdate" method="post">
    <input type="hidden" name="mode" value="arrins">
    <input type="hidden" name="itemgubunarr2" value="">
    <input type="hidden" name="itemidadd2" value="">
    <input type="hidden" name="itemoptionarr2" value="">
    <input type="hidden" name="sellcasharr2" value="">
    <input type="hidden" name="buycasharr2" value="">
    <input type="hidden" name="suplycasharr2" value="">
    <input type="hidden" name="itemnoarr2" value="">
    <input type="hidden" name="itemnamearr2" value="">
    <input type="hidden" name="itemoptionnamearr2" value="">
    <input type="hidden" name="designerarr2" value="">
    <input type="hidden" name="shopid" value="<%=shopid%>">
    <input type="hidden" name="suplyer" value="10x10">
    <input type="hidden" name="idx" value="0">
    <input type="hidden" name="chargeid" value="<%=makerid%>">
    <input type="hidden" name="shopbuypricearr2" value="">
    <input type="hidden" name="isreq" value="Y">
</form>
<form name="frmbag" method="post">
    <input type="hidden" name="onoffgubun">
    <input type="hidden" name="itemgubunarr">
    <input type="hidden" name="itemidarr">
    <input type="hidden" name="itemoptionarr">
    <input type="hidden" name="itemnoarr">
    <input type="hidden" name="makerid">
    <input type="hidden" name="shopid" >
</form>
</table>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">
		<input type="button" class="button" value="�����ٿ�ε�" onclick="exceldownload()">
    </td>
    <td align="right">
    	<input type="button" class="button" value="��ٱ��ϴ��" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="��ٱ��Ϻ���" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
        <% if oshortage.FresultCount>0 then %>
            <input type="button" class="button" value="���ùٷ��ֹ��ۼ�(�ٹ����ٹ���)" onclick="jumundirect()">
        <% end if %>
        <% if oshortage.FresultCount>0 then %>
        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
            	<input type="button" class="button" value="���ùٷ��ֹ��ۼ�(��ü)" onclick="jumundirect_upche()">
            <%' end if %>
        <% end if %>
    </td>
</tr>
</table>
<!-- �׼� �� -->
<%
    set oshortage = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

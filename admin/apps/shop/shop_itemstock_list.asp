<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��ǰ�ֹ��˻�
' Hieditor : 2011.08.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls_2.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
<%
dim page , shopid , isusing , makerid , itemid , itemname , generalbarcode , i , sell7days ,shopsuplycash ,buycash
dim cdl , cdm , cds , shortagetype , comm_cd ,includepreorder ,research , parameter , ipgo , order
    page = request("page")
    research = request("research")
    isusing = request("isusing")
    makerid = request("makerid")
    itemid = request("itemid")
    itemname = request("itemname")
    generalbarcode = request("generalbarcode")
    comm_cd = request("comm_cd")
    cdl = request("cdl")
    cdm = request("cdm")
    cds = request("cds")
    shortagetype = request("shortagetype")
    includepreorder = request("includepreorder")
    sell7days = request("sell7days")
    ipgo = request("ipgo")
	shopid = request("shopid")
    order = request("order")

if page="" then page=1
if (research<>"on") and (includepreorder="") then
    'includepreorder = "on"
end if
if (research<>"on") and (ipgo="") then
    'ipgo = "on"
end if
if (research<>"on") and (shortagetype="") then
    'shortagetype = 7
end if
if (research<>"on") and (order="") then
    'order = "byrecent"
end if
if (research<>"on") and (isusing="") then
    isusing = "Y"
end if

'/�����ϰ�� ���� ���常 ��밡��
if (C_IS_SHOP) then
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
parameter = parameter & "&ipgo="&ipgo&"&order="&order&""

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
    oshortage.Frectmakerid = makerid
    oshortage.Frectitemid = itemid
    oshortage.Frectitemname = itemname
    oshortage.Frectcomm_cd = comm_cd
    oshortage.Frectgeneralbarcode = generalbarcode
    oshortage.Frectshortagetype = shortagetype
    oshortage.Frectipgo = ipgo
    oshortage.Frectorder = order

    if shopid <> "" then
        oshortage.fnewitemstock_list_datamart2
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

<script language='javascript'>

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
        alert('������ �����ϴ�. - ��ü��Ź ��ǰ�� ��� ���� ����.');
        return; //��ü��Ź ��ǰ�� ���?
    <% else %>
        var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popAdmOffrealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
        popwin.focus();
    <% end if %>
}

var jumunitem = null;

//�ٹ����� �ֹ��� �ۼ�
function jumundirect(shopgubun){
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
                    alert('��ü��Ź�̳� ��ü������ �ֹ� �ϽǼ� �����ϴ�.');
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
    //����
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_jumuninput.asp';
	//����
	}else{
		upfrm.action='/admin/fran/jumuninput.asp';
	}
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
function jumundirect_upche(shopgubun){
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
                    alert('�ٹ�������Ź�̳� �������� �ֹ� �ϽǼ� �����ϴ�.');
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
    //����
    if (shopgubun == 'True'){
    	upfrm.action='/common/offshop/shop_ipchulinput.asp';
	//����
	}else{
		upfrm.action='/common/offshop/shop_ipchulinput.asp';
	}

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
        <% if (C_IS_SHOP) then %>
    		<%= shopid %>
    	<% else %>
        	<% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
        <% end if %>
        &nbsp;
        ��뿩��:<% drawSelectBoxUsingYN "isusing", isusing %>
        &nbsp;
        <!-- #include virtual="/common/module/categoryselectbox.asp"-->
        &nbsp;
        ���Ա��� : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
    </td>

    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="javascript:reg('');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        &nbsp;
        ��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg('');">
        &nbsp;
        ��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
        ������ڵ� :
        <input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">

    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
    	
		��ǰ���� :
		<input type="radio" name="order" value="" <% if order="" then response.write " checked" %> onclick="reg('');">��ü
		<input type="radio" name="order" value="byrecent" <% if order="byrecent" then response.write "checked" %> onclick="reg('');">�Ż�ǰ
		<input type="radio" name="order" value="byonbest" <% if order="byonbest" then response.write "checked" %> onclick="reg('');">�¶��� ����Ʈ
		<input type="radio" name="order" value="byoffbest" <% if order="byoffbest" then response.write "checked" %> onclick="reg('');">�������� ����Ʈ
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

    </td>
    <td align="right">
    	<input type="button" class="button" value="��ٱ��ϴ��" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="��ٱ��Ϻ���" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
       
    </td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= oshortage.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oshortage.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td>
    	����ó
    </td>
    <td>�귣��</td>
    <td>��ǰ�ڵ�</td>
    <td>�̹���</td>
    <td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
	
    <td>�ǸŰ�</td>
    <td>����<br>���ް�</td>
    <td>����<br>����</td>
    <td>��<br>�Ǹŷ�</td>
    <td>7��<br>�Ǹŷ�</td>
    <td>3��<br>�Ǹŷ�</td>
    <td>�������</td>
    <td>���</td>
</tr>
<% if oshortage.FresultCount > 0 then %>
<%
for i=0 to oshortage.FresultCount -1

shopsuplycash = oshortage.FItemList(i).GetFranchiseSuplycash
buycash		  = oshortage.FItemList(i).GetFranchiseBuycash
%>
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
<input type="hidden" name="shopsuplycash" value="<%= oshortage.FItemlist(i).fshopsuplycash %>">
<input type="hidden" name="shopbuyprice" value="<%= oshortage.FItemlist(i).fshopbuyprice %>">
<input type="hidden" name="itemname" value="<%= oshortage.FItemlist(i).fshopitemname %>">
<input type="hidden" name="itemoptionname" value="<%= oshortage.FItemlist(i).fshopitemoptionname %>">
<input type="hidden" name="makerid" value="<%= oshortage.FItemlist(i).fmakerid %>">
<input type="hidden" name="shopid" value="<%= oshortage.FItemlist(i).fshopid %>">
    <td >
        <input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
    </td>
    <td>
        <%= GetdeliverGubunName(oshortage.FItemlist(i).fcomm_cd) %><br>(<%= GetJungsanGubunName(oshortage.FItemlist(i).fcomm_cd) %>)
    </td>
    <td>
        <a href="javascript:searchmakerid('<%= oshortage.FItemlist(i).fmakerid %>');" onfocus="this.blur()"><%= oshortage.FItemlist(i).fmakerid %></a>
    </td>
    <td>
        <%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %>
        <% if oshortage.FItemlist(i).Fitemgubun="10" then %>
        	<Br><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=oshortage.FItemlist(i).Fitemid%>" target="_blink" onfocus="this.blur()">[��]</a>
        <% end if %>
    </td>
    <td><img src="<%= oshortage.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0></td>
    <td>
        <%= oshortage.FItemlist(i).fshopitemname %><Br>
        <% if oshortage.FItemlist(i).fshopitemoptionname <> "" then %>
            <font color="blue">[<%=oshortage.FItemlist(i).fshopitemoptionname%>]<font>
        <% end if %>
    </td>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopitemprice,0) %>
    </td>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopbuyprice,0) %>
    </td>
    <td>
		<% if oshortage.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(shopsuplycash/oshortage.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
    </td>
    <td><%= oshortage.FItemlist(i).fsellno %></td><!--���Ǹ���Ȳ -->
    <td><%= oshortage.FItemlist(i).fsell7days %></td>
    <td><%= oshortage.FItemlist(i).fsell3days %></td>
    <td>
        <%= oshortage.FItemlist(i).frealStockno %>
    </td>
    
    <td>
        <% if oshortage.FItemList(i).Fpreorderno>0 then %>
        	���ֹ�:
            <% if oshortage.FItemList(i).Fpreorderno<>oshortage.FItemList(i).Fpreordernofix then response.write CStr(oshortage.FItemList(i).Fpreorderno) + " -> " %>
        	<%= oshortage.FItemList(i).Fpreordernofix %><br>
        <% end if %>
        <img src="/images/cartimage.jpg" style="cursor:pointer" onclick="adminshoppingbagregoneitem('OFF','<%=shopid%>',frmBuyPrc<%= i %>)">
    </td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
    <td colspan="25" align="center">
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
    <td colspan="25" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>


<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">

    </td>
    <td align="right">
    	<input type="button" class="button" value="��ٱ��ϴ��" onclick="adminshoppingbagreg(frmbag,'OFF','<%=shopid%>')">
    	<input type="button" class="button" value="��ٱ��Ϻ���" onclick="adminshoppingbagview(frmbag,'OFF','<%=shopid%>')">
       
    </td>
</tr>
</table>
<!-- �׼� �� -->
<%
    set oshortage = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

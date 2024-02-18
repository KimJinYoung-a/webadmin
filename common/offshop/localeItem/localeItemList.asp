<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ ������ ��ǰ ����
' History : 2010.08.03 ������ ����
'			2010.08.05 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopLocaleItemcls.asp"-->

<%
dim designer,page, pagesize, usingyn ,research,pricediff,imageview ,itemgubun, itemid, itemname , shopitemname , gubun , nameeng
dim cdl, cdm, cds ,shopid , i ,PriceDiffExists , arrexchangerate, currencyUnit_Pos ,multipleRate , exchangeRate, countrylangcd
dim decimalPointLen, decimalPointCut
dim prdcode, generalbarcode , shopdiv , adminok

	designer    	= RequestCheckVar(request("designer"),32)
	page        	= RequestCheckVar(request("page"),9)
	pagesize       	= RequestCheckVar(request("pagesize"),9)
	usingyn     	= RequestCheckVar(request("usingyn"),1)
	research    	= RequestCheckVar(request("research"),9)
	pricediff   	= RequestCheckVar(request("pricediff"),9)
	imageview   	= RequestCheckVar(request("imageview"),9)

	itemgubun   	= RequestCheckVar(request("itemgubun"),2)
	itemid      	= RequestCheckVar(request("itemid"),9)

	itemname    	= RequestCheckVar(request("itemname"),32)
	shopitemname	= RequestCheckVar(request("shopitemname"),32)

	cdl         	= RequestCheckVar(request("cdl"),3)
	cdm         	= RequestCheckVar(request("cdm"),3)
	cds         	= RequestCheckVar(request("cds"),3)
	shopid      	= RequestCheckVar(request("shopid"),32)
	gubun      		= RequestCheckVar(request("gubun"),10)
	nameeng 		= RequestCheckVar(request("nameeng"),10)

	prdcode 		= RequestCheckVar(request("prdcode"),32)
	generalbarcode 	= RequestCheckVar(request("generalbarcode"),32)

''���� session("ssAdminPsn")="6" : �μ���ȣ�� ����Ұ�.
if session("ssBctDiv")="201" or session("ssAdminPsn")="6" then
	shopid = "cafe002"
elseif session("ssBctDiv")="301" or session("ssAdminPsn")="16" then
	shopid = "cafe003"
else
end if

if C_ADMIN_USER then

''���������� �������ΰ�� �ھ� �ִ´�
elseif (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

if page="" then page=1
if pagesize="" then pagesize = 100
''if research<>"on" then usingyn="Y"

decimalPointLen = 0
dim oexchangerate
set oexchangerate = new COffShopLocale
	oexchangerate.frectuserid = shopid

if shopid = "" then
	response.write "<script>alert('������ �����ϼ���');</script>"
else
	oexchangerate.fexchangeratecheck()

	shopdiv = oexchangerate.foneitem.fshopdiv
	currencyUnit_Pos = oexchangerate.foneitem.fcurrencyUnit_Pos
	multipleRate = oexchangerate.foneitem.fmultipleRate
	exchangeRate = oexchangerate.foneitem.fexchangeRate
	decimalPointLen = oexchangerate.foneitem.fdecimalPointLen
	decimalPointCut = oexchangerate.foneitem.fdecimalPointCut
    countrylangcd   = oexchangerate.foneitem.fcountrylangcd

	'/�ؿܸ����� �ƴҰ��
	if shopdiv <> "7" then
		adminok = false

		response.write "<script>"
		response.write "	alert('�����Ͻ� ������ �ؿܸ����� �ƴմϴ�');"
		response.write "</script>"
		response.write "<font color='red'>�ؿܸ��常 ��밡��</font>"
		response.end

	'/�ؿܸ����� ���
	else
		adminok	 = true
	end if

	if oexchangerate.foneitem.fcurrencyUnit_Pos = "" or isnull(oexchangerate.foneitem.fcurrencyUnit_Pos) then response.write "<script>alert('[�ʼ�]�ش���忡 ȭ������� ��ϵǾ� ���� �ʽ��ϴ�\n\n���庰 ȭ������� ����� [OFF]����_�������>>�����޸���Ʈ ���� �Է����ּ���.');</script>"
	if oexchangerate.foneitem.fmultipleRate = "" or isnull(oexchangerate.foneitem.fmultipleRate) then response.write "<script>alert('[�ʼ�]�ش���忡 ��������� ��ϵǾ� ���� �ʽ��ϴ�\n\n���庰 ȭ������� ����� [OFF]����_�������>>�����޸���Ʈ ���� �Է����ּ���.');</script>"
end if

dim ioffitem
set ioffitem  = new COffShopLocale
	ioffitem.FPageSize = pagesize
	ioffitem.FCurrPage = page
	ioffitem.FRectShopId = shopid
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectShopItemName = html2db(shopitemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.frectgubun = gubun
	ioffitem.frectnameeng = nameeng
	ioffitem.FRectPrdCode = prdcode
	ioffitem.FRectGeneralBarcode = generalbarcode
    ioffitem.FRectMultipleRate = MultipleRate
    ioffitem.FRectExchangeRate = exchangeRate

    ioffitem.FRectcountrylangcd = countrylangcd
	if (shopid<>"") then
	    ioffitem.GetLocaleItemList()
	end if


dim isShowMultiLang : isShowMultiLang = (NOT isNULL(countrylangcd) and (countrylangcd<>"") and (countrylangcd<>"KR"))
%>

<script language='javascript'>
function isMayEng(str){
    return (str.length==getbyteLength(str))
}

function getbyteLength (str){
    var retCode = 0;
    var strLength = 0;

    for (i = 0; i < str.length; i++){
        var code = str.charCodeAt(i)
        var ch = str.substr(i,1).toUpperCase()

        code = parseInt(code)

        if ((ch < "0" || ch > "9") && (ch < "A" || ch > "Z") && ((code > 255) || (code < 0)))
            strLength = strLength + 2;
        else
            strLength = strLength + 1;
    }
    return strLength;
}

//������ ��� �����
function CheckThislcprice(frm){
	frm.mrate.value = Math.round(((frm.lcprice.value / frm.ShopItemprice.value) * frm.erate.value) * 100) / 100 ;
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

//������ �ǸŰ� ���
function CheckThismrate(frm){
    var upfrm = document.frm;

    var cutn = upfrm.decimalPointCut.value*1;
	var pown = upfrm.decimalPointLen.value*1;
    var cutnPow = Math.pow(10, cutn)*1;

	//frm.lcprice.value = Math.round(((frm.ShopItemprice.value / frm.erate.value)* frm.mrate.value) * 100) / 100;
	var cc = Math.round(((frm.ShopItemprice.value / frm.erate.value)* frm.mrate.value) * cutnPow) / cutnPow;
	frm.lcprice.value = cc.toFixed(pown);

	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

// ����ǰ �߰� �˾�
function addnewItem(){
	var popup_item;
	popup_item = window.open("pop_localeItem_input.asp", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popup_item.focus();
}

function popForeignPriceBase(shopid){
    var popwin = window.open('/common/offshop/exchangerate/popForeignPriceBase.asp?shopid='+shopid,'popForeignPriceBase','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//ȯ������ ��� & ���� - �������
function popexchangerate(){
    var popexchangerate = window.open('/common/offshop/exchangerate/exchangerate.asp','popexchangerate','width=1024,height=768,scrollbars=yes,resizable=yes');
    popexchangerate.focus();
}

// ȯ�� ��� �ϰ�����
function automulti(upfrm){
    if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}


	var frm;
	var cutn = upfrm.decimalPointCut.value*1;
	var pown = upfrm.decimalPointLen.value*1;
    var cutnPow = Math.pow(10, cutn)*1;

    //3.1234 * 100

		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){

					//if (frm.lcprice.value==''){
					//	alert('�����ǸŰ� �������� �ʾҽ��ϴ�');
					//	frm.lcprice.focus;
					//	return;
					//}

					frm.erate.value = upfrm.exchangeRate.value
					frm.mrate.value = upfrm.multipleRate.value
					//frm.lcprice.value = Math.round(((frm.ShopItemprice.value / upfrm.exchangeRate.value)* upfrm.multipleRate.value) * 100) / 100;

					var cc = Math.round(((frm.ShopItemprice.value / upfrm.exchangeRate.value)* upfrm.multipleRate.value) * cutnPow) / cutnPow;
					frm.lcprice.value = cc.toFixed(pown);
				}
			}
		}
}

//ȯ���ϰ�����
function autoexchangeRate(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){

					if (frm.lcprice.value==''){
						alert('�����ǸŰ� �������� �ʾҽ��ϴ�');
						frm.lcprice.focus;
						return;
					}

					frm.lcprice.value = frm.lcprice.value / frm.exchangeRate.value;


				}
			}
		}
}

//��������ϰ�����
function automultipleRate(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){

					if (frm.lcprice.value==''){
						alert('�����ǸŰ� �������� �ʾҽ��ϴ�');
						frm.lcprice.focus;
						return;
					}

					frm.lcprice.value = frm.lcprice.value * upfrm.multipleRate.value;


				}
			}
		}
}

//�⺻�ǸŰ��ϰ�����
function autoShopItemprice(upfrm){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					frm.lcprice.value = frm.ShopItemprice.value

				}
			}
		}
}

function autoShopItemNameNOptionName(upfrm,tp){
    return; //�̹��� �̻��û
    autoShopItemName(upfrm,tp);
    autoshopitemoptionname(upfrm,tp);
}

//�⺻��ǰ���ϰ�����
function autoShopItemName(upfrm,tp){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
				    if (tp==0){
					    frm.lcitemname.value = frm.ShopItemName.value;
    				}else if (tp==1){
    				    if (frm.multiLang_itemname.value.length>0){
        				    frm.lcitemname.value = frm.multiLang_itemname.value;
        				}
    				}else if (tp==2){

    				    if ((frm.multiLang_itemname.value.length>0)&&(isMayEng(frm.multiLang_itemname.value))){
    				        frm.lcitemname.value = frm.multiLang_itemname.value;
    				    }else if (isMayEng(frm.ShopItemName.value)){
    				        frm.lcitemname.value = frm.ShopItemName.value;
    				    }

    				}
				}
			}
		}
}

//�⺻�ɼǸ��ϰ�����
function autoshopitemoptionname(upfrm,tp){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
				    if (tp==0){
					    frm.lcitemoptionname.value = frm.shopitemoptionname.value;
					}else if (tp==1){
					    if (frm.multiLang_optionname.value.length>0){
    				        frm.lcitemoptionname.value = frm.multiLang_optionname.value;
    				    }
    				}else if (tp==2){

    				    if ((frm.multiLang_optionname.value.length>0)&&(isMayEng(frm.multiLang_optionname.value))){
    				        frm.lcitemoptionname.value = frm.multiLang_optionname.value;
    				    }else if (isMayEng(frm.shopitemoptionname.value)){
    				        frm.lcitemoptionname.value = frm.shopitemoptionname.value;
    				    }

    				}
				}
			}
		}
}

function ModiArr(upfrm){
    if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm1;
	var lina = '';
	var liona = '';

	upfrm.eratea.value = '';
	upfrm.mratea.value = '';
	upfrm.ia.value = '';
	upfrm.ioa.value = '';
	upfrm.iga.value = '';
	upfrm.lina.value = '';
	upfrm.liona.value = '';
	upfrm.lpa.value = '';
		for (var i=0;i<document.forms.length;i++){
			frm1 = document.forms[i];
			if (frm1.name.substr(0,9)=="frmBuyPrc") {
				if (frm1.cksel.checked){
/*
					if (frm1.lcitemname.value == ''){
						alert('��ǰ���� �Է����ּ���');
						frm1.lcitemname.focus();
						return;
					}
*/
					if (frm1.lcprice.value == ''){
						alert('�ǸŰ��� �Է����ּ���');
						frm1.lcprice.focus();
						return;
					}
					upfrm.eratea.value = upfrm.eratea.value + frm1.erate.value + "," ;
					upfrm.mratea.value = upfrm.mratea.value + frm1.mrate.value + "," ;
					upfrm.ia.value = upfrm.ia.value + frm1.itemid.value + "," ;
					upfrm.ioa.value = upfrm.ioa.value + frm1.itemoption.value + "," ;
					upfrm.iga.value = upfrm.iga.value + frm1.itemgubun.value + "," ;

					lina = ''; //frm1.lcitemname.value;
					upfrm.lina.value = upfrm.lina.value + lina.replace(",","") + "," ;
					lina = '';

					liona = ''; //frm1.lcitemoptionname.value
					upfrm.liona.value = upfrm.liona.value + liona.replace(",","") + "," ;
					liona = '';
					upfrm.lpa.value = upfrm.lpa.value + frm1.lcprice.value + "," ;
				}
			}
		}

		upfrm.mode.value = 'litemadd';
		upfrm.method="post";
		upfrm.action = 'localeitem_process.asp';
		upfrm.submit();
}

function reg(page){
    var frm = document.frm;
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

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="currencyUnit_Pos" value="<%= currencyUnit_Pos %>">

<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			<% end if %>
		<% else %>
			���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
		<% end if %>
	    ��ǰ��뱸��:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		������ǰ�� : <input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�����ڵ� :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
		&nbsp;
		������ڵ� :
		<input type="text" class="text" name="generalbarcode" value="<%= generalbarcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�������� : <% drawlocaleitemgubun "gubun" , gubun , "" %>
		&nbsp;
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
		&nbsp;
		<input type="checkbox" name="nameeng" value="on" <% if nameeng="on" then response.write "checked" %> >����(��ǰ��,�ɼǸ�)�� ����
		&nbsp;
		ǥ�ð��� :
		<select class="select" name="pagesize">
			<option value="100">100</option>
			<option value="250" <%= CHKIIF(CLng(pagesize) = 250, "selected", "") %> >250</option>
			<option value="500" <%= CHKIIF(CLng(pagesize) = 500, "selected", "") %> >500</option>
		</select>
	</td>
</tr>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<% if adminok then %>
<tr>
    <td height="30">
    �� ���庰 ȭ������� ����� <input type="button" value="�ؿ� ���� ��� ����" class="button" onClick="popForeignPriceBase('<%= shopid %>');"> ���� �����ϼ���.
    </td>
</tr>
<tr>
	<td align="left">
	    <% if currencyUnit_Pos <> "" and multipleRate <> "" then %>
	    	�ǸŰ� X ȯ��<input type="text" name="exchangeRate" value="<%= exchangeRate %>" size=5 maxlength=6>
	    	X ���<input type="text" name="multipleRate" value="<%= multipleRate %>" size=3 maxlength=4>


			&nbsp;&nbsp;
			(
			�Ҽ���<input type="text" class="text" name="decimalPointLen" value="<%= decimalPointLen %>" size=1 maxlength=2>�ڸ�ǥ��
	    	�Ҽ���<input type="text" class="text" name="decimalPointCut" value="<%= decimalPointCut %>" size=1 maxlength=2>�ݿø�
	    	)
			<!--<input type="button" class="button" value="�⺻�ǸŰ�����" onclick="autoShopItemprice(frm)">
			<input type="button" class="button" value="ȯ������" onclick="autoexchangeRate(frm)">
			<input type="button" class="button" value="�������(X<%= multipleRate %>)" onclick="automultipleRate(frm)">-->
		<% end if %>
	</td>
	<td align="right">
		    <input type="button" class="button" value="�ؿ� �ǸŰ� ���" onclick="automulti(frm)">
			&nbsp;
		<% if (FALSE) then %>
			<input type="button" class="button" value="�⺻��ǰ��/�ɼǸ� ����" onclick="autoShopItemNameNOptionName(frm,0)">
			<% if (isShowMultiLang) then %>
			&nbsp;<input type="button" class="button" value="<%=countrylangcd%> ��ǰ��/�ɼǸ� ����" onclick="autoShopItemNameNOptionName(frm,1)">
			&nbsp;<input type="button" class="button" value="���� �켱  ��ǰ��/�ɼǸ� ����" onclick="autoShopItemNameNOptionName(frm,2)">
		    <% end if %>
		<% end if %>

			<!--
			<input type="button" class="button" value="�⺻��ǰ��" onclick="autoShopItemName(frm)">
			<input type="button" class="button" value="�⺻�ɼǸ�" onclick="autoshopitemoptionname(frm)">
			-->

			&nbsp;<input type="button" class="button" value="�����ϰ�����" onclick="ModiArr(actfrm)">
	</td>
</tr>
<% end if %>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ioffitem.FTotalcount %></b>
		&nbsp;
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		<b><%= page %> / <%= ioffitem.FTotalpage %></b>
	</td>
</tr>
<% if ioffitem.FresultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<% if (imageview<>"") then %>
	<td>�̹���</td>
	<% end if %>
	<td>����<br>����</td>
	<td>�귣��ID<br>������ڵ�</td>
	<td>�����ڵ�<br>�ɼ��߰��ݾ�</td>
	<td>��ǰ��</font><br>������ǰ��</td>
	<td>�ɼǸ�</font><br>�����ɼǸ�</td>
	<td>�Һ��ڰ�(��)<br>�ǸŰ�(��)</td>
	<td>�ؿ��ǸŰ�<br>(<%= currencyUnit_Pos %>)</td>
	<td>ȯ��</td>
	<td>���</td>
	<!-- <td>�ؿܸ���<br>�ǸŰ�(<%= currencyUnit_Pos %>)</td> -->

</tr>

<% for i=0 to ioffitem.FresultCount -1 %>
<form method="get" action="" name="frmBuyPrc<%=i%>">

<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
	<td >
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		<input type="hidden" name="shopid" value="<%=shopid%>">
		<input type="hidden" name="itemid" value="<%=ioffitem.FItemlist(i).FShopitemid%>">
		<input type="hidden" name="itemoption" value="<%=ioffitem.FItemlist(i).Fitemoption%>">
		<input type="hidden" name="itemgubun" value="<%=ioffitem.FItemlist(i).fitemgubun%>">
	</td>
	<% if (imageview<>"") then %>
	<td><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
	<% end if %>
	<td>
		<%= ioffitem.FItemlist(i).fstatus %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FMakerID %>
		<br><%= ioffitem.FItemlist(i).FextBarcode %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).Fitemgubun %><%=  FormatCode(ioffitem.FItemlist(i).Fshopitemid) %><%= ioffitem.FItemlist(i).Fitemoption %>
		<br>
		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopItemName %>
		<% if (isShowMultiLang) then %>
		<p><%= ioffitem.FItemlist(i).FmultiLang_itemname %></p>
	    <% end if %>
		<input type="hidden" name="ShopItemName" value="<%= ioffitem.FItemlist(i).FShopItemName %>">
		<input type="hidden" name="multiLang_itemname" value="<%= ioffitem.FItemlist(i).FmultiLang_itemname %>">
		<% if (FALSE) then %>
		<br><input type="text" name="lcitemname" value="<%= ioffitem.FItemlist(i).flcitemname %>" maxlength=123 size=30 readonly style="background-color:'#EEEEEE'">
		<% end if %>
	</td>
	<td>
	    <%= ioffitem.FItemlist(i).FShopitemOptionname %>
	    <% if (isShowMultiLang) then %>
		<p><%= ioffitem.FItemlist(i).FmultiLang_optionname %></p>
	    <% end if %>

		<input type="hidden" name="shopitemoptionname" value="<%= ioffitem.FItemlist(i).fshopitemoptionname %>">
		<input type="hidden" name="multiLang_optionname" value="<%= ioffitem.FItemlist(i).FmultiLang_optionname %>">
		<% if (FALSE) then %>
		<br><input type="text" name="lcitemoptionname" value="<%= ioffitem.FItemlist(i).flcitemoptionname %>" maxlength=95 size=15 readonly style="background-color:'#EEEEEE'">
	    <% end if %>
	</td>
    <% PriceDiffExists = false %>
    <td>
        <% if (FALSE) then %>
        <%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice,0) %>
		<br>
	    <% end if %>
	    <%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %>
	    <input type="hidden" name="ShopItemprice" value="<%=ioffitem.FItemlist(i).FShopItemprice%>">
    </td>
    <td>
		<input type="text" name="lcprice" value="<%= CHKIIF(IsNULL(ioffitem.FItemlist(i).flcprice),"",NULL2Zero(ioffitem.FItemlist(i).flcprice)) %>" size=5 maxlength=10 onKeyup="CheckThislcprice(frmBuyPrc<%= i %>)">
    </td>
	<td>ȯ��<input type="text" name="erate" value="<%= ioffitem.FItemlist(i).fexchangeRate %>" size=5 maxlength=5 readonly></td>
	<td>X ���<input type="text" name="mrate" value="<%= ioffitem.FItemlist(i).fmultipleRate %>" size=5 maxlength=4 onKeyup="CheckThismrate(frmBuyPrc<%= i %>)"></td>


</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if ioffitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=ioffitem.StartScrollPage-1%>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ioffitem.StartScrollPage to ioffitem.StartScrollPage + ioffitem.FScrollCount - 1 %>
			<% if (i > ioffitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ioffitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>
<form name="actfrm" method="post">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="usingyn" value="<%=usingyn%>">
<input type="hidden" name="designer" value="<%=designer%>">
<input type="hidden" name="itemid" value="<%=itemid%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="shopitemname" value="<%=shopitemname%>">
<input type="hidden" name="prdcode" value="<%=prdcode%>">
<input type="hidden" name="generalbarcode" value="<%=generalbarcode%>">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="imageview" value="<%=imageview%>">
<input type="hidden" name="nameeng" value="<%=nameeng%>">
<input type="hidden" name="currencyUnit_Pos" value="<%= currencyUnit_Pos %>">
<input type="hidden" name="ia">
<input type="hidden" name="ioa">
<input type="hidden" name="iga">
<input type="hidden" name="lina">
<input type="hidden" name="liona">
<input type="hidden" name="lpa">
<input type="hidden" name="eratea">
<input type="hidden" name="mratea">
<input type="hidden" name="mode">
</form>
<%
	set ioffitem = nothing
	set oexchangerate = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

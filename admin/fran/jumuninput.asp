<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ �ֹ��� �ۼ�
' History : 2009.04.07 ������ ����
'			2010.08.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

dim shopid, reguser, divcode,baljuname,regname,comment ,osheetmaster, idx ,suplyer,yyyymmdd ,vatcode ,i,j,cnt,cnt2
dim itemgubunarr, itemidadd, itemoptionarr ,itemnamearr, itemoptionnamearr ,itemnamearr2, itemoptionnamearr2
dim sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr ,itemgubunarr2, itemidadd2, itemoptionarr2
dim sellcasharr2, suplycasharr2, buycasharr2, itemnoarr2, designerarr2 ,itemgubunarr3, itemidadd3, itemoptionarr3
dim itemnamearr3, itemoptionnamearr3 ,sellcasharr3, suplycasharr3, buycasharr3, itemnoarr3, designerarr3
dim isPreExists ,cwflag, shopdiv
dim foreign_sellcasharr, foreign_suplycasharr, foreign_sellcasharr2, foreign_suplycasharr2,foreign_sellcasharr3, foreign_suplycasharr3
dim currencyunit, currencyChar, loginsite, ArrShopInfo
dim addshopid
	cwflag = request("cwflag")
	shopid = request("shopid")
	reguser = shopid
	comment = html2db(request("comment"))
	divcode = "503"
	'baljuname = session("ssBctCname")
	regname = session("ssBctCname")
	idx = request("idx")
	if idx="" then idx=0
	suplyer = request("suplyer")
	yyyymmdd = request("yyyymmdd")
	itemgubunarr = request("itemgubunarr")
	itemidadd	= request("itemidadd")
	itemoptionarr = request("itemoptionarr")
	itemnamearr		= request("itemnamearr")
	itemoptionnamearr = request("itemoptionnamearr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	buycasharr = request("buycasharr")
	foreign_sellcasharr = request("foreign_sellcasharr")
	foreign_suplycasharr = request("foreign_suplycasharr")
	itemnoarr = request("itemnoarr")
	designerarr = request("designerarr")
	itemgubunarr2 = request("itemgubunarr2")
	itemidadd2	= request("itemidadd2")
	itemoptionarr2 = request("itemoptionarr2")
	itemnamearr2	= request("itemnamearr2")
	itemoptionnamearr2 = request("itemoptionnamearr2")
	sellcasharr2 = request("sellcasharr2")
	suplycasharr2 = request("suplycasharr2")
	buycasharr2 = request("buycasharr2")
	itemnoarr2 = request("itemnoarr2")
	designerarr2 = request("designerarr2")
	foreign_sellcasharr2 = request("foreign_sellcasharr2")
	foreign_suplycasharr2 = request("foreign_suplycasharr2")
	addshopid = request("addshopid")
	'chargeid = request("chargeid")
	'shopid = session("ssBctID")
	'vatcode = request("vatcode")
	'divcode  = request("divcode")
	itemgubunarr = split(itemgubunarr,"|")
	itemidadd	= split(itemidadd,"|")
	itemoptionarr = split(itemoptionarr,"|")
	itemnamearr		= split(itemnamearr,"|")
	itemoptionnamearr = split(itemoptionnamearr,"|")
	sellcasharr = split(sellcasharr,"|")
	suplycasharr = split(suplycasharr,"|")
	buycasharr = split(buycasharr,"|")
	foreign_sellcasharr = split(foreign_sellcasharr,"|")
	foreign_suplycasharr = split(foreign_suplycasharr,"|")
	itemnoarr = split(itemnoarr,"|")
	designerarr = split(designerarr,"|")
	itemgubunarr2 = split(itemgubunarr2,"|")
	itemidadd2	= split(itemidadd2,"|")
	itemoptionarr2 = split(itemoptionarr2,"|")
	itemnamearr2		= split(itemnamearr2,"|")
	itemoptionnamearr2 = split(itemoptionnamearr2,"|")
	sellcasharr2 = split(sellcasharr2,"|")
	suplycasharr2 = split(suplycasharr2,"|")
	buycasharr2 = split(buycasharr2,"|")
	foreign_sellcasharr2 = split(foreign_sellcasharr2,"|")
	foreign_suplycasharr2 = split(foreign_suplycasharr2,"|")
	itemnoarr2 = split(itemnoarr2,"|")
	designerarr2 = split(designerarr2,"|")

	cnt = uBound(itemidadd)
	cnt2 = uBound(itemidadd2)

dim sqlStr
if shopid <> "" then
	ArrShopInfo = getoffshopuser(shopid)

	IF isArray(ArrShopInfo) then
		currencyunit = ArrShopInfo(1,0)
		currencyChar = ArrShopInfo(3,0)
		loginsite = ArrShopInfo(2,0)
		shopdiv = ArrShopInfo(12,0)
    END IF
end if

for j=0 to cnt2-1
	isPreExists = false
	for i=0 to cnt-1
		if (itemgubunarr(i)=itemgubunarr2(j)) and (itemidadd(i)=itemidadd2(j)) and (itemoptionarr(i)=itemoptionarr2(j)) then
			itemnoarr(i) = CStr(CLng(itemnoarr(i)) + CLng(itemnoarr2(j)))
			isPreExists = true
			exit for
		end if
	next

	if Not isPreExists then
		itemgubunarr3 = itemgubunarr3 + itemgubunarr2(j) + "|"
		itemidadd3	= itemidadd3 + itemidadd2(j) + "|"
		itemoptionarr3 = itemoptionarr3 + itemoptionarr2(j) + "|"
		itemnamearr3		= itemnamearr3 + itemnamearr2(j) + "|"
		itemoptionnamearr3  = itemoptionnamearr3 + itemoptionnamearr2(j) + "|"
		sellcasharr3 = sellcasharr3 + sellcasharr2(j) + "|"
		suplycasharr3 = suplycasharr3 + suplycasharr2(j) + "|"
		buycasharr3 = buycasharr3 + buycasharr2(j) + "|"
		if ubound(foreign_sellcasharr2) > 0 then
		foreign_sellcasharr3 = foreign_sellcasharr3 + foreign_sellcasharr2(j) + "|"
		foreign_suplycasharr3 = foreign_suplycasharr3 + foreign_suplycasharr2(j) + "|"
	    end if
		itemnoarr3 = itemnoarr3 + itemnoarr2(j) + "|"
		designerarr3 = designerarr3 + designerarr2(j) + "|"
	end if
next

itemgubunarr2 = ""
itemidadd2	= ""
itemoptionarr2 = ""
itemnamearr2	= ""
itemoptionnamearr2 = ""
sellcasharr2 = ""
suplycasharr2 = ""
buycasharr2 = ""
foreign_sellcasharr2 = ""
foreign_suplycasharr2 = ""
itemnoarr2 = ""
designerarr2 = ""

for i=0 to cnt-1
	itemgubunarr2 = itemgubunarr2 + itemgubunarr(i) + "|"
	itemidadd2	= itemidadd2 + itemidadd(i) + "|"
	itemoptionarr2 = itemoptionarr2 + itemoptionarr(i) + "|"
	itemnamearr2	= itemnamearr2 + itemnamearr(i) + "|"
	itemoptionnamearr2 = itemoptionnamearr2 + itemoptionnamearr(i) + "|"
	sellcasharr2 = sellcasharr2 + sellcasharr(i) + "|"
	suplycasharr2 = suplycasharr2 + suplycasharr(i) + "|"
	buycasharr2 = buycasharr2 + buycasharr(i) + "|"
	foreign_sellcasharr2 = foreign_sellcasharr2 + foreign_sellcasharr(i) + "|"
	foreign_suplycasharr2 = foreign_suplycasharr2 + foreign_suplycasharr(i) + "|"
	itemnoarr2 = itemnoarr2 + itemnoarr(i) + "|"
	designerarr2 = designerarr2 + designerarr(i) + "|"
next

itemgubunarr = itemgubunarr2 + itemgubunarr3
itemidadd	= itemidadd2 + itemidadd3
itemoptionarr = itemoptionarr2 + itemoptionarr3
itemnamearr	= itemnamearr2 + itemnamearr3
itemoptionnamearr = itemoptionnamearr2 + itemoptionnamearr3
sellcasharr = sellcasharr2 + sellcasharr3
suplycasharr = suplycasharr2 + suplycasharr3
buycasharr = buycasharr2 + buycasharr3
foreign_sellcasharr = foreign_sellcasharr2 + foreign_sellcasharr3
foreign_suplycasharr = foreign_suplycasharr2 + foreign_suplycasharr3
itemnoarr = itemnoarr2 + itemnoarr3
designerarr = designerarr2 + designerarr3

if cwflag = "" then
	'//��������ϰ�� �⺻���� �����Ź��
	if shopdiv = "13" then
		cwflag = "1"
	else
		cwflag = "0"
	end if
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,foreign_sellcash,foreign_suplycash){
	if (iidx!='0'){
		alert('�ֹ����� ��ġ���� �ʽ��ϴ�. �ٽýõ��� �ּ���.');
		return;
	}

	document.frmMaster.itemgubunarr2.value = igubun;
	document.frmMaster.itemidadd2.value = iitemid;
	document.frmMaster.itemoptionarr2.value = iitemoption;
	document.frmMaster.sellcasharr2.value = isellcash;
	document.frmMaster.suplycasharr2.value = isuplycash;
	document.frmMaster.buycasharr2.value = ibuycash;
	document.frmMaster.itemnoarr2.value = iitemno;
	document.frmMaster.itemnamearr2.value = iitemname;
	document.frmMaster.itemoptionnamearr2.value = iitemoptionname;
	document.frmMaster.designerarr2.value = iitemdesigner;
	document.frmMaster.foreign_sellcasharr2.value = foreign_sellcash;
	document.frmMaster.foreign_suplycasharr2.value = foreign_suplycash;
	document.frmMaster.submit();
}

//��ǰ�߰� ����Ʈ��
function AddItems_locale(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('/common/offshop/localeitem/popshopjumunitem_locale.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=0&cwflag='+cwflag ,'franjumuninputadd','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItems(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('/common/offshop/popshopjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=0&cwflag='+cwflag ,'franjumuninputadd','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddOrderSheet(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('jumunlist.asp?popupyn=Y' ,'franjumuninputaddordersheet','width=1500,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItemsCSV(frm){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('popoffjumunbycsv.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=0&cwflag='+cwflag ,'franjumuninputaddBarcode','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItemsBarCode(frm, digitflag){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('popshopjumunitemBybarcode.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&digitflag=' + digitflag + '&idx=0&cwflag='+cwflag ,'franjumuninputaddBarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function ConFirmIpChulList(bool){
	var msfrm = document.frmMaster;
	var upfrm = document.frmArrupdate;
	var frm; var IsNotpirce=false;

	if (msfrm.yyyymmdd.value.length<1){
		alert('�԰��û���� �Է��� �ּ���.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.foreign_sellcasharr.value = "";
	upfrm.foreign_suplycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (!IsInteger(frm.itemno.value)){
				alert('������ ������ �����մϴ�.');
				frm.itemno.focus();
				return;
			}

			<% if loginsite="WSLWEB" then %>
				<%
				'/Ȧ�����ε� ��ǥȭ�� ��ȭ �ϰ�� �������ݰ� �ؿܰ����� ���ƾ���.
				if (currencyUnit = "KRW" or currencyUnit = "WON") Then
				%>
					if (!IsNotpirce) {
						if (frm.sellcash.value.replace(',','') != frm.foreign_sellcash.value.replace(',','')){
							alert('�ؿܸ����� ��� ȭ�� ��ȭ�ΰ�� �����ǸŰ��� �ؿ��ǸŰ��� �����ؾ� �մϴ�.\n�����Ͻ��� ������忡�� �ݵ�� �ٽ� �������ּ���.');
							IsNotpirce=true;
						}
						if (frm.suplycash.value.replace(',','') != frm.foreign_suplycash.value.replace(',','')){
							alert('�ؿܸ����� ��� ȭ�� ��ȭ�ΰ�� ��������� �ؿ������ �����ؾ� �մϴ�.\n�����Ͻ��� ������忡�� �ݵ�� �ٽ� �������ּ���.');
							IsNotpirce=true;
						}
					}
				<% end if %>
			<% end if %>

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
			upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
			if (frm.foreign_sellcash){  //�����߰� 2016/06/13 eastone ������� �������� ó���Ұ��
    			upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value + frm.foreign_sellcash.value + "|";
    			upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value + frm.foreign_suplycash.value + "|";
    		}
			upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
			upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
		}
	}

	if (!bool) {
		var ret = confirm('������ �ӽ� ���� �Ͻðڽ��ϱ�?');
	}else{
		var ret = confirm('���� �Ͻðڽ��ϱ�?');
	}

	if (ret){
		//�ӽ�����(�ۼ���)
		if (!bool) upfrm.waitflag.value="on"

		upfrm.yyyymmdd.value = msfrm.yyyymmdd.value;
		upfrm.comment.value = msfrm.comment.value;

		<% if shopid = "" or addshopid <> "" then %>
			upfrm.addshopid.value = msfrm.addshopid.value;
		<% end if %>

		var cwflag;
		for (var i =0 ; i < msfrm.cwflag.length ; i++){
			if (msfrm.cwflag[i].checked){
				cwflag = msfrm.cwflag[i].value;
			}
		}

		upfrm.cwflag.value = cwflag;
		upfrm.submit();
	}
}

function chcwflag(shopid){

	if (shopid==''){
		alert('������ �����ϼ���');
		return;
	}

	frmMaster.target='view';
	frmMaster.mode.value='chcwflag';
	frmMaster.action='/common/offshop/inc_shopcwflag_search.asp';
	frmMaster.submit();

	frmMaster.target='';
	frmMaster.mode.value='addmaster';
	frmMaster.action='';
}

// ���� ���� �˾�
function popShopSelect() {
	var frm = document.frmMaster;

	if (frm.shopid.value == '') {
		alert("���� �⺻ ������ �����ϼ���.");
		return;
	}

	var popwin = window.open("/admin/offshop/pop_shopSelect.asp", "popShopSelect","width=460,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// �˾����� ���� ���� �߰�
function addSelectedShop(shopid, shopname)
{
	var frm = document.frmMaster;
	var addshopid = document.getElementById('addshopid');
	var tbl_addshop = document.getElementById('tbl_addshop');


	if (shopid == frm.shopid.value) {
		alert("�̹� �⺻ ���忡 ������ �����Դϴ�.");
		return;
	}

	if (addshopid.value.indexOf(',' + shopid + ',') >= 0) {
		alert("�̹� �߰��� �����Դϴ�.");
		return;
	}

	addSelectedShopNoCheck(shopid, shopname);
}

function addSelectedShopNoCheck(shopid, shopname) {
	var frm = document.frmMaster;
	var addshopid = document.getElementById('addshopid');
	var tbl_addshop = document.getElementById('tbl_addshop');

	var lenRow = tbl_addshop.rows.length;

	// ���߰�
	var oRow = tbl_addshop.insertRow(lenRow);
	oRow.onmouseover=function(){tbl_addshop.clickedRowIndex=this.rowIndex};

	addshopid.value = addshopid.value + shopid + ',';
	var oCell0 = oRow.insertCell(0);
	var oCell1 = oRow.insertCell(1);

	oCell0.id = shopid;
	oCell0.innerHTML = shopid + "/" + shopname;
	oCell1.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdShop()' align=absmiddle>";
}

// ���ø��� ����
function delSelectdShop(){
	var tbl_addshop = document.getElementById('tbl_addshop');
	var addshopid = document.getElementById('addshopid');
	var shopid;

	if(confirm("������ ������ �����Ͻðڽ��ϱ�?")) {
		alert('Before' + addshopid.value);
		shopid = tbl_addshop.rows(tbl_addshop.clickedRowIndex).cells(0).id;
		addshopid.value = addshopid.value.replace(shopid + ',', '')
		tbl_addshop.deleteRow(tbl_addshop.clickedRowIndex);
		alert('After' + addshopid.value);
	}
}
window.onload = function () {
	var i;
	var addshopid = "<%= addshopid %>";
	var addshopidArr = addshopid.split(',');
	for (i = 0; i < addshopidArr.length; i++) {
		if (addshopidArr[i] != '') {
			addSelectedShopNoCheck(addshopidArr[i], '');
		}
	}
	//addshopid
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="">
<input type="hidden" name="mode" value="addmaster">
<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
<input type="hidden" name="itemidadd" value="<%= itemidadd %>">
<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
<input type="hidden" name="buycasharr" value="<%= buycasharr %>">
<input type="hidden" name="foreign_sellcasharr" value="<%=foreign_sellcasharr%>">
<input type="hidden" name="foreign_suplycasharr" value="<%=foreign_suplycasharr%>">
<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
<input type="hidden" name="designerarr" value="<%= designerarr %>">
<input type="hidden" name="itemgubunarr2" value="">
<input type="hidden" name="itemidadd2" value="">
<input type="hidden" name="itemoptionarr2" value="">
<input type="hidden" name="itemnamearr2" value="">
<input type="hidden" name="itemoptionnamearr2" value="">
<input type="hidden" name="sellcasharr2" value="">
<input type="hidden" name="suplycasharr2" value="">
<input type="hidden" name="buycasharr2" value="">
<input type="hidden" name="itemnoarr2" value="">
<input type="hidden" name="designerarr2" value="">
<input type="hidden" name="foreign_sellcasharr2" value="">
<input type="hidden" name="foreign_suplycasharr2" value="">
<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�ֹ�����(OFFSHOP)</strong></font>
			    </td>
			    <td align="right">
				</td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">�ֹ���(SHOP)</td>
	<% if shopid<>"" then %>
	<input type=hidden name="shopid" value="<%= shopid %>">
	<td><%= shopid %></td>
	<% else %>
	<td>
		<% NewDrawSelectBoxDesignerwithNameAndUserDIV "shopid",shopid, "21" %>
		<script>
		 $(function(){
			 $("#shopid").change(function(){
				 chcwflag(this.value);
			 });
		 });
		</script>
	</td>
	<% end if %>
</tr>
<% if shopid = "" or addshopid <> "" then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">�߰�����</td>
	<td>
		<table class=a border="0">
			<tr>
				<td>
					<input type='hidden' id="addshopid" name='addshopid' value=','>
					<table name='tbl_addshop' id='tbl_addshop' class=a>
						<tr onMouseOver='tbl_addshop.clickedRowIndex=this.rowIndex'>
    						<td></td>
    						<td></td>
    					</tr>
					</table>
				</td>
				<td valign="bottom">
					<input type="button" class='button' value="�߰�" onClick="popShopSelect()">
				</td>
			</tr>
		</table>
		<p />
		* <font color="red">������ ����</font>�� �귣�� ��ǰ�� ���庰 �ֹ����� �߰��˴ϴ�.<br />
		* �ؿܸ����� ��� �ֹ����� �ۼ����� �ʽ��ϴ�.
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
	<% if suplyer<>"" then %>
	<input type=hidden name="suplyer" value="<%= suplyer %>">
	<td><%= suplyer %></td>
	<% else %>
	<td><% SelectBoxOffShopSuplyer "suplyer", suplyer, shopid, session("ssBctDiv") %></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�԰��û��</td>
	<td>
		<input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.yyyymmdd);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (���ϴ� �԰� ��¥�� �Է��ϼ���.)
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="divcwflag" name="divcwflag" style="display:none">
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td>
		<input type="radio" name="cwflag" value="0" <% if cwflag="0" then response.write " checked" %>>������
		<input type="radio" name="cwflag" value="1" <% if cwflag="1" then response.write " checked" %>>�����Ź
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ��û����</td>
	<td>
		<textarea name="comment" class="textarea" cols="80" rows="6"><%= comment %></textarea>
	</td>
</tr>
</form>
</table>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
<tr>
	<td align="left">
		<iframe id="view" name="view" src="" width="100%" height=0 frameborder="0" scrolling="no"></iframe>
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->
<%
itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
foreign_sellcasharr = split(foreign_sellcasharr,"|")
foreign_suplycasharr = split(foreign_suplycasharr,"|")

cnt = ubound(itemidadd)

dim selltotal, suplytotal,foreign_selltotal, foreign_suplytotal
	selltotal =0
	suplytotal =0
	foreign_selltotal =0
	foreign_suplytotal =0
%>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�󼼳���</strong></font>
			    </td>
			    <td align="right">
			    	�ѰǼ� : <% if cnt<1 then response.write "0" else response.write cnt end if %>
		        	&nbsp;
					<% if (session("ssBctDiv") < 10) then %>
		        	<input type="button" class="button" value="�ֹ����߰�" onclick="AddOrderSheet(frmMaster)">
					<% end if %>
					<input type="button" class="button" value="��ǰ�߰�" onclick="AddItems(frmMaster)">
					<input type="button" class="button" value="��ǰ�߰�(NEW)" onclick="AddItems_locale(frmMaster)">
					<input type="button" class="button" value="����(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'P')">
					<input type="button" class="button" value="��ǰ(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'M')">
					<input type="button" class="button" value="CSV����" onclick="AddItemsCSV(frmMaster)">
				</td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="120">�귣��ID</td>
	<td width="100">���ڵ�</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="60">�����</td>
	<td width="60">�ֹ�����</td>
	<td width="60">�ǸŰ�</td>
	<td width="60">���</td>
	<td width="70">�ǸŰ���</td>
	<td width="70">�����</td>
	<%IF loginsite = "WSLWEB" THEN%>
	<td>�ؿ��ǸŰ�(<%=currencyunit%>)</td>
	<td>�ؿܰ��ް�(<%=currencyunit%>)</td>
	<td width="70">�ؿ��ǸŰ��հ�(<%=currencyunit%>)</td>
	<td width="70">�ؿ�����հ�(<%=currencyunit%>)</td>
	<%END IF%>
</tr>
<% for i=0 to cnt-1 %>
<%
selltotal  = selltotal + sellcasharr(i) * itemnoarr(i)
suplytotal = suplytotal + suplycasharr(i) * itemnoarr(i)

if ubound(foreign_sellcasharr)>0 then
	foreign_selltotal  = foreign_selltotal + foreign_sellcasharr(i) * itemnoarr(i)
	foreign_suplytotal = foreign_suplytotal + foreign_suplycasharr(i) * itemnoarr(i)
end if

%>
<form name="frmBuyPrc_<%= i %>" method="post" action="">
<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
<input type="hidden" name="itemid" value="<%= itemidadd(i) %>">
<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">
<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">

<%if ubound(foreign_sellcasharr)>0 then%>
	<input type="hidden" name="foreign_sellcash" value="<%= getdisp_price(foreign_sellcasharr(i), currencyChar) %>">
<%end if%>
<%if ubound(foreign_suplycasharr)>0 then%>
	<input type="hidden" name="foreign_suplycash" value="<%= getdisp_price(foreign_suplycasharr(i), currencyChar) %>">
<%end if%>

<input type="hidden" name="buycash" value="<%= buycasharr(i) %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><%= designerarr(i) %></td>
	<td><%= itemgubunarr(i) %><%= CHKIIF(itemidadd(i)>=1000000,format00(8,itemidadd(i)),format00(6,itemidadd(i))) %><%= itemoptionarr(i) %></td>
	<td align="left">
		<%= itemnamearr(i) %>
		<% if itemoptionarr(i) <>"0000" then %>
			<font color="blue"><%= itemoptionnamearr(i) %></font>
		<% end if %>
	</td>
	<td>�����</td>
	<td ><input type="text" class="text" name="itemno" value="<%= itemnoarr(i) %>"  size="4" maxlength="4"></td>
	<td align="right"><%= FormatNumber(sellcasharr(i),0) %></td>
	<td align="right"><%= FormatNumber(suplycasharr(i),0) %></td>
	<td align="right"><%= FormatNumber(sellcasharr(i) * itemnoarr(i),0) %></td>
	<td align="right"><%= FormatNumber(suplycasharr(i) * itemnoarr(i),0) %></td>

	<%IF loginsite = "WSLWEB" THEN%>
		<td align="right">
			<%if ubound(foreign_sellcasharr)>0 then%>
				<%= getdisp_price_currencyChar(foreign_sellcasharr(i), currencyChar) %>
			<% end if %>
		</td>
		<td align="right">
			<%if ubound(foreign_suplycasharr)>0 then%>
				<%= getdisp_price_currencyChar(foreign_suplycasharr(i), currencyChar) %>
			<% end if %>
		</td>
		<td align="right">
			<%if ubound(foreign_sellcasharr)>0 then%>
				<%= getdisp_price_currencyChar(foreign_sellcasharr(i) * itemnoarr(i), currencyChar) %>
			<% end if %>
		</td>
		<td align="right">
			<%if ubound(foreign_suplycasharr)>0 then%>
				<%= getdisp_price_currencyChar(foreign_suplycasharr(i) * itemnoarr(i), currencyChar) %>
			<% end if %>
		</td>
	<%END IF%>
</tr>
</form>
<% next %>

<% if (cnt>0) then %>
<tr bgcolor="#FFFFFF">
	<td align="center">�Ѱ�</td>
	<td colspan="6">
	<td align=right><%= formatNumber(selltotal,0) %></td>
	<td align=right><%= formatNumber(suplytotal,0) %></td>

	<%IF loginsite = "WSLWEB" THEN%>
		<td colspan=2></td>
		<td align=right>
			<%if ubound(foreign_sellcasharr)>0 then%>
				<%= getdisp_price_currencyChar(foreign_selltotal, currencyChar) %>
			<% end if %>
		</td>
		<td align=right>
			<%if ubound(foreign_sellcasharr)>0 then%>
				<%= getdisp_price_currencyChar(foreign_suplytotal, currencyChar) %>
			<% end if %>
		</td>
	<%END IF%>
</tr>
<% end if %>

<!-- �ϴܹ� ���� -->
<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<% if (cnt>0) then %>
    	<input type="button" class="button" value="����Ȯ��(�ֹ�����)" onclick="ConFirmIpChulList(true)">
    	<input type="button" class="button" value="�ӽ�����(�ۼ���)" onclick="ConFirmIpChulList(false)">
    	<% else %>
    	&nbsp;
    	<% end if %>
	</td>
</tr>
<!-- �ϴܹ� �� -->
</table>

<%
'// ����� ���̵� + �ð��� ������ �ߺ��Է� üũ
dim uniqregdate : uniqregdate = getDatabaseTime()
%>

<form name="frmArrupdate" method="post" action="shopjumun_process.asp">
	<input type="hidden" name="mode" value="addshopjumun">
	<input type="hidden" name="waitflag" value="">
	<input type="hidden" name="yyyymmdd" value="">
	<input type="hidden" name="addshopid" value="">
	<input type="hidden" name="baljuid" value="<%= shopid %>">
	<input type="hidden" name="targetid" value="<%= suplyer %>">
	<input type="hidden" name="reguser" value="<%= reguser %>">
	<input type="hidden" name="uniqregdate" value="<%= uniqregdate %>">
	<input type="hidden" name="divcode" value="<%= divcode %>">
	<input type="hidden" name="vatinclude" value="Y">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="regname" value="<%= regname %>">
	<input type="hidden" name="baljuname" value="<%= baljuname %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="foreign_sellcasharr" value="">
	<input type="hidden" name="foreign_suplycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="cwflag">
</form>

<script language='javascript'>
	chcwflag('<%=shopid%>');
</script>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

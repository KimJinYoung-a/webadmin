<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ �ֹ��� �ۼ�
' History : 2009.04.07 ������ ����
'			2010.08.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim idx, isfixed, opage, ourl,oshopid,ostatecd,odesinger, jumunwait, sqlStr
dim ojumunmaster, ojumundetail, oupchemwinfo ,yyyymmdd, IsForeign_confirmed, IsForeignOrder
Dim oprice, ArrShopInfo, currencyunit, currencyChar, loginsite, shopdiv, storemarginrate
	idx = requestCheckVar(getNumeric(request("idx")),10)
	opage = requestCheckVar(request("opage"),10)
	ourl = request("ourl")
	oshopid = requestCheckVar(request("oshopid"),32)
	ostatecd = requestCheckVar(request("ostatecd"),10)
	odesinger = requestCheckVar(request("odesinger"),32)

jumunwait = false
IsForeignOrder = false		'/��ü�����ֹ�
IsForeign_confirmed = false		'/��ü�����ֹ� ���߿ϷῩ��
if idx="" then idx=0

set ojumunmaster = new COrderSheet
	ojumunmaster.FRectIdx = idx
	ojumunmaster.GetOneOrderSheetMaster

set ojumundetail= new COrderSheet
	ojumundetail.FRectIdx = idx
	ojumundetail.FRectShopid = ojumunmaster.FoneItem.FBaljuid
	ojumundetail.GetOrderSheetDetail

'/ �̹� �ű����忡�� ���������̺� ���� �����. ���������� ���� ���� ���� ����. �ֹ����̺� ������ ������ �����;���.	2017.11.01 �ѿ��
'if ojumunmaster.FoneItem.FBaljuid <> "" then
'	ArrShopInfo = getoffshopuser(ojumunmaster.FoneItem.FBaljuid)
'
'	IF isArray(ArrShopInfo) then
'		currencyunit = ArrShopInfo(1,0)
'		currencyChar = ArrShopInfo(3,0)
'		loginsite = ArrShopInfo(2,0)
'		shopdiv = ArrShopInfo(12,0)
'    END IF
'end if
loginsite = ojumunmaster.FOneItem.fsitename
currencyunit = ojumunmaster.FOneItem.fcurrencyUnit

set oupchemwinfo = new CUpcheMwInfo
	oupchemwinfo.FRectdesignerId = ojumunmaster.FOneItem.Ftargetid
	oupchemwinfo.GetDesignerMWInfo

set oprice = new COverSeasItem
	oprice.FRectShopid = ojumunmaster.FOneItem.Fbaljuid

	if (ojumunmaster.FOneItem.fcurrencyUnit = "USD") Then
		oprice.GetOverSeasDefaultPriceInfo
	end if

yyyymmdd = Left(ojumunmaster.FOneItem.Fscheduledate,10)

if ojumunmaster.FOneItem.fforeign_statecd<>"" then
	IsForeignOrder=true

	if ojumunmaster.FOneItem.fforeign_statecd="7" then
		IsForeign_confirmed = true
	end if
else
	IsForeign_confirmed = true
end if
if ojumunmaster.FOneItem.FStatecd=" " then
	jumunwait = true	'/�ֹ����ۼ���
end if
if (ojumunmaster.FOneItem.FStatecd>"5") then
	isfixed = true
else
	isfixed = false
end if

dim tmpcolor

if (  (storemarginrate = "") or (storemarginrate = "0") ) then
	sqlStr = "select IsNull(a.marginrate, 0) as marginrate "
	sqlStr = sqlStr + " from [db_storage].[dbo].vw_acount_user_delivery a "
	sqlStr = sqlStr + " where a.userid = '" +  ojumunmaster.FoneItem.FBaljuid  + "' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		storemarginrate = rsget("marginrate")
	else
		storemarginrate = "0"
	end if
	rsget.close
elseif (storemarginrate = "") then
	storemarginrate = "0"
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">

function DelArr(){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var upfrm = document.frmadd;
	var masterfrm = document.frmMaster;
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
		alert('���� ������ �����ϴ�.');
		return;
	}

	upfrm.detailidxarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + ",";
			}
		}
	}

	if (confirm('���� ������ ���� �Ͻðڽ��ϱ�?')){
		upfrm.targetid.value = masterfrm.suplyer.value;
		upfrm.baljuid.value = masterfrm.shopid.value;
		upfrm.mode.value = "delshopjumunarr";
		upfrm.submit();
	}
}

function SaveArr(){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var upfrm = document.frmadd;
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
	upfrm.buycasharr.value = "";
	upfrm.baljuitemnoarr.value = "";
	upfrm.realitemnoarr.value = "";
	upfrm.commentarr.value = "";
	upfrm.detailidxarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.baljuitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.baljuitemno.focus();
					return;
				}

				if (frm.baljuitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.baljuitemno.focus();
					return;
				}

				if (!IsInteger(frm.realitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.realitemno.focus();
					return;
				}

				if (frm.realitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.realitemno.focus();
					return;
				}

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + "|";
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.baljuitemnoarr.value = upfrm.baljuitemnoarr.value + frm.baljuitemno.value + "|";
				upfrm.realitemnoarr.value = upfrm.realitemnoarr.value + frm.realitemno.value + "|";
				upfrm.commentarr.value = upfrm.commentarr.value + frm.comment.value + "|";
			}
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		upfrm.mode.value = "modeshopjumunarr";
		upfrm.submit();
	}
}

// ��ü����
function SaveALL(){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var masterfrm = document.frmMaster;
	var upfrm = document.frmadd;
	var frm; var IsNotpirce=false;
	var pass = false;

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.baljuitemnoarr.value = "";
	upfrm.realitemnoarr.value = "";
	upfrm.commentarr.value = "";
	upfrm.detailidxarr.value = "";
	upfrm.ipgoflagarr.value = "";
	upfrm.defaultmaginflagarr.value = "";
	upfrm.buymaginflagarr.value = "";
	upfrm.suplymaginflagarr.value = "";
	upfrm.foreign_sellcasharr.value = "";
	upfrm.foreign_suplycasharr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

				if (!IsInteger(frm.baljuitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.baljuitemno.focus();
					return;
				}

				if (frm.baljuitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.baljuitemno.focus();
					return;
				}

				if (!IsInteger(frm.realitemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.realitemno.focus();
					return;
				}

				if (frm.realitemno.value.length<1){
					alert('������ �Է��ϼ���.');
					frm.realitemno.focus();
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
								frm.foreign_sellcash.focus();
								IsNotpirce=true;
							}
							if (frm.suplycash.value.replace(',','') != frm.foreign_suplycash.value.replace(',','')){
								alert('�ؿܸ����� ��� ȭ�� ��ȭ�ΰ�� ��������� �ؿ������ �����ؾ� �մϴ�.\n�����Ͻ��� ������忡�� �ݵ�� �ٽ� �������ּ���.');
								frm.foreign_suplycash.focus();
								IsNotpirce=true;
							}
						}
					<% end if %>
				<% end if %>

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + "|";
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.baljuitemnoarr.value = upfrm.baljuitemnoarr.value + frm.baljuitemno.value + "|";
				upfrm.realitemnoarr.value = upfrm.realitemnoarr.value + frm.realitemno.value + "|";
				upfrm.commentarr.value = upfrm.commentarr.value + frm.comment.value + "|";
				upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value + frm.foreign_sellcash.value + "|";
				upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value + frm.foreign_suplycash.value + "|";

				//alert(frm.ipgoflag.value);
				upfrm.ipgoflagarr.value = upfrm.ipgoflagarr.value + frm.ipgoflag.value + "|";
				//if (frm.ipgoflag.checked){
				//	upfrm.ipgoflagarr.value = upfrm.ipgoflagarr.value + frm.ipgoflag.value + "|";
				//}else{
				//	upfrm.ipgoflagarr.value = upfrm.ipgoflagarr.value + "|";
				//}

				upfrm.defaultmaginflagarr.value = upfrm.defaultmaginflagarr.value + frm.defaultmaginflag.value + "|";
				upfrm.buymaginflagarr.value = upfrm.buymaginflagarr.value + frm.buymaginflag.value + "|";
				upfrm.suplymaginflagarr.value = upfrm.suplymaginflagarr.value + frm.suplymaginflag.value + "|";
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		if (masterfrm.beasongdate!=undefined){
			upfrm.songjangname.value = masterfrm.songjangdiv.options[masterfrm.songjangdiv.selectedIndex].text;
			upfrm.beasongdate.value = masterfrm.beasongdate.value;
			upfrm.songjangdiv.value = masterfrm.songjangdiv.value;
			upfrm.songjangno.value = masterfrm.songjangno.value;
			upfrm.targetid.value = masterfrm.suplyer.value;
			upfrm.baljuid.value = masterfrm.shopid.value;
		}
		upfrm.yyyymmdd.value = masterfrm.yyyymmdd.value;
		upfrm.comment.value = masterfrm.comment.value;

		upfrm.statecd.value = getCheckboxValue(masterfrm,'statecd');
		upfrm.mode.value = "modeshopjumunmasterdetail";
		upfrm.submit();
	}
}

function getCheckboxValue(f,compname){
    for(var i=0;i<f.elements.length;i++){
      if(f.elements[i].name==compname && f.elements[i].checked){
        return f.elements[i].value;
      }
    }
    return false;
}

//��ǰ�߰� ����Ʈ��
function AddItems_locale(frm){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

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

	<% if (ojumunmaster.FOneItem.FStatecd="6") or ((C_ADMIN_AUTH) and (ojumunmaster.FOneItem.FStatecd=" ")) then %>
		for (var i =0 ; i < frm.cwflag.length ; i++){
			if (frm.cwflag[i].checked){
				cwflag = frm.cwflag[i].value;
			}
		}
	<% else %>
		cwflag = frm.cwflag.value;
	<% end if %>

	popwin = window.open('/common/offshop/localeitem/popshopjumunitem_locale.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masteridx.value + '&cwflag='+cwflag ,'franjumuninputadd','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//��ǰ�߰�
function AddItems(frm){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var popwin;
	var suplyer, shopid;

	if (frm.suplyer.value.length<1){
		alert('����ó�� ���� �����ϼ���.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid  = frm.shopid.value;

	var cwflag;

	<% if (ojumunmaster.FOneItem.FStatecd="6") or ((C_ADMIN_AUTH) and (ojumunmaster.FOneItem.FStatecd=" ")) then %>
		for (var i =0 ; i < frm.cwflag.length ; i++){
			if (frm.cwflag[i].checked){
				cwflag = frm.cwflag[i].value;
			}
		}
	<% else %>
		cwflag = frm.cwflag.value;
	<% end if %>

	popwin = window.open('/common/offshop/popshopjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masteridx.value + '&cwflag='+cwflag,'franjumuninputeditadd','width=1280,height=960,scrollbars=yes,status=no');
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

	<% if (ojumunmaster.FOneItem.FStatecd="6") or ((C_ADMIN_AUTH) and (ojumunmaster.FOneItem.FStatecd=" ")) then %>
		for (var i =0 ; i < frm.cwflag.length ; i++){
			if (frm.cwflag[i].checked){
				cwflag = frm.cwflag[i].value;
			}
		}
	<% else %>
		cwflag = frm.cwflag.value;
	<% end if %>

	popwin = window.open('popshopjumunitemBybarcode.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masteridx.value + '&digitflag=' + digitflag + '&cwflag='+cwflag ,'franjumuninputaddBarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiThis(frm){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function chforeign_statecd(){
	var tmpforeign_statecd;
	for (var i =0 ; i < frmMaster.foreign_statecd.length ; i++){
		if (frmMaster.foreign_statecd[i].checked){
			tmpforeign_statecd = frmMaster.foreign_statecd[i].value;
		}
	}

	var ret = confirm('���¸� ���� �Ͻðڽ��ϱ�?');
	if (ret){
		frmedit.foreign_statecd.value=tmpforeign_statecd;
		frmedit.mode.value="chforeign_statecd";
		frmedit.submit();
	}
}

function DelThis(frm){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function DelMaster(frm){
	<% if not(C_ADMIN_AUTH) then %>
		<%
		'/�ֹ��� �ۼ��� ���°� �ƴϰ�
		if not(jumunwait) then
		%>
			<%
			'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
			if (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
			%>
			<% else %>
				alert("�ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�.");
				return;
			<% end if %>
		<% end if %>
	<% end if %>

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){

		frm.targetid.value = frm.suplyer.value;
		frm.baljuid.value = frm.shopid.value;
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function ModiMaster(frm){
	if (frm.beasongdate!=undefined){
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,foreign_sellcash,foreign_suplycash){
	if (iidx!='<%= idx %>'){
		alert('�ֹ����� ��ġ���� �ʽ��ϴ�. �ٽýõ��� �ּ���.');
		return;
	}

	frmadd.itemgubunarr.value = igubun;
	frmadd.itemarr.value = iitemid;
	frmadd.itemoptionarr.value = iitemoption;
	frmadd.sellcasharr.value = isellcash;
	frmadd.suplycasharr.value = isuplycash;
	frmadd.buycasharr.value = ibuycash;
	frmadd.itemnoarr.value = iitemno;
	frmadd.foreign_sellcasharr.value = foreign_sellcash;
	frmadd.foreign_suplycasharr.value = foreign_suplycash;
	frmadd.submit();
}

function ChulgoProc(frm,bool){
	if (frm.yyyymmdd.value.length<1){
		alert('�԰��û���� �Է��� �ּ���.');
		frm.yyyymmdd.focus();
		if (!calendarOpen2(frm.yyyymmdd)) { return };
	}
	if (frm.ipgodate.value.length<1){
		alert('������� �Է��� �ּ���.');
		frm.ipgodate.focus();
		if (!calendarOpen2(frm.ipgodate)) { return };
	}
	if (frm.beasongdate!=undefined){
		frm.songjangname.value = frm.songjangdiv.options[frm.songjangdiv.selectedIndex].text;
	}

	var ret = confirm('���ó�� �Ͻðڽ��ϱ�?');

	if (ret){
		var obj = document.getElementById('btnChulgo');
		if (obj != undefined) {
			obj.disabled = true;
		}
		frm.mode.value="chulgoproc";
		frm.limitflag.value = bool;
		frm.submit();
	}
}

function showSpecialInput(objTarget){
	if(objTarget[objTarget.selectedIndex].id=='special'){
	 	output = window.showModalDialog("/lib/inputpop.html" , null, "dialogwidth:250px;dialogheight:120px;center:yes;scroll:no;resizable:no;status:no;help:no;");

	 	if(output!=''){
	 		objTarget[objTarget.selectedIndex].text=output;
	  		objTarget[objTarget.selectedIndex].value=output;
	 	}else{

	 	}
	 }
}

function IpgoFinish(){
	var imsg = "";

	if (frmMaster.ipgodate.value.length<1){
		var ret1 = calendarOpen2(frmMaster.ipgodate);
		if (!ret1) return;
	}

	var ret2 = confirm('�԰��� : ' + frmMaster.ipgodate.value + ' OK?');
	if (!ret2) return;

	var idivcode = getCheckboxValue(frmMaster,'divcode');

	if (idivcode=="121"){
		imsg = "[�¶�����Ź���->����������Ź] �ΰ�� \r\n�¶��� ������ ���� ������ \r\n���������� ��Ź�԰�˴ϴ�. \r\n�԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}else if(idivcode=="131"){
		imsg = "[�¶�����Ź���->�����������] �ΰ�� \r\n�¶��� ������ ���� ������ \r\n���������� �����԰�˴ϴ�. \r\n�԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}else if(idivcode=="201"){
		imsg = "[�¶��θ������->�����������] �ΰ�� \r\n�¶��� ������ ���� ������ \r\n���������� �����԰�˴ϴ�. \r\n�԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}else{
		imsg = " �԰� Ȯ������ ���� �Ͻðڽ��ϱ�?";
	}

	var ret = confirm(imsg);

	if (ret){

		frmMaster.mode.value= "franupcheipgofinish";
		frmMaster.targetid.value= frmMaster.suplyer.value;
		frmMaster.submit();
	}
}

function DelAlink(frm,alinkcode){
	if (confirm('���õ� ����� ������ ���� �Ͻðڽ��ϱ�?')){
		frmMaster.mode.value = "delalinkipchul";
		frmMaster.alinkcode.value = alinkcode;

		frmMaster.submit();
	}
}

function NotCheckThis(icomp){
	icomp.checked = !(icomp.checked);
	//alert(icomp.checked);
	//if (icomp.checked==true){
	//	icomp.checked = false;
	//}
}

function publicbarreg(barcode){
	//var popwin = window.open('/common/popbarcode_input.asp?itembarcode=' + barcode,'popbarcode_input','width=500,height=400,resizable=yes,scrollbars=yes');
	var popwin = window.open('/admin/stock/popBarcodeManage.asp?itemcode=' + barcode,'popbarcode_input','width=550,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

var exchangeRate = 0;
var multiplerate = 0;
var linkPriceType = 0;
<% If oprice.FResultCount > 0 Then %>
exchangeRate = <%= oprice.FItemList(0).fexchangeRate %>;
multiplerate = <%= oprice.FItemList(0).fmultiplerate %>;
linkPriceType = <%= oprice.FItemList(0).flinkPriceType %>;
<% End If %>

function jsSetForeignPrice() {
	var i, frm;

	if (exchangeRate == 0) {
		alert('�ؿܻ�ǰ �ǸźҰ� ���ó�Դϴ�.');
		return;
	}

	if (confirm("�ؿ��ǸŰ��� �ڵ��Էµ˴ϴ�.\n\n - 90��ǰ�� �Էµ˴ϴ�.\n - �̹� �ؿ��ǸŰ� �Էµ� ��ǰ�� ���ܵ˴ϴ�.\n\n�����Ͻðڽ��ϱ�?") == true) {
		var frm = document.frmMaster;
		frm.mode.value = "insforgnprice";
		frm.submit();
	}
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function ApplyMargin(storemarginrate) {
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frm.suplycash.value = 1 * frm.sellcash.value * (100 - storemarginrate) / 100;
				frm.foreign_suplycash.value = 1 * frm.foreign_sellcash.value * (100 - storemarginrate) / 100;
			}
		}
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="/admin/fran/shopjumun_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="limitflag" value="">
<input type="hidden" name="opage" value="<%= opage %>">
<input type="hidden" name="ourl" value="<%= ourl %>">
<input type="hidden" name="oshopid" value="<%= oshopid %>">
<input type="hidden" name="ostatecd" value="<%= ostatecd %>">
<input type="hidden" name="odesinger" value="<%= odesinger %>">
<input type="hidden" name="masteridx" value="<%= idx %>">
<input type="hidden" name="shopid" value="<%= ojumunmaster.FOneItem.Fbaljuid %>">
<input type="hidden" name="baljuname" value="<%= ojumunmaster.FOneItem.Fbaljuname %>">
<input type="hidden" name="reguser" value="<%= session("ssBctid") %>">
<input type="hidden" name="regname" value="<%= session("ssBctCname") %>">
<input type="hidden" name="orgbaljucode" value="<%= ojumunmaster.FOneItem.FBaljuCode %>">
<input type="hidden" name="targetid" value="">
<input type="hidden" name="baljuid" value="">
<input type="hidden" name="alinkcode" value="">

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�ֹ�����</strong></font>
			        &nbsp;
			        <b>[ <%= ojumunmaster.FOneItem.FBaljuCode %> ]</b>
			        &nbsp;
					<% if (Not IsNULL(ojumunmaster.FOneItem.FALinkCode)) and (ojumunmaster.FOneItem.FALinkCode<>"") then %>
						��������:<%= ojumunmaster.FOneItem.FALinkCode %>
						<% if not IsNULL(ojumunmaster.FOneItem.Fipchuldeldt) then %>
							<font color=red>������</font>
						<% end if %>
						&nbsp;�ѼҺ�:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulsellcash,0) %>
						&nbsp;�Ѱ��ް�:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulsuplycash,0) %>
						&nbsp;�Ѹ��԰�:<%= FormatNumber(ojumunmaster.FOneItem.Fipchulbuycash,0) %>
						<input type="button" class="button" value="���� ����� ����" onClick="DelAlink(frmMaster,'<%= ojumunmaster.FOneItem.FALinkCode %>');">
					<% end if %>
			    </td>
			    <td align="right">
					<input type="button" class="button" value="������� �̵�" onclick="">
				</td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr bgcolor="#FFFFFF">
	<td width="110" bgcolor="<%= adminColor("tabletop") %>" >����ó</td>
	<td width="400">
		<input type="hidden" name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
		<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
	</td>
	<td width="120" bgcolor="<%= adminColor("tabletop") %>" >����ó(OFFSHOP)</td>
	<td>
		<%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)
		&nbsp;&nbsp;/&nbsp;&nbsp;<%= ojumunmaster.FOneItem.fcurrencyUnit %>
		<% if (ojumunmaster.FOneItem.fcurrencyUnit = "USD") then %>
		<input type="button" class="button" value="�ؿ��ǸŰ� �ڵ����" onClick="jsSetForeignPrice('<%= ojumunmaster.FOneItem.fcurrencyUnit %>')">
		<% end if %>

		<% if ojumunmaster.FOneItem.fsitename <> "" then %>
			&nbsp;&nbsp;/&nbsp;&nbsp;<%= ojumunmaster.FOneItem.fsitename %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�ֹ��Ͻ�</td>
	<td><%= ojumunmaster.FOneItem.Fregdate %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�԰��û��</td>
	<td>
		<input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly >
		<a href="javascript:calendarOpen(frmMaster.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
	<td colspan="3">
	    <% if (IsForeignOrder) and ((Not IsForeign_confirmed) or (C_ADMIN_AUTH)) then %>
    		<input type="radio" name="foreign_statecd" value="0" <% if ojumunmaster.FOneItem.fforeign_statecd="0" then response.write "checked" %> >��ü����(������û)
    		<input type="radio" name="foreign_statecd" value="3" <% if ojumunmaster.FOneItem.fforeign_statecd="3" then response.write "checked" %> >��ü����Ȯ��
    		<input type="radio" name="foreign_statecd" value="7" <% if ojumunmaster.FOneItem.fforeign_statecd="7" then response.write "checked" %> >��ü�����Ϸ�(�ֹ��� �ۼ��ߺ���)
    		<input type="button" onclick="chforeign_statecd()" value="���º���<%=CHKIIF((IsForeign_confirmed)," (�����ڱ���)","")%>" class="button">
		<% end if %>

		<% if (Not IsForeignOrder) or (IsForeign_confirmed) or not(isfixed) then %>
		<input type=radio name="statecd" value=" " <% if ojumunmaster.FOneItem.FStatecd=" " then response.write "checked" %> >�ֹ����ۼ���
		<input type=radio name="statecd" value="0" <% if ojumunmaster.FOneItem.FStatecd="0" then response.write "checked" %> >�ֹ�����
		<input type=radio name="statecd" value="1" <% if ojumunmaster.FOneItem.FStatecd="1" then response.write "checked" %> >�ֹ�Ȯ��
		<input type=radio name="statecd" value="6" <% if ojumunmaster.FOneItem.FStatecd="6" then response.write "checked" %> >�����(�����ǰ)

		<% if (ojumunmaster.FOneItem.FStatecd>="7") then %>
			<input type=radio name="statecd" value="7" <% if ojumunmaster.FOneItem.FStatecd="7" then response.write "checked" %> >���Ϸ�
		<% end if %>

		<% 'if (ojumunmaster.FOneItem.FStatecd>="6") then %>
			<% if (not IsNULL(ojumunmaster.FOneItem.Fipchuldeldt)) or (IsNULL(ojumunmaster.FOneItem.Falinkcode)) then %>
			<input type="button" class="button" value="���º���" onClick="ModiMaster(frmMaster)">
			<% else %>
    			<% IF (C_ADMIN_AUTH) THEN %>
    			<input type="button" class="button" value="���º���" onClick="alert('������ ���');ModiMaster(frmMaster)">
    			<% ELSE %>
    			<input type="button" class="button" value="���º���" onClick="alert('���� ����� ������ ��밡���մϴ�.')">
    			<% END IF %>
			<% end if %>
		<% 'end if %>
		<% else %>
		<input type=hidden name="statecd" value="<%=ojumunmaster.FOneItem.FStatecd%>">
		<% end if %>
	</td>
</tr>

<% 'if ojumunmaster.FOneItem.getOrderpaymentstatus<>"" then %>
	<!--<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">wholesale��������</td>
		<td colspan="3">
			<%'= ojumunmaster.FOneItem.getOrderpaymentstatus %>
			<br>���ڹ߼�:
			<br><%'= left(ojumunmaster.FOneItem.fsmssenddate,10) %>
			<br><%'= mid(ojumunmaster.FOneItem.fsmssenddate,12,22) %>
		</td>
	</tr>-->
<% 'end if %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >������Է�</td>
	<td>
		�ù�� : <% drawSelectBoxDeliverCompany "songjangdiv", ojumunmaster.FOneItem.Fsongjangdiv %>
		������ȣ: <input type="text" class="text" name="songjangno" size="16" maxlength="16" value="<%= ojumunmaster.FOneItem.Fsongjangno %>" >
		<input type=hidden name="songjangname" value="">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
	<td>
		<input type="text" class="text" name="beasongdate" value="<%= ojumunmaster.FOneItem.Fbeasongdate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.beasongdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
</tr>

<% if (ojumunmaster.FOneItem.FStatecd="6") or ((C_ADMIN_AUTH) and (ojumunmaster.FOneItem.FStatecd=" ")) then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td>
		<input type="radio" disabled name="cwflag" value="0" <% if ojumunmaster.FOneItem.fcwflag="0" then response.write " checked" %>>������
		<input type="radio" disabled name="cwflag" value="1" <% if ojumunmaster.FOneItem.fcwflag="1" then response.write " checked" %>>�����Ź
		<input type="hidden" name="cwflag" value="<%=ojumunmaster.FOneItem.fcwflag%>">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td>
		<input type="text" class="text" name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" size=10 ><a href="javascript:calendarOpen(frmMaster.ipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		(��� ����� ��������)
	</td>
</tr>
<% elseif (ojumunmaster.FOneItem.FStatecd>"6") then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td colspan="3">
		<input type="hidden" name="cwflag" value="<%= ojumunmaster.FOneItem.fcwflag %>" />
		<%= ojumunmaster.FOneItem.Fipgodate %>
		<input type="hidden" name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" />
	</td>
</tr>
<% else %>
	<input type="hidden" name="cwflag" value="<%= ojumunmaster.FOneItem.fcwflag %>" />
	<input type="hidden" name="ipgodate" value="<%= ojumunmaster.FOneItem.Fipgodate %>" />
<% end if %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ��հ�(�ֹ�)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�����ް��հ�(�ֹ�)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsuplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�Һ��ڰ��հ�(Ȯ��)</td>
	<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></b></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�����ް��հ�(Ȯ��)</td>
	<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsuplycash,0) %></b></td>
</tr>
<% if (IsForeignOrder) then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�ؿܼҺ��ڰ��հ�(�ֹ�)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.fjumunforeign_sellcash,2) %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ؿ������ް��հ�(�ֹ�)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.fjumunforeign_suplycash,2) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�ؿܼҺ��ڰ��հ�(Ȯ��)</td>
	<td><b><%= FormatNumber(ojumunmaster.FOneItem.ftotalforeign_sellcash,2) %></b></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ؿ������ް��հ�(Ȯ��)</td>
	<td><b><%= FormatNumber(ojumunmaster.FOneItem.ftotalforeign_suplycash,2) %></b></td>
</tr>
<% end if %>
<!--
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>�� ���԰�</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalbuycash,0) %> / <%= FormatNumber(ojumunmaster.FOneItem.Fjumunbuycash,0) %> </td>
</tr>
-->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�ֹ��귣��</td>
	<td colspan="3"><textarea class="textarea" cols="80" rows="3"><%= ojumunmaster.FOneItem.FBrandList %></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ��û����</td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6"><%= ojumunmaster.FOneItem.FComment %></textarea>
	</td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="4">
		* 5�ϳ� ��� : ��ü ��� ��ǰ (�������ͷ� �԰� �Ǵ´�� �������� �߼� �ص帮�ڽ��ϴ�.) <br>
		* ��� ���� : �������� ��� �������� ���� ��ü�� ���ְ� �� �ִ� �����Դϴ�. <br>
					2~3�� ���� �԰� �� �� �ִ� ��ǰ �Դϴ�. ���� �����帮�� ������, <B>���� �ֹ��� �߰�(���ֹ�)</B>�� �ּž� �մϴ�.<br>
		* �Ͻ�ǰ�� : ��ü ���������� ���� ��������� ��ǰ�Դϴ�.(�ܱⰣ ���� �԰� �Ǳ� ����� ��ǰ�Դϴ�.)
	</td>
</tr>
<% if (ojumunmaster.FOneItem.FStatecd="6") then %>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
		<% if IsForeign_confirmed then %>
			<input type="button" class="button" value="���ó��" onClick="ChulgoProc(frmMaster,true)" id="btnChulgo">
		<% end if %>

		<input type="button" class="button" value=" ��ü���� " onClick="DelMaster(frmMaster)">
	</td>
</tr>
<% end if %>

<% if (C_ADMIN_AUTH) and (ojumunmaster.FOneItem.FStatecd=" " or ojumunmaster.FOneItem.FStatecd<"7") then %>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
		<% if IsForeign_confirmed then %>
			<input type="button" class="button" value="���ó��(������)" onClick="ChulgoProc(frmMaster,false)" id="btnChulgo">
		<% end if %>
	</td>
</tr>
<% end if %>
</form>
</table>

<%
dim i,selltotal, suplytotal, buytotal ,totalfixno ,selltotalfix, suplytotalfix, buytotalfix
dim foreign_sellcashtotal, foreign_suplycashtotal,foreign_sellcashtotalfix, foreign_suplycashtotalfix
selltotal =0
suplytotal =0
buytotal =0
selltotalfix =0
suplytotalfix =0
buytotalfix =0
totalfixno = 0
foreign_sellcashtotal = 0
foreign_suplycashtotal = 0
foreign_sellcashtotalfix = 0
foreign_suplycashtotalfix = 0

%>
<br>
<!--
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td align="right">
			<table width="300" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align=center bgcolor="#FFDDDD">
				<td></td>
				<td width=120>����</td>
				<td width=120>��ǰ��</td>
			</tr>
			<tr align=center bgcolor="#FFFFFF">
				<td><b>C</b></td>
				<td width=120>��Ź���->�������</td>
				<td width=120>����ǰ->��Ź���</td>
			</tr>
			<tr align=center bgcolor="#FFFFFF">
				<td><b>S</b></td>
				<td width=120>�������->���</td>
				<td width=120>����ǰ->�������</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
-->

<br>

* �ֹ����� ���Ŀ��� �������Ϳ����� ������ ������ �� �ֽ��ϴ�
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="21" align="right">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF">
				<td>
					<%
					'/�ֹ��� �ۼ���
					if jumunwait then
					%>
						<% if IsForeign_confirmed then %>
							<input type="button" class="button" value="���ó�������" onClick="DelArr()">
						<% end if %>
					<%
					'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
					elseif (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
					%>
						<input type="button" class="button" value="���ó�������(�������)" onClick="DelArr()">
					<% elseif C_ADMIN_AUTH then %>
						<input type="button" class="button" value="���ó�������(�����ڸ��)" onClick="DelArr()">
					<% end if %>

		        	<font color="#FF0000">�ٹ�</font>&nbsp;
		        	<font color="#000000">����</font>&nbsp;
		        	<font color="#0000FF">��������</font>
					������ ���� : 
					<input type="text" class="text" style="text-align:right;" name="storemarginrate" id="storemarginrate" value="<%= storemarginrate %>" size="2"> %
			        <input type="button" class="button" value="���� ����������" onclick='ApplyMargin($("#storemarginrate").val());'>
				</td>
				<td align="right">
					�ѰǼ�:  <%= ojumundetail.FResultCount %>
					&nbsp;
					<%
					'/�ֹ��� �ۼ���
					if jumunwait then
					%>
						<% 'if IsForeign_confirmed then %>
							<input type="button" class="button" value="��ǰ�߰�" onclick="AddItems(frmMaster)">
							<input type="button" class="button" value="��ǰ�߰�(NEW)" onclick="AddItems_locale(frmMaster)">
							<input type="button" class="button" value="����(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'P')">
							<input type="button" class="button" value="��ǰ(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'M')">
						<% 'end if %>
					<%
					'/��������, �濵������Ʈ���̻� ������ ����� ������ �ֹ��� ������.
					elseif (C_logics_Part or C_MngPowerUser) and ojumunmaster.FOneItem.Fstatecd<7 then
					%>
						<input type="button" class="button" value="��ǰ�߰�(�������)" onclick="AddItems(frmMaster)">
						<input type="button" class="button" value="��ǰ�߰�(NEW)" onclick="AddItems_locale(frmMaster)">
						<input type="button" class="button" value="����(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'P')">
						<input type="button" class="button" value="��ǰ(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'M')">
					<% elseif C_ADMIN_AUTH then %>
						<input type="button" class="button" value="��ǰ�߰�(�����ڸ��)" onclick="AddItems(frmMaster)">
						<input type="button" class="button" value="��ǰ�߰�(NEW)" onclick="AddItems_locale(frmMaster)">
						<input type="button" class="button" value="����(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'P')">
						<input type="button" class="button" value="��ǰ(���ڵ�)" onclick="AddItemsBarCode(frmMaster,'M')">
					<% end if %>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="ckall" onClick="AnSelectAllFrame(this.checked)"></td>
    <td width="50">�̹���</td>
	<td width="100">��ǰ�ڵ�</td>
	<td width="100">������ڵ�</td>
	<td width="100">��ü�����ڵ�</td>
	<td>�귣��</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="60">�Һ��ڰ�</td>
	<td width="60">���<br>���ް�</td>
	<td width="60">���԰�</td>
	<td width="30">���<br>����</td>
	<td width="30">����<br>����</td>

	<% if IsForeignOrder then %>
		<td width="60">�ֹ���<br>�ؿ��ǸŰ�</td>
		<td width="60">�ֹ���<br>�ؿܰ��ް�</td>
	<% end if %>

	<td width="50">�ֹ���</td>
	<td width="50">���ּ�</td>
	<td width="50">Ȯ����</td>
	<td width="50">��ǰ��</td>
	<td width="30">����<br>����<br>����</td>

	<% if isfixed then %>
		<td>���</td>
		<!-- td width="40" align=center >C:S</td -->
	<% else %>
		<td width="100">���</td>
		<!-- td width="40" align=center >C:S</td -->
	<% end if %>
</tr>
<% for i=0 to ojumundetail.FResultCount-1 %>
<%
selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
buytotal   = buytotal + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Fbaljuitemno
selltotalfix  = selltotalfix + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Frealitemno
suplytotalfix = suplytotalfix + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Frealitemno
buytotalfix   = buytotalfix + ojumundetail.FItemList(i).Fbuycash * ojumundetail.FItemList(i).Frealitemno
totalfixno = totalfixno + ojumundetail.FItemList(i).Frealitemno

foreign_sellcashtotal  = foreign_sellcashtotal + ojumundetail.FItemList(i).fforeign_sellcash * ojumundetail.FItemList(i).Fbaljuitemno
foreign_suplycashtotal = foreign_suplycashtotal + ojumundetail.FItemList(i).fforeign_suplycash * ojumundetail.FItemList(i).Fbaljuitemno
foreign_sellcashtotalfix  = foreign_sellcashtotalfix + ojumundetail.FItemList(i).fforeign_sellcash * ojumundetail.FItemList(i).Frealitemno
foreign_suplycashtotalfix = foreign_suplycashtotalfix + ojumundetail.FItemList(i).fforeign_suplycash * ojumundetail.FItemList(i).Frealitemno
%>
<form name="frmBuyPrc_<%= i %>" method="post" action="shopjumun_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="masteridx" value="<%= idx %>">
<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fidx %>">
<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
<input type="hidden" name="orgsellprice" value="<%= ojumundetail.FItemList(i).Forgsellprice %>">
<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><img src="<%= ojumundetail.FItemList(i).GetImageSmall %>" border="0" width="50" height="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td>
		<%

		if (Not ojumundetail.FItemList(i).IsOnLineItem) then
			tmpcolor = "#0000FF"
		else
			if (ojumundetail.FItemList(i).IsUpchebeasong = True) then
				tmpcolor = "#000000"
			else
				tmpcolor = "#FF0000"
			end if
		end if

		%>

		<font color="<%= tmpcolor %>">
		<%= ojumundetail.FItemList(i).FItemGubun %><%= CHKIIF(ojumundetail.FItemList(i).FItemID>=1000000,format00(8,ojumundetail.FItemList(i).FItemID),format00(6,ojumundetail.FItemList(i).FItemID)) %><%= ojumundetail.FItemList(i).Fitemoption %>
		</font>
	</td>
	<td>
		<a href="javascript:publicbarreg('<%= ojumundetail.FItemList(i).FItemGubun %><%= BF_GetFormattedItemId(ojumundetail.FItemList(i).FItemID) %><%= ojumundetail.FItemList(i).Fitemoption %>');">
		<% if ojumundetail.FItemList(i).FPublicBarcode<>"" then %>
			<font color="#AAAAAA"><b><%= ojumundetail.FItemList(i).FPublicBarcode %></b></font>
		<% else %>
			���>>
		<% end if %>
		</a>
	</td>
	<td align="left"><%= ojumundetail.FItemList(i).FUpcheManageCode %></td>
	<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
	<td align="left">
		<%= ojumundetail.FItemList(i).Fitemname %>
		<% if ojumundetail.FItemList(i).Fitemoption <> "0000" then %>
			<font color="blue">[<%= ojumundetail.FItemList(i).Fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td align="right">
		<% if   (ojumundetail.FItemList(i).FItemDefaultMwDiv<>"W") and (ojumundetail.FItemList(i).Fbuycash>ojumundetail.FItemList(i).Fsuplycash) then %>
		<b><font color=red><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></font></b>(��)
		<% else %>
		<%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %>
		<% end if %>

		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).Fsellcash<>ojumundetail.FItemList(i).Fonlinesellcash) then %>
		<br>
		<div ><font color=red>��:<%= FormatNumber(ojumundetail.FItemList(i).Fonlinesellcash,0) %></font></div>
		<% end if %>
	</td>
	<td align="right">
		<input type="text" class="text" name="suplycash" value="<%= ojumundetail.FItemList(i).Fsuplycash %>" size="7" maxlength="9" style="text-align:right" <% if isfixed then response.write "readonly" %> >
		<% if (ojumundetail.FItemList(i).IsOnLineItem) and (ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine<>ojumundetail.FItemList(i).Fsuplycash) then %>
			<div ><font color=red><%= ojumundetail.FItemList(i).GetOrgShopSuplycashbyMargine %></font></div>
		<% end if %>
	</td>
	<td align="right">
		<% if (((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv<>"4")) or ((ojumundetail.FItemList(i).FItemDefaultMwDiv="M") and (ojumundetail.FItemList(i).FoffChargeDiv="4"))) then %>
		<input type="text" class="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="7" maxlength="9" style="text-align:right; color:#888888" <% if isfixed then response.write "readonly" %>>
		<% else %>
		<input type="text" class="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="7" maxlength="9" style="text-align:right" <% if isfixed then response.write "readonly" %>>
		<% end if %>

		<% if (ojumundetail.FItemList(i).Fbuycash<>ojumundetail.FItemList(i).Fonlinebuycash) and ((ojumundetail.FItemList(i).FItemDefaultMwDiv="W") and (ojumundetail.FItemList(i).FoffChargeDiv="4")) then %>
		<div ><font color="red">��:<%= ojumundetail.FItemList(i).Fonlinebuycash %></font></div>
		<% end if %>
	</td>
	<td align="center">
		<% if (ojumundetail.FItemList(i).Fsellcash <> 0) then %>
			<%= (100-CLng(ojumundetail.FItemList(i).Fsuplycash/ojumundetail.FItemList(i).Fsellcash*100*100)/100) %> %
		<% end if %>
	</td>
	<td align="center">
		<% if (ojumundetail.FItemList(i).Fsellcash <> 0) then %>
			<%= (100-CLng(ojumundetail.FItemList(i).Fbuycash/ojumundetail.FItemList(i).Fsellcash*100*100)/100) %> %
		<% end if %>
	</td>

	<% if IsForeignOrder then %>
		<td align="right">
			<input type="text" class="text" name="foreign_sellcash" value="<%= round(ojumundetail.FItemList(i).fforeign_sellcash,2) %>" size="7" maxlength="9" style="text-align:right" <% if isfixed then response.write "readonly" %>>
		</td>
		<td align="right">
			<input type="text" class="text" name="foreign_suplycash" value="<%= round(ojumundetail.FItemList(i).fforeign_suplycash,2) %>" size="7" maxlength="9" style="text-align:right" <% if isfixed then response.write "readonly" %>>
		</td>
	<% else %>
		<input type="hidden" name="foreign_sellcash" value="<%= ojumundetail.FItemList(i).fforeign_sellcash %>">
		<input type="hidden" name="foreign_suplycash" value="<%= ojumundetail.FItemList(i).fforeign_suplycash %>">
	<% end if %>

	<td align=center><input type="text" class="text" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="5" maxlength="8" style="text-align:right" <% if isfixed then response.write "readonly" %> ></td>
	<td align=center><%= ojumundetail.FItemList(i).Frealbaljuitemno %></td>
	<td align=center><input type="text" class="text" name="realitemno" value="<%= ojumundetail.FItemList(i).Frealitemno %>"  size="5" maxlength="8" style="text-align:right" <% if isfixed then response.write "readonly" %> ></td>
	<td align=center>
		<% if Not IsNull(ojumunmaster.FOneItem.Fcheckusersn) then %>
			<% if (ojumundetail.FItemList(i).Frealitemno <> ojumundetail.FItemList(i).Fcheckitemno) and Not IsNull(ojumunmaster.FOneItem.Fcheckusersn) then %><b><font color=red>&lt;=&nbsp;&nbsp;<% end if %>
			<%= ojumundetail.FItemList(i).Fcheckitemno %>
		<% end if %>
	</td>
	<td align=center><%= ojumundetail.FItemList(i).Fcentermwdiv %></td>
<% if isfixed then %>
	<td>
		<%= ojumundetail.FItemList(i).Fcomment %>
		<input type="hidden" name="comment" value="<%= ojumundetail.FItemList(i).Fcomment %>">
		<input type="hidden" name="ipgoflag" value="<%= ojumundetail.FItemList(i).Fipgoflag %>">
		<div align=center><%= ojumundetail.FItemList(i).GetOn2Off2DivName %></div>
	</td>
	<!-- td align=center>
		<font color=red><%= ojumundetail.FItemList(i).Fipgoflag %></font>
	</td -->
<% else %>
	<td align="center">
		<% DrawMiChulgoDiv "comment", ojumundetail.FItemList(i).Fcomment %>
		<div align=center><%= ojumundetail.FItemList(i).GetOn2Off2DivName %></div>
	</td>
	<% if ((ojumundetail.FItemList(i).Fipgoflag="C") or (IsNULL(ojumundetail.FItemList(i).Fipgoflag) and (ojumundetail.FItemList(i).IsWi2Meaip))) then %>
	<input type="hidden" name="ipgoflag" value="C">
	<% elseif (ojumundetail.FItemList(i).Fipgoflag="S") then %>
	<input type="hidden" name="ipgoflag" value="S">
	<% else %>
	<input type="hidden" name="ipgoflag" value="">
	<% end if %>
	<!-- td align="center">
	<select class="select" name="ipgoflag">
		<option value=""></option>
		<option value="C" <% if ((ojumundetail.FItemList(i).Fipgoflag="C") or (IsNULL(ojumundetail.FItemList(i).Fipgoflag) and (ojumundetail.FItemList(i).IsWi2Meaip))) then response.write "selected" %> >C</option>
		<option value="S" <% if (ojumundetail.FItemList(i).Fipgoflag="S") then response.write "selected" %> >S</option>
	</select>
	</td -->
<% end if %>
	<input type=hidden name="defaultmaginflag" value="<%= ojumundetail.FItemList(i).GetNoinputDefaultmaginflag %>">
	<input type=hidden name="buymaginflag" value="<%= ojumundetail.FItemList(i).GetNoinputBuymaginflag %>">
	<input type=hidden name="suplymaginflag" value="<%= ojumundetail.FItemList(i).GetNoinputSuplymaginflag %>">
</tr>
</form>
<% next %>

<% if (ojumundetail.FResultCount>0) then %>
<tr bgcolor="#FFFFFF">
	<td ></td>
	<td align="center">�Ѱ�</td>
	<td colspan="5" align="center">
	<td align="right">
		<%= formatNumber(selltotal,0) %><br>
		<b><%= formatNumber(selltotalfix,0) %></b>
	</td>
	<td align="right">
		<%= formatNumber(suplytotal,0) %><br>
		<b><%= formatNumber(suplytotalfix,0) %></b>
	</td>
	<td align="right">
		<%= formatNumber(buytotal,0) %><br>
		<b><%= formatNumber(buytotalfix,0) %></b>
	</td>
	<td></td>
	<td></td>
	<% if IsForeignOrder then %>
		<td align="right">
			<%= getdisp_price(foreign_sellcashtotal, currencyChar) %>
			<br><b><%= getdisp_price(foreign_sellcashtotalfix, currencyChar) %></b>
		</td>
		<td align="right">
			<%= getdisp_price(foreign_suplycashtotal, currencyChar) %>
			<br><b><%= getdisp_price(foreign_suplycashtotalfix, currencyChar) %></b>
		</td>
	<% end if %>

	<td></td>
	<td></td>
	<td align=center><%= totalfixno %></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="21" align=center>
	<% if ojumunmaster.FOneItem.FStatecd="9" then %>
		<b>�԰� �Ϸ�� ������ ���� �Ͻ� �� �����ϴ�.</b>
		<% if (C_ADMIN_AUTH) then %>
		<input type="button" class="button" value=" ��ü����(������) " onclick="DelMaster(frmMaster)">
		<% end if %>
	<% elseif (ojumunmaster.FOneItem.FStatecd>"6") then %>
		<b>��� �Ϸ�� ������ ���� �Ͻ� �� �����ϴ�.</b>
		<% if (C_ADMIN_AUTH) then %>
		<input type="button" class="button" value=" ��ü����(������) " onclick="DelMaster(frmMaster)">
		<% end if %>
	<% else %>
		<input type="button" class="button" value=" ��ü����<% if (C_ADMIN_AUTH and Not IsForeign_confirmed) then %>(������)<% end if %>" onclick="SaveALL()">
		<input type="button" class="button" value=" ��ü���� " onclick="DelMaster(frmMaster)">
	<% end if %>
	</td>
</tr>
<form name="frmadd" method="post" action="shopjumun_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="shopjumunitemaddarr">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="detailidxarr" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="foreign_sellcasharr" value="">
	<input type="hidden" name="foreign_suplycasharr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="baljuitemnoarr" value="">
	<input type="hidden" name="realitemnoarr" value="">
	<input type="hidden" name="commentarr" value="">
	<input type="hidden" name="ipgoflagarr" value="">
	<input type="hidden" name="defaultmaginflagarr" value="">
	<input type="hidden" name="buymaginflagarr" value="">
	<input type="hidden" name="suplymaginflagarr" value="">
	<input type="hidden" name="yyyymmdd" value="">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="statecd" value="">
	<input type="hidden" name="beasongdate" value="">
	<input type="hidden" name="songjangdiv" value="">
	<input type="hidden" name="songjangno" value="">
	<input type="hidden" name="songjangname" value="">
	<input type="hidden" name="divcode" value="">
	<input type="hidden" name="targetid" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
	<input type="hidden" name="baljuid" value="<%= ojumunmaster.FOneItem.Fbaljuid %>">
	<input type="hidden" name="foreign_statecd" value="<%= ojumunmaster.FOneItem.fforeign_statecd %>">
</form>
<form name="frmedit" method="post" action="shopjumun_process.asp" style="margin:0px;">
	<input type="hidden" name="mode">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="foreign_statecd">
	<input type="hidden" name="targetid" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
	<input type="hidden" name="baljuid" value="<%= ojumunmaster.FOneItem.Fbaljuid %>">
</form>
</table>

<%
set oupchemwinfo = Nothing
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

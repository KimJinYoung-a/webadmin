<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �������� ��ǰ ��� ����
' Hieditor : 2011.10.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim itemgubun,itemid, itemoption, barcode ,i ,makerid ,ioffitem ,opartner ,ooffontract ,IsOnlineItem
dim editmode , CenterMwDiv ,offList ,offSmall ,OnlineSailYn , IsDirectIpchulContractExistsBrand
dim shopitemname ,shopitemoptionname ,cd1 ,cd2 ,cd3 ,cd1_name ,cd2_name ,cd3_name ,orgsellprice ,shopitemprice
dim shopsuplycash ,shopbuyprice ,isusing ,vatinclude, vatinclude10, extbarcode ,imageList ,offmain ,OnlineOrgprice
dim OnlineBuycash, mwDiv ,OnlineSellcash ,regdate ,updt,stockitemid, itemcopy, isupcheitemreg
	makerid = requestCheckVar(request("makerid"),32)
	barcode	  = request("barcode")

isupcheitemreg = false

'/����
if (C_IS_SHOP) then
	'//�������϶�
	if C_IS_OWN_SHOP then
	else
	end if
else
	'/��ü�� ��� ���̵� �ھƳ���
	if C_IS_Maker_Upche then
		makerid = session("ssBctId")
		IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(makerid)
		isupcheitemreg = getupcheitemregyn(makerid)

		if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then
			if not(isupcheitemreg) then
				response.write "���ѿ���"
				dbget.close() : response.End
			end if
		end if
	end if
end if

editmode = FALSE

'//�����ϰ��
if barcode <> "" and not(isnull(barcode)) then
	editmode = TRUE
    if len(barcode)=14 then
        itemgubun = Left(barcode,2)
	    itemid	  = CLng(Mid(barcode,3,8))
	    itemoption = Right(barcode,4)
    else
	    itemgubun = Left(barcode,2)
	    itemid	  = CLng(Mid(barcode,3,6))
	    itemoption = Right(barcode,4)
    end if

	set ioffitem  = new COffShopItem
		ioffitem.FRectItemgubun = itemgubun
		ioffitem.FRectItemId = itemid
		ioffitem.FRectItemOption = itemoption

		if C_IS_Maker_Upche then
		    ioffitem.FRectMakerid=makerid
		end if

		ioffitem.GetOffOneItem

	if ioffitem.FResultCount > 0 then
		makerid = ioffitem.FOneItem.Fmakerid
		Barcode = ioffitem.FOneItem.GetBarcode
		shopitemname = ioffitem.FOneItem.Fshopitemname
		shopitemoptionname = ioffitem.FOneItem.Fshopitemoptionname
		itemcopy = ioffitem.FOneItem.fitemcopy
		cd1 = ioffitem.FOneItem.FCateCDL
		cd2 = ioffitem.FOneItem.FCateCDM
		cd3 = ioffitem.FOneItem.FCateCDS
		cd1_name = ioffitem.FOneItem.FCateCDLName
		cd2_name = ioffitem.FOneItem.FCateCDMName
		cd3_name = ioffitem.FOneItem.FCateCDSName
		orgsellprice = ioffitem.FOneItem.FShopItemOrgprice
		shopitemprice = ioffitem.FOneItem.Fshopitemprice
		shopsuplycash = ioffitem.FOneItem.Fshopsuplycash
		shopbuyprice = ioffitem.FOneItem.Fshopbuyprice
		ItemGubun = ioffitem.FOneItem.FItemGubun
		isusing = ioffitem.FOneItem.Fisusing
		CenterMwDiv = ioffitem.FOneItem.FCenterMwDiv
		vatinclude = ioffitem.FOneItem.Fvatinclude
		vatinclude10 = ioffitem.FOneItem.Fvatinclude10
		extbarcode = ioffitem.FOneItem.Fextbarcode
		imageList = ioffitem.FOneItem.FimageList
		offmain = ioffitem.FOneItem.FOffImgMain
		offList = ioffitem.FOneItem.FOffImgList
		offSmall = ioffitem.FOneItem.FOffImgSmall
		OnlineSailYn = ioffitem.FOneItem.FOnlineSailYn
		OnlineOrgprice = ioffitem.FOneItem.FOnlineOrgprice
		OnlineBuycash = ioffitem.FOneItem.FOnlineBuycash
		mwDiv = ioffitem.FOneItem.FmwDiv
		OnlineSellcash = ioffitem.FOneItem.FOnlineSellcash
		regdate = ioffitem.FOneItem.Fregdate
		updt = ioffitem.FOneItem.Fupdt
		stockitemid = ioffitem.FOneItem.Fstockitemid
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('�ش�Ǵ� ��ǰ�� �����ϴ�');"
		'response.write "	self.close();"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	IsOnlineItem = (itemgubun="10")

	if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) and isupcheitemreg then
		'// ��ǰ ��� �� 1���ϰ� ��������
		if (DateDiff("d", regdate, Now()) > 14) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��ǰ ����� 2���ϰ��� ���� �����մϴ�.');"
			response.write "</script>"
			Response.write "��ǰ ����� 2���ϰ��� ���� �����մϴ�."
			dbget.close()	:	response.end
		end if
	end if

'/�űԵ��
else
	if makerid <> "" then
		CenterMwDiv = GetDefaultItemMwdivByBrand(makerid)
	end if
end if

set opartner = new CPartnerUser
    opartner.FRectDesignerID = makerid

    if makerid <> "" then
    	opartner.GetOnePartnerNUser
    end if

set ooffontract = new COffContractInfo
    ooffontract.FRectDesignerID = makerid

    if makerid <> "" then
		ooffontract.GetPartnerOffContractInfo
	end if

if vatinclude = "" then
	if opartner.FResultCount > 0 then
		if opartner.FOneItem.fjungsan_gubun="�鼼" then
			vatinclude="N"
		else
			vatinclude="Y"
		end if
	else
		vatinclude="Y"
	end if
end if
if isusing = "" then isusing = "Y"
'C_IS_SHOP = TRUE
%>

<script type="text/javascript">

//�űԵ�϶� �귣�� ����
function ChangeBrand(comp){
	location.href="?makerid=" + comp.value;
}

//����
function EditItem(frm){
	var tmpitemgubuncheck = '';

	<% if editmode then %> var editmode = true; <% else %> var editmode = false; <% end if %>
	<% if C_IS_Maker_Upche then %> var C_IS_Maker_Upche = true; <% else %> var C_IS_Maker_Upche = false; <% end if %>
	<% if C_ADMIN_USER then %> var C_ADMIN_USER = true; <% else %> var C_ADMIN_USER = false; <% end if %>
	<% if C_IS_SHOP then %> var C_IS_SHOP = true; <% else %> var C_IS_SHOP = false; <% end if %>
	<% if C_IS_OWN_SHOP then %> var C_IS_OWN_SHOP = true; <% else %> var C_IS_OWN_SHOP = false; <% end if %>

	//��ǰ���� ���ð� üũ
	if (editmode){
		tmpitemgubuncheck = frm.itemgubun.value;
	}else{
		var itemgubun = document.getElementsByName("itemgubun");
		for(var i=0; i < itemgubun.length ; i++){
			if (itemgubun[i].checked){
				tmpitemgubuncheck = frm.itemgubun[i].value;
			}
		}
	}

	if (!editmode){
		if (tmpitemgubuncheck == ''){
			alert('��ǰ������ �����ϼ���.');
			return;
		}
	}

	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	}

	//itemgubun 00 ������Ի�ǰ
	if (tmpitemgubuncheck!='00'){
		if (frm.cd1.value.length<1){
			alert('ī�װ��� �����ϼ���.');
			return;
		}
	}

	if (editmode){
	    if (frm.orgsellprice.value.length<1){
			alert('�Һ��ڰ��� �Է��ϼ���.');
			frm.orgsellprice.focus();
			return;
		}
	}

	if (frm.shopitemprice.value.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.shopitemprice.focus();
		return;
	}

	if (C_ADMIN_USER || C_IS_OWN_SHOP){
		if (frm.shopsuplycash.value.length<1){
			alert('���԰��� �Է��ϼ���.');
			frm.shopsuplycash.focus();
			return;
		}
	}

	if (tmpitemgubuncheck=='85') {
		alert('�¶��� ����ǰ�� ������ �� �����ϴ�.\n\n[ON]��ǰ����>>�ٹ� ����ǰ ���� �޴��� �̿��ϼ���.');
		return;
	}

	//���α�
	if (tmpitemgubuncheck=='60'){
		if (editmode){
		    if (frm.orgsellprice.value.substr(0,1) != '-'){
				frm.orgsellprice.value = "-"+frm.orgsellprice.value
			}
		}
	    if (frm.shopitemprice.value.substr(0,1) != '-'){
			frm.shopitemprice.value = "-"+frm.shopitemprice.value
		}
	//����ǰ
	}else if (tmpitemgubuncheck=='80'){
	    if (frm.shopitemprice.value > 0){
			alert("����ǰ�� �ǸŰ��� 0���Ͽ��� �մϴ�.");
			frm.shopitemprice.focus();
			return;
		}
		if (editmode){
		    if (frm.orgsellprice.value > 0){
				alert("����ǰ�� �Һ��ڰ� 0���Ͽ��� �մϴ�.");
				frm.orgsellprice.focus();
				return;
			}
		}
		if (editmode){
			if (frm.shopitemname.value.match(/^\[����ǰ\] /) == null) {
				alert("����ǰ ������ ������ �� �����ϴ�.");
				return;
			}
		}
		if (!editmode){
			if (frm.shopitemname.value.match(/����ǰ/) != null) {
				alert("����ǰ ������ ��ǰ�� �ڵ��Էµ˴ϴ�. ����ǰ ������ ���켼��.");
				return;
			}
		}

	//itemgubun 00 ������Ի�ǰ
	}else if (tmpitemgubuncheck!='00'){
	    if (!IsDigit(frm.shopitemprice.value)){
			alert('�ǸŰ��� ���ڸ� �����մϴ�.');
			frm.shopitemprice.focus();
			return;
		}
		if (editmode){
		    if (!IsDigit(frm.orgsellprice.value)){
				alert('�Һ��ڰ��� ���ڸ� �����մϴ�.');
				frm.orgsellprice.focus();
				return;
			}
		}
	}else{
		if (!IsInteger(frm.shopitemprice.value)){
			alert('�ǸŰ��� ���ڸ� �����մϴ�.');
			frm.shopitemprice.focus();
			return;
		}
		if (editmode){
		    if (!IsInteger(frm.orgsellprice.value)){
				alert('�Һ��ڰ��� ���ڸ� �����մϴ�.');
				frm.orgsellprice.focus();
				return;
			}
		}
	}

	if (frm.extbarcode.value != ''){
		str = frm.extbarcode.value;
		for (j=0; j<str.length; j++){
			checkStr = str.charAt(j);
			if(/\W/.test(checkStr) && /[^\s]/.test(checkStr)){
				alert("������ڵ忡 Ư�����ڴ� ������� �ʽ��ϴ�.");
				frm.extbarcode.focus();
				return;
			}
		}

		if (frm.extbarcode.value.length < 8){
			alert('������ڵ� ���̰� �ʹ� ª���ϴ�(8�� �̻�).\n���� ���ڵ尡 �ִ°�츸 �Է��� �ּ���');
			frm.extbarcode.focus();
			return;
		}
	}

	if (!editmode){
		if (C_ADMIN_USER){
			if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
				if (!confirm('!! �⺻ ��� ������ �ٸ� ��쿡�� ���԰� ���ް��� �Է� �ϼž� �մϴ�. \n\n��� �Ͻðڽ��ϱ�?')){
					return;
				}
			}
		}
	}

	//����ǰ�� �ƴҰ��
	if (editmode){
		if (tmpitemgubuncheck!='80'){
		    if (frm.orgsellprice.value*1<frm.shopitemprice.value*1){
		        alert('�Һ��ڰ����� �� �ǸŰ��� Ŭ �� �����ϴ�. �ٽ� �Է��ϼ���.');
				frm.shopitemprice.focus();
				return;
		    }
		}
	}

	if (tmpitemgubuncheck!='10'){
		if (editmode){
			if (frm.tmpoffmain.value.length<1 && frm.file1.value.length<1){
				alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
				frm.file1.focus();
				return;
			}
		}else{
			if (frm.file1.value.length<1){
				alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
				frm.file1.focus();
				return;
			}
		}
	}

	if (!C_IS_Maker_Upche && frm.centermwdiv.length>0){
	    if ((!frm.centermwdiv[0].checked)&&(!frm.centermwdiv[1].checked)){
	        alert('���� ���� ������ ���� �ϼ���.');
			frm.centermwdiv[0].focus();
			return;
	    }
	}

    if ((!frm.vatinclude[0].checked)&&(!frm.vatinclude[1].checked)){
        alert('���� ������ ���� �ϼ���.');
		frm.vatinclude[0].focus();
		return;
    }

	if (tmpitemgubuncheck == "10") {
		if ((frm.mwdiv.value != GetCenterMWDiv(frm)) && (frm.mwdiv.value != "U")) {
			alert("��ü��� ��ǰ�� ���Ա����� �¶��ΰ� �ٸ��� ������ �� �ֽ��ϴ�.");
			return;
		}

		if (frm.vatinclude10.value != GetVatinclude(frm)) {
			alert("���������� �¶��ΰ� �ٸ��� ������ �� �����ϴ�.");
			return;
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		if (tmpitemgubuncheck=='80') {

			if (frm.shopitemname.value.match(/����ǰ/) == null) {
				frm.shopitemname.value = "[����ǰ] " + frm.shopitemname.value;
			}
		}

		frm.submit();
	}
}

function jsSetNoDisp() {
	setCategory("999","999","999","���þ���","���þ���","���þ���");
}

function GetCenterMWDiv(frm) {
	if (frm.centermwdiv.value == undefined) {
		if (frm.centermwdiv[0].checked == true) {
			return frm.centermwdiv[0].value;
		} else if (frm.centermwdiv[1].checked == true) {
			return frm.centermwdiv[1].value;
		} else {
			return "";
		}
	} else {
		return frm.centermwdiv.value;
	}
}

function GetVatinclude(frm) {
	if (frm.vatinclude[0].checked == true) {
		return frm.vatinclude[0].value;
	} else if (frm.vatinclude[1].checked == true) {
		return frm.vatinclude[1].value;
	} else {
		return "";
	}
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo",'width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ī�װ����
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//ī�װ� ����
function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frmedit;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}

function showCateCode() {
	var frm = document.frmedit;
	alert("�� ī�װ� �ڵ� : " + frm.cd1.value + "\n" + "�� ī�װ� �ڵ� : " + frm.cd2.value + "\n" + "�� ī�װ� �ڵ� : " + frm.cd3.value);
}

</script>

<!-- ����Ʈ ���� -->
>>�������� ��ǰ ���
<form name="frmedit" method="post" action="<%=uploadImgUrl%>/linkweb/offshop/item/itemedit_off.asp" enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<input type="hidden" name="editmode" value="<%=editmode%>">
<input type="hidden" name="barcode" value="<%=barcode%>">
<input type="hidden" name="offmain" value="<%=offmain%>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="C_IS_SHOP" value="<%= C_IS_SHOP %>">
<input type="hidden" name="C_IS_Maker_Upche" value="<%= C_IS_Maker_Upche %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<% if NOT(editmode) then %>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width="100">�귣��ID</td>
	<td bgcolor="#FFFFFF">
		<% if C_IS_Maker_Upche then %>
			<%= makerid %>
			<input type="hidden" name="makerid" value="<%= makerid %>">
		<% else %>
			<%
			''drawOffContractBrandChangeEvent "makerid",makerid
			''2015-09-02, skyer9
			drawSelectBoxDesignerwithName "makerid",makerid
			%>
			&nbsp;
			&nbsp;
			&nbsp;
			<input type="button" class="button" value=" �� �� " onClick="ChangeBrand(document.frmedit.makerid)">
			&nbsp;
			�ؽű� ����Ͻ� ��ǰ�� �귣�带 ������ �ּ���.
		<% end if %>
	</td>
</tr>
<%
end if

'/�귣�� ������ ������� �������� �ʰ�, ������ �귣�� �����ϵ���..
if makerid = "" then dbget.close() : response.write "</table>" : response.end
%>

<% if (opartner.FResultCount>0) then %>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>�귣��������</td>
	<td bgcolor="#FFFFFF">
		<%= makerid %> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td>�¶���</td>
	<td bgcolor="#FFFFFF">
		<%= FormatNumber(OnlineOrgprice,0) %> / <%= FormatNumber(OnlineBuycash,0) %>
		&nbsp;&nbsp;
		<font color="<%= mwdivColor(mwDiv) %>"><%= mwdivName(mwDiv) %></font>
		&nbsp;
		<% if OnlineSellcash<>0 then %>
		<%= CLng((1- OnlineBuycash/OnlineOrgprice)*100) %> %
		<% end if %>

		<% if (OnlineSailYn="Y") then %>
		<br>
		<font color="red">
		<%= FormatNumber(OnlineSellcash,0) %> / <%= FormatNumber(OnlineBuycash,0) %>
		&nbsp;&nbsp;
			<% if (OnlineOrgprice<>0) then %>
		        <%= CLng((OnlineOrgprice - OnlineSellcash)/OnlineOrgprice*100) %>%
		    <% end if %>
		    ����
		</font>
		&nbsp;&nbsp;
		<font color="<%= mwdivColor(mwDiv) %>"><%= mwdivName(mwDiv) %></font>
		&nbsp;
			<% if OnlineSellcash<>0 then %>
				<%= CLng((1- OnlineBuycash/OnlineSellcash)*100) %> %
			<% end if %>

		<% end if %>

	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>��������<br>[������]</td>
	<td bgcolor="#FFFFFF">
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop000','<%= makerid %>')"><b>��������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="1") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td width=60><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td width=60><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>��������<br>[������]</td>
	<td bgcolor="#FFFFFF">
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop800','<%= makerid %>')"><b>����������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3")  then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>

		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>��������<br>[�ؿܰ���]</td>
	<td bgcolor="#FFFFFF">
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop870','<%= makerid %>')"><b>�ؿܰ��޴�ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop870") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop870") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="5")  then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>

<% if (editmode) then %>
	<input type="hidden" name="itemgubun" value="<%=itemgubun%>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">��ǰ�ڵ�</td>
		<td bgcolor="#FFFFFF">
			<%= Barcode %>
			<%if left(Barcode,2) = "10" then %>
				�¶��ΰ����ǰ
			<% elseif left(Barcode,2) = "55" then %>
				CS����
			<% elseif left(Barcode,2) = "90" then %>
				�������������ǰ
			<% elseif left(Barcode,2) = "95" then %>
				���������������ǸŻ�ǰ
			<% elseif left(Barcode,2) = "80" then %>
				OFF����ǰ
			<% elseif left(Barcode,2) = "85" then %>
				ON����ǰ
			<% elseif left(Barcode,2) = "75" then %>
				������
			<% elseif left(Barcode,2) = "70" then %>
				�Ҹ�ǰ
			<% elseif left(Barcode,2) = "76" then %>
				�ΰŽ� ������
			<% end if %>
			<br><font color="#AAAAAA">(90������������, 80OFF����ǰ, 85ON����ǰ, 70�Ҹ�ǰ, 75������, 76�ΰŽ�������, 95���������������Ǹ�, 60���α�, 55CS����, 00������Ի�ǰ)</font>
		</td>
	</tr>
<% else %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width=100>��ǰ����</td>
		<td bgcolor="#FFFFFF">
		<input type="radio" name="itemgubun" value="90" <% if itemgubun = "90" then response.write " checked" %>>������ �����ǰ(90)
		<input type="radio" name="itemgubun" value="70" <% if itemgubun = "70" then response.write " checked" %> disabled>�Ҹ�ǰ(70)
		<% if NOT(C_IS_Maker_Upche) then %>
			<input type="radio" name="itemgubun" value="75" <% if itemgubun = "75" then response.write " checked" %>>������(75)
			<input type="radio" name="itemgubun" value="76" <% if itemgubun = "76" then response.write " checked" %>>�ΰŽ�������(76)
			<input type="radio" name="itemgubun" value="80" <% if itemgubun = "80" then response.write " checked" %>>OFF����ǰ(80)					<!-- �����ڸ� ��� �����ϰ� �����Ѱ� ����, skyer9, 2017-06-07 -->
			<input type="radio" name="itemgubun" value="85" <% if itemgubun = "85" then response.write " checked" %>>ON����ǰ(85)
			<input type="radio" name="itemgubun" value="60" <% if itemgubun = "60" then response.write " checked" %> disabled>���α�(60)
			<input type="radio" name="itemgubun" value="55" <% if itemgubun = "55" then response.write " checked" %>>CS����(55)
		<% end if %>
		<br><font color="#AAAAAA">(90������������, 80OFF����ǰ, 85ON����ǰ, 70�Ҹ�ǰ, 75������, 95���������������Ǹ�, 60���α�, 55CS����, 00������Ի�ǰ)</font>
		</td>
	</tr>
<% end if %>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰ��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="80" maxlength="90">
		<br>�� ����ǰ�� ��ǰ�� "[����ǰ]" ������ �ڵ����� �ٽ��ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�ɼǸ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopitemoptionname" value="<%= shopitemoptionname %>" size="40" maxlength="40" class="input_01">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰī��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemcopy" value="<%= itemcopy %>" size="80" maxlength="255">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>ī�װ�</td>
	<td bgcolor="#FFFFFF">
	  <input type="hidden" name="cd1" value="<%= cd1 %>">
	  <input type="hidden" name="cd2" value="<%= cd2 %>">
	  <input type="hidden" name="cd3" value="<%= cd3 %>">

      <input type="text" class="text" name="cd1_name" value="<%= cd1_name %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" class="text" name="cd2_name" value="<%= cd2_name %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" class="text" name="cd3_name" value="<%= cd3_name %>" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" class="button" value="����" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">

	  &nbsp;
	  <a href="javascript:showCateCode('aa')">�ڵ庸��</a>

	  &nbsp;
	  <input type="button" class="button" name="" value="���þ��� ����" onClick="jsSetNoDisp()">
	</td>
</tr>

<% if editmode then %>
	<tr bgcolor="<%= adminColor("tabletop") %>" >
		<td>�Һ��ڰ�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="orgsellprice" value="<%= orgsellprice %>" <% if C_IS_SHOP or C_IS_Maker_Upche then response.write " readonly" %> size=8 maxlength=9 class="input_right" >
		</td>
	</tr>
<% else %>
	<input type="hidden" name="orgsellprice" value="<%= orgsellprice %>">
<% end if %>

<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>�ǸŰ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopitemprice" value="<%= shopitemprice %>" <% if editmode and C_IS_Maker_Upche then response.write " readonly" %> size=8 maxlength=9 class="input_right" >
        <Br>�ؿ¶��� �Ǹ� ��ǰ�� ��� ���� ������ �¶��� �ǸŰ��� �����ϰ� �����˴ϴ�.</b>
	</td>
</tr>

<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	<tr bgcolor="<%= adminColor("tabletop") %>" >
		<td>���԰�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="shopsuplycash" value="<%= shopsuplycash %>" <% if not(C_ADMIN_USER) then response.write " readonly" %> size=8 maxlength=9 class="input_right">
			<Br>��0�ΰ�� �⺻���� �ڵ� ����
			<Br>�ػ���ǰ�� ��� ���� ������ �������
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" >
		<td>������ް�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="shopbuyprice" value="<%= shopbuyprice %>" <% if not(C_ADMIN_USER) then response.write " readonly" %> size=8 maxlength=9 class="input_right" >
			<Br>��0�ΰ�� �⺻���� �ڵ� ����
			<Br>�ػ���ǰ�� ��� ���� ������ �������
		</td>
	</tr>
<% end if %>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�������</td>
	<td bgcolor="#FFFFFF">
		<% if isusing = "Y" then %>
		<input type=radio name=isusing value="Y" checked >�����
		<input type=radio name=isusing value="N">������
		<% else %>
		<input type=radio name=isusing value="Y"  >�����
		<input type=radio name=isusing value="N" checked >������
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>���͸��Ա���</td>
	<td bgcolor="#FFFFFF" height="25">
		<%IF stockitemid = 0 or C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH THEN %>
    		<input type="radio" name="centermwdiv" value="W" <%= ChkIIF(centermwdiv="W","checked","") %> >Ư��
    		<input type="radio" name="centermwdiv" value="M" <%= ChkIIF(centermwdiv="M","checked","") %> >����
			&nbsp;
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
			[�����ں�]
			<% end if  %>
			<% if (stockitemid = 0) then  %>
			[����������]
			<% end if %>
		<%ELSE%>
		<%= fnColor(centermwdiv,"mw") %>
		<input type="hidden" name="centermwdiv" value="<%=centermwdiv%>">
		<%END IF%>

		<input type="hidden" name="mwdiv" value="<%= mwDiv %>">
		<% if (itemgubun = "10") then %>
		&nbsp;&nbsp;
		(�¶��� : <%= mwDiv %>)
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>��������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="vatinclude" value="Y" <%= ChkIIF(vatinclude = "Y","checked","") %>  >����
		<input type="radio" name="vatinclude" value="N" <%= ChkIIF(vatinclude = "N","checked","") %> > <font color="<%= ChkIIF(vatinclude = "N","blue","#000000") %>">�鼼</font>

		<input type="hidden" name="vatinclude10" value="<%= vatinclude10 %>">
		<% if (itemgubun = "10") then %>
			&nbsp;&nbsp;
			(�¶��� : <%= vatinclude10 %>)
		<% end if %>
		<% if opartner.FOneItem.fjungsan_gubun<>"" and not(isnull(opartner.FOneItem.fjungsan_gubun)) then %>
			&nbsp;&nbsp;
			(�귣���� : <%= opartner.FOneItem.fjungsan_gubun %>)
		<% end if %>
		
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>������ڵ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="extbarcode" value="<%= extbarcode %>" size="20" maxlength="20" class="input_01" >
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�̹���</td>
	<td bgcolor="#FFFFFF">
		<% if IsOnlineItem then %>
			<img src="<%= imageList %>" width="50" height="50">
		<% else %>
			<input type="file" name="file1" class="button" size=20 >
			<Br>�� �⺻ �̹����� �ݵ�� 400x400 , jpg ���Ϸ� �÷��ֽñ� �ٶ��ϴ�.
			<Br>�� 400x400 �̹����� ���� �Ͻø�, �ڵ����� 100x100 , 50x50 �� ���� �˴ϴ�.
			<input type="hidden" name="tmpoffmain" value="<%= offmain %>">
   				<% IF offmain <> "" THEN %>
	   				<BR><img src="<%=offmain%>" border="0" width=400 height=400> 400x400
   				<% END IF %>
   				<% if offlist <> "" then %>
   					<BR><img src="<%=offlist%>" border="0" width=100 height=100> 100x100
   				<% end if %>
   				<% if offsmall <> "" then %>
   					<BR><img src="<%=offsmall%>" border="0" width=50 height=50> 50x50
   				<% end if %>
		<% end if %>
	</td>
</tr>

<% if editmode then %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>�����</td>
	<td bgcolor="#FFFFFF"><%= regdate %></td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>����������</td>
	<td bgcolor="#FFFFFF"><%= updt %></td>
</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center>
		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<input type="button" class="button" value="<% if editmode then %>����<% else %>�ű�����<% end if %>" onclick="EditItem(frmedit)">
		<% end if %>
	</td>
</tr>
<% end if %>

</table>
</form>

<%
function drawOffContractBrandChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onchange="ChangeBrand(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select c.userid, c.socname_kor from [db_user].[dbo].tbl_user_c c with (nolock)"
   query1 = query1 & " , [db_shop].[dbo].tbl_shop_designer s"
   query1 = query1 & " where c.userid = s.makerid "
   query1 = query1 & " and s.shopid='streetshop000'"
   query1 = query1 & " order by c.userid"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Function

set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

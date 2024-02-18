<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/shopify/shopifycls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, shopifyGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, deliverytype, mwdiv
Dim expensive10x10, diffPrc, shopifyYes10x10No, shopifyNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing, cisextusing, rctsellcnt
Dim page, i, research
Dim oshopify
Dim priceOption
Dim exctrans
Dim startMargin, endMargin
Dim isusing, itemweight, deliverOverseas, sitename
page    				= request("page")
research				= request("research")
itemid  				= request("itemid")
makerid					= request("makerid")
itemname				= request("itemname")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
startMargin				= request("startMargin")
endMargin				= request("endMargin")
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")
infoDiv					= request("infoDiv")
morningJY				= request("morningJY")
extsellyn				= request("extsellyn")
shopifyGoodNo			= request("shopifyGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
shopifyYes10x10No		= request("shopifyYes10x10No")
shopifyNo10x10Yes		= request("shopifyNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
isextusing				= requestCheckVar(request("isextusing"), 1)
isusing					= request("isusing")
itemweight				= request("itemweight")
deliverOverseas			= request("deliverOverseas")
sitename				= request("sitename")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	bestOrd = "on"
	sellyn = "Y"
End If

'�ٹ����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
'shopify ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If shopifyGoodNo <> "" then
	Dim iA2, arrTemp2, arrshopifyGoodNo
	shopifyGoodNo = replace(shopifyGoodNo,",",chr(10))
	shopifyGoodNo = replace(shopifyGoodNo,chr(13),"")
	arrTemp2 = Split(shopifyGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrshopifyGoodNo = arrshopifyGoodNo & trim("'"&arrTemp2(iA2)&"'") & ","
		End If
		iA2 = iA2 + 1
	Loop
	shopifyGoodNo = left(arrshopifyGoodNo,len(arrshopifyGoodNo)-1)
End If

Set oshopify = new Cshopify
	oshopify.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oshopify.FPageSize					= 50
Else
	oshopify.FPageSize					= 20
End If
	oshopify.FRectCDL					= request("cdl")
	oshopify.FRectCDM					= request("cdm")
	oshopify.FRectCDS					= request("cds")
	oshopify.FRectItemID				= itemid
	oshopify.FRectItemName				= itemname
	oshopify.FRectSellYn				= sellyn
	oshopify.FRectLimitYn				= limityn
	oshopify.FRectSailYn				= sailyn
'	oshopify.FRectonlyValidMargin		= onlyValidMargin
	oshopify.FRectStartMargin			= startMargin
	oshopify.FRectEndMargin				= endMargin
	oshopify.FRectMakerid				= makerid
	oshopify.FRectshopifyGoodNo			= shopifyGoodNo
	oshopify.FRectMatchCate				= MatchCate
	oshopify.FRectIsMadeHand			= isMadeHand
	oshopify.FRectIsOption				= isOption
	oshopify.FRectIsReged				= isReged
	oshopify.FRectDeliverytype			= deliverytype
	oshopify.FRectMwdiv					= mwdiv
	oshopify.FRectIsextusing			= isextusing

	oshopify.FRectExtNotReg				= ExtNotReg
	oshopify.FRectExpensive10x10		= expensive10x10
	oshopify.FRectdiffPrc				= diffPrc
	oshopify.FRectshopifyYes10x10No		= shopifyYes10x10No
	oshopify.FRectshopifyNo10x10Yes		= shopifyNo10x10Yes
	oshopify.FRectExtSellYn				= extsellyn
	oshopify.FRectFailCntOverExcept		= ""
	oshopify.FRectFailCntExists			= failCntExists
	oshopify.FRectReqEdit				= reqEdit

	oshopify.FRectIsUsing				= isusing
	oshopify.FRectItemweight			= itemweight
	oshopify.FRectDeliverOverseas		= deliverOverseas
	oshopify.FRectSitename				= sitename
If (bestOrd = "on") Then
    oshopify.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oshopify.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oshopify.getshopifyreqExpireItemList
Else
	oshopify.getshopifyRegedItemList		'�� �� ����Ʈ
End If

' dotnet �׽�Ʈ��.
Dim dotnetApiURL
if application("Svr_Info")="Dev" then
	dotnetApiURL = "https://testscmplay.10x10.co.kr"
else
	dotnetApiURL = "https://scmplay.10x10.co.kr"
end if
' dotnetApiURL = "http://local.10x10.co.kr:5000"
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function checkisReged(comp){
    if (comp.name=="isReged"){
    	if(document.getElementById("QR").checked == true){ //��ϻ�ǰ �ǸŰ���
    		comp.form.ExtNotReg.value = "D"
   			//comp.form.ExtNotReg.disabled = true;
			comp.form.extsellyn.value = "N";
			comp.form.sellyn.value = "Y";
   		}else{
			if (document.getElementById("NR").checked == false){
			}else{
				comp.form.sitename.value = "Y";
				comp.form.isusing.value = "Y";
				comp.form.itemweight.value = "Y";
				comp.form.deliverOverseas.value = "Y";
//				comp.form.sellyn.value = "Y";
//				comp.form.MatchCate.value = "Y";
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="shopifyYes10x10No")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	//document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.isReged.checked = true;
			//comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "N";
			comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="shopifyNo10x10Yes")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	//document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			//comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "Y";
			comp.form.extsellyn.value = "N";
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.shopifyYes10x10No.checked){
            comp.form.shopifyYes10x10No.checked = false;
        }
        if (comp.checked){
        	//document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			//comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
	        comp.form.sellyn.value = "Y";
	        comp.form.onlyValidMargin.value="";
	        comp.form.extsellyn.value = "Y";
    	}
    }
	if ((comp.name=="diffPrc")){
        if (comp.checked){
        	//document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			//comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
        }
	}

	if (comp.name=="reqEdit"){
		if (comp.checked){
			//document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			//comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="shopifyYes10x10No")&&(frm.shopifyYes10x10No.checked)){ frm.shopifyYes10x10No.checked=false }
	if ((comp.name!="shopifyNo10x10Yes")&&(frm.shopifyNo10x10Yes.checked)){ frm.shopifyNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//��Ͽ��� ���� Reset
function ckeckReset(){
	document.frm.ExtNotReg.disabled = false;
	document.frm.wReset.checked=false;
	//document.getElementById("AR").checked=false;
	document.getElementById("NR").checked=false;
	document.getElementById("RR").checked=false;
	document.getElementById("QR").checked=false;
}
function onlyJY(comp){
     if ((comp.name=="morningJY")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	//document.getElementById("AR").checked=true;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.value="D"
			//comp.form.ExtNotReg.disabled = true;
			comp.form.sellyn.value = "A";
			comp.form.extsellyn.value = "";
			comp.form.onlyValidMargin.value="";
    	}
    }

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="shopifyYes10x10No")&&(frm.shopifyYes10x10No.checked)){ frm.shopifyYes10x10No.checked=false }
	if ((comp.name!="shopifyNo10x10Yes")&&(frm.shopifyNo10x10Yes.checked)){ frm.shopifyNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// TEST SubCategory
function shopifyCategoryProcess() {
    if (confirm('shopify�� ī�װ� Ȯ��?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "SubCategory";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function shopifySelectRegProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function checkshopifyItemDelete() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "DELETE";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// ��ǰ ��ȸ
function checkshopifyItemConfirm() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
    }
}

function collectionRefresh(v){
    if (confirm(v + ' collection�� �����ϰڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "collectionRefresh";
		document.frmSvArr.collectionType.value = v;
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
	}
}
// ���õ� ��ǰ �Ǹſ��� ����
function shopifySellYnProcess(chkYn) {
	var chkSel=0;
	var strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}
	if(chkYn == "Y"){
		strSell = "�Ǹ���";
	}else if(chkYn == "N"){
		strSell = "ǰ��";
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
	}
}

// ���õ� ��ǰ �ϰ� ��� dotnet
function shopifySelectRegProcess_2() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		// itemid ����� ����
		var itemIds = [];
		for(var i=0;i<frmSvArr.cksel.length;i++) {
			if (frmSvArr.cksel[i].checked) {
				itemIds.push(frmSvArr.cksel[i].value);
			}
		}

		if (frmSvArr.cksel.length == undefined) {
			itemIds.push(frmSvArr.cksel.value)
		}

		$('#shopifyResultDiv').html('');
		for (var i = 0; i < itemIds.length; i++) {
			(function(itemId) {
				$.ajax({
					type: 'POST',
					contentType: "application/json",
					url: '<%=dotnetApiURL%>/api/outmall/ShopifyItem/Create',
					xhrFields: {
						withCredentials: true
					},
					dataType:'json',
					data: JSON.stringify({
						"itemid": itemId
					}),
					success: function(data) {
						console.log(data);
						var message = data.message || itemId + " �Ϸ�";
						$('#shopifyResultDiv').append(message + "<br>");
					},
					error: function(error) {
						console.log(error);
						var message = "";
						try {
							var responseJson = JSON.parse(error.responseText);
							message = responseJson.errorMessage;
						} catch(e) {
							console.log("parse error : " + error.responseText);
							message = itemId + "����";
						}

						$('#shopifyResultDiv').append(message + "<br>");
					}
				});
			})(itemIds[i])
		}
    }
}

// ��ǰ ��ȸ for dotnet
function checkshopifyItemConfirm_2() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}
    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
		var itemIds = [];
		for(var i=0;i<frmSvArr.cksel.length;i++) {
			if (frmSvArr.cksel[i].checked) {
				itemIds.push(frmSvArr.cksel[i].value);
			}
		}
		if (frmSvArr.cksel.length == undefined) {
			itemIds.push(frmSvArr.cksel.value)
		}

		$('#shopifyResultDiv').html('');
		for (var i = 0; i < itemIds.length; i++) {
			(function(itemId) {
				$.ajax({
					type: 'POST',
					contentType: "application/json",
					url: '<%=dotnetApiURL%>/api/outmall/ShopifyItem/CheckStatItem',
					xhrFields: {
						withCredentials: true
					},
					dataType:'json',
					data: JSON.stringify({
						"itemId": itemId
					}),
					success: function(data) {
						console.log(data);
						var message = data.message || itemId + " �Ϸ�";
						$('#shopifyResultDiv').append(message + "<br>");
					},
					error: function(error) {
						console.log(error);
						var message = "";
						try {
							var responseJson = JSON.parse(error.responseText);
							message = responseJson.errorMessage;
						} catch(e) {
							console.log("parse error : " + error.responseText);
							message = itemId + "����";
						}

						$('#shopifyResultDiv').append(message + "<br>");
					}
				});
			})(itemIds[i])
		}
    }
}

// ���õ� ��ǰ �ϰ� ����
function shopifyEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
	}
}

// ���õ� ��ǰ �ϰ� ����
function shopifyEditProcess_2(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		// itemid ����� ����
		var itemIds = [];
		for(var i=0;i<frmSvArr.cksel.length;i++) {
			if (frmSvArr.cksel[i].checked) {
				itemIds.push(frmSvArr.cksel[i].value);
			}
		}
		if (frmSvArr.cksel.length == undefined) {
			itemIds.push(frmSvArr.cksel.value)
		}

		$('#shopifyResultDiv').html('');
		for (var i = 0; i < itemIds.length; i++) {
			(function(itemId) {
				$.ajax({
					type: 'POST',
					contentType: "application/json",
					url: '<%=dotnetApiURL%>/api/outmall/ShopifyItem/EditItem',
					xhrFields: {
						withCredentials: true
					},
					dataType:'json',
					data: JSON.stringify({
						"itemId": itemId
					}),
					success: function(data) {
						console.log(data);
						var message = data.message || itemId + " �Ϸ�";
						$('#shopifyResultDiv').append(message + "<br>");
					},
					error: function(error) {
						console.log(error);
						var message = "";
						try {
							var responseJson = JSON.parse(error.responseText);
							message = responseJson.errorMessage;
						} catch(e) {
							console.log("parse error : " + error.responseText);
							message = itemId + "����";
						}

						$('#shopifyResultDiv').append(message + "<br>");
					}
				});
			})(itemIds[i])
		}
	}
}

// ���õ� ��ǰ ���� �ϰ� ����
function shopifyPriceEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ������ ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ��� �ϰ� ����
function shopifyEditQtyProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��� ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "QTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��ȸ �� ��� �ϰ� ����
function shopifySelectEditQtyProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('shopify�� �����Ͻ� ' + chkSel + '�� ��� ���� �Ͻðڽ��ϱ�?\n\n��shopify���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDITQTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// shopify _Collection ����
function pop_CollectionManager() {
	var pCM = window.open("/admin/etc/shopify/pop_CollectionManager.asp","pop_CollectionManager","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}

function pop_NewCollectionManager() {
	var url = "<%=dotnetApiURL%>/OutMall/Shopify#/mapping_collections";

	var pCM = window.open(url,"pop_NewCollectionManager","width=1000,height=675,scrollbars=yes,resizable=yes");
	pCM.focus();
}

function pop_CollectionsManager() {
	var url = "<%=dotnetApiURL%>/OutMall/Shopify#/collections";

	var pCM = window.open(url,"pop_CollectionsManager","width=1000,height=675,scrollbars=yes,resizable=yes");
	pCM.focus();
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=shopify&ml=EN','itemWeightEdit','width=1024,height=768,scrollbars=yes,resizable=yes')
	popwin.focus();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=shopify&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

//function PopAttributes(iitemid, iitemoption, icatekey){
//	var popwin = window.open('/admin/etc/shopify/popAttribute.asp?itemid=' + iitemid +'&itemoption='+iitemoption+'&catekey='+icatekey,'itemWeightEdit','width=1024,height=500,scrollbars=yes,resizable=yes')
//	popwin.focus();
//}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="https://10x10-co-kr.myshopify.com/admin" target="_blank">shopify_Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then

			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		shopify ��ǰ�ڵ� : <textarea rows="2" cols="20" name="shopifyGoodNo" id="itemid"><%=Replace(replace(shopifyGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >shopify ��Ͽ����̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >shopify ��ϿϷ�(����)
		</select>&nbsp;
		<!-- <input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">��ü</label>&nbsp; -->
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">�̵��</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">ǰ��ó�����</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">��ϻ�ǰ �ǸŰ���</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">��Ͽ�������Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
		����ѵ�Ͽ��� : 
		<select name="sitename" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= CHkIIF(sitename="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(sitename="N","selected","") %>>N</option>
		</select>
		&nbsp;
		��뿩�� : 
		<select name="isusing" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= CHkIIF(isusing="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(isusing="N","selected","") %>>N</option>
		</select>
		&nbsp;
		���Ե�� : 
		<select name="itemweight" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= CHkIIF(itemweight="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(itemweight="N","selected","") %>>N</option>
		</select>
		&nbsp;
		�ؿܹ�� : 
		<select name="deliverOverseas" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= CHkIIF(deliverOverseas="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(deliverOverseas="N","selected","") %>>N</option>
		</select>
		&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>shopify ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="shopifyYes10x10No" <%= ChkIIF(shopifyYes10x10No="on","checked","") %> ><font color=red>shopify�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="shopifyNo10x10Yes" <%= ChkIIF(shopifyNo10x10Yes="on","checked","") %> ><font color=red>shopifyǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="shopifyitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">

<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="LEFT">
			    <input class="button" type="button" value="�귣��Collection Refresh" onclick="collectionRefresh('brand');">
				<input class="button" type="button" value="ī�װ�Collection" onclick="collectionRefresh('category');">
			</td>
			<td align="right">
			    <% If (FALSE)  Then %>
				<input class="button" type="button" value="SubCategory" onclick="shopifyCategoryProcess();">&nbsp;&nbsp;
				<% End If %>
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('shopify');">&nbsp;&nbsp;

			    <% If (FALSE)  Then %>
				<input class="button" type="button" value="�ݷ��� ����" onclick="pop_CollectionsManager();">
				<input class="button" type="button" value="ī�װ� ���� ����" onclick="pop_NewCollectionManager();">
				<input class="button" type="button" value="�ݷ��� ����(��)" onclick="pop_CollectionManager();">
				<% End If %>
			</td>
		</tr>
		</table>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
				������ǰ ��� :
				<input class="button" type="button" id="btnRegSel" value="���" onClick="shopifySelectRegProcess();">
				<br /><br />
				������ǰ �˻� :
				<input class="button" type="button" id="btnSelectGoodNo" value="��ȸ" onClick="checkshopifyItemConfirm();">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnDelSel" value="����" onClick="shopifyEditProcess();">
				&nbsp;
				<input class="button" type="button" id="btnDelSel" value="����" onClick="checkshopifyItemDelete();">
			</td>

			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="shopifySellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="collectionType">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= FormatNumber(oshopify.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oshopify.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">shopify�����<br>shopify����������</td>
	<td width="90">���ǸŰ�<br /><font color='BLUE'>�Ǹŵɰ���</font></td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">��ǰ<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">shopify<br>����</td>
	<td width="70">shopify<br>�Ǹ�</td>
	<td width="70">shopify���</td>
	<td width="100">shopify<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="70">��ǰ��</td>
	<td width="70">3����<br>�Ǹŷ�</td>
	<td width="60">����</td>
    <% if (FALSE) then %>
    <td width="60">ī�װ�<br>��Ī����</td>
    <% end if %>
</tr>
<% For i=0 to oshopify.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oshopify.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oshopify.FItemList(i).Fsmallimage %>" width="50"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oshopify.FItemList(i).FItemID%>" target="_blank"><%= oshopify.FItemList(i).FItemID %></a>
		<% If oshopify.FItemList(i).FshopifyStatcd <> 7 Then %>
		<br><%= oshopify.FItemList(i).getshopifyStatName %>
		<% End If %>
		<%= oshopify.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left">
		<%= oshopify.FItemList(i).FMakerid %> <%= oshopify.FItemList(i).getDeliverytypeName %>
		<br /><%= oshopify.FItemList(i).FItemName %>
		<br /><font color="BLUE"><%= oshopify.FItemList(i).FChgitemname %>
	</td>
	<td align="center"><%= oshopify.FItemList(i).FRegdate %><br><%= oshopify.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oshopify.FItemList(i).FshopifyRegdate %><br><%= oshopify.FItemList(i).FshopifyLastUpdate %></td>
	<td align="right">
		<%= FormatNumber(oshopify.FItemList(i).FOrgprice,0) %><br>
		<%
			If oshopify.FItemList(i).FMaySellPrice <> "" Then
				response.write "<font color='BLUE'>"&formatNumber(oshopify.FItemList(i).FMaySellPrice,0)&"</font>"
			End If
		%>
	</td>
	<td align="center">
	<%
		If oshopify.FItemList(i).Fsellcash <> 0 Then
			response.write CLng(10000-oshopify.FItemList(i).Fbuycash / oshopify.FItemList(i).Fsellcash*100*100)/100 & "%" &" <br>"
		End If
	%>
	</td>
	<td align="center"><%= FormatNumber((oshopify.FItemList(i).FitemWeight/1000),3) %>kg</td>
	<td align="center">
	<%= oshopify.FItemList(i).getSellStateTitle  %>
	</td>
	<td align="center">
	<%
		If (oshopify.FItemList(i).FshopifyStatCd > 0) Then
			If Not IsNULL(oshopify.FItemList(i).FshopifyPrice) Then
				If (oshopify.FItemList(i).FOrgprice <> oshopify.FItemList(i).FRegOrgprice) Then
	%>
					<strong><%= CDBL(formatNumber(oshopify.FItemList(i).FshopifyPrice,2)) %></strong><br>
	<%
				Else
					response.write CDBL(formatNumber(oshopify.FItemList(i).FshopifyPrice,2))&"<br>"
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oshopify.FItemList(i).FSellyn="Y" and oshopify.FItemList(i).FshopifySellYn<>"Y") or (oshopify.FItemList(i).FSellyn<>"Y" and oshopify.FItemList(i).FshopifySellYn="Y") Then
			response.write "<strong>" & oshopify.FItemList(i).FshopifySellYn & "</strong>"
		Else
			response.write oshopify.FItemList(i).FshopifySellYn
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oshopify.FItemList(i).FQuantity)) Then
			response.write oshopify.FItemList(i).FQuantity
		End If
	%>
	</td>

	<td align="center">
	<%
		If Not(IsNULL(oshopify.FItemList(i).FshopifyGoodNo)) Then
			Response.Write "<a target='_blank' href='https://shopify.com/en-sg/product/details/"&oshopify.FItemList(i).FshopifyGoodNo&"'>"&oshopify.FItemList(i).FshopifyGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oshopify.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oshopify.FItemList(i).FItemID%>','0');"><%= oshopify.FItemList(i).FoptionCnt %>:<%= oshopify.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"></td>
	<td><input type="button" class="button" value="����" onclick="PopItemContent('<%=oshopify.FItemList(i).FItemid%>')"></td>
	<% if (FALSE) then %>
	<td align="center">
	<%
		If oshopify.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
		If (oshopify.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oshopify.FItemList(i).FlastErrStr) &"'>ERR:"& oshopify.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
	<% End If %>
</tr>
<% Next %>
<tr height="20">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oshopify.HasPreScroll then %>
		<a href="javascript:goPage('<%= oshopify.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oshopify.StartScrollPage to oshopify.FScrollCount + oshopify.StartScrollPage - 1 %>
    		<% if i>oshopify.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oshopify.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<div id="shopifyResultDiv">
</div>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
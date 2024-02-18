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
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	bestOrd = "on"
	sellyn = "Y"
End If

'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
'shopify 상품코드 엔터키로 검색되게
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

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oshopify.getshopifyreqExpireItemList
Else
	oshopify.getshopifyRegedItemList		'그 외 리스트
End If

' dotnet 테스트용.
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
    	if(document.getElementById("QR").checked == true){ //등록상품 판매가능
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
//등록여부 조건 Reset
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
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// TEST SubCategory
function shopifyCategoryProcess() {
    if (confirm('shopify에 카테고리 확인?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "SubCategory";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 삭제
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 상품을 일괄 삭제 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "DELETE";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// 상품 조회
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 상품조회 하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
    }
}

function collectionRefresh(v){
    if (confirm(v + ' collection을 갱신하겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "collectionRefresh";
		document.frmSvArr.collectionType.value = v;
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
	}
}
// 선택된 상품 판매여부 변경
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}
	if(chkYn == "Y"){
		strSell = "판매중";
	}else if(chkYn == "N"){
		strSell = "품절";
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
	}
}

// 선택된 상품 일괄 등록 dotnet
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		// itemid 목록을 만듦
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
						var message = data.message || itemId + " 완료";
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
							message = itemId + "실패";
						}

						$('#shopifyResultDiv').append(message + "<br>");
					}
				});
			})(itemIds[i])
		}
    }
}

// 상품 조회 for dotnet
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}
    if (confirm('shopify에 선택하신 ' + chkSel + '개 상품조회 하시겠습니까?')){
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
						var message = data.message || itemId + " 완료";
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
							message = itemId + "실패";
						}

						$('#shopifyResultDiv').append(message + "<br>");
					}
				});
			})(itemIds[i])
		}
    }
}

// 선택된 상품 일괄 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if (confirm('shopify에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "/admin/etc/shopify/actShopifyReq.asp"
		document.frmSvArr.submit();
	}
}

// 선택된 상품 일괄 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if (confirm('shopify에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		// itemid 목록을 만듦
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
						var message = data.message || itemId + " 완료";
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
							message = itemId + "실패";
						}

						$('#shopifyResultDiv').append(message + "<br>");
					}
				});
			})(itemIds[i])
		}
	}
}

// 선택된 상품 가격 일괄 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 가격을 수정 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 재고 일괄 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 재고 수정 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "QTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 조회 후 재고 일괄 수정
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
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('shopify에 선택하신 ' + chkSel + '개 재고 수정 하시겠습니까?\n\n※shopify와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDITQTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shopify/actshopifyReq.asp"
		document.frmSvArr.submit();
    }
}

// shopify _Collection 관리
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
//옵션 수 팝업
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
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="https://10x10-co-kr.myshopify.com/admin" target="_blank">shopify_Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then

			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		shopify 상품코드 : <textarea rows="2" cols="20" name="shopifyGoodNo" id="itemid"><%=Replace(replace(shopifyGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >shopify 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >shopify 등록완료(전시)
		</select>&nbsp;
		<!-- <input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp; -->
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">미등록</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">품절처리요망</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">등록상품 판매가능</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">등록여부조건Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
		언어팩등록여부 : 
		<select name="sitename" class="select">
			<option value="">전체</option>
			<option value="Y" <%= CHkIIF(sitename="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(sitename="N","selected","") %>>N</option>
		</select>
		&nbsp;
		사용여부 : 
		<select name="isusing" class="select">
			<option value="">전체</option>
			<option value="Y" <%= CHkIIF(isusing="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(isusing="N","selected","") %>>N</option>
		</select>
		&nbsp;
		무게등록 : 
		<select name="itemweight" class="select">
			<option value="">전체</option>
			<option value="Y" <%= CHkIIF(itemweight="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(itemweight="N","selected","") %>>N</option>
		</select>
		&nbsp;
		해외배송 : 
		<select name="deliverOverseas" class="select">
			<option value="">전체</option>
			<option value="Y" <%= CHkIIF(deliverOverseas="Y","selected","") %>>Y</option>
			<option value="N" <%= CHkIIF(deliverOverseas="N","selected","") %>>N</option>
		</select>
		&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>shopify 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="shopifyYes10x10No" <%= ChkIIF(shopifyYes10x10No="on","checked","") %> ><font color=red>shopify판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="shopifyNo10x10Yes" <%= ChkIIF(shopifyNo10x10Yes="on","checked","") %> ><font color=red>shopify품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
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
			    <input class="button" type="button" value="브랜드Collection Refresh" onclick="collectionRefresh('brand');">
				<input class="button" type="button" value="카테고리Collection" onclick="collectionRefresh('category');">
			</td>
			<td align="right">
			    <% If (FALSE)  Then %>
				<input class="button" type="button" value="SubCategory" onclick="shopifyCategoryProcess();">&nbsp;&nbsp;
				<% End If %>
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('shopify');">&nbsp;&nbsp;

			    <% If (FALSE)  Then %>
				<input class="button" type="button" value="콜렉션 관리" onclick="pop_CollectionsManager();">
				<input class="button" type="button" value="카테고리 매핑 관리" onclick="pop_NewCollectionManager();">
				<input class="button" type="button" value="콜렉션 관리(구)" onclick="pop_CollectionManager();">
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
				실제상품 등록 :
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="shopifySelectRegProcess();">
				<br /><br />
				실제상품 검색 :
				<input class="button" type="button" id="btnSelectGoodNo" value="조회" onClick="checkshopifyItemConfirm();">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnDelSel" value="수정" onClick="shopifyEditProcess();">
				&nbsp;
				<input class="button" type="button" id="btnDelSel" value="삭제" onClick="checkshopifyItemDelete();">
			</td>

			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="shopifySellYnProcess(frmReg.chgSellYn.value);">
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
		검색결과 : <b><%= FormatNumber(oshopify.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oshopify.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">shopify등록일<br>shopify최종수정일</td>
	<td width="90">원판매가<br /><font color='BLUE'>판매될가격</font></td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">상품<br>무게</td>
	<td width="70">품절여부</td>
	<td width="70">shopify<br>가격</td>
	<td width="70">shopify<br>판매</td>
	<td width="70">shopify재고</td>
	<td width="100">shopify<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="70">단품수</td>
	<td width="70">3개월<br>판매량</td>
	<td width="60">관리</td>
    <% if (FALSE) then %>
    <td width="60">카테고리<br>매칭여부</td>
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
	<td><input type="button" class="button" value="관리" onclick="PopItemContent('<%=oshopify.FItemList(i).FItemid%>')"></td>
	<% if (FALSE) then %>
	<td align="center">
	<%
		If oshopify.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
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
<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Session.codepage="65001"
Response.ContentType="text/html;charset=UTF-8"
Response.AddHeader "Pragma", "no-cache"
Response.CacheControl = "no-cache"
Response.Expires = -1
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 상품 등록"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<%
Dim vImgURL
vImgURL=""

Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = request.cookies("partner")("userid")
npartner.GetAcademyPartnerList

Dim waititemid
waititemid = requestCheckVar(request("waititemid"),10)
%>
<script>
<!--
function fnVodLinkReg(){
	var vodlink = Base64.encode($("#itemvideo").val());
	fnAPPpopupVod("<%=g_AdminURL%>/apps/academy/itemmaster/popup/popVodAdd.asp?waititemid="+$("#waititemid").val() + "&param=" + encodeURIComponent(vodlink));
}

function fnCategorySet(callbackval){
	var catearr = callbackval.replace(/ /g, "");
	var catearr2 = catearr.replace(/!/g, "','");
	var catearr3=eval("['" + catearr2 + "']");
	$(document).find("#selectcate").empty().append("<span class='setContView'>" + catearr3[0] + "</span>");
	$("#catecode").val(catearr3[1]);
	$("#catedepth").val(catearr3[2]);
	$("#isDefault").val(catearr3[3]);
	$("#catename").val(catearr3[4]);
}

function fnCategoryReg(){
	var catecode = $("#catecode").val();
	var catedepth = $("#catedepth").val();
	var isDefault = $("#isDefault").val();
	var catename = $("#catename").val();
	var paramtxt = catecode + "!" +  catedepth + "!" +  isDefault + "!" +  catename;
	fnAPPpopupCategory("<%=g_AdminURL%>/apps/academy/itemmaster/popup/popCategorySelectWait.asp?waititemid="+$("#waititemid").val() + "&param=" + encodeURIComponent(paramtxt));
}

function fnMakeUnusualReg(){
    fnAPPpopupReqContents('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popSpecialNoteWait.asp?param='+encodeURIComponent(Base64.encode($('#requirecontents').val())));
	//fnAPPpopupReqContents('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popSpecialNoteWait.asp?param='+encodeURIComponent($('#requirecontents').val()));
}

function fnDeliveryInfoReg(){
    fnAPPpopupDeliveryInfo('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popDeliveryInfoWait.asp?param='+encodeURIComponent(Base64.encode($('#ordercomment').val())));
	//fnAPPpopupDeliveryInfo('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popDeliveryInfoWait.asp?param='+encodeURIComponent($('#ordercomment').val()));
}

function fnOptionSet(callbackval){
	callbackval = Base64.decode(callbackval);
	var jSonTXT = JSON.parse(callbackval);
	//alert(jSonTXT.mode);
	$("#optsetend").addClass("setContView");
	if(jSonTXT.mode=="editOption"){
		$("#waititemid").val(jSonTXT.waititemid);
		$("#itemoptioncode2").val(jSonTXT.optioncode);
		$("#itemoptioncode3").val(jSonTXT.optionname);
		$("#itemoptioncode4").val(jSonTXT.optionaddprice);
		$("#itemoptioncode5").val(jSonTXT.optionbuyprice);
		$("#opttype").val(jSonTXT.mode);
		$("#optionTypename1").val(jSonTXT.optiontypename);
		$("#useoptionyn").val("Y");
		$("#tempSaveYn").val("Y");
	}else{
		$("#opttype").val(jSonTXT.mode);
		$("#optionTypename1").val(jSonTXT.optiontypename1);
		$("#optionTypename2").val(jSonTXT.optiontypename2);
		$("#optionTypename3").val(jSonTXT.optiontypename3);
		if(jSonTXT.optiontypename1!=""){
			$("input[name=optname1]").each(function(idx){
				$("#optname1").remove();
				$("#optaddprice1").remove();
				$("#optbuyprice1").remove();
			});
			for(var i=0; i < jSonTXT.optionname1.length; i++){
				$('#fopt').append('<input type="hidden" id="optname1" name="optname1" value="' + jSonTXT.optionname1[i] + '">');
				$('#fopt').append('<input type="hidden" id="optaddprice1" name="optaddprice1" value="' + jSonTXT.optionaddprice1[i] + '">');
				$('#fopt').append('<input type="hidden" id="optbuyprice1" name="optbuyprice1" value="' + jSonTXT.optionbuyprice1[i] + '">');
			}
		}
		if(jSonTXT.optiontypename2!=""){
			$("input[name=optname2]").each(function(idx){
				$("#optname2").remove();
				$("#optaddprice2").remove();
				$("#optbuyprice2").remove();
			});
			for(var i=0; i < jSonTXT.optionname2.length; i++){
				$('#fopt').append('<input type="hidden" id="optname2" name="optname2" value="' + jSonTXT.optionname2[i] + '">');
				$('#fopt').append('<input type="hidden" id="optaddprice2" name="optaddprice2" value="' + jSonTXT.optionaddprice2[i] + '">');
				$('#fopt').append('<input type="hidden" id="optbuyprice2" name="optbuyprice2" value="' + jSonTXT.optionbuyprice2[i] + '">');
			}
		}
		if(jSonTXT.optiontypename3!=""){
			$("input[name=optname3]").each(function(idx){
				$("#optname3").remove();
				$("#optaddprice3").remove();
				$("#optbuyprice3").remove();
			});
			for(var i=0; i < jSonTXT.optionname3.length; i++){
				$('#fopt').append('<input type="hidden" id="optname3" name="optname3" value="' + jSonTXT.optionname3[i] + '">');
				$('#fopt').append('<input type="hidden" id="optaddprice3" name="optaddprice3" value="' + jSonTXT.optionaddprice3[i] + '">');
				$('#fopt').append('<input type="hidden" id="optbuyprice3" name="optbuyprice3" value="' + jSonTXT.optionbuyprice3[i] + '">');
			}
		}
		$("#waititemid").val(jSonTXT.waititemid);
		$("#useoptionyn").val("Y");
		$("#tempSaveYn").val("Y");
	}
}

function fnOptionReg(){
	var param;
	if($("#optlevel").val()=="1"){
		var itemoptioncode3 = $("#itemoptioncode3").val();
		var itemoptioncode4 = $("#itemoptioncode4").val();
		var itemoptioncode5 = $("#itemoptioncode5").val();
		var itemoptioncode6 = $("#itemoptioncode6").val();
		var optionTypename = $("#optionTypename1").val();
		param = optionTypename + "!" + itemoptioncode3 + "!" + itemoptioncode4 + "!" + itemoptioncode5 + "!" + itemoptioncode6;
		//alert(param);
		fnAPPpopupOptionWaitSet('sellcash=' + $('#sellcash').val() + '&buycash=' + $('#buycash').val() + '&dmargin=<%= npartner.FPartnerList(0).Fdiy_margin %>&param='+encodeURIComponent(param)+"&limityn="+$('#limityn').val()+"&waititemid="+$('#waititemid').val(),$('#optlevel').val());
	}else{
		var jSonTXT;
		var arrOptName1, arrOptAddPrice1, arrOptBuyPrice1;
		var arrOptName2, arrOptAddPrice2, arrOptBuyPrice2;
		var arrOptName3, arrOptAddPrice3, arrOptBuyPrice3;
		arrOptName1 = new Array();
		arrOptAddPrice1 = new Array();
		arrOptBuyPrice1 = new Array();
		arrOptName2 = new Array();
		arrOptAddPrice2 = new Array();
		arrOptBuyPrice2 = new Array();
		arrOptName3 = new Array();
		arrOptAddPrice3 = new Array();
		arrOptBuyPrice3 = new Array();
		if($("input[name=optionTypename1]").val()!=""){
			$("input[name=optname1]").each(function(idx){
				 arrOptName1[idx] = $("input[name=optname1]:eq(" + idx + ")").val();
				 arrOptAddPrice1[idx] = $("input[name=optaddprice1]:eq(" + idx + ")").val();
				 arrOptBuyPrice1[idx] = $("input[name=optbuyprice1]:eq(" + idx + ")").val();
			});
		}
		if($("input[name=optionTypename2]").val()!=""){
			$("input[name=optname2]").each(function(idx){
				 arrOptName2[idx] = $("input[name=optname2]:eq(" + idx + ")").val();
				 arrOptAddPrice2[idx] = $("input[name=optaddprice2]:eq(" + idx + ")").val();
				 arrOptBuyPrice2[idx] = $("input[name=optbuyprice2]:eq(" + idx + ")").val();
			});
		}
		if($("input[name=optionTypename3]").val()!=""){
			$("input[name=optname3]").each(function(idx){
				 arrOptName3[idx] = $("input[name=optname3]:eq(" + idx + ")").val();
				 arrOptAddPrice3[idx] = $("input[name=optaddprice3]:eq(" + idx + ")").val();
				 arrOptBuyPrice3[idx] = $("input[name=optbuyprice3]:eq(" + idx + ")").val();
			});
		}
		jSonTXT = JSON.stringify({"optiontypename1":$("input[name=optionTypename1]").val(), "optionname1":arrOptName1, "optionaddprice1":arrOptAddPrice1, "optionbuyprice1":arrOptBuyPrice1, "optiontypename2":$("input[name=optionTypename2]").val(), "optionname2":arrOptName2, "optionaddprice2":arrOptAddPrice2, "optionbuyprice2":arrOptBuyPrice2, "optiontypename3":$("input[name=optionTypename3]").val(), "optionname3":arrOptName3, "optionaddprice3":arrOptAddPrice3, "optionbuyprice3":arrOptBuyPrice3});
		jSonTXT = Base64.encode(jSonTXT);
		fnAPPpopupOptionWaitSet('sellcash=' + $('#sellcash').val() + '&buycash=' + $('#buycash').val() + '&dmargin=<%= npartner.FPartnerList(0).Fdiy_margin %>&waititemid='+$('#waititemid').val() + '&param=' + jSonTXT,$('#optlevel').val());
	}
}

function fnOptionMultiItemCountReg(){
	//alert($('#waititemid').val());
	if($('#waititemid').val()==""){
		alert("옵션 수량 변경은 옵션 항목/가격 설정 후 가능합니다.");
	}else{
		if($("#optlevel").val()=="2"){
			popOptionMultiCountWaitSet('waititemid=' + $('#waititemid').val() + '&limityn=' + $('#limityn').val());
		}else{
			popOptionCountWaitSet('waititemid=' + $('#waititemid').val() + '&limityn=' + $('#limityn').val());
		}
	}
}

function fnOptionNoEditSet(callbackval){
	$("#optionedit").val(callbackval);
}

function fnMultipleStateOptionEditEnd(TotalOptLimitNo){
	if(TotalOptLimitNo > 0){
		chgodr('LimitCnt',1,'limityn','Y');
		$("#blimity").addClass("selected");
		$("#blimitn").removeClass("selected")
		$("#limitno").val(TotalOptLimitNo);
		$("#optlimitset").addClass("setContView");
	}
}

function fnKeyWordReg(){
	//var keyword = encSpecialCharNativeFun($("#keywords").val());
	var keyword = Base64.encode($("#keywords").val());
	fnAPPpopupKeyWord('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popKeywordWait.asp?param='+encodeURIComponent(keyword));
}

function fnSafeInfoSet(callbackval){
	var catearr = callbackval.replace(/ /g, "");
	var catearr2 = catearr.replace(/!/g, "','");
	var catearr3=eval("['" + catearr2 + "']");
	$(document).find("#safeinfotxt").empty().append("<span class='setContView'>" + catearr3[3] + "</span>");
	var safetyNum = Base64.decode(catearr3[2]);
	$("#safetyYn").val(catearr3[0]);
	$("#safetyDiv").val(catearr3[1]);
	$("#safetyNum").val(safetyNum);
}
function fnSafeInfoReg(){
	var safetyYn = $("#safetyYn").val();
	var safetyDiv = $("#safetyDiv").val();
	var safetyNum = Base64.encode($("#safetyNum").val());
	var param = safetyYn + "," + safetyDiv + "," + safetyNum
	fnAPPpopupSafeInfo('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popArtSafeWait.asp?param='+encodeURIComponent(param));
}

function fnItemInfoDivReg(){
	var infoDiv = $("#infoDiv").val();
	var infoCont = $("#infoCont").val();
	var infoChk = $("#infoChk").val();
	var infoCd = $("#infoCd").val();
	var param = infoDiv + "!" + infoCont + "!" + infoChk + "!" + infoCd;
	fnAPPpopupItemInfoDiv('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popArtInfoWait.asp?param='+encodeURIComponent(param)+'&waititemid=' + $('#waititemid').val());
}

function fnWaitItemInfoDivSet(callbackval){
	//alert(callbackval);
	callbackval = callbackval.replace(/ /g, "");
	callbackval = callbackval.replace(/!/g, "','");
	var arriteminfo=eval("['" + callbackval + "']");
	var infoDiv = arriteminfo[0];
	var setContView = arriteminfo[1];
	var waititemid = arriteminfo[2];
	$(document).find("#iteminfotxt").empty().append("<span class='setContView'>" + setContView + "</span>");
	$("#tempSaveYn").val("Y");
	$("#infoDiv").val(infoDiv);
	$("#waititemid").val(waititemid);
	CheckFormConfirmbtnShow();
}

function fnAppCallWinConfirm(){
//상품등록 콜
	var itemreg = document.itemreg;
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}
	if ($("input[name='catecode']").val() == 0){
		alert("[기본] 전시 카테고리를 선택해 주세요.");
		return;
	}
	if (itemreg.itemname.value == ""){
		alert("상품명을 입력해 주세요.");
		$("input[name='itemname']").focus();
		return;
	}
	//상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
	if (itemreg.itemdiv.value == ""){
		alert("상품 구분을 선택해 주세요.");
		itemreg.itemdiv.focus();
		return;
	}
	if (itemreg.itemdiv.value == "16" || itemreg.itemdiv.value == "06") {
		if(itemreg.cstodr.value == ""){
			alert("발송 구분을 선택해 주세요(즉시 발송/제작 후 발송).");
            return;
		}
	    if (itemreg.cstodr.value == "2" && itemreg.requireMakeDay.value>2000000000) {
	        alert("제작 기간은 최대 2,000,000,000 이하로 입력해주세요.");
			itemreg.requireMakeDay.focus();
            return;
        }
		if (itemreg.cstodr.value == "2" && itemreg.requireMakeDay.value == "") {
	        alert("제작 기간을 입력해 주세요.");
            return;
        }
		if (itemreg.requireimgchk.value == "Y" && itemreg.requireMakeEmail.value == "") {
			alert("이미지 수신 메일을 입력해 주세요.");
			// itemreg.useoptionyn.focus();
			return;
		}
	}
	if (getByteLength(itemreg.makername.value)>64){
	    alert("제작자명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.makername.focus();
		return;
	}
	if (itemreg.makername.value == ""){
		alert("제작자를 입력해 주세요.");
		itemreg.makername.focus();
		return;
	}
	if (getByteLength(itemreg.sourcearea.value)>128){
	    alert("원산지는 최대 128byte 이하로 입력해주세요.(한글64자 또는 영문128자)");
		itemreg.sourcearea.focus();
		return;
	}
	if (itemreg.sourcearea.value == ""){
		alert("원산지를 입력해 주세요.");
		itemreg.sourcearea.focus();
		return;
	}
	if (getByteLength(itemreg.itemsource.value)>128){
	    alert("재질은 최대 128byte 이하로 입력해주세요.(한글64자 또는 영문128자)");
		itemreg.itemsource.focus();
		return;
	}
	if (itemreg.itemsource.value == ""){
		alert("재질을 입력해 주세요.");
		itemreg.itemsource.focus();
		return;
	}
	if (getByteLength(itemreg.itemsize.value)>128){
	    alert("상품 크기는 최대 128byte 이하로 입력해주세요.(한글64자 또는 영문128자)");
		itemreg.itemsize.focus();
		return;
	}
	if (itemreg.itemsize.value == ""){
		alert("상품 크기를 입력해주세요.");
		itemreg.itemsize.focus();
		return;
	}
	if (itemreg.itemWeight.value>2000000000){
	    alert("상품 무게는 최대 2,000,000,000 이하로 입력해주세요.");
		itemreg.itemWeight.focus();
		return;
	}
	if (itemreg.itemWeight.value == ""){
		alert("상품 무게를 입력해주세요.");
		itemreg.itemWeight.focus();
		return;
	}
	if(!is_number($("input[name='itemWeight']").val())){
		alert("상품 무게는 숫자로 입력해 주세요.");
		return;
	}
	if (getByteLength(itemreg.keywords.value)>512){
	    alert("상품 크기는 최대 512byte 이하로 입력해주세요.(한글256자 또는 영문512자)");
		itemreg.keywords.focus();
		return;
	}
	if (itemreg.keywords.value == ""){
		alert("검색 키워드를 입력해주세요.");
		itemreg.keywords.focus();
		return;
	}
	if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}
	if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("마일리지는 판매가보다 클 수 없습니다.");
        return;
    }
    if(itemreg.limityn.value == "Y" && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }
    if (itemreg.useoptionyn.value == "Y") {
	    if (itemreg.optlevel.value == "1") {
	    //단일옵션
    	    if (itemreg.itemoptioncode2.value =="") {
                alert("추가된 옵션이 없습니다.");
                // itemreg.useoptionyn.focus();
                return;
            }
        }else if (itemreg.optlevel.value == "2") {
        //이중옵션
            if (itemreg.optionTypename1.value =="") {
                alert("추가된 옵션이 없습니다.");
                // itemreg.useoptionyn.focus();
                return;
            }
        }
	}
    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype.value=="9"){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            return;
        }
    }
    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype.value=="7")){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        return;
    }
	//배송비 설정
    if (!itemreg.deliverytype.value){
        alert('배송 설정을 선택해주세요.');
        return;
    }
	//상품 품목정보
    if (!itemreg.infoDiv.value){
        alert('상품정보제공고시를 입력해 주세요.');
        return;
    }
	//안전인증정보
    if (itemreg.safetyYn.value=="Y"){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해 주세요.');
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해 주세요.');
	        return;
	    }
    }
	if (itemreg.imgbasic.value == ""){
		alert("기본 이미지 한개는 등록 하셔야 합니다.");
		return;
	}
    // 정상가격
	if (confirm("소비자가(" + itemreg.sellcash.value + ")/공급가(" + itemreg.buycash.value + ")가 정확히 입력되었습니까?") == false) {
		itemreg.sellcash.focus();
		return;
    }
    if(confirm("상품을 올리시겠습니까? \n담당MD 승인후 반영 됩니다.") == true){
		$("#currstate").val("1");
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
		if(itemreg.tempSaveYn.value=="Y"){
			itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_Noimg_Edit_Process_App.asp";
		}else{
			itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_Noimg_Process_App.asp";
		}
		itemreg.itemregYn.value="Y";
		itemreg.target = "FrameCKP";
        itemreg.submit();
    }
}
function is_number(x){
    var reg = /^\d+$/;
    return reg.test(x);
}

function fnTempSave(){
	//alert($("#useoptionyn").val() + "/" + $("#optionedit").val() );
	if(itemreg.itemname.value==""){
		alert("상품명을 입력해주세요.");
		return;
	}else if (itemreg.requireMakeDay.value != "" && itemreg.requireMakeDay.value>2000000000) {
		alert("제작 기간은 최대 2,000,000,000 이하로 입력해주세요.");
		return;
	}else if (itemreg.makername.value != "" && getByteLength(itemreg.makername.value)>64){
		alert("제작자명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		return;
	}else if (itemreg.sourcearea.value != "" && getByteLength(itemreg.sourcearea.value)>128){
		alert("원산지는 최대 128byte 이하로 입력해주세요.(한글64자 또는 영문128자)");
		return;
	}else if (itemreg.itemsource.value != "" && getByteLength(itemreg.itemsource.value)>128){
		alert("재질은 최대 128byte 이하로 입력해주세요.(한글64자 또는 영문128자)");
		return;
	}else if (itemreg.itemsize.value != "" && getByteLength(itemreg.itemsize.value)>128){
		alert("상품 크기는 최대 128byte 이하로 입력해주세요.(한글64자 또는 영문128자)");
		return;
	}else if (itemreg.itemWeight.value != "" && itemreg.itemWeight.value>2000000000){
		alert("상품 무게는 최대 2,000,000,000 이하로 입력해주세요.");
		return;
	}else if (itemreg.keywords.value != "" && getByteLength(itemreg.keywords.value)>512){
		alert("상품 크기는 최대 512byte 이하로 입력해주세요.(한글256자 또는 영문512자)");
		return;
	}else if(itemreg.itemdiv.value==""){
		alert("상품 구분(일반 상품/주문제작 상품)을 선택해주세요.");
		return;
	}else if(!is_number($("input[name='itemWeight']").val()) && $("input[name='itemWeight']").val()!=""){
		alert("상품 무게는 숫자로 입력해 주세요.");
		return;
	}else{
		itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
		if(itemreg.tempSaveYn.value=="Y"){
			itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_Temp_Noimg_Edit_Process_App.asp";
		}else{
			itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_Temp_Noimg_Process_App.asp";
		}
		itemreg.target = "FrameCKP";
		itemreg.submit();
	}
}

function fntempSaveEnd(waititemid,OptionSaveYN){
	$("#tempSaveYn").val("Y");
	$("#waititemid").val(waititemid);
	$('#alert1').fadeIn(800).css("display","");
	fnAPPParentsWinReLoad();
	setTimeout(function(){
			$("#alert1").fadeOut(1000);
		}, 5000);
	$('#alert1').fadeIn(800).css("display","none");
	if(OptionSaveYN=="Y"){
		$("#setoptcnt").css("display","");
	}
}

function fnSaveEnd(waititemid){
	$("#savetime").append("<%=FormatDate(now(),"0000.00.00-00:00")%>");
	$('#alert3').fadeIn(800).css("display","");
	fnAPPParentsWinReLoad();
	setTimeout(function(){
			$("#alert3").fadeOut(1000);
		}, 5000);
	$('#alert3').fadeIn(800).css("display","none");
	setTimeout(function(){
			fnAPPclosePopup();
		}, 300);
}

function fnPreviewItem(){
	if($('#waititemid').val()==""){
		alert("기본항목을 입력하고 임시 저장 후 확인 할 수 있습니다.");
	}else{
		fnAPPpopupItemRegPreview('<%=g_AdminURL%>/apps/academy/preview/shop_prd_wait_app.asp?itemid='+$('#waititemid').val());
	}
}

//등록 대기 상품 수정페이지 강제 호출
function fnMoveWaitItemEdit(url){
	self.location.replace(url);
	return false;
}
//-->
</script>
<script type="text/javascript" src="/apps/academy/lib/waititemreg.js"></script>
<script type="text/javascript" src="/apps/academy/lib/confirm.js"></script>
<style type="text/css">
.selectBtn2, .list li.selectBtn2 {display:table; width:100%; padding:1.5rem 1rem;}
.selectBtn2 div {display:table-cell; vertical-align:middle;}
.selectBtn2 .selected {background-color:#a6d216; color:#fff;}
.selectBtn2 .grid2, .selectBtn2 .grid3, .selectBtn2 .grid4 {padding-left:0.25rem; padding-right:0.25rem;}
.selectBtn2 .grid2:first-child, .selectBtn2 .grid3:first-child, .selectBtn2 .grid4:first-child {padding-left:0;}
.selectBtn2 .grid2:last-child, .selectBtn2 .grid3:last-child, .selectBtn2 .grid4:last-child {padding-right:0;}

.selectBtn1, .list li.selectBtn1 {display:table; width:100%; padding:1.5rem 1rem;}
.selectBtn1 div {display:table-cell; vertical-align:middle;}
.selectBtn1 .selected {background-color:#a6d216; color:#fff;}
.selectBtn1 .grid2, .selectBtn1 .grid3, .selectBtn1 .grid4 {padding-left:0.25rem; padding-right:0.25rem;}
.selectBtn1 .grid2:first-child, .selectBtn1 .grid3:first-child, .selectBtn1 .grid4:first-child {padding-left:0;}
.selectBtn1 .grid2:last-child, .selectBtn1 .grid3:last-child, .selectBtn1 .grid4:last-child {padding-right:0;}

.selectBtn3, .list li.selectBtn3 {display:table; width:100%; padding:1.5rem 1rem;}
.selectBtn3 div {display:table-cell; vertical-align:middle;}
.selectBtn3 .selected {background-color:#a6d216; color:#fff;}
.selectBtn3 .grid2, .selectBtn3 .grid3, .selectBtn3 .grid4 {padding-left:0.25rem; padding-right:0.25rem;}
.selectBtn3 .grid2:first-child, .selectBtn3 .grid3:first-child, .selectBtn3 .grid4:first-child {padding-left:0;}
.selectBtn3 .grid2:last-child, .selectBtn3 .grid3:last-child, .selectBtn3 .grid4:last-child {padding-right:0;}
</style>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<form name="itemreg" method="post" id="fopt">
		<input type="hidden" name="waititemid" id="waititemid" value="<%=waititemid%>">
		<input type="hidden" name="itemregYn" id="itemregYn" value="N">
		<input type="hidden" name="imgbasic" id="imgbasic">
		<input type="hidden" name="imgadd1" id="imgadd1">
		<input type="hidden" name="imgadd2" id="imgadd2">
		<input type="hidden" name="itemvideo" id="itemvideo">
		<input type="hidden" name="optionedit" id="optionedit">
		<input type="hidden" name="opttype" id="opttype">
		<input type="hidden" name="optionTypename1" id="optionTypename1">
		<input type="hidden" name="optionTypename2" id="optionTypename2">
		<input type="hidden" name="optionTypename3" id="optionTypename3">
		<input type="hidden" name="designerid" value="<%= request.cookies("partner")("userid") %>">
		<input type="hidden" name="defultmargine" value="<%= npartner.FPartnerList(0).Fdiy_margin %>">
		<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FPartnerList(0).Fmaeipdiv %>">
		<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FPartnerList(0).FdefaultFreeBeasongLimit %>">
		<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FPartnerList(0).FdefaultDeliverPay %>">
		<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FPartnerList(0).FdefaultDeliveryType %>">
		<input type="hidden" name="cd1" value="999">
		<input type="hidden" name="cd2" value="999">
		<input type="hidden" name="cd3" value="999">
		<input type="hidden" name="catecode" id="catecode">
		<input type="hidden" name="catedepth" id="catedepth">
		<input type="hidden" name="isDefault" id="isDefault">
		<input type="hidden" name="catename" id="catename">
		<input type="hidden" name="itemdiv" id="itemdiv">
		<input type="hidden" name="cstodr" id="cstodr">
		<input type="hidden" name="reqMsg" id="reqMsg">
		<input type="hidden" name="requireimgchk" id="requireimgchk" value="N">
		<input type="hidden" name="vatYn" id="vatYn" value="Y">
		<input type="hidden" name="limityn" id="limityn" value="N">
		<input type="hidden" name="useoptionyn" id="useoptionyn" value="N">
		<input type="hidden" name="optlevel" id="optlevel" value="0">
		<input type="hidden" name="optwintitle" id="optwintitle">
		<input type="hidden" name="keywords" id="keywords">
		<input type="hidden" name="safetyYn" id="safetyYn" value="N">
		<input type="hidden" name="safetyDiv" id="safetyDiv" value="0">
		<input type="hidden" name="safetyNum" id="safetyNum">
		<input type="hidden" name="infoCd" id="infoCd">
		<input type="hidden" name="infoChk" id="infoChk">
		<input type="hidden" name="infoCont" id="infoCont">
		<input type="hidden" name="infoDiv" id="infoDiv">
		<input type="hidden" name="tempSaveYn" id="tempSaveYn" value="N">
		<input type="hidden" name="itemoptioncode2" id="itemoptioncode2">
		<input type="hidden" name="itemoptioncode3" id="itemoptioncode3">
		<input type="hidden" name="itemoptioncode4" id="itemoptioncode4">
		<input type="hidden" name="itemoptioncode5" id="itemoptioncode5">
		<input type="hidden" name="itemoptioncode6" id="itemoptioncode6">
		<input type="hidden" name="itemoptioncode7" id="itemoptioncode7">
		<input type="hidden" name="deliverytype" id="deliverytype" value="9">
		<input type="hidden" name="currstate" id="currstate" value="8">
		<input type="hidden" name="delmode" id="delmode">
		<input type="hidden" name="delfilename" id="delfilename">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">작품 등록</h1>
			<div class="artRegist">
				<div class="registUnit"><!-- for dev msg : 비활성화 시 class : disabled 붙여주세요 -->
					<div class="getBtn"><button type="button" class="btnM1 btnGrn" onClick="fnAPPpopupItemListCall('<%=g_AdminURL%>/apps/academy/itemmaster/artWaitDummyList.asp');">작품 정보 불러오기</button></div>
					<div class="basicImgRegist">
						<div class="swiper-container">
							<div class="swiper-wrapper">
								<div class="swiper-slide" id="imgspan1"><button type="button" onclick="fnAPPuploadImage('imgbasic','basic');">이미지 등록1</button></div>
								<div class="swiper-slide" id="imgspan2"><button type="button" onclick="fnAPPuploadImage('imgadd1','add1');">이미지 등록2</button></div>
								<div class="swiper-slide" id="imgspan3"><button type="button" onclick="fnAPPuploadImage('imgadd2','add2');">이미지 등록3</button></div>
								<div class="swiper-slide" id="imgspan4"><button type="button" onclick="fnVodLinkReg();">동영상 등록</button></div>
							</div>
						</div>
					</div>
					<ul class="list">
						<li class="critical" onclick="fnCategoryReg();">
							<dfn><b>카테고리 설정</b></dfn>
							<div class="listButton btnCtgySet" id="selectcate"></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="critical">
							<dfn><b>상품명</b></dfn>
							<div><input type="text" name="itemname" maxlength="64" placeholder="22자 이하로 입력해주세요" id="[on,off,off,off][상품명]" /></div>
						</li>
						<li class="selectBtn">
							<div class="grid2"><button type="button" value="01" class="btnM1 btnGry" onclick="chgodr('CustomOrder',1,'itemdiv','01');chgodr('DetailShowHide',2,'','');">일반 상품</button></div>
							<div class="grid2"><button type="button" value="06" class="btnM1 btnGry" onclick="chgodr('CustomOrder',2,'itemdiv','16');chgodr('DetailShowHide',2,'','');">주문제작 상품</button></div>
						</li>
					</ul>
				</div>
				<div id="DetailShowHide" style="display:none">
				<!-- for dev msg : 주문제작 상품 선택시 노출됩니다. -->
				<div class="registUnit orderArt" id="CustomOrder" style="display:none">
					<h2 class="critical"><b>주문제작 설정</b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid2"><button type="button" onclick="chgodr('MakeDay',1,'cstodr',1);chgodr('MakeDay2',1,'','');chgodr('MakeDay3',1,'','');" class="btnM1 btnGry">즉시 발송</button></div>
							<div class="grid2"><button type="button" onclick="chgodr('MakeDay',2,'cstodr',2);chgodr('MakeDay2',2,'','');chgodr('MakeDay3',2,'','');" class="btnM1 btnGry">제작 후 발송</button></div>
						</li>
						<li class="critical" id="MakeDay" style="display:none">
							<dfn><b>제작 기간</b></dfn>
							<div><input type="number" name="requireMakeDay" maxlength="2" value="" placeholder="3" /></div>
							<div style="width:1.6rem">일</div>
						</li>
						<li class="" onclick="fnMakeUnusualReg();" id="MakeDay2" style="display:none">
							<dfn><b>특이사항</b><input type="hidden" id="requirecontents" name="requirecontents"></dfn>
							<div class="listButton btnCtgySet" id="requirecontentstxt"></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="selectBtn2" id="MakeDay3" style="display:none">
							<div class="grid2"><button type="button" name="prdmsg" id="prdmsg" class="btnM1 btnGry ckBtn" onclick="MultiSelectButton('prdmsg','reqMsg','06');checkItemDiv();">제작 메시지 필요</button></div>
							<div class="grid2"><button type="button" name="prdimg" id="prdimg" class="btnM1 btnGry ckBtn" onclick="MultiSelectButton('prdimg','requireimgchk','Y');chgodr3();">제작 이미지 필요</button></div>
						</li>
						<li class="critical" onclick="#" id="MakeDay4" style="display:none"><!-- for dev msg : 제작 이미지 필요 선택시 노출됩니다. -->
							<dfn><b>이미지 수신 메일</b></dfn>
							<div><input type="email" name="requireMakeEmail" placeholder="id1234@example.com" /></div>
						</li>
					</ul>
				</div>
				<!--// for dev msg : 주문제작 상품 선택시 노출됩니다. -->

				<div class="registUnit basicInfo">
					<h2>기본 정보</h2>
					<ul class="list">
						<li class="critical">
							<dfn><b>제작자</b></dfn>
							<div><input type="text" name="makername" maxlength="32" placeholder="작가명/법인을 입력해주세요" id="[on,off,off,off][제작자]" /></div>
						</li>
						<li class="critical">
							<dfn><b>원산지</b></dfn>
							<div><input type="text" name="sourcearea" maxlength="64" placeholder="국가명을 입력해주세요" id="[on,off,off,off][원산지]" /></div>
						</li>
						<li class="critical">
							<dfn><b>재질</b></dfn>
							<div><input type="text" name="itemsource" maxlength="64" placeholder="예) 플라스틱, 합금, 은" id="[on,off,off,off][재질]" /></div>
						</li>
						<li class="critical">
							<dfn><b>크기</b></dfn>
							<div><input type="text" name="itemsize" id="[on,off,off,off][크기]" maxlength="64" placeholder="예) 7.5 * 7.5" /></div>
							<div style="width:2.4rem">cm</div>
						</li>
						<li class="critical">
							<dfn><b>무게</b></dfn>
							<div><input type="number" name="itemWeight" id="[on,off,off,off][무게]" maxlength="12" placeholder="예) 785" pattern="[0-9]*" inputmode="numeric" min="0" /></div>
							<div style="width:1.4rem">g</div>
						</li>
						<li class="critical" onclick="fnKeyWordReg()">
							<dfn><b>검색 키워드</b></dfn>
							<div class="listButton btnCtgySet" id="keywordtxt"></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="registUnit salePrice">
					<h2 class="critical"><b>판매 가격 <span>(부가세 포함)</span></b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid2"><button type="button" name="bvatYn" value="Y" class="btnM1 btnGry" onclick="TnGoClear(this.form);chgodr('',1,'vatYn','Y');">과세</button></div>
							<div class="grid2"><button type="button" name="bvatYn" value="N" class="btnM1 btnGry" onclick="TnGoClear(this.form);chgodr('',2,'vatYn','N');">면세</button></div>
						</li>
						<li>
							<dfn><b>공급 마진</b></dfn>
							<div><input type="number" name="margin" id="margin" maxlength="32" value="<% =npartner.FPartnerList(0).Fdiy_margin %>" readonly placeholder="100" /></div>
							<div style="width:1.8rem">%</div>
						</li>
						<li class="critical">
							<dfn><b>판매가</b><input type="hidden" id="sellvat" name="sellvat"></dfn>
							<div><input type="number" name="sellcash" id="sellcash" onKeyUp="CalcuAuto(itemreg);" maxlength="7" placeholder="판매가(소비자가)를 입력해주세요" /></div>
						</li>
						<li>
							<dfn><b>공급가</b><input type="hidden" name="buyvat"></dfn>
							<div><input type="number" name="buycash" id="buycash" maxlength="16" placeholder="0" readonly /></div>
						</li>
						<input type="hidden" name="mwdiv" value="U"> <!-- 매입위탁구분 :업체배송 -->
						<input type="hidden" name="sellyn" value="N">
						<input type="hidden" name="isusing" value="Y">
						<input type="hidden" name="mileage" value="0">
					</ul>
				</div>
				<div class="registUnit quantity">
					<h2 class="critical"><b>수량 설정</b></h2>
					<ul class="list">
						<li class="selectBtn3">
							<div class="grid2"><button type="button" class="btnM1 btnGry" id="limitbtn1" name="blimityn" value="Y" onclick="fnLimitCheckOption('Y');">한정 수량</button></div>
							<div class="grid2"><button type="button" class="btnM1 btnGry" id="limitbtn2" name="blimityn" value="N" onclick="fnLimitCheckOption('N');">무제한</button></div>
						</li>
						<li id="LimitCnt" style="display:none"><!--for dev msg : 한정수량 선택시 노출됩니다. -->
							<dfn><b>수량</b></dfn>
							<div><input type="number" name="limitno" id="limitno" placeholder="수량을 입력해주세요" /></div>
							<div style="width:1.6rem">개</div>
						</li>
					</ul>
				</div>
				<div class="registUnit option">
					<h2 class="critical"><b>옵션 설정</b></h2>
					<ul class="list">
						<li class="selectBtn1">
							<div class="grid3"><button type="button" name="boptlevel" id="optbtn0" value="0" class="btnM1 btnGry" onClick="fnOptionCheckEdit(0);">사용안함</button></div>
							<div class="grid3"><button type="button" name="boptlevel" id="optbtn1" value="1" class="btnM1 btnGry" onClick="fnOptionCheckEdit(1);">단일 옵션</button></div>
							<div class="grid3"><button type="button" name="boptlevel" id="optbtn2" value="2" class="btnM1 btnGry" onClick="fnOptionCheckEdit(2);">이중 옵션</button></div>
						</li>
						<li class="critical" id="setopt" onclick="fnOptionReg();" style="display:none"><!--for dev msg : 단일 옵션 or 이중 옵션 선택시 노출됩니다. -->
							<dfn><b>항목/가격</b></dfn>
							<div class="listButton btnCtgySet"><span id="optsetend" class="">설정됨</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="critical" id="setoptcnt" onclick="fnOptionMultiItemCountReg();" style="display:none">
							<dfn><b>수량</b></dfn>
							<div class="listButton btnCtgySet"><span id="optlimitset" class="">설정됨</span></div>
						</li>
					</ul>
				</div>
				<div class="registUnit delivery">
					<h2 class="critical"><b>배송 설정 <span>(부가세 포함)</span></b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid4"><button type="button" name="bdeliverytype" value="2" class="btnM1 btnGry" onClick="chgodr('',1,'deliverytype',2);">무료</button></div>
							<div class="grid4"><button type="button" name="bdeliverytype" value="9" class="btnM1 btnGry selected" onClick="chgodr('',1,'deliverytype',9);">조건부</button></div>
							<div class="grid4"><button type="button" name="bdeliverytype" value="7" class="btnM1 btnGry" onClick="chgodr('',1,'deliverytype',7);">착불</button></div>
						</li>
						<li class="" onclick="fnDeliveryInfoReg();">
							<dfn><b>배송비 안내</b><input type="hidden" id="ordercomment"  name="ordercomment"></dfn>
							<div class="listButton btnCtgySet" id="deliveryInfotxt"></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li style="display:none">
							<dfn><b>교환 / 환불 정책</b></dfn>
							<div><p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize" name="refundpolicy">- 반품/환불은 상품수령일로부터 7일 이내만 가능합니다. 
- 출고 이후 환불요청 시 상품 회수 후 처리됩니다. 
- 변심 반품의 경우 왕복배송비를 차감한 금액이 환불되며, 제품 및 포장 상태가 재판매 가능하여야 합니다. 
- 상품 불량인 경우는 배송비를 포함한 전액이 환불됩니다.
- 완제품으로 수입된 상품의 경우 A/S가 불가합니다. 
- 교환/환불/배송비안내/AS에 대한 개별기준이 상품페이지에 있는 경우 작가님의 개별기준이 우선 적용 됩니다.</textarea></p></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="registUnit law">
					<h2 class="critical"><b>관련법 필수 입력 항목</b></h2>
					<ul class="list">
						<li class="critical" onclick="fnItemInfoDivReg();">
							<dfn><b>상품정보제공고시</b></dfn>
							<div class="listButton btnCtgySet" id="iteminfotxt"></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
						<li class="critical" onclick="fnSafeInfoReg();">
							<dfn><b>안전인증대상</b></dfn>
							<div class="listButton btnCtgySet" id="safeinfotxt"></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
						</li>
					</ul>
				</div>
				<div class="detail" id="DetailInfo">
					<div class="registUnit">
						<h2 class="critical"><b>상세 정보</b><input type="hidden" name="dicheckcnt" id="dicheckcnt" value="1"></h2>
						<ul class="list">
							<li id="DetailList1">
								<p id="imgArea1"><button type="button" class="btnImgRegist" onclick="fnAPPuploadAddImage('addimgname1','1');">이미지 등록</button><input type="hidden" name="addimgname" id="addimgname"></p>
								<p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize" name="addimgtext"></textarea></p>
							</li>
						</ul>
					</div>
					<div class="addBtn">
						<button type="button" class="btnB1 btnDkGry" id="addbtn" onClick="AddDetailInfo()"><span class="itemAdd">추가</span></button>
						<p class="tPad2r">최대 15개까지 등록 가능합니다.</p>
					</div>
				</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<!-- 알림 메세지 -->
		<!-- 알림 메세지 -->
		<div class="attentionBar" style="display:none" id="alert1">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_dot.png" alt="필수표시" style="width:0.4rem; height:0.4rem; margin:0.3rem 0.3rem 0 0" /> 임시저장 되었습니다. 저장된 데이터는 ‘등록대기’ 탭에서 확인 가능합니다.</p>
		</div>
		<div class="attentionBar" style="display:none" id="alert2">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_dot.png" alt="필수표시" style="width:0.4rem; height:0.4rem; margin:0.3rem 0.3rem 0 0" /> 표기는 필수 선택/입력 항목입니다. 꼭 입력해주세요.</p>
		</div>
		<div class="attentionBar" style="display:none" id="alert3">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_save.png" alt="저장표시" style="width:1.2rem; height:1.2rem; margin:0.3rem 0.3rem 0 0" /> <span id="savetime"></span>에 저장되었습니다.</p>
		</div>
		<!-- 하단 플로팅 버튼 -->
		<div class="floatingBar">
			<p><button type="button" class="btnV16a btnWishV16a" onClick="fnTempSave();">임시저장</button></p>
			<!-- <p><button type="button" class="btnV16a btnWishV16a" onClick="fnAppCallWinRegister();">임시저장</button></p> -->
			<p><button type="button" class="btnV16a btnRed2V16a" onClick="fnPreviewItem()">미리보기</button></p>
			<!-- <p><button type="button" class="btnV16a btnRed2V16a" onClick="document.location.reload();">미리보기</button></p> -->
		</div>
		</form>
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
</body>
</html>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
	$('#alert2').fadeIn(800).css("display","");
	setTimeout(function(){
			$("#alert2").fadeOut(1000);
		}, 5000);
	$('#alert2').fadeIn(800).css("display","none");
fnAPPShowRightRegisterBtns();
});
//-->
</script>
<%
Set npartner = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
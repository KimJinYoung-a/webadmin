$(function() {
	// main banner
	var swiper01 = new Swiper(".basicImgRegist .swiper-container", {
		pagination:false,
		slidesPerView:'auto',
		spaceBetween:5
	});

	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

	// textarea auto size
	$("textarea.autosize").keyup(function () {
		$(this).css("height","1.96rem").css("height",($(this).prop("scrollHeight"))+"px");
	});
});

function chgodr(hidediv,v,formname,formdata){
	if(hidediv!=''){
		if (v == 1){
			eval("$('#"+hidediv+"')").css("display","none");
		}else{
			eval("$('#"+hidediv+"')").css("display","");
		}
	}
	if(formname!=''){
		eval("$('#"+formname+"')").val(formdata);
	}
}

function chgodr2(hidediv,v){
	if (v == 1){
		eval("$('#"+hidediv+"')").css("display","none");
	}else{
		eval("$('#"+hidediv+"')").css("display","");
	}
}

function chgodr3(){
	if ($('#requireimgchk').val() == "Y"){
		eval("$('#MakeDay4')").css("display","");
	}else{
		eval("$('#MakeDay4')").css("display","none");
	}
}

function TnGoClear(frm){
	frm.buyvat.value = "";
}

function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;
	
	if(frm.vatYn.value=="Y"){
		isvatYn = true;
	}else{
		isvatYn = false;
	}
	

	if (imargin.length<1){
		alert('마진을 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('판매가를 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('마진은 숫자로 입력하세요.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('판매가는 숫자로 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (isvatYn==true){
		isellvat = parseInt(parseInt(1/11 * parseInt(isellcash)));
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.01) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = 0;
		imileage = parseInt(isellcash*0.01) ;
	}

	frm.sellvat.value = isellvat;
	frm.buycash.value = ibuycash;
	frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
}

function CheckNumber(frm){
	var itemWeight = frm.itemWeight.value;
	if (!IsDigit(itemWeight)){
		alert('무게는 숫자로 입력하세요.');
		frm.itemWeight.focus();
		return;
	}
}

function MultiSelectButton(clsid,formname,formval){
	if(eval("$('#"+formname+"')").val() == formval){
		eval("$('#"+clsid+"')").removeClass("selected");
		eval("$('#"+formname+"')").val("");
	}
	else{
		eval("$('#"+clsid+"')").addClass("selected");
		eval("$('#"+formname+"')").val(formval);
	}
}

// 주문제작상품 문구
function checkItemDiv(){
	if ($("#reqMsg").val()=="06"){
		$("#itemdiv").val("06");
	}else{
		$("#itemdiv").val("16");
	}
}

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
 }

function AddDetailInfo(){
	if($("#dicheckcnt").val()>14){
		alert("상세 정보의 추가 갯수는 15개 입니다.");
	}else{
		// 행추가
		var oRow;
		oRow = "							<li id='DetailList" + Number(Number($("#dicheckcnt").val())+1) + "'>"
		oRow += "								<p id='imgArea" + Number(Number($("#dicheckcnt").val())+1) + "'><button type='button' class='btnImgRegist' onclick=fnAPPuploadAddImage('addimgname"+ Number(Number($("#dicheckcnt").val())+1) +"','"+ Number(Number($("#dicheckcnt").val())+1) +"');>이미지 등록</button></p>"
		oRow += "								<p class='tMar1-5r'><textarea placeholder='내용을 입력해주세요' class='autosize' id='addimgtext' name='addimgtext'></textarea><input type='hidden' name='addimgname' id='addimgname'></p>"
		oRow += "							</li>"
		$("#DetailInfo ul").append(oRow);
		$("#dicheckcnt").val(Number(Number($("#dicheckcnt").val())+1));//추가 수량 카운트
		//alert($("#dicheckcnt").val());
	}
}

function Num2Str(inum,olen,cChr,oalign){
	var i, ilen, strChr;
	ilen = String(inum);
	ilen= ilen.length;
	strChr="";
	if(ilen < olen){
		for(i=0;i < olen-ilen; i++){
			strChr = strChr + "0";
		}
	}
	if(oalign=="L"){
		return inum + strChr;
	}else{
		return strChr + inum;
	}
}

function fnLimitCheckOption(limityn){
	if(limityn=="N"){
		if($("#useoptionyn").val()=="Y" && limityn != $("#limityn").val() && $("#limityn").val()=="Y"){
			if(confirm("저장된 옵션의 수량을\n무제한으로 변경하시겠습니까?")){
				chgodr('LimitCnt',1,'limityn','N');
				eval("$('#limitbtn2')").addClass("selected");
				eval("$('#limitbtn1')").removeClass("selected");
			}
		}else{
			chgodr('LimitCnt',1,'limityn','N');
			eval("$('#limitbtn2')").addClass("selected");
			eval("$('#limitbtn1')").removeClass("selected");
		}
	}else{
		if($("#useoptionyn").val()=="Y"){
			if(confirm("한정수량 변경 시 저장된 옵션의 수량 설정이\n 필요합니다. 변경하시겠습니까?")){
				chgodr('LimitCnt',1,'limityn','Y');
				eval("$('#limitbtn1')").addClass("selected");
				eval("$('#limitbtn2')").removeClass("selected");
			}
		}else{
			chgodr('LimitCnt',2,'limityn','Y');
			eval("$('#limitbtn1')").addClass("selected");
			eval("$('#limitbtn2')").removeClass("selected");
		}	
	}
}

function fnOptionCheckEdit(optiondiv){
	if($("#optlevel").val() != optiondiv && optiondiv!="0" && $("#optlevel").val()!="0" && $("#useoptionyn").val()=="Y"){
		if(confirm("옵션 설정 변경 시\n기존 저장된 옵션이 초기화됩니다.\n변경하시겠습니까?")){
			if(optiondiv=="1"){
				$("#optionTypename1").val("");
				$("#optionTypename2").val("");
				$("#optionTypename3").val("");
				$("#optionName1").val("");
				$("#optionName2").val("");
				$("#optionName3").val("");
				$("#optaddprice1").val("");
				$("#optaddprice2").val("");
				$("#optaddprice3").val("");
				$("#optaddbuyprice1").val("");
				$("#optaddbuyprice2").val("");
				$("#optaddbuyprice3").val("");
				chgodr('setopt',0,'optlevel',optiondiv);
				$("#setoptcnt").css("display","");
			}else if(optiondiv=="2"){
				$("#optionTypename1").val("");
				chgodr('setopt',0,'optlevel',optiondiv);
				$("#setoptcnt").css("display","");
			}
			$("#optsetend").removeClass("setContView");
			$("#optsetend").removeClass("setContView");
			for (var ix = 0; ix < 3; ix++){
				if(ix==optiondiv){
					eval("$('#optbtn" + optiondiv + "')").addClass("selected");
				}else{
					eval("$('#optbtn" + ix + "')").removeClass("selected");
				}
			}
		}
	}else{
		if(optiondiv=="0"){
			chgodr('setopt',1,'optlevel',0);chgodr('','','useoptionyn','N');
			$("#setoptcnt").css("display","none");
		}else if(optiondiv=="1"){
			$("#optionTypename1").val("");
			$("#optionTypename2").val("");
			$("#optionTypename3").val("");
			$("#optionName1").val("");
			$("#optionName2").val("");
			$("#optionName3").val("");
			$("#optaddprice1").val("");
			$("#optaddprice2").val("");
			$("#optaddprice3").val("");
			$("#optaddbuyprice1").val("");
			$("#optaddbuyprice2").val("");
			$("#optaddbuyprice3").val("");
			chgodr('setopt',0,'optlevel',optiondiv);
			$("#setoptcnt").css("display","");
		}else if(optiondiv=="2"){
			$("#optionTypename1").val("");
			chgodr('setopt',0,'optlevel',optiondiv);
			$("#setoptcnt").css("display","");
		}
		for (var ix = 0; ix < 3; ix++){
			if(ix==optiondiv){
				eval("$('#optbtn" + optiondiv + "')").addClass("selected");
			}else{
				eval("$('#optbtn" + ix + "')").removeClass("selected");
			}
		}
		if($("#limityn").val()=="Y"){
			chgodr('LimitCnt',1,'','');
		}
	}
}

function fnMakeUnusualSet(requirecontents){
    //requirecontents = decSpecialCharNativeFun(requirecontents); //
	requirecontents = Base64.decode(requirecontents);
	$(document).find("#requirecontentstxt").empty().append("<span class='setContView'>"+requirecontents + "</span>");
	$("#requirecontents").val(requirecontents);
}

function fnDeliveryInfoSet(deliveryInfo){
	//deliveryInfo = decSpecialCharNativeFun(deliveryInfo);
	deliveryInfo = Base64.decode(deliveryInfo);
	$(document).find("#deliveryInfotxt").empty().append("<span class='setContView'>" + deliveryInfo + "</span>");
	$("#ordercomment").val(deliveryInfo);
}

function fnItemInfoDivSet(callbackval){
	//alert(callbackval);
	callbackval = callbackval.replace(/ /g, "");
	callbackval = callbackval.replace(/!/g, "','");
	var arriteminfo=eval("['" + callbackval + "']");
	$(document).find("#iteminfotxt").empty().append("<span class='setContView'>" + arriteminfo[1] + "</span>");
	$("#infoDiv").val(arriteminfo[0]);
	CheckFormConfirmbtnShow();
}

function fnKeyWordSet(callbackval){
	//callbackval = decSpecialCharNativeFun(callbackval); //
	//var keyword=encSpecialCharNativeFun(callbackval);
	callbackval= Base64.decode(callbackval); //
	var keyword= callbackval; //
	keyword = keyword.replace(/ /g, "");
	keyword = keyword.replace(/\'/g, "");  //추가
	keyword = keyword.replace(/,/g, "','");
	var arrkeyword=eval("['" + keyword + "']");
	var keywordcnt = arrkeyword.length;
	$(document).find("#keywordtxt").empty().append("<span class='setContView'>" + keywordcnt + "건 등록</span>");
	$("#keywords").val(callbackval);
}

function fnVodLinkSet(vodlink){
	vodlink = Base64.decode(vodlink);
	//alert(vodlink);
	$("#itemvideo").val(vodlink);
	$("#imgspan4").addClass("done");
}

function fnVodDelSet(vodlink){
	$("#itemvideo").val("");
	$("#imgspan4").removeClass("done");
}

function getId(url,voddiv) {
	var regExp = '';
	var match = '';
	if(voddiv=="youtube"){
		regExp = /^.*(youtu.be\/|v\/|u\/\w\/|embed\/|watch\?v=|\&v=)([^#\&\?]*).*/;
		match = url.match(regExp);
		if (match && match[2].length == 11) {
			return match[2];
		} else {
			return 'error';
		}
	}else{
		regExp = /^.*(player\/|vimeo.com\/|v\/|u\/\w\/|video\?v=|\&v=)([^#\&\?]*).*/;
		match = url.match(regExp);
		if (match && match[2].length == 9) {
			return match[2];
		} else {
			return 'error';
		}
	}
}
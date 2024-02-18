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

function TnCheckOptionYN(optwintitle){
	var wintitle;
	if(optwintitle==""){
		$('#itemoptioncode2').val("");
		$('#itemoptioncode3').val("");
	}else{
		if(optwintitle==1){
			wintitle="단일 옵션 설정"
		}else{
			wintitle="이중 옵션 설정"
		}
		$('#optwintitle').val(wintitle);
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

function fnTempSave(frm){
	if(frm.itemname.value==""){
		alert("상품명을 입력해주세요.");
		return false;
	}else{
		if(frm.tempSaveYn.value=="Y"){
			frm.action="/apps/academy/itemmaster/WaitDIYItemRegister_Noimg_Edit_Process_App.asp";
		}else{
			frm.action="/apps/academy/itemmaster/WaitDIYItemRegister_Noimg_Process_App.asp";
		}
		frm.target = "FrameCKP";
		frm.submit();
	}
}

function fnAppCallWinRegister(){
//상품등록 콜
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}
	if ($("input[name='catecode']").val() == 0){
		alert("[기본] 전시 카테고리를 선택하세요.");
		return;
	}
	//상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
	if (validate(itemreg)==false) {
        return;
    }
	if (itemreg.itemsize.value == ""){
		alert("상품 크기를 입력해주세요.");
		itemreg.itemsize.focus();
		return;
	}
	if (itemreg.itemWeight.value == ""){
		alert("상품 무게를 입력해주세요.");
		itemreg.itemWeight.focus();
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
	//배송비 안내
    if (!itemreg.ordercomment.value){
        alert('배송비 안내를 입력해주세요.');
        return;
    }
	//상품 품목정보
    if (!itemreg.infoDiv.value){
        alert('상품에 해당하는 품목을 선택해주십시요.');
        return;
    }
	//안전인증정보
    if (itemreg.safetyYn.value=="Y"){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해주세요.');
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해주세요.');
	        return;
	    }
    }
    // 정상가격
	if (confirm("소비자가(" + itemreg.sellcash.value + ")/공급가(" + itemreg.buycash.value + ")가 정확히 입력되었습니까?") == false) {
		itemreg.sellcash.focus();
		return;
    }
    if(confirm("상품을 올리시겠습니까? \n담당MD 승인후 반영 됩니다.") == true){
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


function AddDetailInfo(){
	if($("#dicheckcnt").val()>14){
		alert("상세 정보의 추가 갯수는 15개 입니다.");
	}else{
		// 행추가
		var oRow;
		oRow = "							<li id='DetailList" + Number($("#dicheckcnt").val())+1 + "'>"
		oRow += "								<p id='imgArea'><button type='button' class='btnImgRegist'>이미지 등록</button></p>"
		oRow += "								<p class='tMar1-5r'><textarea placeholder='내용을 입력해주세요' class='autosize' name='addimgtext'></textarea><input type='hidden' name='addimgname' id='addimgname'></p>"
		oRow += "							</li>"
		$("#DetailInfo ul").append(oRow);
		$("#dicheckcnt").val(Number($("#dicheckcnt").val())+1);//추가 수량 카운트
	}
}

function fntempSaveEnd(waititemid){
	if($("#itemregYn").val()=="Y"){
		fnAPPclosePopup();
	}else{
		$("#tempSaveYn").val("Y");
		$("#waititemid").val(waititemid);
		$('#alert1').fadeIn(800).css("display","");
		setTimeout(function(){
				$("#alert1").fadeOut(1000);
			}, 5000);
		$('#alert1').fadeIn(800).css("display","none");
	}
}
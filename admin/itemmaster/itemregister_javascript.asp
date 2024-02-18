function UseTemplate() {
	var popwin = window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
    popwin.focus();
}

function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_ItemAttribSelect.asp?itemid=0&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value               = varArray[0];
	document.itemreg.margin.value                   = varArray[1];
	document.itemreg.defaultmargin.value            = varArray[1];  //업체기본마진.
	document.itemreg.defaultmaeipdiv.value          = varArray[2];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[3];
    document.itemreg.defaultDeliverPay.value        = varArray[4];
    document.itemreg.defaultDeliveryType.value      = varArray[5];

    if (document.itemreg.defaultmaeipdiv.value=="M"){
        document.itemreg.mwdiv[0].checked = true; //매입
    }else if (document.itemreg.defaultmaeipdiv.value=="W"){
        document.itemreg.mwdiv[1].checked = true; //위탁
    }else if (document.itemreg.defaultmaeipdiv.value=="U"){
        document.itemreg.mwdiv[2].checked = true; //업체
    }

    TnCheckUpcheYN(document.itemreg);
}

// ============================================================================
// 업체마진자동입력
function TnDesignerNMargineAppl2(){
	var varArray;
	varArray = document.itemreg.marginData.value.split(',');

	document.itemreg.designerid.value = document.itemreg.makerid.value;
	document.itemreg.margin.value = varArray[0];

    document.itemreg.defaultmargin.value            = varArray[0];  //업체기본마진.
	document.itemreg.defaultmaeipdiv.value          = varArray[1];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[2];
    document.itemreg.defaultDeliverPay.value        = varArray[3];
    document.itemreg.defaultDeliveryType.value      = varArray[4];

    if (document.itemreg.defaultmaeipdiv.value=="M"){
        document.itemreg.mwdiv[0].checked = true; //매입
    }else if (document.itemreg.defaultmaeipdiv.value=="W"){
        document.itemreg.mwdiv[1].checked = true; //위탁
    }else if (document.itemreg.defaultmaeipdiv.value=="U"){
        document.itemreg.mwdiv[2].checked = true; //업체
    }

    TnCheckUpcheYN(document.itemreg);
}

function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;

	isvatinclude = frm.vatinclude[0].checked;

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

	if (isvatinclude==true){
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round로 변경
		imileage = parseInt(isellcash*0.005) ;
	}else{
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round로 변경
		imileage = parseInt(isellcash*0.005) ;
	}

	frm.buycash.value = ibuycash;
	frm.mileage.value = imileage;

	//최대구매수량 조정(가격비례)
	if(isellcash<100) {
		frm.orderMaxNum.value="1";
	} else if(isellcash<10000) {
		frm.orderMaxNum.value="500";
	} else if(isellcash<100000) {
		frm.orderMaxNum.value="200";
	} else {
		frm.orderMaxNum.value="100";
	}
}

// ============================================================================
// 카테고리등록
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.itemreg;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;

	EnDisableFlowerShop();
}

// ============================================================================
	// 카태고리 선택 팝업
	function popCateSelect(iid){
	    var dftDiv = "";
	    var chk = 0;

	    //기본 카테고리인지 추가인지 체크

	    if (!document.all.cate_div){
	        dftDiv = "D";
	    }else{
	        if (document.all.cate_div.length==undefined){
	            if (document.all.cate_div.value=="D") chk++;
	        }else{
        	    for(l=0;l<document.all.cate_div.length;l++)	{
        			if (document.all.cate_div[l].value=="D") chk++;
        		}
        	}
		}

		if (chk<1) dftDiv="D";

		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid + "&dftDiv=" + dftDiv, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var designerid = document.all.itemreg.designerid.value;
		if(designerid == ""){
			alert("업체를 선택하세요.");
			return;
		}

		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 레이어에서 전시카테고리 추가
	function addDispCateItem(dcd,cnm,div,dpt) {
		// 기존에 값에 중복 카테고리 여부 검사
		if(tbl_DispCate.rows.length>0)	{
			if(tbl_DispCate.rows.length>1)	{
				for(l=0;l<document.all.isDefault.length;l++)	{
				    if((document.all.catecode[l].value==dcd)) {
						alert("이미 지정된 같은 카테고리가 있습니다..");
						return;
					}
				}
			}
			else {
			    if((document.all.catecode.value==dcd)) {
					alert("이미 지정된 같은 카테고리가 있습니다..");
					return;
				}
			}
		}

		// 행추가
		var oRow = tbl_DispCate.insertRow();
		oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="y") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'>";
		}
		$(cnm).each(function(i){
			if(dpt>i) {
				if(i>0) oCell2.innerHTML += " >> ";
				oCell2.innerHTML += $(this).text();
			}
		});
		oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
		oCell2.innerHTML += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
		$("#lyrDispCateAdd").fadeOut();

		//상품속성 출력
		printItemAttribute();
	}

	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//상품속성 출력
			printItemAttribute();
		}
	}

// ============================================================================
// 옵션수정
function editItemOption(itemid, waityn) {
	var param = "itemid=" + itemid + "&waityn=" + waityn;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function doEditItemOption(itemid, waityn, arrmode, arritemoption, arritemoptionname, arroptuseyn, arroptsellyn, arroptlimityn, arroptlimitno, arroptlimitsold) {
	alert("a");
	// var param = "itemid=" + itemid + "&waityn=" + waityn;

	// popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=700,height=400,scrollbars=yes,resizable=yes');
	// popwin.focus();
}

function popNormalOptionAdd() {
	popwin = window.open('/common/module/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popEtcOptionAdd(){
	popwin = window.open('/common/module/etcitemoptionadd.asp' ,'normalitemoptionadd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 옵션을 추가한다
function InsertOptionWithGubun(ioptTypeName, ft, fv) {
	var frm = document.itemreg;

	//옵션값이 같은것이 있으면 skip ,전용옵션인경우 제외
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

    frm.optTypeNm.value = ioptTypeName;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

// 선택된 옵션 삭제
function delItemOptionAdd()
{
	var frm = document.itemreg;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0){
		alert("삭제할 옵션을 선택해주십오.");
	}else{
	    for(i=0; i<frm.realopt.options.length; i++){
    		if(frm.realopt.options[i].selected){
    			frm.realopt.options[i] = null;
    			i=i-1;
    		}
    	}

		if (frm.realopt.options.length<1){
		    frm.optTypeNm.value = '';
		}

		//frm.realopt.options[sidx] = null;
	}
}


// ============================================================================
// 이미지표시
function ClearImage(img,fsize,wd,ht) {
    img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";
}

function ClearImage2(img,fsize,wd,ht) {
    img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +");\" class='text' size='"+ fsize +"'>";
}

function CheckExtension(imgname, allowext) {
    var ext = imgname.lastIndexOf(".");
    if (ext < 0) {
        return false;
    }

    ext = imgname.toLowerCase().substring(ext + 1);
    allowext = "," + allowext + ",";
    if (allowext.indexOf(ext) < 0) {
        return false;
    }

    return true;
}

function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
{
    var ext;
    var filename;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
        ClearImage(img,fsize,imagewidth,imageheight);
        return false;
    }

    return true;
}


// ============================================================================
// 저장하기
function SubmitSave() {
	var itemreg = document.all.itemreg;

	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.makerid.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if(!$("input[name='isDefault'][value='y']").length) {
		alert("전시 카테고리를 선택하세요.\n※ 전시 기본 카테고리는 필수 있습니다.");
		return;
	}

	// 입력한 마진과 다를경우 체크
    if (itemreg.margin.value.length>0){
        if (Math.round((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[소비자가*마진 = 공급가]");


    		if (!confirm('입력한 마진과 입력된 판매가 대비 매입가 금액이 상이 합니다. 계속 진행 하시겠습니까?')){
    		    itemreg.sellcash.focus();
    			return;
    		}
        }
	}

	// 업체 기본마진과 다를경우 체크
	if (itemreg.defaultmargin.value.length>0){
        if (Math.round((itemreg.sellcash.value*1) * (itemreg.defaultmargin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {

    		if (!confirm('업체 기본 마진과 입력된 판매가 대비 매입가 금액이 상이 합니다. 계속 진행 하시겠습니까?')){
    			return;
    		}
        }
	}

    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            return;
        }
    }

    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[4].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        itemreg.deliverytype[4].focus();
        return;
    }

    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
      //  if (itemreg.deliverOverseas.checked){
       //     alert('텐바이텐 배송일 경우에만 해외배송을 하실 수 있습니다.');
      //      return;
      //  }
    }

    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
    }

    //업체배송만 주문제작 가능.
    if ((!itemreg.mwdiv[2].checked)&&(itemreg.itemdiv[1].checked)){
        alert('주문 제작상품은 업체배송인경우만 가능합니다.');
        itemreg.itemdiv[0].focus();
        return;
    }

	if(document.itemreg.deliverfixday[1].checked) {
		if(document.itemreg.freight_min.value<=0||document.itemreg.freight_max.value<=0) {
            alert('화물배송 비용을 입력해주세요.');
            document.itemreg.freight_min.focus();
            return;
		}
	}

	// 배송방법 해외직구 체크
	if (itemreg.deliverfixday[3].checked == true){
		if (itemreg.mwdiv[2].checked == false){
			alert('해외직구는 업체배송만 선택 가능 합니다.');
			return;
		}
		if ( !(itemreg.deliverytype[1].checked == true || itemreg.deliverytype[3].checked == true) ){
			alert('해외직구는 업체무료배송과 업체조건배송만 선택 가능 합니다.');
			return;
		}
		if (itemreg.deliverarea[0].checked == false){
			alert('해외직구는 전국배송만 선택 가능 합니다.');
			return;
		}
	}

    //==================================================================================


	if(!itemreg.itemdiv[3].checked) { //Present상품은 판매가 0원 가능
	    if (itemreg.buycash.value*1>itemreg.sellcash.value*1){
	        alert("매입가격이 판매가 보다 큽니다.");
			itemreg.sellcash.focus();
			return;
	    }

		if (itemreg.sellcash.value*1 < 0 || itemreg.sellcash.value*1 >= 20000000){
			alert("판매 가격은 20,000,000원 미만으로 등록 가능합니다.");
			itemreg.sellcash.focus();
			return;
		}

		if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
	        alert("마일리지는 판매가보다 클 수 없습니다.");
	        itemreg.mileage.focus();
	        return;
	    }

		if((itemreg.sellcash.value*0.05) <= itemreg.mileage.value*1){
		  	alert("마일리지는 1% 이상 5% 이하로만 등록 가능합니다.");
		  	itemreg.mileage.focus();
		  	return;
		}
	}

	if((itemreg.sellyn[0].checked)&&(itemreg.isusing[1].checked)) {
        alert('판매여부와 사용여부를 확인해주세요.\n\n※사용하지 않는 상품은 판매중을 선택할 수 없습니다.');
        return;
	}

	//상품 품목정보
    if (!itemreg.infoDiv.value){
        alert('상품에 해당하는 품목을 선택해주십시요.');
        itemreg.infoDiv.focus();
        return;
    } else if(itemreg.infoDiv.value=="35") {
    	if(!itemreg.itemsource.value) {
	        alert('상품의 재질을 입력해주세요.');
	        itemreg.itemsource.focus();
	        return;
    	}
    	if(!itemreg.itemsize.value) {
	        alert('상품의 크기를 입력해주세요.');
	        itemreg.itemsize.focus();
	        return;
    	}
    }

	//안전인증정보.
    if (itemreg.safetyYn[0].checked){
  		if($("#real_safetynum").val() == ""){
  			alert("안전인증구분을 선택하고 인증번호를 입력후 추가버튼을 클릭해주세요.");
  			return;
  		}
    }

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }

	if(itemreg.orderMinNum.value<1||document.itemreg.orderMinNum.value>32000) {
        alert('최소판매수는 1~32,000 범위의 숫자로 입력해주세요.');
        itemreg.orderMinNum.focus();
        return;
	}
	if(itemreg.orderMaxNum.value<1||document.itemreg.orderMaxNum.value>32000) {
        alert('최대판매수는 1~32,000 범위의 숫자로 입력해주세요.');
        itemreg.orderMaxNum.focus();
        return;
	}
	if(parseInt(itemreg.orderMinNum.value)>parseInt(itemreg.orderMaxNum.value)) {
        alert('최대판매수보다 최소판매수가 클 수 없습니다.');
        itemreg.orderMinNum.focus();
        return;
	}


	if (itemreg.useoptionyn[0].checked == true) {
	    if (itemreg.optlevel[0].checked == true) {
	    //단일옵션
    	    if (itemreg.realopt.length < 1) {
                alert("추가된 옵션이 없습니다.");
                // itemreg.useoptionyn.focus();
                return;
            }

    	    if (itemreg.realopt.length < 2) {
                alert("옵션은 두개 이상이어야 합니다.(옵션별로 한정/전시설정이 가능합니다.)");
                // itemreg.useoptionyn.focus();
                return;
            }
        }else if (itemreg.optlevel[1].checked == true) {
        //이중옵션
            if ((itemreg.optionTypename1.value.length<1)||(itemreg.optionTypename2.value.length<1)){
                alert("이중옵션을 사용할 경우 옵션구분명 은 최소 2개 이상 등록하셔야 합니다.");
                itemreg.optionTypename2.focus();
                return;
            }

            var chkCnt=0;
            for (var i=0;i<itemreg.optionName1.length;i++){
                if (itemreg.optionName1[i].value.length>0) chkCnt++;
            }

            if (chkCnt<2){
                alert("옵션은 각 구분당 2개 이상이어야 합니다.");
                itemreg.optionName1[1].focus();
                return;
            }

            chkCnt=0;

            for (var i=0;i<itemreg.optionName2.length;i++){
                if (itemreg.optionName2[i].value.length>0) chkCnt++;
            }

            if (chkCnt<2){
                alert("옵션은 각 구분당 2개 이상이어야 합니다.");
                itemreg.optionName2[1].focus();
                return;
            }

            if (itemreg.optionTypename3.value.length>0){
                chkCnt=0;

                for (var i=0;i<itemreg.optionTypename3.length;i++){
                    if (itemreg.optionTypename3[i].value.length>0) chkCnt++;
                }

                if (chkCnt<2){
                    alert("옵션은 각 구분당 2개 이상이어야 합니다.");
                    itemreg.optionName3[1].focus();
                    return;
                }

            }
        }
	}

    var optiont = "";
    var optionv = "";
    var optvalue = 11; // 전용옵션(11 - 99)
    for(var i = 0; i < itemreg.realopt.options.length; i++) {
        optiont += (itemreg.realopt.options[i].text + "|");

        // 전용옵션추가
        if (itemreg.realopt.options[i].value == "0000") {
            if (optvalue > 99) {
                alert("너무많은 옵션을 추가하셨습니다.");
                return;
            }
            itemreg.realopt.options[i].value = "00" + optvalue;
            optvalue = optvalue + 1;
        }

        optionv += (itemreg.realopt.options[i].value + "|");
    }

	// 기본 색상선택
	if(!itemreg.DFcolorCD.value) {
        alert("상품의 기본 색상을 선택해주세요.");
        return;
	}
    if (itemreg.imgDFColor.value != "") {
        if (CheckImage(itemreg.imgDFColor, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40) != true) {
            return;
        }
    }

	//=== 상품 이미지 ================================
    //if(itemreg.regimg.checked==false) {
	    if (itemreg.imgbasic.value=="") {
	        alert("기본이미지는 필수입니다.");
	        return;
	    } else {
	        if (CheckImage(itemreg.imgbasic, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40) != true) {
	            return;
	        }
	    }
	//}

    if (itemreg.imgmask.value != "") {
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40) != true) {
            return;
        }
    }

	var imgupcnt = document.all.imgIn.rows.length;
	var tmp = "";
	var tmpvalue = "";
	for(var a=0;a<imgupcnt;a++){
		tmp = itemreg.addimgname[a];
	    if (tmp.value != "") {
	        if (CheckImage(tmp, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40) != true) {
	            return;
	        }
	    }
	}

    if(confirm("상품을 올리시겠습니까?") == true){
		//안전인증 api로 조회 후 받은 데이터 db저장 후 생성idx값 받아 셋팅
		if(itemreg.safetyYn[0].checked) {
			$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"o",$("#real_safetydiv").val()));
		}

        itemreg.itemoptioncode2.value = optionv;
        itemreg.itemoptioncode3.value = optiont;

		itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.deliverytype[3].disabled=false;
        itemreg.deliverytype[4].disabled=false;

        itemreg.target = "FrameCKP";
        itemreg.submit();
    }
}

//매입구분 체크에 따른 배송구분 체크
function TnCheckUpcheYN(frm){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// 기본체크
		// 배송구분 지정(텐바이텐)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
        frm.deliverytype[3].disabled=true;  //업체개별배송(9)
        frm.deliverytype[4].disabled=true;  //업체착불배송(7)
        frm.deliverOverseas.checked=true;	// 해외배송체크
       // frm.optlevel[0].checked=true;
       // frm.optlevel[1].disabled=true;
	}
	else if(frm.mwdiv[2].checked){
	    // 배송구분 지정(업체배송)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// 업체조건배송 기본 체크
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// 업체착불배송 기본 체크
	    }else{
	        frm.deliverytype[1].checked=true;	// 기본 체크
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;
        frm.deliverytype[4].disabled=false;  //업체착불배송(7)
        frm.deliverOverseas.checked=false;	// 해외배송체크해제
      //  frm.optlevel[1].disabled=false;
	}

	if (frm.deliverytype[1].checked==true || frm.deliverytype[3].checked==true){
		frm.deliverfixday[3].disabled=false;	// 해외직구
	}
}

function TnGoClear(frm){
	frm.buycash.value = "";
	frm.mileage.value = "";
}

// 배송구분
function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("매입위탁 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입위탁구분을 확인해주세요!!");
			frm.mwdiv[0].checked=true;
			//frm.optlevel[1].checked=false;
			//frm.optlevel[1].disabled=true;
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked || frm.deliverytype[4].checked){
	//else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("매입위탁 구분이 매입이나 위탁일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입위탁구분을 확인해주세요!!");
			frm.mwdiv[2].checked=true;
			//frm.optlevel[1].disabled=false;
		}
	}
}

function TnChkIsUsing(frm) {
	if(frm.isusing[0].checked) {
		frm.sellyn[0].disabled=false;
	} else {
		if(frm.sellyn[0].checked) {
			alert("사용여부를 사용안함으로 선택하셨습니다.\n판매여부가 [판매안함]으로 자동설정됩니다.");
		}
		frm.sellyn[1].checked=true;
		frm.sellyn[0].disabled=true;
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // 옵션사용

        opttype.style.display="";

        if (frm.optlevel[1].checked==true){
            optlist.style.display ="none";
            optlist2.style.display ="";
        }else{
            optlist.style.display="";
            optlist2.style.display="none";
        }

	} else {
	    // 옵션없음
	    while (frm.realopt.length > 0) {
	        frm.realopt.options[0] = null;
        }
        opttype.style.display="none";
        document.all.optlist2.style.display="none";
		document.all.optlist.style.display="none";

        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";
    }
}


function EnDisableFlowerShop(){

    var frm = document.itemreg;
    if ((frm.cd1.value=="110")&&(frm.cd2.value=="060")){
        frm.deliverarea[1].disabled = false;
        frm.deliverarea[2].disabled = false;

        frm.deliverfixday.disabled = false;
    }else{
        frm.deliverarea[1].disabled = true;
        frm.deliverarea[2].disabled = true;

        frm.deliverfixday.disabled = true;
        frm.deliverfixday.checked = false;
    }
}

function TnAutoChkDeliver() {
	var frm = document.itemreg
	switch(frm.defaultmaeipdiv.value) {
		case "M" :
			frm.mwdiv[0].checked=true;
			break;
		case "W" :
			frm.mwdiv[1].checked=true;
			break;
		case "U" :
			frm.mwdiv[2].checked=true;
			break;
	}
	TnCheckUpcheYN(frm);
}

// 배송방법
function TnCheckFixday(frm) {
	if(frm.deliverfixday[0].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="none";
	} else if(frm.deliverfixday[1].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="";

	// 해외직구
	} else if(frm.deliverfixday[3].checked) {
		frm.mwdiv[2].checked=true;
		frm.deliverarea[0].checked=true;

		document.getElementById("lyrFreightRng").style.display="none";
	} else {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=false;
		frm.deliverarea[2].disabled=false;
		document.getElementById("lyrFreightRng").style.display="none";
	}
}

//색상코드 선택
function selColorChip(cd) {
	var i;
	itemreg.DFcolorCD.value= cd;
	for(i=0;i<=31;i++) {
		document.all("cline"+i).bgColor='#DDDDDD';
	}
	if(!cd) document.all("cline0").bgColor='#DD3300';
	else document.all("cline"+cd).bgColor='#DD3300';
}

function ClearVal(comp){
    comp.value = "";
}

function checkItemDiv(comp){
    var frm = comp.form;

    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
        }
    }

	//티켓 상품일 경우 배송방법 클래스 선택 활성화(2018-05-10 정태훈)
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[2].checked){
            frm.deliverfixday[4].disabled=false;
        }else{
            frm.deliverfixday[4].disabled=true;
			frm.deliverfixday[0].checked=true;
        }
    }

    //주문제작 상품인경우.
    if (frm.itemdiv[1].checked){
        if (frm.reqMsg.checked){
            frm.itemdiv[1].value="06";
        }else{
            frm.itemdiv[1].value="16";
        }
    }

	//렌탈상품인 경우
	if (frm.itemdiv[4].checked){
		frm.orderMaxNum.value=1;
	}
}

//품목 선택 / 품목내용 표시
function chgInfoDiv(v) {
	$("#itemInfoList").empty();

	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "act_itemInfoDivForm.asp",
			data: "ifdv="+v,
			dataType: "html",
			async: false
		}).responseText;

		if(str!="") {
			$("#itemInfoList").html(str);
		}
	}
	if(v=="35") {
		$("#lyItemSrc").show();
		$("#lyItemSize").show();
	} else {
		$("#lyItemSrc").hide();
		$("#lyItemSize").hide();
	}

	// 안전인증체크. 전안법
	jsSafetyCheck('','');
}

//단순 라디오 선택자
function chgInfoChk(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
}

//문구 라디오 선택자
function chgInfoSel(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
	$(fm).parent().parent().find('[name="infoCont"]').val($(fm).attr("msg"));

	if($(fm).val()=="Y") {
		$(fm).parent().parent().find('[name="infoCont"]').removeAttr("readonly");
		$(fm).parent().parent().find('[name="infoCont"]').removeClass("text_ro");
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text");
	} else {
		$(fm).parent().parent().find('[name="infoCont"]').attr("readonly", true);
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text_ro");
	}
}

//상품설명이미지추가
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 6){
		alert("이미지는 최대 7개까지 가능합니다.");
		return;
	}

	var i = rowLen;
	var r  = f.imgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	r.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '15%';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = 'PC상품설명이미지 #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 800, 1600)"> (선택,800X1600, Max 800KB,jpg,gif)';
}

//모바일상품상세이미지추가
function InsertMobileImageUp() {
	var f = document.all;
	var rowLen = f.MobileimgIn.rows.length;

	if(rowLen > 11){
		alert("이미지는 최대 12개까지 가능합니다.");
		return;
	}

	var i = rowLen;
	var r  = f.MobileimgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	r.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '15%';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = '모바일상품상세이미지 #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addmoblieimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)';
}

//상품군에 따른 원산지 설명 표기
function jsSetArea(iValue){
	var i;
	for(i=0;i<=4;i++) {
 		eval("document.all.dvArea"+i).style.display = "none";
	}
 	eval("document.all.dvArea"+iValue).style.display = "";
}

function jsCallAPIsafety(certnum,isSave,safetydiv){
	var returnmsg = "";
	$.ajax({
		url: "/admin/itemmaster/safety_api_auth_proc.asp?issave="+isSave+"&certnum="+certnum+"&safetydiv="+safetydiv+"&statusmode=real",
		cache: false,
		async: false,
		success: function(message)
		{
			returnmsg = message;
		}
	});
	return returnmsg;
}

//전시카테고리(안전인증값)에 따른 alert 메세지.
function jsAlertCatecodeSafety(){
	var auth_go_catecode = "";
	if(typeof itemreg.catecode != "undefined"){
		if(itemreg.catecode.length == undefined){
			auth_go_catecode = itemreg.catecode.value;
		}else{
			for(si=0; si<itemreg.catecode.length; si++){
				auth_go_catecode = auth_go_catecode + itemreg.catecode[si].value + ",";
			}
		}

		if(auth_go_catecode != ""){
			$("#auth_go_catecode").val(auth_go_catecode);

			var ccode = $("#auth_go_catecode").val();
			$.ajax({
					url: "/common/item/catecode_safety_info_ajax.asp?catecode="+ccode,
					cache: false,
					async: false,
					success: function(msgc)
					{
						if(msgc != ""){
							msgc = msgc.replace(/br/gi,"\n");
							alert(msgc);
						}
					}
			});
		}
	}else{
		alert("전시카테고리를 선택해주세요.");
	}
}

//추가된 안전인증 리스트 개별 삭제
function jsSafetyDivListDel(listnum){
	var realvalue = $("#real_safetydiv").val();
	var jbSplit = $("#real_safetydiv").val().split(",");
	var jbSplitnum = $("#real_safetynum").val().split(",");
	var resultDiv = "";
	var resultNum = "";

	for(var i in jbSplit){
		if(jbSplit[i] != listnum){
			resultDiv = resultDiv + jbSplit[i] + ",";
			resultNum = resultNum + jbSplitnum[i] + ",";
		}
	}

	if(resultDiv.substr(resultDiv.length-1, 1) == ","){
		resultDiv = resultDiv.substr(0, resultDiv.length-1);
		resultNum = resultNum.substr(0, resultNum.length-1);
	}
	$("#real_safetydiv").val(resultDiv);
	$("#real_safetynum").val(resultNum);

	$("#l"+listnum+"").remove();
}

$(function(){
	// 로딩후 상품속성 내용 출력
	printItemAttribute();
});

function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_ItemAttribSelect.asp?itemid=<%=itemid%>&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

function popMultiLangEdit(iid) {
	window.open("/common/item/pop_MultiLangItemCont.asp?itemid="+iid+"&lang=EN", "multiLang_win", "width=600, height=500, scrollbars=yes, resizable=yes");
}

// ============================================================================
// 업체 기본마진 자동입력 - 브랜드변경시만.
function TnDesignerNMargineAppl(){
	var varArray;
	varArray = document.itemreg.marginData.value.split(',');

	document.itemreg.designerid.value = document.itemreg.designer.value;
	document.itemreg.margin.value = varArray[0];
    
    document.itemreg.defaultmargin.value            = varArray[0];  //업체기본마진.
	document.itemreg.defaultmaeipdiv.value          = varArray[1];
    document.itemreg.defaultFreeBeasongLimit.value  = varArray[2];
    document.itemreg.defaultDeliverPay.value        = varArray[3];
    document.itemreg.defaultDeliveryType.value      = varArray[4];
    
    if(document.itemreg.mwdiv.length>0){
        if (document.itemreg.defaultmaeipdiv.value=="M"){
            document.itemreg.mwdiv[0].checked = true; //매입
        }else if (document.itemreg.defaultmaeipdiv.value=="W"){
            document.itemreg.mwdiv[1].checked = true; //위탁
        }else if (document.itemreg.defaultmaeipdiv.value=="U"){
            document.itemreg.mwdiv[2].checked = true; //업체
        }
    }else{
        document.itemreg.mwdiv.value=document.itemreg.defaultmaeipdiv.value;
    }
    
    TnCheckUpcheYN(document.itemreg);
}

function CalcuAuto(frm){
	var isvatinclude, imileage;
	var isellcash, ibuycash, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatinclude = frm.vatinclude[0].checked;

	if (frm.sailyn[0].checked == true) {
	    // 정상가격
	    isellcash = frm.sellcash.value;
	    imargin = frm.margin.value;

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
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);   //parseInt-> round로 변경 
			imileage = parseInt(isellcash*0.005) ;
    	}else{
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);   //parseInt-> round로 변경 
			imileage = parseInt(isellcash*0.005) ;
    	}

    	frm.buycash.value = ibuycash;
    	frm.mileage.value = imileage;
	} else {
	    // 세일가격
	    isailprice = frm.sailprice.value;
	    isailmargin = frm.sailmargin.value;

    	if (isailmargin.length<1){
    		alert('세일마진을 입력하세요.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (isailprice.length<1){
    		alert('세일판매가를 입력하세요.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (!IsDouble(isailmargin)){
    		alert('세일마진은 숫자로 입력하세요.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (!IsDigit(isailprice)){
    		alert('세일판매가는 숫자로 입력하세요.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (isvatinclude==true){
    		isailpricevat = parseInt(parseInt(1/11 * parseInt(isailprice)));
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);         //parseInt-> round로 변경 
    		isailsuplycashvat = parseInt(parseInt(1/11 * parseInt(isailsuplycash)));
			imileage = parseInt(isailprice*0.005) ;
    	}else{
    		isailpricevat = 0;
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);         //parseInt-> round로 변경 
    		isailsuplycashvat = 0;
			imileage = parseInt(isailprice*0.005) ;
    	}

    	frm.sailpricevat.value = isailpricevat;
    	frm.sailsuplycash.value = isailsuplycash;
    	frm.sailsuplycashvat.value = isailsuplycashvat;
    	frm.mileage.value = imileage;
    }

	//할인율 계산
	if (frm.sailyn[0].checked == true) {
		document.getElementById("lyrPct").innerHTML = "";
	} else {
		isellcash = frm.sellcash.value;
		isailprice = frm.sailprice.value;
		var isalePercent = parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10;
		document.getElementById("lyrPct").innerHTML = "할인율: <font color='#EE0000'><strong>" + isalePercent + "%</strong></font>";
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
}

// ============================================================================
// 옵션수정
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function doEditItemOption(optioncnt, optlimitno, optlimitsold, optlimitstock) {
    // 옵션창에서 오픈창으로
    itemreg.optioncnt.value = optioncnt;

    itemreg.limitno.value = optlimitno;
    itemreg.limitsold.value = optlimitsold;
    itemreg.limitstock.value = optlimitstock;
}

function popNormalOptionAdd() {
	popwin = window.open('/common/module/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=800,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 옵션을 추가한다
function InsertOption(ft, fv) {
	var frm = document.itemreg;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

// 선택된 옵션 삭제
function delItemOptionAdd()
{
	var frm = document.itemreg;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0)
		alert("삭제할 옵션을 선택해주십오.");
	else
	{
		frm.realopt.options[sidx] = null;
	}
}


// ============================================================================
// 이미지표시
function ClearImage(img,fsize,wd,ht) {
	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

    document.getElementById("div"+ img.name).style.display = "none";

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "del";
}

function ClearImage2(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");\" class='text' size='"+ fsize +"'>";
	$("#divaddimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "del";
    }else{
    	document.itemreg.addimgdel.value = "del";
    }
}

function oldClearImage(img,fsize,wd,ht) {
	$("#divimg"+img+"").hide();
	$("input[name='"+img+"']").val("del");
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
        alert("이미지파일은 다음의 파일만 사용하세요.[" + extname + "]");
        ClearImage(img,fsize, imagewidth, imageheight, num);
        return false;
    }

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "";

    return true;
}

function CheckImage2(img, filesize, imagewidth, imageheight, extname, fsize, num)
{
    var ext;
    var filename;
    var imgcnt = $('input[name="addimgname"]').length;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지파일은 다음의 파일만 사용하세요.[" + extname + "]");
        ClearImage2(img,fsize, imagewidth, imageheight, num);
        return false;
    }

	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "";
    }else{
    	document.itemreg.addimgdel.value = "";
    }

    return true;
}


// ============================================================================
// 저장하기
function SubmitSave() {
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[기본] 전시 카테고리를 선택하세요.\n※ [추가] 전시 카테고리만 넣을 수 없습니다.");
		return;
	}

    // 카테고리 지정여부 검사
	if(tbl_Category.rows.length>0)	{
		if(tbl_Category.rows.length>1)	{
			var chk=0;
			for(l=0;l<document.all.cate_div.length;l++)	{
				if(document.all.cate_div[l].value=="D") chk++;
			}
			if(chk==0) {
				alert("카테고리에 기본 카테고리를 선택해주세요.\n※기본 카테고리는 필수항목입니다.");
				return;
			} else if(chk>1) {
				alert("카테고리에 기본 카테고리를 한개만 선택해주세요.");
				return;
			}
		}
		else {
			if(document.all.cate_div.length){
				if(document.all.cate_div[0].value!="D") {
					alert("카테고리에 기본 카테고리를 선택해주세요.\n※기본 카테고리는 필수항목입니다.");
					return;
				}
			} else {
				if(document.all.cate_div.value!="D") {
					alert("카테고리에 기본 카테고리를 선택해주세요.\n※기본 카테고리는 필수항목입니다.");
					return;
				}
			}
		}
	} else {
		alert("카테고리를 선택해주세요.");
		return;
	}
	
	//배송구분 체크 =========================================================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('배송 구분을 확인해주세요. 개별배송 업체가 아닙니다.');
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
        if(document.itemreg.mwdiv.length>0){
            if ((document.itemreg.mwdiv[0].checked)||(document.itemreg.mwdiv[1].checked)){
                alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
                return;
            }
        }else{
            if ((document.itemreg.mwdiv.value=="M")||(document.itemreg.mwdiv.value=="W")){
	            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
	            return;
	        }
        }
        
       // if (itemreg.deliverOverseas.checked){
       //     alert('텐바이텐 배송일 경우에만 해외배송을 하실 수 있습니다.');
       //     return;
       // }
    }
    
    if(document.itemreg.mwdiv.length>0){
        if (document.itemreg.mwdiv[2].checked){
            if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
                alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
                return;
            }
        }
    }else{
        if (document.itemreg.mwdiv.value=="U"){
	        if ((document.itemreg.deliverytype[0].checked)||(document.itemreg.deliverytype[2].checked)){
	            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
	            return;
	        }
	    }
    }
    
    //업체배송만 주문제작 가능.
    if(document.itemreg.mwdiv.length>0){
        if ((!document.itemreg.mwdiv[2].checked)&&(document.itemreg.itemdiv[1].checked)){
            alert('주문 제작상품은 업체배송인경우만 가능합니다.');
            return;
        }
    }else{
        if ((document.itemreg.mwdiv.value!="U")&&(document.itemreg.itemdiv[1].checked)){
            alert('주문 제작상품은 업체배송인경우만 가능합니다.');
            return;
        }
    }
    
	if(document.itemreg.deliverfixday[1].checked) {
		if(document.itemreg.freight_min.value<=0||document.itemreg.freight_max.value<=0) {
            alert('화물배송 비용을 입력해주세요.');
            document.itemreg.freight_min.focus();
            return;
		}
	}

    //==================================================================================
    
   //-------------------------------------------------------------------------------- 2014.02.14 정윤정 추가
	//1.사업자가 [간이과세자] 인 경우, 매입상품 등록 불가 / 업체,위탁 상품만 등록가능
	if(document.itemreg.mwdiv.length>0){
    	if((document.itemreg.jungsangubun.value =="간이과세")&&(document.itemreg.mwdiv[0].checked)){
    		alert("사업자가 [간이과세자]인 경우, [매입]상품은 등록불가능합니다. \n[위탁],[업체배송]상품만 등록가능합니다. ");
    		document.itemreg.mwdiv[0].focus();
    		return;
    	}
    }else{
        if((document.itemreg.jungsangubun.value =="간이과세")&&(document.itemreg.mwdiv.value=="M")){
    		alert("사업자가 [간이과세자]인 경우, [매입]상품은 등록불가능합니다. \n[위탁],[업체배송]상품만 등록가능합니다. ");
    		return;
    	}
    }
	
	//2.사업자가 [면세사업자] 인 경우, 면세상품으로만 등록가능 
	if((document.itemreg.jungsangubun.value =="면세")&&(document.itemreg.vatinclude[0].checked)){
		alert("사업자가 [면세사업자]인 경우, [과세]상품은 등록불가능합니다. \n[면세]상품만 등록가능합니다. ");
		document.itemreg.vatinclude[1].focus();
		return; 
	}
	
	//3.사업자가 [텐바이텐]인 경우, 매입상품만 등록 가능
	if(document.itemreg.mwdiv.length>0){
    	if((document.itemreg.companyno.value =="211-87-00620")&&(!document.itemreg.mwdiv[0].checked)){
    		alert("사업자가 [텐바이텐]인 경우, [매입]상품만 등록가능합니다. ");
    		document.itemreg.mwdiv[0].focus();
    		return;
    	}
    }else{
        if((document.itemreg.companyno.value =="211-87-00620")&&(!document.itemreg.mwdiv.value=="M")){
    		alert("사업자가 [텐바이텐]인 경우, [매입]상품만 등록가능합니다. ");
    		return;
    	}
    }
	 //--------------------------------------------------------------------------------  
    if (validate(itemreg)==false) {
        return;
    }
    
    //상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
	
    if (itemreg.sailyn[0].checked == true) {
        // 정상가격
        if (Math.round((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[소비자가*마진 = 공급가]");
    		itemreg.sellcash.focus();

    		if (!confirm('마진율로 계산 할 수 없을때 공급가만 입력하면 마진율은 공급가에 맞춰 계산됩니다. \n계속 진행 하시겠습니까?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
            alert("마일리지는 판매가보다 클 수 없습니다.");
            itemreg.mileage.focus();
            return;
        }

        if(!itemreg.itemdiv[3].checked) { //Present상품은 판매가 0원 가능
	        if (itemreg.sellcash.value*1 < 300 || itemreg.sellcash.value*1 >= 20000000){
				alert("판매 가격은 300원 이상 20,000,000원 미만으로 등록 가능합니다.");
				itemreg.sellcash.focus();
				return;
			}
		}

    } else {
        // 할인가격
        if (Math.round((itemreg.sailprice.value*1) * (itemreg.sailmargin.value*1) / 100) != ((itemreg.sailprice.value*1) - (itemreg.sailsuplycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[할인소비자가*할인마진 = 할인공급가]");
    		itemreg.sailprice.focus();

    		if (!confirm('계속 진행 하시겠습니까?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sailprice.value*1){
            alert("마일리지는 판매가보다 클 수 없습니다.");
            itemreg.mileage.focus();
            return;
        }

        if(!itemreg.itemdiv[3].checked) { //Present상품은 판매가 0원 가능
	        if (itemreg.sailprice.value*1 < 300 || itemreg.sailprice.value*1 >= 20000000){
				alert("판매 가격은 300원 이상 20,000,000만원 미만으로 등록 가능합니다.");
				itemreg.sailprice.focus();
				return;
			}
		}
    }

    //세일가격이 정상가격 보다 클 수 없음.
    if (itemreg.sailprice.value*1>itemreg.sellcash.value*1){
        alert('세일가격이 정상가보다 클 수 없습니다.');
        return;
    }
    
    if (itemreg.sailsuplycash.value*1>itemreg.buycash.value*1){
        alert('세일매입가가 정상 매입가보다 클 수 없습니다.');
        return;
    }

	// 원래입력된 가격보다 수정된 가격이 20%이상 차이가 날때 확인 메시지
	if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.8)%>) {
		if(!confirm("\n\n\n\n입력하신 소비자가 수정하기 전의 가격보다 매우 많이 차이납니다.\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sellcash.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	}
	<% if oitem.FOneItem.Fsailyn="Y" then %>
	if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.8)%>) {
		if(!confirm("\n\n\n\n입력하신 할인가가 수정하기 전의 가격보다 매우 많이 차이납니다.\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sailprice.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	}
	<% end if %>

	// 할인율 검사(50%이상 경고)
	if (document.itemreg.sailyn[1].checked == true) {
		if(((document.itemreg.sellcash.value-document.itemreg.sailprice.value)/document.itemreg.sellcash.value*100)>50) {
			if(!confirm("\n\n할인율이 매우 높게 설정되어있습니다.\n\n입력하신 내용이 정확합니까?\n\n")) {
				return;
			}
		}
	}

	if((itemreg.sellyn[0].checked||itemreg.sellyn[1].checked)&&(itemreg.isusing[1].checked)) {
        alert('판매여부와 사용여부를 확인해주세요.\n\n※사용하지 않는 상품은 판매중을 선택할 수 없습니다.');
        return;
	}

		itemreg.chkModSR.value = "N"; //기본값 설정, 상태 변경시 값 변경
 //오픈예약상태일때 판매여부 변경처리
 if(itemreg.sellreservedate.value != ""){
 	 if(itemreg.sellyn[0].checked){
 	 	if(confirm(itemreg.sellreservedate.value+"로 오픈예약된 상품입니다. 판매중으로 상태 변경하시면, 상품오픈예약설정은 취소됩니다. 계속하시겠습니까? ")){
 	 		itemreg.sellreservedate.value = "";
 	 		itemreg.chkModSR.value = "Y";
 	 	}else{
 	 		itemreg.sellyn[0].focus();
 	 		return;
 	 	}
 	}
 	
 	if(itemreg.sellyn[1].checked){
 	 	if(confirm(itemreg.sellreservedate.value+"로 오픈예약된 상품입니다. 일시품절로 상태 변경하시면, 상품오픈예약설정은 취소됩니다. 계속하시겠습니까? ")){
 	 		itemreg.sellreservedate.value = "";
 	 		itemreg.chkModSR.value = "Y";
 	 	}else{
 	 		itemreg.sellyn[1].focus();
 	 		return;
 	 	}
 	}
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

    if (itemreg.basic.value == "del") {
        alert("기본이미지는 필수입니다.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage(itemreg.imgbasic, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
                return;
            }
        }
    }

    if (itemreg.imgmask.value != "") {
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage(itemreg.imgadd1, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage(itemreg.imgadd2, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage(itemreg.imgadd3, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage(itemreg.imgadd4, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage(itemreg.imgadd5, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }


    if(confirm("상품을 올리시겠습니까?") == true){
		// 안전인증 api로 조회 후 받은 데이터 db저장 후 생성idx값 받아 셋팅
		if(itemreg.safetyYn[0].checked) {
			$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"u",$("#real_safetydiv").val()));
		}

        itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.deliverytype[3].disabled=false;
        itemreg.deliverytype[4].disabled=false;
        itemreg.target = "FrameCKP";
        itemreg.submit();
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
 
function TnCheckUpcheYN(frm){
    if(frm.mwdiv.length>0){
    	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
    		frm.deliverytype[0].checked=true;	// 기본체크
    		// 배송구분 지정(텐바이텐)
    		frm.deliverytype[0].disabled=false;
    		frm.deliverytype[1].disabled=true;
    		frm.deliverytype[2].disabled=false;
    		frm.deliverytype[3].disabled=true;  //업체조건배송(9)
    		frm.deliverytype[4].disabled=true;  //업체착불배송(7) 
    	}
    	else if(frm.mwdiv[2].checked){
    	    // 배송구분 지정(업체조건배송)
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
            
    	}
    }else{
        if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
    		frm.deliverytype[0].checked=true;	// 기본체크
    		// 배송구분 지정(텐바이텐)
    		frm.deliverytype[0].disabled=false;
    		frm.deliverytype[1].disabled=true;
    		frm.deliverytype[2].disabled=false;
    		frm.deliverytype[3].disabled=true;  //업체조건배송(9)
    		frm.deliverytype[4].disabled=true;  //업체착불배송(7) 
    	}
    	else if(frm.mwdiv.value=="U"){
    	    // 배송구분 지정(업체조건배송)
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
            
    	}
    }
}

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
	} else {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=false;
		frm.deliverarea[2].disabled=false;
		document.getElementById("lyrFreightRng").style.display="none";
	}
}

function TnCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readonly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readonly=true;
		frm.limitsold.style.background='#E6E6E6';
		
		document.all.dvDisp.style.display = 'none';
		this.form.limitdispyn[0].checked = false;
		this.form.limitdispyn[1].checked = true;
	}
	else {
		// 한정
		if ((frm.optioncnt.value*1) > 0) {
		    // 옵션사용중
		    alert("옵션을 사용할경우 한정수량은 옵션창에서 수정가능합니다.");
		    frm.limityn[0].checked = true;
		    return;
        }

		frm.limitno.readonly = false;
		frm.limitno.style.background = '#FFFFFF';

		frm.limitsold.readonly = false;
		frm.limitsold.style.background = '#FFFFFF';
		
		document.all.dvDisp.style.display = '';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readonly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readonly=true;
		frm.limitsold.style.background='#E6E6E6';
	}
	else {
		// 한정
		if ((frm.optioncnt.value*1) > 0) {
		    // 옵션사용중
		    // alert("한정수량은 옵션창에서 수정가능합니다.");
		    return;
        }

		frm.limitno.readonly=false;
		frm.limitno.style.background='#FFFFFF';

		frm.limitsold.readonly=false;
		frm.limitsold.style.background='#FFFFFF';
	}
}

function TnGoClear(frm){
	frm.buycash.value = "";
	frm.mileage.value = "";
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
	    if(frm.mwdiv.length>0){
    		if (frm.mwdiv[2].checked){
    			alert("매입위탁 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입위탁구분을 확인해주세요!!");
    			frm.mwdiv[0].checked=true;
    		}
    	}else{
    	    if (frm.mwdiv.value=="U"){
    			alert("매입위탁 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입위탁구분을 확인해주세요!!");
    		}
    	}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	//else if(frm.deliverytype[1].checked ){
	    if(frm.mwdiv.length>0){
    		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
    			alert("매입위탁 구분이 매입이나 위탁일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입위탁구분을 확인해주세요!!");
    			frm.mwdiv[2].checked=true;
    		}
    	}else{
    	    if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
    			alert("매입위탁 구분이 매입이나 위탁일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입위탁구분을 확인해주세요!!");
    		}
    	}
	}
}

function TnChkIsUsing(frm) {
	if(frm.isusing[0].checked) {
		frm.sellyn[0].disabled=false;
		frm.sellyn[1].disabled=false;
	} else {
		if(frm.sellyn[0].checked||frm.sellyn[1].checked) {
			alert("사용여부를 사용안함으로 선택하셨습니다.\n판매여부가 [판매안함]으로 자동설정됩니다.");
		}
		frm.sellyn[2].checked=true;
		frm.sellyn[0].disabled=true;
		frm.sellyn[1].disabled=true;
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // 옵션사용
        frm.btnoptadd.disabled = false;
        frm.btnoptdel.disabled = false;
	} else {
	    // 옵션없음
	    while (frm.realopt.length > 0) {
	        frm.realopt.options[0] = null;
        }
        frm.btnoptadd.disabled = true;
        frm.btnoptdel.disabled = true;
    }
}

function TnCheckSailYN(frm){
	CheckSailEnDisabled(frm);
    CalcuAuto(frm);
}

function CheckSailEnDisabled(frm){
	if (frm.sailyn[0].checked == true) {
	    // 정상가격
        frm.sellcash.readonly = false;
        frm.margin.readonly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readonly = true;
        frm.sailmargin.readonly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // 세일가격
        frm.sellcash.readonly = true;
        frm.margin.readonly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readonly = false;
        frm.sailmargin.readonly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}

function ClearVal(comp){
    comp.value = "";
}

// ============================================================================
	// 카태고리 선택 팝업
	function popCateSelect(iid){
		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 팝업에서 선택 카테고리 추가
	function addCateItem(lcd,lnm,mcd,mnm,scd,snm,div)
	{
		// 기존에 값에 중복 카테고리 여부 검사 - 플라워 전국배송은 제외함;
		if(tbl_Category.rows.length>0)	{
			if(tbl_Category.rows.length>1)	{
				for(l=0;l<document.all.cate_div.length;l++)	{
				    if (!((document.all.cate_large[l].value=="110")&&(document.all.cate_mid[l].value=="060"))){
    					if((document.all.cate_large[l].value==lcd)&&(document.all.cate_mid[l].value==mcd)) {
    						alert("같은 중분류에 이미 지정된 카테고리가 있습니다.\n기존 카테고리를 삭제하고 다시 선택해주세요.");
    						return;
    					}
    				}
				}
			}
			else {
			    if (!((document.all.cate_large.value=="110")&&(document.all.cate_mid.value=="060"))){
    				if((document.all.cate_large.value==lcd)&&(document.all.cate_mid.value==mcd)) {
    					alert("같은 중분류에 이미 지정된 카테고리가 있습니다.\n※기존 카테고리를 삭제하고 다시 선택해주세요.");
    					return;
    				}
    			}
			}
		}
		
		// 행추가
		var oRow = tbl_Category.insertRow();
		oRow.onmouseover=function(){tbl_Category.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="D") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='cate_div' value='D'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='cate_div' value='A'>";
		}
		oCell2.innerHTML = lnm + " >> " + mnm + " >> " + snm
					+ "<input type='hidden' name='cate_large' value='" + lcd + "'>"
					+ "<input type='hidden' name='cate_mid' value='" + mcd + "'>"
					+ "<input type='hidden' name='cate_small' value='" + scd + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle>";
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

	// 선택 카테고리 삭제
	function delCateItem()
	{
		if(confirm("선택한 카테고리를 삭제하시겠습니까?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}

	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//상품속성 출력
			printItemAttribute();
		}
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
    
    //주문제작 상품인경우.
    if (frm.itemdiv[1].checked){
        if (frm.reqMsg.checked){
            frm.itemdiv[1].value="06";
        }else{
            frm.itemdiv[1].value="16";
        }
    }

    //렌탈 상품인경우.
    if (frm.itemdiv[7].checked){
		frm.reserveItemTp[1].checked = true;
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
			data: "itemid=<%=itemid%>&ifdv="+v,
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
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 800, 1600, '+parseInt(rowLen-1)+')"> (선택,800X1600, Max 800KB,jpg,gif)';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
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
		url: "/admin/itemmaster/safety_api_auth_proc.asp?itemid=<%=itemid%>&issave="+isSave+"&certnum="+certnum+"&safetydiv="+safetydiv+"&statusmode=real",
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
	var del_safetynum = "";
	var del_safetydiv = "";
	
	for(var i in jbSplit){
		if(jbSplit[i] != listnum){
			resultDiv = resultDiv + jbSplit[i] + ",";
			resultNum = resultNum + jbSplitnum[i] + ",";
		}else{
			del_safetynum = jbSplitnum[i];
			del_safetydiv = jbSplit[i];
		}
	}
	
	if(resultDiv.substr(resultDiv.length-1, 1) == ","){
		resultDiv = resultDiv.substr(0, resultDiv.length-1);
		resultNum = resultNum.substr(0, resultNum.length-1);
	}
	$("#real_safetydiv").val(resultDiv);
	$("#real_safetynum").val(resultNum);
	
	$("#l"+listnum+"").remove();
	
	var tmp_num = $("#real_safetynum_delete").val();
	var tmp_div = $("#real_safetydiv_delete").val();
	if(tmp_num == ""){
		$("#real_safetynum_delete").val(del_safetynum);
		$("#real_safetydiv_delete").val(del_safetydiv);
	}else{
		$("#real_safetynum_delete").val(tmp_num + "," + del_safetynum);
		$("#real_safetydiv_delete").val(tmp_div + "," + del_safetydiv);
	}
}

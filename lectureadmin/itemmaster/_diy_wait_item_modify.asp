<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemRegCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<%

'CONST CBASIC_IMG_MAXSIZE = 180   'KB
'CONST CMAIN_IMG_MAXSIZE = 500   'KB
CONST CBASIC_IMG_MAXSIZE = 600   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB


'==============================================================================
Dim oitemdetail,oitemreg,optiontotal,ix
Dim oitemvideo
Dim fingerson : fingerson = "on" '//상품고시용 fingersflag

set oitemdetail = new CWaitItemDetail

oitemdetail.FRectDesignerID = session("ssBctID")
oitemdetail.WaitProductDetail requestCheckvar(request("itemid"),10) '임시등록 데이터 불러오기
oitemdetail.WaitProductDetailOption requestCheckvar(request("itemid"),10) '옵션 2번 넘버,이름 불러오기

'상품이미지
Dim itemaddimage

if (IsNull(oitemdetail.Fimgadd) or (oitemdetail.Fimgadd="")) then oitemdetail.Fimgadd = ",,,,"

itemaddimage = split(oitemdetail.Fimgadd,",")

'==============================================================================
set oitemreg = new CItemReg

'if oitemdetail.FResultCount <> 0 then
'	oitemreg.SearchOptionNameBig left(oitemdetail.FItemList(ix).Fitemoption,2) '옵션 1번 불러오기
'end if

oitemreg.SearchCategoryNameLarge oitemdetail.Flarge '카테고리 1번 불러오기
oitemreg.SearchCategoryNameMid oitemdetail.Flarge,oitemdetail.FMid '카테고리 2번 불러오기
oitemreg.SearchCategoryNameSmall oitemdetail.Flarge,oitemdetail.FMid,oitemdetail.Fsmall '카테고리 3번 불러오기
'==============================================================================
dim imgsubdir

imgsubdir = GetImageSubFolderByItemid(requestCheckvar(request("itemid"),10))
'==============================================================================
Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetAcademyPartnerList

'//동영상
Set oitemvideo = New CItem
oitemvideo.FRectItemId = requestCheckvar(request("itemid"),10)
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetWaitItemContentsVideo
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
$(function(){
	if($("input[name='catecode']").length>1){
		$("#btnAddDispCate").hide();
	}
});

function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// 업체마진자동입력
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var isvatYn, imileage;
	var isellcash, ibuycash, isellvat, ibuyvat, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatYn = frm.vatYn[0].checked;

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

// ============================================================================
// 카테고리등록

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		if($("input[name='catecode']").length>1){
			alert("카테고리는 2개까지 지정 가능합니다.");
			return;
		}

		var designerid = document.all.itemreg.designerid.value;
		if(designerid == ""){
			alert("업체를 선택하세요.");
			return;
		}
		
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/academy/comm/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt+"&isUpche=upche",
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

		if($("input[name='catecode']").length>1){
			$("#btnAddDispCate").hide();
		}

		//상품속성 출력
		//printItemAttribute();
	}
	
	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			if($("input[name='catecode']").length<2){
				$("#btnAddDispCate").show();
			}

			//상품속성 출력
			//printItemAttribute();
		}
	}
// ============================================================================
// 옵션수정
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_itemoption.asp?' + param ,'editItemOption','width=800,height=400,scrollbars=yes,resizable=yes');
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
	popwin = window.open('/academy/comm/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=800,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 옵션을 추가한다
function InsertOption(ft, fv) {
	var frm = document.itemreg;

	//옵션값이 같은것이 있으면 skip ,전용옵션인경우 제외
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

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
function ClearImage(img) {
    var e = eval("itemreg." + img);

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    }

	document.getElementById("div"+img).style.display='none';

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "del";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "del";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "del";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "del";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "del";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "del";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "del";
    }
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

function CheckImage2(img, filesize, imagewidth, imageheight, extname){
    
}

function CheckImage(img, filesize, imagewidth, imageheight, extname)
{
    var preview;
    var e;
    var ext;
    var filename;

    e = eval("itemreg." + img);
    filename = e.value;

    e = eval("itemreg." + img);
    if (e.value == "") { return false; }
    
    if (CheckExtension(filename, extname) != true) {
        alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
        ClearImage(img);
        return false;
    }
    
    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "";
    }

    return true;
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

//임시저장

function IMSISave() {
    if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}
	

    //if (validate(itemreg)==false) {
    //    return;
    //}
    
    //상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
    
    //배송구분 체크 ================================================================ 
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[1].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            itemreg.deliverytype[1].focus();
            return;
        }
    }
    
    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        itemreg.deliverytype[2].focus();
        return;
    }

    //==================================================================================
    
    
    

    if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("마일리지는 판매가보다 클 수 없습니다.");
        itemreg.mileage.focus();
        return;
    }

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}
    
    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }
    
    // 정상가격
	if (confirm("소비자가(" + itemreg.sellcash.value + ")/공급가(" + itemreg.buycash.value + ")가 정확히 입력되었습니까?") == false) {
		itemreg.sellcash.focus();
		return;
    }

    
    if(confirm("상품을 임시저장하시겠습니까??") == true){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }

		itemreg.target = "FrameCKP";
        itemreg.submit();
    }
}

// ============================================================================
// 저장하기
function SubmitSave(istate) {
//alert('현재 서버 작업 중으로 상품 등록/ 변경이 불가합니다.');
//return;

	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}
	
	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[기본] 전시 카테고리를 선택하세요.\n");
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }
    
    //상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
    
    //배송구분 체크 ================================================================ 
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[1].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            itemreg.deliverytype[1].focus();
            return;
        }
    }
    
    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        itemreg.deliverytype[2].focus();
        return;
    }

    //==================================================================================
    
    
    

    if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("마일리지는 판매가보다 클 수 없습니다.");
        itemreg.mileage.focus();
        return;
    }

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}
    
    if (!IsDigit(itemreg.itemWeight.value)){
		alert('무게는 숫자로 입력하세요.');
		itemreg.itemWeight.focus();
		return;
	}
	
    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }
    
    // 정상가격
	if (confirm("소비자가(" + itemreg.sellcash.value + ")/공급가(" + itemreg.buycash.value + ")가 정확히 입력되었습니까?") == false) {
		itemreg.sellcash.focus();
		return;
    }
    if (itemreg.basic.value == "del") {
        alert("기본이미지는 필수입니다.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
                return;
            }
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

//    if (itemreg.imgadd4.value != "") {
//        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgadd5.value != "") {
//        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgmain.value != "") {
//        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
//            return;
//        }
//    }

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

	//안전인증정보
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }

    
    if(confirm("상품을 올리시겠습니까?") == true){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
        if (istate){
            itemreg.CurrState.value = istate;	// 2016/12/08
        }
		itemreg.target = "FrameCKP";
        itemreg.submit();
    }

}

// 재요청하기
function SubmitReSave()
{
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

    if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[기본] 전시 카테고리를 선택하세요.\n");
		return;
	}
	
    if (validate(itemreg)==false) {
        return;
    }
    
    //상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
	
	//배송구분 체크 ================================================================ 
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[1].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            itemreg.deliverytype[1].focus();
            return;
        }
    }
    
    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        itemreg.deliverytype[2].focus();
        return;
    }

    //==================================================================================
    
    // 정상가격
	if (confirm("소비자가(" + itemreg.sellcash.value + ")/공급가(" + itemreg.buycash.value + ")가 정확히 입력되었습니까?") == false) {
		itemreg.sellcash.focus();
		return;
    }

    if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("마일리지는 판매가보다 클 수 없습니다.");
        itemreg.mileage.focus();
        return;
    }

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 5000000){
		alert("판매 가격은 400원 이상 5,000,000만원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }

    if (itemreg.basic.value == "del") {
        alert("기본이미지는 필수입니다.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
                return;
            }
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

//    if (itemreg.imgadd4.value != "") {
//        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgadd5.value != "") {
//        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgmain.value != "") {
//        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
//            return;
//        }
//    }

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

	//안전인증정보
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }
    
    // 재요청 내용 쓰기 팝업
    reMsg = prompt("재등록 요청시 전달할 내용을 간략히 써주세요.","");
    if(reMsg){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
        itemreg.reRegMsg.value = reMsg;
        itemreg.CurrState.value = "5";	// 상태정보를 '등록재요청'으로 수정
        itemreg.target = "FrameCKP"; //추가
        itemreg.submit();
    }
    else {
    	return;
    }

}

function TnCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readOnly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readOnly=true;
		frm.limitsold.style.background='#E6E6E6';
	}
	else {
		// 한정
		if ((frm.optioncnt.value*1) > 0) {
		    // 옵션사용중
		    alert("옵션을 사용할경우 한정수량은 옵션창에서 수정가능합니다.");
		    frm.limityn[0].checked = true;
		    return;
        }

		frm.limitno.readOnly=false;
		frm.limitno.style.background='#FFFFFF';

		frm.limitsold.readOnly=false;
		frm.limitsold.style.background='#FFFFFF';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readOnly=true;
		frm.limitno.style.background='#E6E6E6';

		frm.limitsold.readOnly=true;
		frm.limitsold.style.background='#E6E6E6';
	}
	else {
		// 한정
		if ((frm.optioncnt.value*1) > 0) {
		    // 옵션사용중
		    // alert("한정수량은 옵션창에서 수정가능합니다.");
		    return;
        }

		frm.limitno.readOnly=false;
		frm.limitno.style.background='#FFFFFF';

		frm.limitsold.readOnly=false;
		frm.limitsold.style.background='#FFFFFF';
	}
}

function TnGoClear(frm){
	frm.buyvat.value = "";
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("매입특정 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[0].checked=true;
		}
	}
	//else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	else if(frm.deliverytype[1].checked ){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("매입특정 구분이 매입이나 특정일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[2].checked=true;
		}
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
        frm.sellcash.readOnly = false;
        frm.margin.readOnly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readOnly = true;
        frm.sailmargin.readOnly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // 세일가격
        frm.sellcash.readOnly = true;
        frm.margin.readOnly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readOnly = false;
        frm.sailmargin.readOnly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}
// ============================================================================
// 미리보기
function ViewItemDetail(itemno){
	//window.open('viewDIYitem.asp?itemid='+itemno ,'ViewItemDetail','width=790,height=600,scrollbars=yes,status=no');
	var popwin = window.open('/academy/itemmaster/viewDIYitem/viewDIYitem.asp?itemid='+itemno ,'window1','width=1024,height=960,scrollbars=yes,status=no');
	popwin.focus();
}

function ClearVal(comp){
    comp.value = "";
}

function popWaitItemOptionEdit(iitemid){
    var popwin = window.open('/academy/comm/pop_DIYwaititemoptionedit.asp?itemid=' + iitemid,'popWaitItemOptionEdit','width=790,height=600,scrollbars=yes,status=no');
    popwin.focus();
}

//=====리뉴얼
function chgodr(v){
	if (v == 1){
		$("#customorder").css("display","none");
	}else{
		$("#customorder").css("display","");
	}
}

function chgodr2(v){
	if (v == 1){
		$("#subodr").css("display","none");
	}else{
		$("#subodr").css("display","");
	}
}

// 안전인증정보 선택
function chgSafetyYn(frm) {
	if(frm.safetyYn[0].checked) {
		frm.safetyDiv.disabled=false;
		frm.safetyNum.disabled=false;
	} else {
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
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
			url: "/admin/itemmaster/act_waititemInfoDivForm.asp",
			data: "itemid=<%=request("itemid")%>&ifdv="+v+"&fingerson=on",
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
}


//상품설명이미지추가
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 14){
		alert("이미지는 최대 15개까지 가능합니다.");
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
	c0.innerHTML = '상품상세이미지 #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');CheckImageSize(this);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 1000, 667, '+parseInt(rowLen-1)+')"> (선택,1000X667, Max 600KB,jpg,gif)';
	c1.innerHTML += '<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
}

//상품상세이미지지우기
function ClearImage2(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");CheckImageSize(this);\" class='text' size='"+ fsize +"'>";
	$("#divaddimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "del";
    }else{
    	document.itemreg.addimgdel.value = "del";
    }
}

function requireimg(){
	var frm = document.itemreg;
	if (frm.requireimgchk.checked){
		$("#rmemail").css("display","");
	}else{
		$("#rmemail").css("display","none");
	}
}

function CheckImageSize(obj) {
	var MaxSize=600;
	if((obj.files[0].size/1024) > MaxSize){
		alert("이미지는 600kb 까지 올리실 수 있습니다. (" + ((obj.files[0].size/1024)-MaxSize).toFixed(2) + "kb 초과)" );
		obj.value="";
		return;
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>상품수정</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>등록대기중인 상품을 수정합니다.</b>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- 표 상단바 끝-->


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>기본정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<% if (TRUE) or (session("ssBCTid")="fingertest01") then %>
<form name="itemreg" method="post" action="<%= uploadImgUrl %>/linkweb/academy/items/WaitDIYItemModify_Process.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0px;">
<% else %>
<form name="itemreg" method="post" action="<%= UploadImgFingers %>/linkweb/items/WaitDIYItemModify_Process.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0px;">
<% end if %>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemid" value="<%= request("itemid") %>">
<input type="hidden" name="designerid" value="<%= oitemdetail.FRectDesignerID %>">
<input type="hidden" name="defultmargine" value="<% =npartner.FPartnerList(0).Fdiy_margin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FPartnerList(0).Fmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FPartnerList(0).FdefaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FPartnerList(0).FdefaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FPartnerList(0).FdefaultDeliveryType %>">

<input type="hidden" name="isusing" value="N">
<input type="hidden" name="dispyn" value="N">
<input type="hidden" name="sellyn" value="N">
<input type="hidden" name="reRegMsg" value="">
<input type="hidden" name="CurrState" value="<%=oitemdetail.FCurrState%>">

<input type="hidden" name="cd1" value="<%= oitemdetail.Flarge %>">
<input type="hidden" name="cd2" value="<%= oitemdetail.Fmid %>">
<input type="hidden" name="cd3" value="<%= oitemdetail.Fsmall %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= request("itemid") %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="미리보기" onclick="ViewItemDetail('<%= request("itemid") %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">카테고리 구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategoryWait(trim(request("itemid")))%></td>
			<td valign="bottom"><input id="btnAddDispCate" type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" <% if oitemdetail.Fitemdiv ="01" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">일반상품
	  <input type="radio" name="itemdiv" value="06" <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","checked","")%> onclick="checkItemDiv(this);chgodr(2);">주문제작상품
      <input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitemdetail.Fitemdiv="06","checked","")%> <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문제작 메세지가 필요한 경우)</font>
	  <input type="checkbox" name="requireimgchk" value="Y" <%=chkIIF(oitemdetail.Frequirechk="Y","checked","")%> onClick="requireimg();">주문제작 이미지 필요
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" <% if oitemdetail.Fitemdiv ="20" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">추가전용상품 -->
<!--       <font color="red">(상품목록에서는 제외, 추가옵션에서만 보여짐)</font><br> -->
  	</td>
  </tr>
   <!-- 주문 제작 이메일 -->
  <tr id="rmemail" style="display:<%=chkiif(oitemdetail.Frequirechk="Y","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문제작 이메일 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="requireMakeEmail" value="<%=oitemdetail.FrequireEmail%>" size="50" maxlength="100"> (ex)작가님의 메일 주소)
  	</td>
  </tr>
  <!-- 주문 제작 이메일 -->
  <tr align="left" id="customorder" style="display:<%=chkiif(oitemdetail.Fitemdiv="06" Or oitemdetail.Fitemdiv="16","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문제작 추가옵션</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="cstodr" value="1" onclick="chgodr2(1)" <%=chkiif(oitemdetail.Fcstodr="1","checked","")%>>즉시발송
      <input type="radio" name="cstodr" value="2" onclick="chgodr2(2)" <%=chkiif(oitemdetail.Fcstodr="2","checked","")%>>제작후 발송<br>
	  <div id="subodr" style="display:<%=chkiif(oitemdetail.Fcstodr="2","block","none")%>;">
		제작후 발송 기간 <input type="text" name="requireMakeDay" value="<%=oitemdetail.FrequireMakeDay%>" size="3" maxlength="2">일<br>
		&lt--특이사항을 입력 해주세요--&gt;<br><textarea name="requirecontents" rows="5" cols="80"><%=oitemdetail.Frequirecontents%></textarea>
	  </div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="50" id="[on,off,off,off][상품명]" value="<%= oitemdetail.Fitemname %>">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][상품재질]" value="<%= oitemdetail.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsize" maxlength="64" size="50" id="[on,off,off,off][상품사이즈]" value="<%= oitemdetail.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][상품무게]" value="<%= oitemdetail.FitemWeight %>">g&nbsp;(무게는 g단위로 입력)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][원산지]" value="<%= oitemdetail.Fsourcearea %>">&nbsp;(ex:한국,중국,중국OEM,일본...)
      <br>( 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][제조사]" value="<%= oitemdetail.Fmakername %>">&nbsp;(제조업체명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][검색키워드]" value="<%= oitemdetail.Fkeywords %>">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" value="<%= oitemdetail.Fupchemanagecode %>" size="20" maxlength="32" id="[off,off,off,off][업체상품코드]">
  	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="radio" name="usinghtml" value="N" <% if oitemdetail.Fusinghtml = "N" then response.write "checked" %>>일반TEXT -->
<!--       <input type="radio" name="usinghtml" value="H" <% if oitemdetail.Fusinghtml = "H" then response.write "checked" %>>TEXT+HTML -->
<!--       <input type="radio" name="usinghtml" value="Y" <% if oitemdetail.Fusinghtml = "Y" then response.write "checked" %>>HTML사용 -->
<!--       <br> -->
<!--       <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][아이템설명]"><%= oitemdetail.Fitemcontent %></textarea> -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :<br/>[배송비 안내]</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][유의사항]"><%= oitemdetail.Fordercomment %></textarea><br>
      <font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">교환 / 환불 정책</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][환불정책]"><%=oitemdetail.Frefundpolicy%></textarea><br>
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">업체코멘트 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][업체코멘트]" value="<%= oitemdetail.Fdesignercomment %>"><br> -->
<!--       상품에관한 스토리나 재미난 이야기를 적어주세요... -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 동영상 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][아이템동영상]"><%= db2html(oitemvideo.FOneItem.FvideoFullUrl) %></textarea>
		<br>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)
	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>가격정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <input type="hidden" name="sellvat" value="<%= oitemdetail.Fsellvat %>">
  <input type="hidden" name="buyvat" value="<%= oitemdetail.Fbuyvat %>">
  <tr align="left">
    <td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatYn" value="Y" onclick="TnGoClear(this.form);" <% if oitemdetail.FvatYn = "Y" then response.write "checked" %>>과세
      <input type="radio" name="vatYn" value="N" onclick="TnGoClear(this.form);" <% if oitemdetail.FvatYn = "N" then response.write "checked" %>>면세
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">공급 마진 :</td>
  	<td bgcolor="#FFFFFF">
  	    <% if (oitemdetail.Fsellcash=0) then ''2016/12/08 추가 임시저장상품 판매가 0 있음. %>
  	    <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][마진]" value="<%= npartner.FPartnerList(0).Fdiy_margin %>" readonly style="background-color:#E6E6E6;">%
  	    <% else %>
        <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][마진]" value="<%= oitemdetail.FMargin %>" readonly style="background-color:#E6E6E6;">%
        <% end if %>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가) :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" size="12" id="[on,on,off,off][소비자가]" onKeyUp="CalcuAuto(itemreg);" maxlength="7" value="<%= oitemdetail.Fsellcash %>" >원
      <!--<input type="button" value="공급가 자동계산" onclick="CalcuAuto(itemreg);">-->
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">공급가 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][공급가]" value="<%= oitemdetail.Fbuycash %>" >원
      (<b>부가세 포함가</b>)
  	</td>
  </tr>
  <tr>
  	<td bgcolor="#DDDDFF"></td>
  	<td bgcolor="#FFFFFF" colspan="3">
      - 공급가는 <b>부가세 포함가</b>입니다.<br>
      - 소비자가(할인가)와 마진(할인마진)을 입력하고 [공급가자동계산] 버튼을 누르면 공급가와 마일리지가 자동계산됩니다.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitemdetail.Fdeilverytype = "2" then response.write "checked" %>>업체(무료)배송&nbsp;
      <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitemdetail.Fdeilverytype = "9" then response.write "checked" %>>업체조건배송(개별 배송비부과)&nbsp;
  	  <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitemdetail.Fdeilverytype = "7" then response.write "checked" %>>업체착불배송
  	</td>
  </tr>
  <input type="hidden" name="mileage" id="[on,off,off,off][마일리지]" value="<%= oitemdetail.Fmileage %>">
  <input type="hidden" name="mwdiv" value="U">
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>옵션정보/한정정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">옵션구분 :</td>
  	<input type="hidden" name="optioncnt" value="<%= oitemdetail.Foptioncnt %>">
  	<td width="35%" bgcolor="#FFFFFF">
      <% if oitemdetail.Foptioncnt < 1 then %>
      옵션사용안함
      <% else %>
      옵션사용중(<%= oitemdetail.Foptioncnt %>개)
      <% end if %>
      &nbsp;&nbsp;<input type="button" class="button" value="옵션수정" onClick="popWaitItemOptionEdit('<%= oitemdetail.FWaitItemID %>');">
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">한정판매구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="limityn" value="N" onClick="TnCheckLimitYN(itemreg)" <% if oitemdetail.Flimityn = "N" then response.write "checked" %>>비한정판매&nbsp;&nbsp;
  	  <input type="radio" name="limityn" value="Y" onClick="TnCheckLimitYN(itemreg)" <% if oitemdetail.Flimityn = "Y" then response.write "checked" %>>한정판매
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">한정수량 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="limitno" maxlength="32" size="8" readOnly style="background-color:#E6E6E6;" id="[off,on,off,off][한정수량]" value="<%= oitemdetail.Flimitno %>">
      <input type="hidden" name="limitsold" value="0">
      <input type="hidden" name="limitstock" value="<%= oitemdetail.Flimitno %>">
  	</td>
  </tr>
  <tr align="left">
  	<td width="15%" bgcolor="#DDDDFF">옵션설정 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <br>
      - 옵션정보는 옵션창에서 수정가능합니다.<br>
      - 옵션은 추가는 가능하지만 삭제는 불가능합니다. 정확히 입력하세요.<br>
      - 한정수량은 옵션이 있을 경우, 옵션창에서 수정이 가능합니다.(위의 정보는 부정확할수 있습니다.)<br>
      <br>
  	</td>
  </tr>
</table>


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>이미지정보
          <br>- 텐바이텐에서 이미지를 등록할 경우 따로 입력하지 마시기 바랍니다.
          <br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
          <br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
          <br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (oitemdetail.Fimgbasic <> "") then %>
      <div id="divimgbasic" style="display:block;">
      <table id="imgbasic"  style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/basic/<%= imgsubdir  %>/<%= oitemdetail.Fimgbasic %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgbasic" style="display:none;">
      <table id="imgbasic" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);" size="40"> (<font color=red>필수</font>,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">아이콘이미지<br>(자동생성) :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	실서버 등록시 자동생성
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (itemaddimage(0) <> "") then %>
      <div id="divimgadd1" style="display:block;">
      <table id="imgadd1" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/add1/<%= imgsubdir  %>/<%= itemaddimage(0) %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd1" style="display:none;">
      <table id="imgadd1" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%"" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (선택,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (itemaddimage(1) <> "") then %>
      <div id="divimgadd2" style="display:block;">
      <table id="imgadd2" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/add2/<%= imgsubdir  %>/<%= itemaddimage(1) %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd2" style="display:none;">
      <table id="imgadd2" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (선택,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (itemaddimage(2) <> "") then %>
      <div id="divimgadd3" style="display:block;">
      <table id="imgadd3" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="<%=UploadImgFingers%>/diyItem/waitimage/add3/<%= imgsubdir  %>/<%= itemaddimage(2) %>">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd3" style="display:none;">
      <table id="imgadd3" style="background-repeat: no-repeat;width:400px;height:300px;background-size:100%" background="">
        <tr>
          <td></td>
        </tr>
      </table>
      </div>
<% end if %>
      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (선택,1000X667,MAX <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3">
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- <% if (itemaddimage(3) <> "") then %> -->
<!--       <div id="divimgadd4" style="display:block;"> -->
<!--       <table id="imgadd4" width="400" height="400" style="background-repeat: no-repeat" background="<%=UploadImgFingers%>/diyItem/waitimage/add4/<%= imgsubdir  %>/<%= itemaddimage(3) %>"> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgadd4" style="display:none;"> -->
<!--       <table id="imgadd4" width="400" height="400" style="background-repeat: no-repeat" background=""> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--       <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> (선택,1000X667,jpg,gif) -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- <% if (itemaddimage(4) <> "") then %> -->
<!--       <div id="divimgadd5" style="display:block;"> -->
<!--       <table id="imgadd5" width="400" height="400" style="background-repeat: no-repeat" background="<%=UploadImgFingers%>/diyItem/waitimage/add5/<%= imgsubdir  %>/<%= itemaddimage(4) %>"> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgadd5" style="display:none;"> -->
<!--       <table id="imgadd5" width="400" height="400" style="background-repeat: no-repeat" background=""> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--       <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> (선택,1000X667,jpg,gif) -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- <% if (oitemdetail.Fimgmain <> "") then %> -->
<!--       <div id="divimgmain" style="display:block;"> -->
<!--       <table id="imgmain" width="400" height="400" style="background-repeat: no-repeat" background="<%=UploadImgFingers%>/diyItem/waitimage/main/<%= imgsubdir  %>/<%= oitemdetail.Fimgmain %>"> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgmain" style="display:none;"> -->
<!--       <table id="imgmain" width="400" height="400" style="background-repeat: no-repeat" background=""> -->
<!--         <tr> -->
<!--           <td></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--       <input type="file" name="imgmain" onchange="CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');" size="40"> (선택,600X2000 Max <%= CMAIN_IMG_MAXSIZE %>Kb 이하,jpg) -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> -->
<!--   	</td> -->
<!--   </tr> -->

</table>

<!-- 품목상세정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">품목상세정보 &nbsp;<font color=gray>상품정보제공고시 관련 법안 추진에 따라 아래 내용을 정확히 입력해주시기 바랍니다.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목선택 :</td>
	<td bgcolor="#FFFFFF">
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::상품품목::</option>
		<option value="01">의류</option>
		<option value="02">구두/신발</option>
		<option value="03">가방</option>
		<option value="04">패션잡화(모자/벨트/액세서리)</option>
		<option value="05">침구류/커튼</option>
		<option value="06">가구(침대/소파/싱크대/DIY제품)</option>
		<!--
		<option value="07">영상가전(TV류)</option>
		<option value="08">가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option>
		<option value="09">계절가전(에어컨/온풍기)</option>
		<option value="10">사무용기기(컴퓨터/노트북/프린터)</option>
		<option value="11">광학기기(디지털카메라/캠코더)</option>
		<option value="12">소형전자(MP3/전자사전 등)</option>
		<option value="14">내비게이션</option>
		-->
		<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
		<!-- <option value="16">의료기기</option>-->
		<option value="17">주방용품</option>
		<option value="18">화장품</option>
		<option value="19">귀금속/보석/시계류</option>
		<option value="20">식품(농수산물)</option>
		<option value="21">가공식품</option>
		<option value="22">건강기능식품/체중조절식품</option>
		<option value="23">영유아용품</option>
		<option value="24">악기</option>
		<option value="25">스포츠용품</option>
		<option value="26">서적</option>
		<!--<option value="27">호텔/펜션예약</option>
		<option value="28">여행상품</option>
		<option value="29">항공권</option>
		-->
		<option value="35">기타</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitemdetail.FinfoDiv%>";
		//setTimeout(function(){
			chgInfoDiv(<%=oitemdetail.FinfoDiv%>);
		//},0);
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="") Then
			Server.Execute("/admin/itemmaster/act_waititemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<!-- <tr align="left" id="lyItemSrc" _style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...) -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr align="left" id="lyItemSize" _style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsize %>">&nbsp;(ex:7.5x15(cm)) -->
<!-- 	</td> -->
<!-- </tr> -->
</table>
<!-- 안전인증정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">안전인증정보</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">안전인증대상 :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitemdetail.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> 대상</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> 대상아님</label> /
		<select name="safetyDiv" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::안전인증구분::</option>
		<option value="10" <%=chkIIF(oitemdetail.FsafetyDiv="10","selected","")%>>국가통합인증(KC마크)</option>
		<option value="20" <%=chkIIF(oitemdetail.FsafetyDiv="20","selected","")%>>전기용품 안전인증</option>
		<option value="30" <%=chkIIF(oitemdetail.FsafetyDiv="30","selected","")%>>KPS 안전인증 표시</option>
		<option value="40" <%=chkIIF(oitemdetail.FsafetyDiv="40","selected","")%>>KPS 자율안전 확인 표시</option>
		<option value="50" <%=chkIIF(oitemdetail.FsafetyDiv="50","selected","")%>>KPS 어린이 보호포장 표시</option>
		</select>
		인증번호 <input type="text" name="safetyNum" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" class="text" value="<%=oitemdetail.FsafetyNum%>" />
		
		<font color="darkred">유아용품이나 전기용품일 경우 필수 입력</font>
	</td>
</tr>
</table>

<%
	Dim cImg, k, vArr, j, txtBuf
	set cImg = new CItemAddImage
	cImg.FRectItemID = request("itemid")
	vArr = cImg.GetWaitAddImageList
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
	<% If isArray(vArr) Then
			If vArr(3,UBound(vArr,2)) > 0 Then
			For k = 1 To vArr(3,UBound(vArr,2))
	%>
			  <tr align="left">
			  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #<%= (k) %> :</td>
			  	<td bgcolor="#FFFFFF">
		  		<%
		  		If cImg.IsImgExist(vArr,k) Then
		    		For j = 0 To UBound(vArr,2)
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & UploadImgFingers & "/diyItem/waitcontentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ width=""400""></div>"
							Exit For
		    			End If
		    		Next
				Else
					Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
				End If
				%>
			      <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40, <%= (k-1) %>);CheckImageSize(this);" class="text" size="40">
			      <input type="button" value="#<%= (k) %> 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname<%=CHKIIF(vArr(3,UBound(vArr,2))=1,"","["&(k-1)&"]")%>,40, 1000, 667, <%= (k-1) %>)"> (선택,1000X667, Max 800KB,jpg,gif)
				  <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/>
				  <%
				  txtBuf=""
				  For j = 0 To UBound(vArr,2)
	    			If CStr(vArr(3,j)) = CStr(k) Then
	    			    txtBuf = db2html(vArr(5,j))
						Exit For
	    			End If
	    		  Next
	    		  %>
				  <textarea name="addimgtext" cols="70" rows="5"><%=txtBuf%></textarea>
			      <input type="hidden" name="addimggubun" value="<%= (k) %>">
			      <input type="hidden" name="addimgdel" value="">
			  	</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #1 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname1" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,0);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 667, 0)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="1">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #2 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname2" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,1);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 667, 1)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="2">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #3 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname3" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,2);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 667, 2)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="3">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #4 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname4" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,3);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#4 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[3],40, 1000, 667, 3)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="4">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #5 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname5" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,4);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#5 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[4],40, 1000, 667, 4)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="5">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #6 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname6" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,5);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#6 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[5],40, 1000, 667, 5)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="6">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #7 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname7" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,6);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#7 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[6],40, 1000, 667, 6)"> (선택,1000X667, ,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="7">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
	<%
	   End IF %>
</table>
<%	set cImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF" height="30">
      <input type="button" value="상품상세이미지추가" class="button" onClick="InsertImageUp()">
      <font color="red">* 업로드가 된 이미지가 제대로 안나오면 새로고침(CTRL + F5(콘트롤 F5 버튼))을 해주세요.</font>
  	</td>
  </tr>
</table>

</form>


    <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"8" then %>
    <table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
      <tr align="left">
      	<td height="30" width="15%" bgcolor="#F8DDFF">등록보류 사유 :</td>
      	<td bgcolor="#FFFFFF" colspan="3">
      		<%=oitemdetail.Frejectmsg & " [" & oitemdetail.FrejectDate & "]"%>
      	</td>
      </tr>
    </table>
    <% end if %>
    <% if oitemdetail.FCurrState="5" and Not(isNull(oitemdetail.FreRegMsg)) then %>
    <table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
      <tr align="left">
      	<td height="30" width="15%" bgcolor="#F8DDFF">재요청 메시지 :</td>
      	<td bgcolor="#FFFFFF" colspan="3">
      		<%=oitemdetail.FreRegMsg & " [" & oitemdetail.FreRegDate & "]"%>
      	</td>
      </tr>
    </table>
    <% end if %>
    
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <% if (oitemdetail.FCurrState="8") then %>
          <input type="button" value="임시저장" onClick="IMSISave()">
          <input type="button" value="등록요청" onClick="SubmitSave('1')">
          <% else %>
              <% if oitemdetail.FCurrState="1" or oitemdetail.FCurrState="5" then %>
              <input type="button" value="저장하기" onClick="SubmitSave('<%=oitemdetail.FCurrState%>')">
              <% else %>
              <input type="button" value="재요청하기" onClick="SubmitReSave()">
              <% end if %>
          <% end if %>
          <input type="button" value="목록으로" onClick="history.back()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<%
	set oitemreg = Nothing
	Set oitemvideo = Nothing
	set oitemdetail = Nothing
	set npartner = Nothing
%>
<p>
<script>
// 한정
TnSilentCheckLimitYN(itemreg);
// 세일
// TnCheckSailYN(itemreg);
</script>
<% if (application("Svr_Info")	= "Dev") or (session("ssBctId")="fingertest01") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
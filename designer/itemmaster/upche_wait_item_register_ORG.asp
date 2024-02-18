<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->

<%
CONST CBASIC_IMG_MAXSIZE = 180   'KB
CONST CMAIN_IMG_MAXSIZE = 500   'KB

Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetOnePartnerNUser


dim i,j,k 
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript" >
function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "UseTemplate", "width=700, height=450, scrollbars=yes, resizable=yes");
}

// 업체마진자동입력
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;
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

	//카테고리에 따른 Enable 설정 -플라워
	EnDisableFlowerShop();
	
}

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

//2008년 용
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

// 이미지표시
function ClearImage(img) {
    var e = eval("itemreg." + img);

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');\" size='40'>";
    }

    e = eval("document.all.div" + img);
    e.style.display = "none";

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

function ShowImage(img) {
	var e = eval("document.all.div" + img);
    e.style.display = "";

    var filename;
    e = eval("itemreg." + img );
    filename = e.value;

	eval("document.all." + img + "_img").src=filename;
    //document.getElementById(img).src=filename;


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

function pause(numberMillis) {
     var now = new Date();
     var exitTime = now.getTime() + numberMillis;


     while (true) {
          now = new Date();
          if (now.getTime() > exitTime)
              return;
     }
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

	ShowImage(img);

    if (CheckExtension(filename, extname) != true) {
        alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
        ClearImage(img);
        return false;
    }
    
    try{
        // iframe 속에 이미지를 넣고, 사이즈/크기를 체크한다.
        document.imgpreview.document.getElementById("imgpreview").src = filename;
        // 시간차이로 이미지 로딩전에 넘어갈수 있음
        preview = document.imgpreview.document.getElementById("imgpreview");
    
        if(preview.fileSize > (filesize * 1024)){
            alert("파일사이즈는 " + filesize + "Kbyte를 넘기실 수 없습니다.");
            ClearImage(img);
            return false;
        }
    
        if(preview.width > (imagewidth)){
            alert("가로폭은 " + imagewidth + "픽셀을 넘기실 수 없습니다.");
            ClearImage(img);
            return false;
        }
    
        if(preview.height > (imageheight)){
            alert("세로폭은 " + imageheight + "픽셀을 넘기실 수 없습니다.");
            ClearImage(img);
            return false;
        }
    }catch(ex){
        // nothing;
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


// 저장하기
function SubmitSave() {
//alert('현재 서버 작업 중으로 상품 등록/ 변경이 불가합니다.');
//return;
	
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
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
	
	if (itemreg.itemsize.value!=''){
		if (itemreg.unit.value!=''){
			itemreg.itemsize.value=itemreg.itemsize.value + '(' + itemreg.unit.value + ')';
		}
	}
	
    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            itemreg.deliverytype[3].focus();
            return;
        }
    }
    
    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[4].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        itemreg.deliverytype[4].focus();
        return;
    }
    
    //배송구분 업체이나 매입구분이 업체가 아닌것.
    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            itemreg.deliverytype[1].focus();
            return;
        }
    }
    
    //매입구분이 업체이나 배송구분이 업체가 아닌것.
    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            itemreg.deliverytype[0].focus();
            return;
        }
    }
    
	if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000원 미만으로 등록 가능합니다.");
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

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }

	
	
    if (itemreg.imgbasic.value == "") {
        // alert("기본이미지는 필수입니다.");
        // return;
    } else {
        if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
            return;
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

    if (itemreg.imgadd4.value != "") {
        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgmain.value != "") {
        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
            return;
        }
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
            
            if ((fnTrim(itemreg.optionTypename1.value)==fnTrim(itemreg.optionTypename2.value))||(fnTrim(itemreg.optionTypename2.value)==fnTrim(itemreg.optionTypename3.value))||(fnTrim(itemreg.optionTypename1.value)==fnTrim(itemreg.optionTypename3.value))){
                alert('이중옵션은 옵션 구분명을 서로 다르게 지정해야 합니다.');
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
            
                for (var i=0;i<itemreg.optionName3.length;i++){
                    if (itemreg.optionName3[i].value.length>0) chkCnt++;
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
    
    // 정상가격
	if (confirm("소비자가(" + itemreg.sellcash.value + ")/공급가(" + itemreg.buycash.value + ")가 정확히 입력되었습니까?") == false) {
		itemreg.sellcash.focus();
		return;
    }
    
    if(confirm("상품을 올리시겠습니까? \n담당MD 승인후 반영 됩니다.") == true){
        itemreg.itemoptioncode2.value = optionv;
        itemreg.itemoptioncode3.value = optiont;

		itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.deliverytype[3].disabled=false;
        itemreg.deliverytype[4].disabled=false;
        
        itemreg.submit();
    }

}

function TnCheckUpcheYN(frm){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// 기본체크
		// 배송구분 지정(텐바이텐)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=true;  //텐바이텐 무료배송은 체크 할 수 없음.
		frm.deliverytype[3].disabled=true;  //업체개별배송(9)
		frm.deliverytype[4].disabled=true;  //업체착불배송(7) : 업체에서 설정불가
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
        frm.deliverytype[3].disabled=false; //업체개별배송(9)
        frm.deliverytype[4].disabled=false;  //업체착불배송(7) : 업체에서 설정불가
	}
}

function TnGoClear(frm){
	frm.sellvat.value = "";
	frm.buycash.value = "";
	frm.buyvat.value = "";
	frm.mileage.value = "";
}

function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("매입특정 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[0].checked=true;
		}
	}
	//else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	else if(frm.deliverytype[1].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("매입특정 구분이 매입이나 특정일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[2].checked=true;
		}
	}
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // 옵션사용
        
        opttype.style.display="inline";
        
        if (frm.optlevel[1].checked==true){
            optlist.style.display ="none";
            optlist2.style.display ="inline";
        }else{
            optlist.style.display="inline";
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

// 이미지 알림창
function PopImageInformation(){
	window.open("itemreg_info_win.asp","PopImageInformation","width=920,height=600,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no");
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

function ClearVal(comp){
    comp.value = "";
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>기본정보</strong>
        </td>
        <td align="right">
          <input type="button" value="기본틀생성" onClick="UseTemplate();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/WaitUpcheItemRegister_Process.asp" onsubmit="return false;" enctype="multipart/form-data">
  <input type="hidden" name="itemoptioncode2">
  <input type="hidden" name="itemoptioncode3">
  <input type="hidden" name="designerid" value="<%= session("ssBctID") %>">
  <input type="hidden" name="defultmargine" value="<%= npartner.FOneItem.Fdefaultmargine %>">
  <input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FOneItem.Fmaeipdiv %>">
  <input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FOneItem.FdefaultFreeBeasongLimit %>">
  <input type="hidden" name="defaultDeliverPay" value="<%= npartner.FOneItem.FdefaultDeliverPay %>">
  <input type="hidden" name="defaultDeliveryType" value="<%= npartner.FOneItem.FdefaultDeliveryType %>">
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">카테고리 구분 :</td>
    <input type="hidden" name="cd1" value="">
    <input type="hidden" name="cd2" value="">
    <input type="hidden" name="cd3" value="">
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="cd1_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">

      <input type="button" value="카테고리 선택" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" checked>일반상품
      <input type="radio" name="itemdiv" value="06">주문제작상품
      <font color="red">(주문제작 메세지가 필요한 경우, 예를들어 고객요청 이니셜을 넣어줄경우)</font>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="50" id="[on,off,off,off][상품명]">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][상품재질]">&nbsp;(ex:플라스틱,비즈,금,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsize" maxlength="64" size="50" id="[on,off,off,off][상품사이즈]">
      	<select name="unit">
			<option value="">직접입력</option>
			<option value="mm">mm</option>
			<option value="cm" selected>cm</option>
			<option value="m²">m²</option>
			<option value="km">km</option>
			<option value="m²">m²</option>
			<option value="km²">km²</option>
			<option value="ha">ha</option>
			<option value="m³">m³</option>
			<option value="cm³">cm³</option>
			<option value="L">L</option>
			<option value="g">g</option>
			<option value="Kg">Kg</option>
			<option value="t">t</option>
		</select>
      &nbsp;(ex:7.5x15(cm))
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][원산지]">&nbsp;(ex:한국,중국,중국OEM,일본...)
      <br>( 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][제조사]">&nbsp;(제조업체명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][검색키워드]">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" value="" size="20" maxlength="32" id="[off,off,off,off][업체상품코드]">
  	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
  	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>가격정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

  <tr align="left">
    <td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatinclude" value="Y" checked onclick="TnGoClear(this.form);">과세
      <input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);">면세
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본 공급 마진 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][마진]" value="<% =npartner.FOneItem.Fdefaultmargine %>" readonly style="background-color:#E6E6E6;">%
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가) :</td>
  	<input type="hidden" name="sellvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="sellcash" size="12" id="[on,on,off,off][소비자가]" onKeyUp="CalcuAuto(itemreg);" maxlength="7">원
      <!--<input type="button" value="공급가 자동계산" onclick="CalcuAuto(itemreg);">-->
  	</td>
  	<td width="15%" bgcolor="#DDDDFF">공급가 :</td>
  	<input type="hidden" name="buyvat">
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][공급가]" >원
      (<b>부가세 포함가</b>)
  	</td>
  </tr>
  <input type="hidden" name="mileage" value="0">
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>판매정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">매입특정구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="mwdiv" value="M" checked onclick="TnCheckUpcheYN(this.form);">매입
      <input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);">특정
      <input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);">업체배송
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
  	<td width="85%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverytype" value="1" checked  onclick="TnCheckUpcheDeliverYN(this.form);">텐바이텐배송&nbsp;
      <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);">업체(무료)배송&nbsp;
      <input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);">텐바이텐무료배송&nbsp;
      <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);">업체조건배송(개별 배송비부과)&nbsp;
      <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);">업체착불배송
     
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">배송지역 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="deliverarea" value="" checked>전국택배(일반)&nbsp;
      <input type="radio" name="deliverarea" value="C" disabled >수도권배송&nbsp;
      <input type="radio" name="deliverarea" value="S" disabled >서울배송&nbsp;
      <input type="checkbox" name="deliverfixday" value="C" disabled >플라워지정일
      <br>&nbsp;<font color="red">(플라워 상품인 경우만 수도권배송, 서울배송, 플라워지정일 옵션이 사용가능합니다.)</font>
  	</td>
  </tr>
  <!-- 사용안함
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">판매종료(예정)일 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sellEndDate" id="[off,off,off,off][판매종료(예정)일]"  size="10" value="" > 
  	    <a href="javascript:calendarOpen(itemreg.sellEndDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
  	    <a href="javascript:ClearVal(itemreg.sellEndDate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
  	</td>
  </tr>
  -->
  <input type="hidden" name="pojangok" value="N">
  <input type="hidden" name="sellyn" value="N">
  <input type="hidden" name="dispyn" value="N">
  <input type="hidden" name="isusing" value="Y">
</table>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>상품설명</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 설명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="usinghtml" value="N" checked >일반TEXT
      <input type="radio" name="usinghtml" value="H">TEXT+HTML
      <input type="radio" name="usinghtml" value="Y">HTML사용
      <br>
      <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][아이템설명]"></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][유의사항]"></textarea><br>
      <font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체코멘트 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][업체코멘트]"><br>
      상품에관한 스토리나 재미난 이야기를 적어주세요...
  	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>옵션정보/한정정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">옵션구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);">옵션사용함&nbsp;&nbsp;
      <input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>옵션사용안함
  	</td>
  </tr>

  <tr align="left">
  	<td width="15%" bgcolor="#DDDDFF" rowspan="2">한정판매구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="limityn" value="N" onClick="this.form.limitno.readonly=true; this.form.limitno.value=''; this.form.limitno.style.background='#E6E6E6'; this.form.limitno.readOnly=true" checked>비한정판매&nbsp;&nbsp;
  	  <input type="radio" name="limityn" value="Y" onClick="this.form.limitno.readonly=false;this.form.limitno.style.background='#FFFFFF'; this.form.limitno.readOnly=false">한정판매
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">한정수량 :</td>
  	<td width="35%" bgcolor="#FFFFFF" >
      <input type="text" name="limitno" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" id="[off,on,off,off][한정수량]">(개)
  	</td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** 옵션이 있는경우 옵션별로 한정수량이 일괄 설정됩니다.(개별설정은 등록후 수정가능)</font></td>
  </tr>
</table>

<div id="opttype" style="display:none">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr height="40">
    <td width="15%" bgcolor="#DDDDFF">옵션 구분  :</td>
    <td width="85%" bgcolor="#FFFFFF">
        <input type="radio" name="optlevel" value="1" onClick="TnCheckOptionYN(this.form);" checked >단일 옵션 (옵션 구분 1개)
        <input type="radio" name="optlevel" value="2" onClick="TnCheckOptionYN(this.form);">이중 옵션 (옵션 구분 최대 3개)
    </td>
  </tr>
</table>
</div>

<div id="optlist" style="display:none" >
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr>
    <td width="15%" bgcolor="#DDDDFF">옵션 설정 :</td>
  	<td width="85%" bgcolor="#FFFFFF" colspan="3">
      	<table width="500" border="0" cellspacing="0" cellpadding="0" class="a" >
      	<tr>
      	    <td width="100">옵션 구분명 :</td>
      	    <td width="400"><input type="text" name="optTypeNm" value="" size="20" maxlength="20" id="[off,off,off,off][옵션 구분명]"></td>
      	</tr>
      	<tr>
      	    <td colspan="2">
              <select multiple name="realopt" size="10" style="width:400">
              </select>
              <br>
            </td>
        </tr>
        <tr>
            <td colspan="2">
              <input type="button" value="기본옵션추가" name="btnoptadd" onclick="popNormalOptionAdd();" >
              <input type="button" value="전용옵션추가" name="btnetcoptadd" onclick="popEtcOptionAdd();">
              <input type="button" value="선택옵션삭제" name="btnoptdel" onclick="delItemOptionAdd()" >
              <br><br>
              - 기본옵션추가 : 색상, 사이즈등 기본적으로 정의된 옵션을 추가 하실 수 있습니다.<br>
              - 전용옵션추가 : 기본옵션에 정의되지 않은 상품전용옵션을 지정하실 수 있습니다.<br>
              - 선택옵션삭제 : 선택된 옵션을 삭제합니다.<br>
              - 주의사항 : 한번 저장된 옵션은 <font color=red>삭제가 불가능</font>합니다.<br>
              <br>
            </td>
        </tr>
        </table>
  	</td>
  </tr>
</table>
</div>


<%
dim iMaxCols : iMaxCols = 3
dim iMaxRows : iMaxRows = 9
%>
<div id="optlist2" style="display:none">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr>
    <td width="15%" bgcolor="#DDDDFF">옵션설정 :</td>
    <td width="85%" bgcolor="#FFFFFF" colspan="3">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
        <tr align="center"  bgcolor="#DDDDFF">
            <td width="100">옵션구분명</td>
            <% for j=0 to iMaxCols-1 %>
            <td>
                <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20" id="[off,off,off,off][옵션 구분명<%= j %>]">
            </td>
            <% Next %>
            <td width="80">(등록예시)<br>색상</td>
            <td width="80">(등록예시)<br>사이즈</td>
        </tr>
        <tr height="2" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <% for i=0 to iMaxRows-1 %>
<tr align="center"  bgcolor="#FFFFFF">
<td>옵션명 <%= i+1 %></td>
<% for j=0 to iMaxCols-1 %>
<td><input type="hidden" name="itemoption<%= j+1 %>" value="">
    <input type="text" name="optionName<%= j+1 %>" size="18" maxlength="20" id="[off,off,off,off][옵션명<%= i %><%= j %>]"></td>
<% next %>
<td>
    <% if i=0 then %>
    빨강
    <% elseif i=1 then %>
    파랑
    <% elseif i=2 then %>
    노랑
    <% elseif i=3 then %>
    베이지
    <% end if %>
</td>
<td>
    <% if i=0 then %>
    XL
    <% elseif i=1 then %>
    L
    <% elseif i=2 then %>
    S
    <% end if %>
</td>
</tr>
        <% next %>
        </table>
     </td>
   </tr>
 </table>
</div>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="100">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
			<img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>이미지등록</strong>
			<br>- 텐바이텐에서 이미지를 등록할 경우 따로 입력하지 마시기 바랍니다.
			<br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
			<br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
			<br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
        </td>
        <td align="right" valign="bottom">
        	<a href="javascript:PopImageInformation()"><b><font color=red>[필독]이미지 등록요령</font></b> <img src="/images/icon_help.gif" border="0" align="absbottom"></a>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<iframe name="imgpreview" src="iframe_imagepreview.asp" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic"> (<font color=red>필수</font>,400X400,<b><font color="red">jpg</font></b>)
	  <div id="divimgbasic" style="display:none;">
      <table width="400" height="400" >
        <tr>
          <td>
          	<img id="imgbasic_img" src=""> 
          </td>
        </tr>
      </table>
      </div>
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1"> (선택,400X400,jpg,gif)
	  <div id="divimgadd1" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd1_img" src=""></td>
        </tr>
        
      </table>
      </div>
  	</td>
  </tr>
  <tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <div id="divimgadd2" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd2_img" src=""></td>
        </tr>
       
      </table>
      </div>

      <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2"> (선택,400X400,jpg,gif)
  	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
	  <div id="divimgadd3" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd3_img" src=""></td>
        </tr>
        
      </table>
      </div>

      <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3"> (선택,400X400,jpg,gif)

  	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">

      <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> (선택,400X400,jpg,gif)

	  <div id="divimgadd4" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd4_img" src=""></td>
        </tr>
        
      </table>
      </div>
  	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg,gif');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> (선택,400X400,jpg,gif)

      <div id="divimgadd5" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd5_img" src=""></td>
        </tr>
       
      </table>
      </div>
   	</td>
  </tr>
  <tr height="2" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="file" name="imgmain" onchange="CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');" size="40">
      <input type="button" value="이미지지우기" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> (선택,600X2000, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
	  <div id="divimgmain" style="display:none;">
      <table width="400" height="400">
        <tr>
          <td>
          <img id="imgmain_img" src="">
          </td>
        </tr>
      </table>
      </div>
  	</td>
  </tr>
  </form>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" onClick="SubmitSave()">
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

<script language='javascript'>
function getOnload(){
    EnDisableFlowerShop();
}
window.onload = getOnload;
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
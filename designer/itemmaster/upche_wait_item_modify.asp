<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemregcls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"--> 
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB
' 추가한부분
CONST CMOBILE_IMG_MAXSIZE = 400   'KB
'// 추가한부분

Dim itemid  
itemid =  requestCheckvar(Request("itemid"),16)
'==============================================================================
Dim oitemdetail,oitemreg,optiontotal,ix,ooimage, colorCnt

set oitemdetail = new CWaitItemDetail

oitemdetail.FRectDesignerID = session("ssBctID")
oitemdetail.WaitProductDetail itemid '임시등록 데이터 불러오기
oitemdetail.WaitProductDetailOption itemid '옵션 2번 넘버,이름 불러오기

if oitemdetail.Foptioncnt>0 then
	colorCnt = oitemdetail.FItemList(0).FcolorCnt
else
	colorCnt = 0
end if


'==============================================================================
set oitemreg = new CItemReg

'if oitemdetail.FResultCount <> 0 then
'	oitemreg.SearchOptionNameBig left(oitemdetail.FItemList(ix).Fitemoption,2) '옵션 1번 불러오기
'end if

oitemreg.SearchCategoryNameLarge oitemdetail.Flarge '카테고리 1번 불러오기
oitemreg.SearchCategoryNameMid oitemdetail.Flarge,oitemdetail.FMid '카테고리 2번 불러오기
oitemreg.SearchCategoryNameSmall oitemdetail.Flarge,oitemdetail.FMid,oitemdetail.Fsmall '카테고리 3번 불러오기


'==============================================================================
set ooimage = new CWaitItemImagelist
ooimage.WaitProductImageList itemid  '이미지 데이터 불러오기

Dim itemaddimage,itemaddcontent, itemstoryimage

if (IsNull(ooimage.Fimgadd) or (ooimage.Fimgadd="")) then ooimage.Fimgadd = ",,,,"
if (IsNull(ooimage.Fitemaddcontent) or (ooimage.Fitemaddcontent="")) then ooimage.Fitemaddcontent = "||||"
if (IsNull(ooimage.Fimgstory) or (ooimage.Fimgstory="")) then ooimage.Fimgstory = ",,,,"


itemaddimage = split(ooimage.Fimgadd,",")
itemaddcontent = split(ooimage.Fitemaddcontent,"|")
itemstoryimage = split(ooimage.Fimgstory,",")


'==============================================================================
dim imgsubdir
dim arrOld
imgsubdir = GetImageSubFolderByItemid(itemid)

'--- 등록진행정보 
Dim clsWait, arrList, intLoop
 set clsWait = new CWaitItemlist2014
 	clsWait.Fitemid = itemid
 	arrList = clsWait.fnGetWaitItemLog
 	IF not isArray(arrList) THEN
 		arrOld = clsWait.fnGetOldWaitItemLog
	END IF
 set clsWait = nothing

'==============================================================================
'--계약 진행 내역
dim ocontract,ctrState
set ocontract = new CPartnerContract
ocontract.FRectGroupId =  session("ssGroupid")
ocontract.FRectMakerId =  session("ssBctID")
ocontract.fnGetCtrState
ctrState = ocontract.FCtrState
set ocontract = nothing

'==============================================================================
Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = session("ssBctID")
npartner.GetOnePartnerNUser

 
%>
<!--
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
 //호환성 오류 
-->
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
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
	var isvatinclude, imileage;
	var isellcash, ibuycash, isellvat, ibuyvat, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatinclude = frm.vatinclude[0].checked;

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
		isellvat = parseInt(parseInt(1/11 * parseInt(isellcash)));
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.005) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = 0;
		imileage = parseInt(isellcash*0.005) ;
	}

	frm.sellvat.value = isellvat;
	frm.buycash.value = ibuycash;
	frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
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
function ClearImage(img,fsize,wd,ht) {
	//img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <!--%= CBASIC_IMG_MAXSIZE %-->, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

    document.getElementById("div"+ img.name).style.display = "none";

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "del";
}

function ClearImage2(img,fsize,wd,ht) {
	//img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <!--%= CBASIC_IMG_MAXSIZE %-->, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

	if (img=="mobile")
	{
	    document.getElementById("divmobileimgmain").style.display = "none";
	}
	if (img=="mobile2")
	{
	    document.getElementById("divmobileimgmain2").style.display = "none";
	}
	if (img=="mobile3")
	{
	    document.getElementById("divmobileimgmain3").style.display = "none";
	}
	if (img=="mobile4")
	{
	    document.getElementById("divmobileimgmain4").style.display = "none";
	}
	if (img=="mobile5")
	{
	    document.getElementById("divmobileimgmain5").style.display = "none";
	}
	if (img=="mobile6")
	{
	    document.getElementById("divmobileimgmain6").style.display = "none";
	}
	if (img=="mobile7")
	{
	    document.getElementById("divmobileimgmain7").style.display = "none";
	}
	// 20160601추가한부분
	if (img=="mobile8")
	{
	    document.getElementById("divmobileimgmain8").style.display = "none";
	}
	if (img=="mobile9")
	{
	    document.getElementById("divmobileimgmain9").style.display = "none";
	}
	if (img=="mobile10")
	{
	    document.getElementById("divmobileimgmain10").style.display = "none";
	}
	if (img=="mobile11")
	{
	    document.getElementById("divmobileimgmain11").style.display = "none";
	}
	if (img=="mobile12")
	{
	    document.getElementById("divmobileimgmain12").style.display = "none";
	}
	// 20160601추가한부분
	var e = eval("itemreg."+img);
	e.value = "del";
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
        ClearImage(img,fsize, imagewidth, imageheight);
        return false;
    }

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "";

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

// ============================================================================
// 저장하기
function SubmitSave() {
//alert('현재 서버 작업 중으로 상품 등록/ 변경이 불가합니다.');
//return;
	var itemreg = document.itemreg;

	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length){
		alert("전시 카테고리를 선택하세요.\n※ 전시 [기본] 카테고리는 반드시 지정되어야 합니다.");
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

//상품무게 숫자체크
 if (!IsDigit(document.itemreg.itemWeight.value)){
		alert('상품무게는  숫자로 입력하세요.');
		itemreg.itemWeight.focus();
		return;
	}
	
    //배송구분 체크 ================================================================ 
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
    
    //배송구분 업체이나 매입구분이 업체가 아닌것.
    if (itemreg.deliverytype[1].checked){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
    }
    
    //매입구분이 업체이나 배송구분이 업체가 아닌것.
    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
    }

    //업체배송만 주문제작 가능.
    if ((!itemreg.mwdiv[2].checked)&&(itemreg.itemdiv[1].checked)){
        alert('주문제작 상품은 업체배송인경우만 가능합니다.');
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
	
		if(document.itemreg.optionaddprice.value >0 && document.itemreg.deliverOverseas.checked){
			alert("옵션에 추가가격이 있을 경우 해외배송이 불가능합니다. 해외배송체크를 해제해주세요" );
			document.itemreg.deliverOverseas.focus();
			 return;
		}

    if(document.itemreg.deliverOverseas.checked){
	    if(document.itemreg.itemWeight.value<=0){
	        alert("해외배송시 배송비 산출을 위해 상품무게를 꼭 입력해주세요")
	        document.itemreg.itemWeight.focus();
	        return;
	    }
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

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}

	//상품 설명에 불가항목 검사
	var cntRe = /.js["'>\s]/gi;
	if(cntRe.test(itemreg.itemcontent.value)) {
        alert('상품설명에는 js파일을 넣을 수 없습니다.');
        itemreg.itemcontent.focus();
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

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
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
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
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

    if (itemreg.imgmain.value != "") {
        if (CheckImage(itemreg.imgmain, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgmain2.value != "") {
        if (CheckImage(itemreg.imgmain2, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgmain3.value != "") {
        if (CheckImage(itemreg.imgmain3, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

   //추가 이미지 2장이상 등록 권유 2014.09.19 추가 정윤정
     var strWMsg = "";
    
    //계약완료  확인
    if (itemreg.ctrState.value != 7){
    	strWMsg = strWMsg + "계약서가 완료된 후 상품 오픈이 가능합니다. \n계약서를 우편으로 발송해주세요\n\n"
    }
    
    if (itemreg.imgadd1.value=="" || itemreg.imgadd2.value=="") {
    	 strWMsg = strWMsg + "편리한 PC&모바일 쇼핑을 위해서 추가이미지를 2장 이상 등록바랍니다.\n\n"; 
	 } 
	 
    if(confirm(strWMsg + "상품을 올리시겠습니까?") == true){

		$("#regbtn").hide();
		$("#lyLoading").show();

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

    if (validate(itemreg)==false) {
        return;
    }
    
	//상품 설명에 불가항목 검사
	var cntRe = /.js["'>\s]/gi;
	if(cntRe.test(itemreg.itemcontent.value)) {
        alert('상품설명에는 js파일을 넣을 수 없습니다.');
        itemreg.itemcontent.focus();
        return;
	}

//상품무게 숫자체크
 if (!IsDigit(document.itemreg.itemWeight.value)){
		alert('상품무게는  숫자로 입력하세요.');
		itemreg.itemWeight.focus();
		return;
	}
	
	//배송구분 체크 ================================================================ 
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[3].checked){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            return;
        }
    }
    //업체착불배송
    if ((itemreg.defaultDeliveryType.value!="7")&&(itemreg.deliverytype[4].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송] 업체가 아닙니다.');
        return;
    }
    
    //배송구분 업체이나 매입구분이 업체가 아닌것.
    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
    }
    
    //매입구분이 업체이나 배송구분이 업체가 아닌것.
    if (itemreg.mwdiv[2].checked){
        if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
    }
    
    if(document.itemreg.optionaddprice.value >0 && document.itemreg.deliverOverseas.checked){
			alert("옵션에 추가가격이 있을 경우 해외배송이 불가능합니다. 해외배송체크를 해제해주세요" );
			document.itemreg.deliverOverseas.focus();
			 return;
		}
		
    if(document.itemreg.deliverOverseas.checked){
	    if(document.itemreg.itemWeight.value<=0){
	        alert("해외배송시 배송비 산출을 위해 상품무게를 꼭 입력해주세요")
	        document.itemreg.itemWeight.focus();
	        return;
	    }
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

    if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000만원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}

    if(itemreg.limityn[1].checked == true && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }

	//상품 설명에 불가항목 검사
	var cntRe = /.js["'>\s]/gi;
	if(cntRe.test(itemreg.itemcontent.value)) {
        alert('상품설명에는 js파일을 넣을 수 없습니다.');
        itemreg.itemcontent.focus();
        return;
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
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
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

    if (itemreg.imgmain.value != "") {
        if (CheckImage(itemreg.imgmain, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgmain2.value != "") {
        if (CheckImage(itemreg.imgmain2, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgmain3.value != "") {
        if (CheckImage(itemreg.imgmain3, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40) != true) {
            return;
        }
    }

    // 재요청 내용 쓰기 팝업
    reMsg = prompt("재등록 요청시 전달할 내용을 간략히 써주세요.","");
    if(reMsg){
        itemreg.reRegMsg.value = reMsg;
        itemreg.CurrState.value = "5";	// 상태정보를 '등록재요청'으로 수정
        itemreg.target = "FrameCKP";
        itemreg.submit();
    }
    else {
    	return;
    }

}

function TnCheckUpcheYN(frm){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// 기본체크
		// 배송구분 지정(텐바이텐)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=true;  //텐바이텐 무료배송은 업체에서 체크 할 수 없음.
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
		// 배송구분 지정(업체배송)
		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
		frm.deliverytype[3].disabled=false; //업체개별배송(9)
        frm.deliverytype[4].disabled=false;  //업체착불배송(7) : 업체에서 설정불가
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
		frm.limitno.readOnly=true;
		frm.limitno.className='text_ro';
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
		frm.limitno.className='text';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readOnly=true;
		frm.limitno.className='text_ro';
	}
	else {
		// 한정
		if ((frm.optioncnt.value*1) > 0) {
		    // 옵션사용중
		    // alert("한정수량은 옵션창에서 수정가능합니다.");
		    return;
        }

		frm.limitno.readOnly=false;
		frm.limitno.className='text';
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
	else if(frm.deliverytype[1].checked ){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("매입특정 구분이 매입이나 특정일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입특정구분을 확인해주세요!!");
			frm.mwdiv[2].checked=true;
		}
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

// ============================================================================
// 미리보기
function ViewItemDetail(itemno){
	window.open('viewitem.asp?itemid='+itemno ,'ViewItemDetail','width=790,height=600,scrollbars=yes,status=no');
}

function ClearVal(comp){
    comp.value = "";
}

function popWaitItemOptionEdit(iitemid){
    var popwin = window.open('/common/pop_upchewaititemoptionedit.asp?itemid=' + iitemid,'popWaitItemOptionEdit','width=790,height=600,scrollbars=yes,status=no');
    popwin.focus();
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
			url: "/admin/itemmaster/act_waitItemInfoDivForm.asp",
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

<!--2013 리뉴얼 추가 ------->
$(function(){
	// 로딩후 상품속성 내용 출력
	printItemAttribute();
});

// 상품속성 출력
function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_waitItemAttribSelect.asp?itemid=<%=itemid%>&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

// 전시카테고리 선택 팝업
function popDispCateSelect(){
	var dCnt = $("input[name='isDefault'][value='y']").length;
	var cCnt = $("input[name='isDefault']").length;
	$.ajax({
		url: "/common/module/act_DispCategorySelectUpche.asp?isDft="+dCnt+"&chk="+cCnt,
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

	if($("input[name='isDefault']").length>1) {
		$("#btnAddCate").hide();
	}

	//상품속성 출력
	printItemAttribute();
}

// 선택 전시카테고리 삭제
function delDispCateItem() {
	if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
		tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

		if($("input[name='isDefault']").length<2) {
			$("#btnAddCate").show();
		}

		//상품속성 출력
		printItemAttribute();
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
<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/WaitUpcheItemModify_Process.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<!--<input type="hidden" name="designerid" value="<%= oitemdetail.FRectDesignerID %>">-->
<input type="hidden" name="defultmargine" value="<% =npartner.FOneItem.Fdefaultmargine %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FOneItem.Fmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FOneItem.FdefaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FOneItem.FdefaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FOneItem.FdefaultDeliveryType %>">
<input type="hidden" name="DFcolorCD" value="<%=oitemdetail.FDFcolorCd%>">

<input type="hidden" name="pojangok" value="N">
<input type="hidden" name="sellyn" value="N">
<input type="hidden" name="dispyn" value="N">
<input type="hidden" name="isusing" value="Y">
<input type="hidden" name="reRegMsg" value="">
<input type="hidden" name="CurrState" value="<%=oitemdetail.FCurrState%>">
<input type="hidden" name="adminid" value="<%= session("ssBctID")%>">
<input type="hidden" name="ctrState" value="<%=ctrState%>">
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


<!-- 1.일반정보 --> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>1.일반정보</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designerid"  value="<%= oitemdetail.FMakerid %>" class="text_ro" readonly size="30" id="[on,off,off,off][브랜드ID]">&nbsp;
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= itemid %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="미리보기" onclick="ViewItemDetail('<%= itemid %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemname %>" id="[on,off,off,off][상품명]">&nbsp;
	</td>
</tr>
</table>

<!-- 2.구분 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.구분</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="재고/매출 등의 관리 카테고리" style="cursor:help;">관리 카테고리 :</td>
		<input type="hidden" name="cd1" value="<%= oitemdetail.Flarge %>">
		<input type="hidden" name="cd2" value="<%= oitemdetail.Fmid %>">
		<input type="hidden" name="cd3" value="<%= oitemdetail.Fsmall %>">
		<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" name="cd1_name" value="<%= oitemreg.largename %>" id="[on,off,off,off][카테고리]" size="20" readonly class="text_ro">
		<input type="text" name="cd2_name" value="<%= oitemreg.midname %>"  id="[on,off,off,off][카테고리]" size="20" readonly class="text_ro">
		<input type="text" name="cd3_name" value="<%= oitemreg.smallname %>" id="[on,off,off,off][카테고리]" size="20" readonly class="text_ro">
		
		<input type="button" value="카테고리 선택" class="button" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td id="lyrDispList"><%=getDispCategoryWait(itemid)%></td>
			<td valign="bottom"><input id="btnAddCate" type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
	<td bgcolor="#FFFFFF" >
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitemdetail.Fitemdiv ="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">일반상품</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitemdetail.Fitemdiv %>" <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">주문 제작상품</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitemdetail.Fitemdiv="06","checked","")%> <%=chkIIF(oitemdetail.Fitemdiv="06" or oitemdetail.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문시 이니셜등 제작문구가 필요한경우 체크)</font>
		<br>
      <!--업체등록시에는 사용안함(MD만 등록가능)
      	<label><input type="radio" name="itemdiv" value="08" <%=chkIIF(oitemdetail.Fitemdiv ="08","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';">티켓상품</label>
      //-->
		
	</td>
	<td bgcolor="#FFFFFF" >
	    <div id="lyRequre" style="<%=chkIIF(oitemdetail.Fitemdiv ="06" or oitemdetail.Fitemdiv ="16","","display:none;")%>padding-left:22px;">
			예상제작소요일 <input type="text" name="requireMakeDay" value="<%=oitemdetail.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][예상제작소요일]">일
			<font color="red">(상품발송전 상품제작 기간)</font>
		</div>
    </td>
</tr>
</table>

<!-- 3.가격정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>3.가격정보</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="radio" name="vatinclude" value="Y" onclick="TnGoClear(this.form);" <%=chkIIF(oitemdetail.Fvatinclude="Y","checked","")%>>과세
		<input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);" <%=chkIIF(oitemdetail.Fvatinclude="N","checked","")%>>면세
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">기본 공급 마진 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][마진]" value="<%= oitemdetail.FMargin %>"   class="text" onKeyUp="CalcuAuto(itemreg);">%
	</td>
</tr>
<tr align="left">
<input type="hidden" name="sellvat" value="<%= oitemdetail.Fsellvat %>">
<input type="hidden" name="buyvat" value="<%= oitemdetail.Fbuyvat %>">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가) :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="sellcash" size="12" id="[on,on,off,off][소비자가]" onKeyUp="CalcuAuto(itemreg);" maxlength="8" class="text" value="<%= oitemdetail.Fsellcash %>">원
	</td>
	<td width="15%" bgcolor="#DDDDFF">공급가 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][공급가]" readonly class="text_ro" value="<%= oitemdetail.Fbuycash %>">원
		(<b>부가세 포함가</b>)
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#F8F8F8" colspan="3">
		- 공급가는 <b>부가세 포함가</b>입니다.<br>
		- 소비자가(할인가)와 마진(할인마진)을 입력하면 공급가와 마일리지가 자동계산됩니다.
	</td>
</tr>
<input type="hidden" name="mileage" value="<%= oitemdetail.Fmileage %>">
</table>

<!-- 4.관리정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>4.관리정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="상품 상세 속성" style="cursor:help;">상품속성 :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="upchemanagecode" value="<%= oitemdetail.Fupchemanagecode %>" size="20" maxlength="32" class="text" id="[off,off,off,off][업체상품코드]">
	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
	</td>
</tr>
</table>

<!-- 5.기본정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>5.기본정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][제조사]" value="<%= oitemdetail.Fmakername %>">&nbsp;(제조업체명)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="sourcearea" maxlength="64" size="25" class="text" id="[on,off,off,off][원산지]" value="<%= oitemdetail.Fsourcearea %>">&nbsp;(ex:한국,중국,중국OEM,일본 등 / 식품일 경우 국내: 국내산 또는 시군구명, 수입: 미국산, 중국산 등)
	  <br>( 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][상품무게]" style="text-align:right" value="<%= oitemdetail.Fitemweight %>">g &nbsp;(그램단위로 입력, ex:1.5kg→ 1500) / 숫자만 입력 / 해외배송시 배송비 산출을 위한 것이므로 정확히 입력.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="text" name="keywords" maxlength="256" size="120" class="text" id="[on,off,off,off][검색키워드]" value="<%= oitemdetail.Fkeywords %>">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
	</td>
</tr>
</table>

<!-- 5-1.품목상세정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 품목상세정보 </strong> &nbsp;<font color=gray>상품정보제공고시 관련 법안 추진에 따라 아래 내용을 정확히 입력해주시기 바랍니다.</font></td>
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
		<option value="07">영상가전(TV류)</option>
		<option value="08">가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option>
		<option value="09">계절가전(에어컨/온풍기)</option>
		<option value="10">사무용기기(컴퓨터/노트북/프린터)</option>
		<option value="11">광학기기(디지털카메라/캠코더)</option>
		<option value="12">소형전자(MP3/전자사전 등)</option>
		<option value="14">내비게이션</option>
		<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
		<option value="16">의료기기</option>
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
		<option value="35">기타</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitemdetail.FinfoDiv%>";
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="") then
			Server.Execute("/admin/itemmaster/act_waitItemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
	<td bgcolor="#FFFFFF">
	  <input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitemdetail.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitemdetail.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
	</td>
</tr>
</table>
<!-- 5-2.안전인증정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 안전인증정보</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">안전인증대상 :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitemdetail.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)">대상</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)">대상아님</label><br />
		<select name="safetyDiv" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::안전인증구분::</option>
		<option value="10" <%=chkIIF(oitemdetail.FsafetyDiv="10","selected","")%>>국가통합인증(KC마크)</option>
		<option value="20" <%=chkIIF(oitemdetail.FsafetyDiv="20","selected","")%>>전기용품 안전인증</option>
		<option value="30" <%=chkIIF(oitemdetail.FsafetyDiv="30","selected","")%>>KPS 안전인증 표시</option>
		<option value="40" <%=chkIIF(oitemdetail.FsafetyDiv="40","selected","")%>>KPS 자율안전 확인 표시</option>
		<option value="50" <%=chkIIF(oitemdetail.FsafetyDiv="50","selected","")%>>KPS 어린이 보호포장 표시</option>
		</select>
		인증번호 <input type="text" name="safetyNum" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> size="18" maxlength="18" class="text" value="<%=oitemdetail.FsafetyNum%>" />
		<font color="darkred">유아용품이나 전기용품일 경우 필수 입력</font>
	</td>
</tr>
</table>

<!-- 6.배송정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>6.배송정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">매입특정구분 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(oitemdetail.Fmwdiv="M","checked","")%>>매입
	  <input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(oitemdetail.Fmwdiv="W","checked","")%>>특정
	  <input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(oitemdetail.Fmwdiv="U","checked","")%>>업체배송
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="1","checked","")%>>텐바이텐배송&nbsp;
	  <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="2","checked","")%>>업체(무료)배송&nbsp;
	  <input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="4","checked","")%>>텐바이텐무료배송&nbsp;
	  <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="9","checked","")%>>업체조건배송(개별 배송비부과)&nbsp;
	  <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="7","checked","")%>>업체착불배송
	 </td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송방법 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverfixday" value="" onclick="TnCheckFixday(this.form)" <%=chkIIF(Trim(oitemdetail.Fdeliverfixday)="" or IsNull(oitemdetail.Fdeliverfixday),"checked","")%>>택배(일반)&nbsp;
	  <input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)" <%=chkIIF(oitemdetail.Fdeliverfixday="X","checked","")%>>화물&nbsp;
	  <input type="radio" name="deliverfixday" value="C" onclick="TnCheckFixday(this.form)" <%=chkIIF(oitemdetail.Fdeliverfixday="C","checked","")%>>플라워지정일
		<span id="lyrFreightRng" style="display:<%=chkIIF(oitemdetail.Fdeliverfixday="X","","none")%>;">
			<br />&nbsp;
			반품/교환 시 화물배송 비용(편도) :
			최소 <input type="text" name="freight_min" class="text" size="6" value="<%=oitemdetail.Ffreight_min%>" style="text-align:right;">원 ~
			최대 <input type="text" name="freight_max" class="text" size="6" value="<%=oitemdetail.Ffreight_max%>" style="text-align:right;">원
		</span>
	  <br>&nbsp;<font color="red">(플라워 상품인 경우만 수도권배송, 서울배송, 플라워지정일 옵션이 사용가능합니다.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송지역 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(oitemdetail.Fdeliverarea)="" or IsNull(oitemdetail.Fdeliverarea),"checked","")%>>전국배송&nbsp;
	  <input type="radio" name="deliverarea" value="C" <%=chkIIF(oitemdetail.Fdeliverarea="C","checked","")%>>수도권배송&nbsp;
	  <input type="radio" name="deliverarea" value="S" <%=chkIIF(oitemdetail.Fdeliverarea="S","checked","")%>>서울배송&nbsp;
	  <label><input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF(oitemdetail.FdeliverOverseas="Y","checked","")%> title="해외배송은 상품무게가 입력이 돼야 완료됩니다.">해외배송
	  	(+ 옵션추가 가격이 있을경우 해외배송 불가, 상품무게입력 필수)
	  	</label>
	</td>
</tr>
</table>

<!-- 7.옵션정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>7.옵션정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" rowspan="2">옵션구분 :</td>
	<input type="hidden" name="optioncnt" value="<%= oitemdetail.Foptioncnt %>">
	<input type="hidden" name="optionaddprice" value="<%= oitemdetail.fnGetWaitOptAddPrice(itemid) %>">
	<td width="85%" bgcolor="#FFFFFF">
	  <% if oitemdetail.Foptioncnt < 1 then %>
	  옵션사용안함
	  <% else %>
	  옵션사용중(<%= oitemdetail.Foptioncnt %>개)
	  <% end if %>
	  &nbsp;&nbsp;<input type="button" class="button" value="옵션수정" onClick="popWaitItemOptionEdit('<%= oitemdetail.FWaitItemID %>');">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF">
      - 옵션정보는 옵션창에서 수정가능합니다.<br>
      - 옵션은 추가는 가능하지만 삭제는 불가능합니다. 정확히 입력하세요.
	</td>
</tr>
<tr id="lyDFColor" height="30" style="display:;">
	<td colspan="2" bgcolor="#FFFFFF" style="padding:0px;">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="0">
		<tr>
			<td width="15%" bgcolor="#DDDDFF">기본 색상선택 :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-left:1px solid <%= adminColor("tablebg") %>;"><%=FnSelectColorBar(oitemdetail.FDFColorCD,25)%></td>
		</tr>
		<tr>
			<td width="15%" rowspan="2" bgcolor="#DDDDFF" style="border-top:1px solid <%= adminColor("tablebg") %>;">색상별 상품이미지 :</td>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
			  <% if (oitemdetail.FDFcolorImg <> "") then %>
				<div id="divimgDFColor" style="display:block;">
				<img src="<%=partnerUrl%>/waitimage/color/<%=imgsubdir%>/<%=oitemdetail.FDFcolorImg %>" width="200">
				</div>
			  <% end if %>
				<input type="file" size="40" name="imgDFColor" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif', 40);" class="text">
				<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgDFColor, 40, 1000, 1000)"> (선택,1000X1000,<b><font color="red">jpg</font></b>)
				<input type="hidden" name="DFColor">
			</td>
		</tr>
		<tr>
			<td width="85%" bgcolor="#FFFFFF" style="border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;">
		      - 색상별 이미지는 별도로 등록을 하지않으면 상품 기본이미지가 사용됩니다.
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- 8.한정정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>8.한정정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF" rowspan="2">한정판매구분 :</td>
	<td width="35%" bgcolor="#FFFFFF">
	  <input type="radio" name="limityn" value="N" onClick="TnCheckLimitYN(itemreg)" <%=chkIIF(oitemdetail.Flimityn="N","checked","")%>>비한정판매&nbsp;&nbsp;
	  <input type="radio" name="limityn" value="Y" onClick="TnCheckLimitYN(itemreg)" <%=chkIIF(oitemdetail.Flimityn="Y","checked","")%>>한정판매
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">한정수량 :</td>
	<td width="35%" bgcolor="#FFFFFF" >
	  <input type="text" name="limitno" maxlength="32" size="8" readonly class="text_ro" id="[off,on,off,off][한정수량]" value="<%= oitemdetail.Flimitno %>">(개)
      <input type="hidden" name="limitsold" value="0">
      <input type="hidden" name="limitstock" value="<%= oitemdetail.Flimitno %>">
	</td>
</tr>
<tr>
	<td colspan="3" bgcolor="#FFFFFF"><font color="red">** 한정수량은 옵션이 있을 경우, 옵션창에서 수정이 가능합니다.(위의 정보는 부정확할수 있습니다.)</font></td>
</tr>
</table>

<!-- 9.상품설명 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>9.상품설명</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="usinghtml" value="N" <%=chkIIF(oitemdetail.Fusinghtml="N","checked","")%>>일반TEXT
	  <input type="radio" name="usinghtml" value="H" <%=chkIIF(oitemdetail.Fusinghtml="H","checked","")%>>TEXT+HTML
	  <input type="radio" name="usinghtml" value="Y" <%=chkIIF(oitemdetail.Fusinghtml="Y","checked","")%>>HTML사용
	  <br>
	  <textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][상품설명]"><%= oitemdetail.Fitemcontent %></textarea>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][유의사항]"><%= oitemdetail.Fordercomment %></textarea><br>
	  <font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
	</td>
</tr>
</table>

<!-- 10.이미지정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>10.이미지정보</strong>
		<br>- 텐바이텐에서 이미지를 등록할 경우에는 필수항목인 기본이미지만 입력하시기 바랍니다.
		<br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
		<br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
		<br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 기본이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgbasic <> "") then %>
		<div id="divimgbasic" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/basic/<%=imgsubdir%>/<%= ooimage.Fimgbasic %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgbasic" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>필수</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="basic">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(0) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add1/<%=imgsubdir%>/<%=itemaddimage(0) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd1" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(1) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add2/<%=imgsubdir%>/<%=itemaddimage(1) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(2) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add3/<%=imgsubdir%>/<%=itemaddimage(2) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(3) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add4/<%=imgsubdir%>/<%=itemaddimage(3) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (itemaddimage(4) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add5/<%=imgsubdir%>/<%=itemaddimage(4) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 흰배경(누끼)이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmask <> "") then %>
		<div id="divimgmask" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mask/<%=imgsubdir%>/<%=ooimage.Fimgmask %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgmask" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="mask">
	</td>
</tr>
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 기존의 제품설명이미지는 사용하지 않고 상품설명이미지를 사용합니다. 기존에 등록된 제품설명이미지는 사용은 하되 추가 수정은 되지않고 삭제만 됩니다.</strong></font>
 	</td>
 </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmain <> "") then %>
		<div id="divimgmain" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main/<%=imgsubdir%>/<%=ooimage.Fimgmain %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgmain" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmain, 40, 800, 1600)"> (선택,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #2:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmain2 <> "") then %>
		<div id="divimgmain2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main2/<%=imgsubdir%>/<%=ooimage.Fimgmain2 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgmain2" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmain2, 40, 800, 1600)"> (선택,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #3:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmain3 <> "") then %>
		<div id="divimgmain3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main3/<%=imgsubdir%>/<%=ooimage.Fimgmain3 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgmain3" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmain3, 40, 800, 1600)"> (선택,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main3">
	</td>
</tr>

 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 모바일 상품상세 이미지는 앞으로 이 영역으로 대체 됩니다. html은 사용하지 않을 예정이오니 이쪽으로 업로드 해주시기 바랍니다.<br>※ 모바일 상품상세에는 이미지를 잘라서 올려주시기 바랍니다.</strong></font>
 	</td>
 </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain <> "") then %>
		<div id="divmobileimgmain" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #2:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain2 <> "") then %>
		<div id="divmobileimgmain2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile2/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain2 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain2" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile2', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #3:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain3 <> "") then %>
		<div id="divmobileimgmain3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile3/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain3 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain3" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile3', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #4:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain4 <> "") then %>
		<div id="divmobileimgmain4" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile4/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain4 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain4" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile4', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #5:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain5 <> "") then %>
		<div id="divmobileimgmain5" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile5/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain5 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain5" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile5', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile5">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #6:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain6 <> "") then %>
		<div id="divmobileimgmain6" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile6/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain6 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain6" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain6" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile6', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile6">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #7:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain7 <> "") then %>
		<div id="divmobileimgmain7" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile7/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain7 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain7" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain7" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile7', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile7">
	</td>
</tr>
<!-- 20160601추가한부분 -->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #8:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain8 <> "") then %>
		<div id="divmobileimgmain8" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile8/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain8 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain8" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain8" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile8', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile8">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #9:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain9 <> "") then %>
		<div id="divmobileimgmain9" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile9/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain9 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain9" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain9" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile9', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile9">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #10:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain10 <> "") then %>
		<div id="divmobileimgmain10" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile10/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain10 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain10" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain10" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile10', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile10">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #11:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain11 <> "") then %>
		<div id="divmobileimgmain11" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile11/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain11 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain11" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain11" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile11', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile11">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">모바일 상품상세이미지 #12:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fmobileimgmain12 <> "") then %>
		<div id="divmobileimgmain12" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mobile12/<%=imgsubdir%>/<%=ooimage.Fmobileimgmain12 %>" width="400">
		</div>
	  <% else %>
	  <div id="divmobileimgmain12" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="mobileimgmain12" onchange="ClearImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile12', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile12">
	</td>
</tr>
<!--// 20160601추가한부분 -->


</table> 
<!-- 11.기타 정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>11.등록정보</strong></td>
  	<td align="right">3회 이상 보류시, 반려 처리(재등록불가) 되므로 참고 부탁드립니다.</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
			<td height="30" width="15%" bgcolor="#DDDDFF">진행일자</td>
			<td bgcolor="#FFFFFF"> 
				<%IF isARray(arrList) THEN%>
			 <% dim count2, strMsg, sMsgCd2, sMsgcd0
			 count2 = 0 
			 sMsgCd2 = ""
			 sMsgcd0 = ""
			 For intLoop = 0 To UBound(arrList,2)
			 strMsg = ""
			 		IF arrList(2,intLoop) = 2 THEN
			 			count2 = count2 + 1
			 			strMsg = count2&"차"
			 			sMsgCd2 = sMsgCd2 + "^" + arrList(6,intLoop)
			 		ELSEIF arrList(2,intLoop) = 0 THEN
			 				sMsgCd0 = sMsgCd0 + "^" + arrList(6,intLoop)	
			 		END IF	
			 %> 
			 <div style="padding:3"><font color="<%=GetCurrStateColor(arrList(2,intLoop))%>"><%=strMsg%><%=fnGetCurrStateShortName(arrList(2,intLoop))%></font>: <%=arrList(4,intLoop)%> &nbsp;<%IF arrList(3,intLoop) <> "" THEN%>[<%=replace(arrList(3,intLoop),"^","/")%>]<%END IF%></div>
			 <%Next 
			 ELSEIF isArray(arrOld) THEN  
				 IF arrold(4,0) = 5 THEN
	  	%>
	  	 <div style="padding:3">보류:<%=arrold(0,0)%> &nbsp;[<%=arrold(1,0)%>] </div>
	  	 <div style="padding:3"><font color="<%=GetCurrStateColor(arrold(4,0))%>"><%=fnGetCurrStateShortName(arrold(4,0))%></font>: <%=arrold(2,0)%> &nbsp;[<%=arrold(3,0)%>]</div>
	  	<%ELSE%>
			 <font color="<%=GetCurrStateColor(arrold(4,0))%>"><%=fnGetCurrStateShortName(arrold(4,0))%></font>: <%=arrold(0,0)%> &nbsp;[<%=arrold(1,0)%>] 
			 <%END IF%> 
			  <%END IF%> 
			</td>
		</tr> 
</table>
 

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center" style="padding:10px 0 0 0;">
	<% if oitemdetail.FCurrState="1" or oitemdetail.FCurrState="5" then %>
		<span id="regbtn"><input type="button" value="임시저장" class="button_s" onClick="SubmitSave()"></span>
		<span id="lyLoading" style="display:none;text-align:center;"><img src="http://fiximage.10x10.co.kr/icons/loading16.gif" style="width:16px;height:16px;" />&nbsp;&nbsp;<font color="blue"><strong>저장 중 입니다. 잠시만 기다려주세요.</strong></font></span>
	<% elseif oitemdetail.FCurrState="2" then %>
		<input type="button" value="재요청" class="button" onClick="SubmitReSave()">
	<% end if %>
		<input type="button" value="목록으로" class="button" onClick="history.back()">
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
</form>
<p>
<script>
// 매입특정구분 및 배송구분세팅
TnCheckUpcheYN(itemreg);
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements(i).name == "deliverytype") {
        if (itemreg.elements(i).value == "<%= oitemdetail.Fdeilverytype %>") {
            itemreg.elements(i).checked = true;
        }
    }
}

// 한정
TnSilentCheckLimitYN(itemreg);
// 세일
// TnCheckSailYN(itemreg);
</script>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  온라인 승인대기상품
' History : 서동석 생성
'			2023.08.11 한용민 수정(isbn 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemregcls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 640   'KB

CONST CMAIN_IMG_MAXSIZE = 1230   'KB
CONST CMAIN_IMG_MAXWIDTH = 1000   'Px
CONST CMAIN_IMG_MAXHEIGHT = 3000   'Px

CONST CMOBILE_IMG_MAXSIZE = 500   'KB

dim arrold, deliverfixday, mwdiv, deliverytype, purchaseType, deliverarea
Dim clsWait, itemid ,makerid,arrlist, intLoop, i
makerid	= requestCheckvar(Request("designer"),32)
itemid =  requestCheckvar(Request("itemid"),16)

Dim oitemdetail,oitemreg,optiontotal,ix,ooimage, mainImg(10)

set oitemdetail = new CWaitItemDetail

oitemdetail.FRectDesignerID = request("designer")
oitemdetail.WaitProductDetail request("itemid") '임시등록 데이터 불러오기
oitemdetail.WaitProductDetailOption request("itemid") '옵션 2번 넘버,이름 불러오기


if oitemdetail.FTotalCount>0 then
	purchaseType = oitemdetail.fpurchaseType		' 구매유형

	' 구매유형이 해외직구 일경우 강제 고정
	if purchaseType="9" then
		deliverfixday = "G"	' 해외직구
		mwdiv = "U"
		deliverarea = ""

		' 업체(무료)배송 일경우
		if oitemdetail.Fdeliverytype="2" then
			deliverytype = oitemdetail.Fdeliverytype
		else
			deliverytype = "9"
		end if
	else
		deliverfixday = oitemdetail.Fdeliverfixday	' 해외직구
		mwdiv = oitemdetail.Fmwdiv
		deliverarea = oitemdetail.Fdeliverarea
		deliverytype = oitemdetail.Fdeliverytype
	end if
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
ooimage.WaitProductImageList request("itemid")  '이미지 데이터 불러오기

Dim itemaddimage,itemaddcontent, itemstoryimage

if (IsNull(ooimage.Fimgadd) or (ooimage.Fimgadd="")) then ooimage.Fimgadd = ",,,,"
if (IsNull(ooimage.Fitemaddcontent) or (ooimage.Fitemaddcontent="")) then ooimage.Fitemaddcontent = "||||"
if (IsNull(ooimage.Fimgstory) or (ooimage.Fimgstory="")) then ooimage.Fimgstory = ",,,,"


itemaddimage = split(ooimage.Fimgadd,",")
itemaddcontent = split(ooimage.Fitemaddcontent,"|")
itemstoryimage = split(ooimage.Fimgstory,",")


'==============================================================================
dim imgsubdir

imgsubdir = GetImageSubFolderByItemid(request("itemid"))


'==============================================================================
Dim npartner
Dim npt_defaultmargine, npt_defaultFreeBeasongLimit, npt_defaultDeliverPay, npt_defaultDeliveryType
Dim npt_jungsan_gubun, npt_company_no
set npartner = new CPartnerUser
npartner.FRectDesignerID = oitemdetail.FMakerid

if Not(oitemdetail.FMakerid="" or isNull(oitemdetail.FMakerid)) then
	npartner.GetOnePartnerNUser
	if npartner.FResultCount > 0 THEN
	npt_defaultmargine	 = npartner.FOneItem.Fdefaultmargine
	npt_defaultFreeBeasongLimit	= npartner.FOneItem.FdefaultFreeBeasongLimit
	npt_defaultDeliverPay	= npartner.FOneItem.FdefaultDeliverPay
	npt_defaultDeliveryType	= npartner.FOneItem.FdefaultDeliveryType
	npt_jungsan_gubun = npartner.FOneItem.Fjungsan_gubun '2014.02.14 정윤정 추가
	npt_company_no = npartner.FOneItem.Fcompany_no
	end if
end if
set npartner = Nothing

'--- 등록진행정보
 set clsWait = new CWaitItemlist2014
 	clsWait.Fitemid = itemid
 	arrList = clsWait.fnGetWaitItemLog
 	IF not isArray(arrList) THEN
 		arrOld = clsWait.fnGetOldWaitItemLog
	END IF
 set clsWait = nothing

function getFileSize(fsz)
	if fsz>1024 then
		getFileSize = formatNumber(fsz/1024,2) & "Mb"
	else
		getFileSize = fsz & "Kb"
	end if
end function
%>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style>
	FORM {display:inline;}
	#dialog {display:none; position:absolute;z-index:9100;background:#efefef; padding:10px;width:650;}
	#mask {display:none; position:absolute; left:0; top:0; z-index:9000; background:url(http://fiximage.10x10.co.kr/web2013/common/mask_bg.png) left top repeat;}
	#boxes .window {position:fixed; left:0; top:0; display:none; z-index:99999;}
</style>
<script type="text/javascript">
function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_waitItemAttribSelect.asp?itemid=<%=request("itemid")%>&arrDispCate="+arrDispCd,
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
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round로 변경
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.005) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - Math.round(isellcash*imargin/100);       //parseInt-> round로 변경
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

$(function(){
	// 로딩후 상품속성 내용 출력
	printItemAttribute();
});

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
	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";

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
function SubmitSave(processstatus) {
	var optionv="";
	var optiont = "";
	var ctrstate = "<%=oitemdetail.Fctrstate%>";
//	if (ctrstate != "7"){
//		alert("계약미완료된 브랜드는 승인이 불가능합니다.\n계약확인 후 처리해주세요");
// 	  	return;
//	}

	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

	if (processstatus==true&&!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("전시 카테고리를 선택하세요.\n※ 전시 기본 카테고리는 필수 선택되어야 합니다.");
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
	//-------------------------------------------------------------------------------- 2014.02.14 정윤정 추가
	//1.사업자가 [간이과세자] 인 경우, 매입상품 등록 불가 / 업체,위탁 상품만 등록가능
	if((itemreg.jungsangubun.value =="간이과세")&&(itemreg.mwdiv[0].checked)){
		alert("사업자가 [간이과세자]인 경우, [매입]상품은 등록불가능합니다. \n[위탁],[업체배송]상품만 등록가능합니다. ");
		itemreg.mwdiv[0].focus();
		return;
	}

	//2.사업자가 [면세사업자] 인 경우, 면세상품으로만 등록가능
	if((itemreg.jungsangubun.value =="면세")&&(itemreg.vatinclude[0].checked)){
		alert("사업자가 [면세사업자]인 경우, [과세]상품은 등록불가능합니다. \n[면세]상품만 등록가능합니다. ");
		itemreg.vatinclude[1].focus();
		return;
	}

	//3.사업자가 [텐바이텐]인 경우, 매입상품만 등록 가능
	if((itemreg.companyno.value =="211-87-00620")&&(!itemreg.mwdiv[0].checked)){
		alert("사업자가 [텐바이텐]인 경우, [매입] 상품만 등록가능합니다. ");
		itemreg.mwdiv[0].focus();
		return;
	}
	 //--------------------------------------------------------------------------------
    //배송구분 체크 =========================================================================
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
    if ((itemreg.deliverytype[1].checked)||(itemreg.deliverytype[3].checked)||(itemreg.deliverytype[4].checked)){
        if ((itemreg.mwdiv[0].checked)||(itemreg.mwdiv[1].checked)){
            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
            return;
        }
//        if (itemreg.deliverOverseas.checked){
//            alert('텐바이텐 배송일 경우에만 해외배송을 하실 수 있습니다.');
//            return;
//        }
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

	//자체배송인경우 판매안함 전시안함

	if ((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
		if (itemreg.sellyn[0].checked){
			alert('자체배송인경우 판매여부는 N로 선택하세요.');
			itemreg.sellyn[1].focus();
			return;
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

	//안전인증정보. 유예 기간을 두고 풀어줌. 차후에 다시 막아야함. 유예기간 : 2018년 1월1일??
    if (itemreg.safetyYn[0].checked){
  		if($("#real_safetynum").val() == ""){
  			alert("안전인증구분을 선택하고 인증번호를 입력후 추가버튼을 클릭해주세요.");
  			return;
  		}
    }

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

    if (itemreg.sellcash.value*1 < 200 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 200원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
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

    if (itemreg.imgmain.value == "" && (itemreg.imgmain2.value != "" || itemreg.imgmain3.value != "")) {
        alert("설명이미지는 #1부터 차례로 넣어주세요.");
        return;
    }

    if (itemreg.imgmain2.value == "" && itemreg.imgmain3.value != "") {
        alert("설명이미지는 #2부터 차례로 넣어주세요.");
        return;
    }

    if (itemreg.imgmain.value != "") {
        if (CheckImage(itemreg.imgmain, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain2.value != "") {
        if (CheckImage(itemreg.imgmain2, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain3.value != "") {
        if (CheckImage(itemreg.imgmain3, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain4.value != "") {
        if (CheckImage(itemreg.imgmain4, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain5.value != "") {
        if (CheckImage(itemreg.imgmain5, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain6.value != "") {
        if (CheckImage(itemreg.imgmain6, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }
    if (itemreg.imgmain7.value != "") {
        if (CheckImage(itemreg.imgmain7, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png',40) != true) {
            return;
        }
    }

   if(typeof(itemreg.chkSR)=="object"){
   	if(itemreg.chkSR.checked){
	    if(itemreg.dSR.value==""){
	    	alert("오픈예약이 설정되어있습니다. 날짜를 입력해주세요");
	    	itemreg.dSR.focus();
	    	return;
	    }

     if((itemreg.deliverytype[0].checked)||(itemreg.deliverytype[2].checked)){
    	alert("[텐바이텐(무료)배송]상품의 경우 오픈예약이 불가능합니다.");
    	itemreg.chkSR.focus();
    	return;
    }
 	 }
   }

    if (processstatus==true){
	    if(typeof(itemreg.chkSR)=="object"){
	    	if (itemreg.chkSR.checked) {
	    		strMsg = itemreg.dSR.value+" 오픈예약된 상품입니다.";
	    	}else{
	    		strMsg = "업체배송상품의 경우 프론트에 바로 적용되며,\n텐바이텐배송상품은 입고 완료 후 상품이 오픈됩니다.";
	    	}
	    }else{
	    		strMsg = "업체배송상품의 경우 프론트에 바로 적용되며,\n텐바이텐배송상품은 입고 완료 후 상품이 오픈됩니다.";
	    }
		if(confirm("["+itemreg.itemname.value+"]을 승인 하시겠습니까?\n"+strMsg) == true){
			<% ''안전인증 api로 조회 후 받은 데이터 db저장 후 생성idx값 받아 셋팅 %>
			if(itemreg.safetyYn[0].checked) {
				$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"u"));
			}

			itemreg.action = "<%= ItemUploadUrl %>/linkweb/items/doWaitItemToReg_byadmin.asp";
			itemreg.mode.value = "realupload";
			itemreg.itemoptioncode2.value=optionv;
			itemreg.itemoptioncode3.value=optiont;
			itemreg.target = "FrameCKP";
			itemreg.submit();
		}
	}else{
		if(confirm("상품을 임시 저장 하시겠습니까?") == true){
			<% ''안전인증 api로 조회 후 받은 데이터 db저장 후 생성idx값 받아 셋팅 %>
			if(itemreg.safetyYn[0].checked) {
				$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"u"));
			}

			itemreg.action = "<%= ItemUploadUrl %>/linkweb/items/doWaitItemToReg_byadmin.asp";
			itemreg.mode.value = "waititemmodi";
			itemreg.itemoptioncode2.value=optionv;
			itemreg.itemoptioncode3.value=optiont;
			itemreg.target = "FrameCKP";
			itemreg.submit();
		}
	}
}

// 매입위탁구분
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
	}
	else if(frm.mwdiv[2].checked){
	    // 배송구분 지정(업체배송)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// 기본 체크
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// 업체착불배송 기본 체크
	    }else{
	        frm.deliverytype[1].checked=true;	// 기본 체크
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;

        <%
        ' 해외직구 일경우
        if deliverfixday="G" then
        %>
        	frm.deliverytype[4].disabled=true;  //업체착불배송(7)
        <% else %>
			frm.deliverytype[4].disabled=false;  //업체착불배송(7)
		<% end if %>

       // frm.deliverOverseas.checked=false;	// 해외배송체크해제
	}

	if (frm.deliverytype[1].checked==true || frm.deliverytype[3].checked==true){
		frm.deliverfixday[3].disabled=false;	// 해외직구
	}
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

function TnCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readOnly=true;
		frm.limitno.className='text_ro';

		frm.limitsold.readOnly=true;
		frm.limitsold.className='text_ro';
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

		frm.limitsold.readOnly=false;
		frm.limitsold.className='text';
	}
}

function TnSilentCheckLimitYN(frm){
	if (frm.limityn[0].checked == true) {
		// 비한정
		frm.limitno.readOnly=true;
		frm.limitno.className='text_ro';

		frm.limitsold.readOnly=true;
		frm.limitsold.className='text_ro';
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

		frm.limitsold.readOnly=false;
		frm.limitsold.className='text';
	}
}

function TnGoClear(frm){
	frm.sellvat.value = "";
	frm.buycash.value = "";
	frm.buyvat.value = "";
	frm.mileage.value = "";
}

// 배송구분
function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if (frm.mwdiv[2].checked){
			alert("매입위탁 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입위탁구분을 확인해주세요!!");
			frm.mwdiv[0].checked=true;
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
		if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
			alert("매입위탁 구분이 매입이나 위탁일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입위탁구분을 확인해주세요!!");
			frm.mwdiv[2].checked=true;
		}
	}

	if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)&&(!frm.deliverytype[3].checked)){
	    alert('업체 조건 배송 업체입니다. 배송구분을 확인하세요.')
	    frm.deliverytype[3].focus();
	}

	if (((frm.defaultFreeBeasongLimit.value*1<1)||(frm.defaultDeliverPay.value*1<1))&&(frm.deliverytype[3].checked)){
	    alert('업체 조건 배송 업체가 아닙니다. 배송구분을 확인하세요.')
	    frm.deliverytype[3].focus();
	}

	if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(frm.deliverytype[4].checked)){
	    alert('업체 착불 배송 업체가 아닙니다. 배송구분을 확인하세요.')
	    frm.deliverytype[4].focus();
	}
}
function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // 옵션사용
	    if (confirm("옵션은 같은가격/디자인에 색상 또는 사이즈가 2개 이상인경우 사용 가능합니다. 진행하시겠습니까?") == true) {
            frm.btnoptadd.disabled = false;
            frm.btnetcoptadd.disabled = false;
            frm.btnoptdel.disabled = false;

            optlist.style.display="";
        } else {
            frm.useoptionyn[1].checked = true;
            TnCheckOptionYN(frm);
        }
	} else {
	    // 옵션없음
	    // while (frm.realopt.length > 0) {
	    //     frm.realopt.options[0] = null;
        // }
        frm.btnoptadd.disabled = true;
        frm.btnetcoptadd.disabled = true;
        frm.btnoptdel.disabled = true;

		optlist.style.display="none";

        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";
    }
}
function TnCheckSailYN(frm){
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
    CalcuAuto(frm);
}


//색상코드 선택
function selColorChip(cd) {
	var i;
	itemreg.DFcolorCD.value= cd;
	for(i=0;i<=30;i++) {
		document.all("cline"+i).bgColor='#DDDDDD';
	}
	if(!cd) document.all("cline0").bgColor='#DD3300';
	else document.all("cline"+cd).bgColor='#DD3300';
}


// ============================================================================
// 미리보기
function ViewItemDetail(itemno){
	window.open('/designer/itemmaster/viewitem.asp?itemid='+itemno ,'window1','width=790,height=600,scrollbars=yes,status=no');
}


function jsUniWaitState(currstate, count2){
	if(typeof(itemreg.chkSR) == "object"){
		if(itemreg.chkSR.checked){
			alert("상품을 승인하지 않으면, 예약된 날짜에 상품이 오픈될 수 없습니다.\n반려 또는 보류를 하시려면 상품오픈예약 설정을 취소해주세요.");
			itemreg.chkSR.focus();
			return;
		}
	}

 $("#dv2").hide();
 $("#dv0").hide();
	if(count2>=2){
		 document.all.chkV0[4].checked = true;
		 document.all.sM0.value = "3회 이상 보류시, 반려처리(재등록 불가)됩니다.";
	}
	var maskHeight = $(document).height();
	var maskWidth = $(document).width();

	$('#mask').css({'width':maskWidth,'height':maskHeight});
	$('#boxes').show();
	$('#mask').show();
		var winH = $(document).height()-800;
		var winW = $(document).width();
		$("#dialog").css('top', winH-$("#dialog").height());
		$("#dialog").css('left', winW/2-$("#dialog").width()/2);
		$("#dialog").show();
		$("#dv"+currstate).show();
}

//승인처리
 function jsConfirm(currstate){
 	var chkCount = 0;
 	var iMsgcd = "";
 	var sMsg = "";
 	var iNo = "";
 	for(i=0;i<eval("document.all.chkV"+currstate).length;i++){
 		if(eval("document.all.chkV"+currstate)[i].checked){
 		chkCount = chkCount + 1;
 		iNo = eval("document.all.chkV"+currstate)[i].value;
 		if (iMsgcd==""){
 			iMsgcd = eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = $("#sp"+currstate+iNo).text();
 			}
 		}else{
 		iMsgcd = iMsgcd +"^"+ eval("document.all.chkV"+currstate)[i].value;
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = sMsg +"^"+eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = sMsg +"^"+ $("#sp"+currstate+iNo).text();
 			}
 		}
 	}
 	}
 	if(chkCount == 0){
 		alert("승인 거부 사유를 한개 이상 선택해주세요");
 		return;
 	}
 	document.borufrm.sMsgcd.value= iMsgcd;
 	document.borufrm.sMsg.value = sMsg;
 	document.borufrm.hidM.value = "U";
 	document.borufrm.sCS.value = currstate;
  document.borufrm.submit();
}

  function jsCancel(){
  	document.borufrm.sMsgcd.value= "";
 		document.borufrm.sMsg.value = "";
  	 $( "#dialog" ).hide();
  	 $('#mask').hide();
  	 $('#boxes').hide();
  }

function GetRejectMsg(falg){
    var tmp = window.showModalDialog('pop_rejectMsg.asp?falg=' + falg,null, "dialogwidth:250px;dialogheight:230px;center:yes;scroll:no;resizable:no;status:no;help:no;");
    return tmp;
}

function ClearVal(comp){
    comp.value = "";
}

function popWaitItemOptionEdit(iitemid){
    var popwin = window.open('/common/pop_upchewaititemoptionedit.asp?itemid=' + iitemid,'popWaitItemOptionEdit','width=790,height=600,scrollbars=yes,status=no');
    popwin.focus();
}

function EnDisableFlowerShop(){
    var frm = document.itemreg;
    if ((frm.cd1.value=="110")&&(frm.cd2.value=="060")){
        frm.deliverarea[1].disabled = false;
        frm.deliverarea[2].disabled = false;

        deliverfixday.disabled = false;
    }else{
        frm.deliverarea[1].disabled = true;
        frm.deliverarea[2].disabled = true;

        deliverfixday.disabled = true;
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

//품목 선택 / 품목내용 표시
function chgInfoDiv(v) {
	$("#itemInfoList").empty();

	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "act_waitItemInfoDivForm.asp",
			data: "itemid=<%=request("itemid")%>&ifdv="+v,
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

	//달력
	function jsPopCal(sName){
	 if(!document.all.chkSR.checked){
	 	 document.all.chkSR.checked= true;
	 	}
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();

	}

	//오픈예약
	function jsChkSellReserve(){
		if(!document.all.chkSR.checked){
			document.all.dSR.value = "";
		}
	}


//상품군에 따른 원산지 설명 표기
function jsSetArea(iValue){
	var i;
	for(i=0;i<=4;i++) {
 		eval("document.all.dvArea"+i).style.display = "none";
	}
 	eval("document.all.dvArea"+iValue).style.display = "";
}

//안전인증 추가 버튼 액션
function jsSafetyAuth(){
	//var cnum = $("#safetyNum").val();
	var cnum = itemreg.safetyNum.value.ltrim().rtrim();
	var listbody = "";
	var safetyvalue = "";
	var safetynum = "";

	if(typeof itemreg.catecode == "undefined"){
		alert("카테고리를 선택해 주세요.");
		return;
	}

	if($("#safetyDiv").val() == ""){
		alert("안전인증구분을 선택해 주세요.");
		return;
	}

	var isExist = $("#real_safetydiv").attr("value").indexOf($("#safetyDiv").val()) > -1;
	if(isExist){
		alert("이미 선택된 안전인증구분 입니다.");
		return;
	}
//	var isExistsafetynum = $("#real_safetynum").attr("value").indexOf(cnum) > -1;
//	if(isExistsafetynum){
//		alert("이미 선택된 안전인증번호 입니다.");
//		return;
//	}

	if($("#safetyDiv").val() == "30" || $("#safetyDiv").val() == "60" || $("#safetyDiv").val() == "90"){
		$("#issafetyauth").val("ok");

		safetyvalue = $("#real_safetydiv").val();
		if(safetyvalue == ""){
			$("#real_safetydiv").val($("#safetyDiv").val());
		}else{
			$("#real_safetydiv").val(safetyvalue + "," + $("#safetyDiv").val())
		}

		safetynum = $("#real_safetynum").val();
		if(safetynum == ""){
			$("#real_safetynum").val("x");
		}else{
			$("#real_safetynum").val(safetynum + "," + "x");
		}


		listbody = $("#safetyDivList").html();
		$("#safetyDivList").html(listbody + "<p id='l"+$("#safetyDiv").val()+"'>- " + $("#safetyDiv option:selected").text() + "(인증번호 없음) <input type='button' value='삭제' onClick='jsSafetyDivListDel("+$("#safetyDiv").val()+");' class='button'><p>");
	}else{

		var msgg = jsCallAPIsafety(cnum,"x");

		if(msgg == "적합" || msgg == "변경" || msgg == "개선명령" || msgg == "청문실시"){
			$("#issafetyauth").val("ok");

			safetyvalue = $("#real_safetydiv").val();
			if(safetyvalue == ""){
				$("#real_safetydiv").val($("#safetyDiv").val());
			}else{
				$("#real_safetydiv").val(safetyvalue + "," + $("#safetyDiv").val())
			}

			safetynum = $("#real_safetynum").val();
			if(safetynum == ""){
				$("#real_safetynum").val(cnum);
			}else{
				$("#real_safetynum").val(safetynum + "," + cnum);
			}


			listbody = $("#safetyDivList").html();
			$("#safetyDivList").html(listbody + "<p id='l"+$("#safetyDiv").val()+"'>- " + $("#safetyDiv option:selected").text() + "("+cnum+") <input type='button' value='삭제' onClick='jsSafetyDivListDel("+$("#safetyDiv").val()+");' class='button'><p>");
		}else{
			alert("인증번호의 현재 상태 : " + msgg);
			return;
		}
	}
	jsSafetyDefault();
}

function jsCallAPIsafety(certnum,isSave){
	var returnmsg = "";
	$.ajax({
		url: "/admin/itemmaster/safety_api_auth_proc.asp?itemid=<%=itemid%>&issave="+isSave+"&certnum="+certnum+"&statusmode=wait",
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
		<font color="red"><strong>승인대기상품정보</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>승인대기중인 상품을 정식등록합니다.</b>
			<br><br>- 잘못된 부분은 임시저장 기능을 이용하여 수정하실수 있습니다.
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
<form name="itemreg" method="post" action="" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemid" value="<%= request("itemid") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="defaultmaeipdiv" value="<%= npt_defaultmargine %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npt_defaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= npt_defaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= npt_defaultDeliveryType %>">
<input type="hidden" name="DFcolorCD" value="<%=oitemdetail.FDFcolorCd%>">
<input type="hidden" name="jungsangubun" value="<%=npt_jungsan_gubun%>">
<input type="hidden" name="companyno" value="<%=npt_company_no%>">

<input type="hidden" name="pojangok" value="Y">
<input type="hidden" name="itemoptioncode2" value="">
<input type="hidden" name="itemoptioncode3" value="">
<input type="hidden" name="isusing" value="Y">
<input type="hidden" name="adminid" value="<%=session("ssBctId") %>">
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
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td  height="30" width="15%" bgcolor="#DDDDFF">브랜드 계약상태</td>
	<td bgcolor="#FFFFFF" colspan="3"><% IF oitemdetail.Fctrstate = "7" then%>계약완료<%else%>미계약<%end if%></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designerid"  value="<%= oitemdetail.FMakerid %>" class="text_ro" readonly size="30" id="[on,off,off,off][브랜드ID]">&nbsp;
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= request("itemid") %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="미리보기" onclick="ViewItemDetail('<%= request("itemid") %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="50" class="text" value="<%= Replace(oitemdetail.Fitemname,"""","&quot;") %>" id="[on,off,off,off][상품명]">&nbsp;
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
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td id="lyrDispList"><%=getDispCategoryWait(request("itemid"))%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
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
		<label><input type="radio" name="itemdiv" value="23" <%=chkIIF(oitemdetail.Fitemdiv ="23","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B상품</label>
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
		<input type="text" name="margin" maxlength="32" size="5" id="[off,off,off,off][마진]" value="<%= oitemdetail.FMargin %>" class="text">%
		<% if (CStr(npt_defaultmargine)<>CStr(oitemdetail.FMargin)) then %>
		<font Color="red">(업체 기본 마진 : <%= npt_defaultmargine %>)</font>
		<% end if %>
	</td>
</tr>
<tr align="left">
<input type="hidden" name="sellvat" value="<%= oitemdetail.Fsellvat %>">
<input type="hidden" name="buyvat" value="<%= oitemdetail.Fbuyvat %>">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매가(소비자가) :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="sellcash" size="12" id="[on,on,off,off][소비자가]" onKeyUp="CalcuAuto(itemreg);" maxlength="8" class="text" value="<%= oitemdetail.Fsellcash %>">원
		<input type="button" value="공급가 자동계산" onclick="CalcuAuto(itemreg);" class="button" style="width:100px;">
	</td>
	<td width="15%" bgcolor="#DDDDFF">공급가 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<input type="text" name="buycash" maxlength="16" size="12" id="[on,on,off,off][공급가]" class="text" value="<%= oitemdetail.Fbuycash %>">원
		(<b>부가세 포함가</b>)
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="3">
		- 공급가는 <b>부가세 포함가</b>입니다.
		<br>- 판매가(소비자가)를 입력하면 지정된 마진으로 공급가가 자동계산됩니다.
		<br>- 별도로 마진과 상관없이 공급가를 입력하실수 있습니다.
	</td>
</tr>
<input type="hidden" name="mileage" id="[on,off,off,off][마일리지]" value="<%= oitemdetail.Fmileage %>">
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
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3">지정된 전시카테고리가 없습니다.</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="upchemanagecode" value="<%= oitemdetail.Fupchemanagecode %>" size="20" maxlength="32" class="text" id="[off,off,off,off][업체상품코드]">
	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="<%= oitemdetail.Fisbn13 %>" size="13" maxlength="13">
		/ 부가기호 <input type="text" name="isbn_sub" class="text" value="<%= oitemdetail.FisbnSub %>" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="<%= oitemdetail.Fisbn10 %>" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매여부 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y" <% if oitemdetail.Fsellyn = "Y" then response.write "checked" %>>판매함</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" <% if oitemdetail.Fsellyn = "N" then response.write "checked" %>>판매안함</label>
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
	  <p>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" <%if isNull(oitemdetail.Fsourcekind) or oitemdetail.Fsourcekind="0" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 식품 외</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1" <%if oitemdetail.Fsourcekind="1" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 농축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2" <%if oitemdetail.Fsourcekind="2" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 수산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3" <%if oitemdetail.Fsourcekind="3" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4" <%if oitemdetail.Fsourcekind="4" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 농수산가공품</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][원산지]"  value="<%= oitemdetail.Fsourcearea %>"/></p>
	  <div id="dvArea0" style="display:<%if isNull(oitemdetail.Fsourcekind) or oitemdetail.Fsourcekind="0" then%>block<%else%>none<%end if%>;">
	  <p><strong>ex: 한국, 중국, 중국OEM, 일본 등 </strong></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea1" style="display:<%if oitemdetail.Fsourcekind ="1" then%>block<%else%>none<%end if%>;">
	  <p><strong>국내산 :</strong> 국산, 국내산 또는 시·도명, 시·군명(대한민국, 한국X)  <span style="margin-right:10px;">ex. 쌀(국산)</span></BR>
	   <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 곶감(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea2" style="display:<%if oitemdetail.Fsourcekind ="2" then%>block<%else%>none<%end if%>;">
	  <p><strong>국내산 :</strong> 국산,국내산 또는 연근해산(양식 수산물은 시·군명 가능)   <span style="margin-right:10px;">ex. 갈치(국산), 오징어(연근해산)</span> </BR>
	  	<strong>원양산 :</strong> 원양산 또는 원양산(해역명)   <span style="margin-right:10px;">ex. 참치[원양산(대서양)]</span> </BR>
	    <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 농어(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea3" style="display:<%if oitemdetail.Fsourcekind ="3" then%>block<%else%>none<%end if%>;">
	  <p>소고기의 경우 식육의 종류(한우/육우/젖소구분) 및 원산지   <span style="margin-right:10px;">ex. 쇠고기(횡성산 한우), 쇠고기(호주산)</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea4" style="display:<%if oitemdetail.Fsourcekind ="4" then%>block<%else%>none<%end if%>;">
	  <p><strong>98%이상 원료가 있는 경우:</strong>  한가지 원료만 표시 가능    <span style="margin-right:10px;">ex. 쇠고기(미국산)</span> </BR>
	  	<strong>복합 원료를 사용한 경우:</strong> 혼합비율이 높은 순으로 2개 국가   <span style="margin-right:10px;">ex. 고추장[밀가루(미국산),고춧가루(국내산)]</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][상품무게]" style="text-align:right" value="<%= oitemdetail.Fitemweight %>">g &nbsp;(그램단위로 입력, ex:1.5kg→ 1500) / 해외배송시 배송비 산출을 위한 것이므로 정확히 입력.
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
		<% DrawInfoDiv "infoDiv", oitemdetail.FinfoDiv, " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitemdetail.FinfoDiv) or oitemdetail.FinfoDiv="") then
			Server.Execute("act_waitItemInfoDivForm.asp")
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
<%
dim arrAuth, r, real_safetydiv, real_safetynum, safetyDivList
arrAuth = oitemdetail.FAuthInfo
if isArray(arrAuth) THEN
	For r =0 To UBound(arrAuth,2)
		real_safetydiv = real_safetydiv & arrAuth(0,r)
		if r <> UBound(arrAuth,2) then real_safetydiv = real_safetydiv & "," end if

		real_safetynum = real_safetynum & arrAuth(1,r)
		if r <> UBound(arrAuth,2) then real_safetynum = real_safetynum & "," end if

		safetyDivList = safetyDivList & "<p class='tPad05' id='l"&arrAuth(0,r)&"'>"
		safetyDivList = safetyDivList & "- "&fnSafetyDivCodeName(arrAuth(0,r))&"("&CHKIIF(arrAuth(1,r)="x","인증번호 없음",arrAuth(1,r))&")"
		safetyDivList = safetyDivList & " <input type='button' value='삭제' class='btn3 btnIntb' onClick='jsSafetyDivListDel("&arrAuth(0,r)&");'>"
		safetyDivList = safetyDivList & "</p>"
	Next
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 안전인증정보</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		안전인증대상 :
		<input type="button" value="안전인증 필수 품목 확인" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitemdetail.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 대상</label>
				<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitemdetail.FsafetyYn="N","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 대상아님</label>
				<label><input type="radio" name="safetyYn" value="I" <%=chkIIF(oitemdetail.FsafetyYn="I","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 상품설명에 표기</label>
				<label><input type="radio" name="safetyYn" value="S" <%=chkIIF(oitemdetail.FsafetyYn="S","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 안전기준준수</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="<%=real_safetydiv%>">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="<%=real_safetynum%>">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
				<input type="hidden" name="real_safetynum_delete" id="real_safetynum_delete" value="">
				<input type="hidden" name="real_safetydiv_delete" id="real_safetydiv_delete" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", oitemdetail.FsafetyYn, "" %>
				인증번호 <input type="text" name="safetyNum" id="[off,off,off,off][안전인증 인증번호]" <%=chkIIF(oitemdetail.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" value="" /><%'=oitemdetail.FsafetyNum%>
				<input type="button" id="safetybtn" value="추   가" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList">
					<%=safetyDivList%>
				</div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">상품 설명에 표기(표기대상 상품인경우 상품 상세 페이지에 인증번호와 모델명, KC 마크를 꼭 표기해주세요.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* 인증정보를 입력 안 하거나, 잘못된 인증정보를 입력한 경우 발견 <strong><font color='red'>즉시 판매정지 또는 삭제</font></strong> 됩니다.<br>
		* <strong><font color='red'>안전기준준수</font></strong> 대상일경우 인증번호가 없으며, KC마크를 표시하지 않아야 됩니다.<br>
		* 입력한 인증정보는 제품안전정보센터에서 제공된 정보를 기준으로 조회되며, <strong><font color='red'>검증되지 않은 정보는 등록이 불가</font></strong>능합니다.<br>
		* 정상적인 인증정보를 입력했음에도 불구하고 등록이 안될경우에 "상품설명에 표기"로 설정이 가능하며, 상품 상세 페이지에 모델명과 표기대상 상품인경우 인증번호,KC마크를 표기해야 합니다.<br>
		* 안전인증정보 관련 문의는 홈페이지(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)로 확인해 주시기 바랍니다.
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
	<td height="30" width="15%" bgcolor="#DDDDFF">매입위탁구분 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(mwdiv="M","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >매입
	  <input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(mwdiv="W","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >위탁
	  <input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <%=chkIIF(mwdiv="U","checked","")%>>업체배송
	  &nbsp;&nbsp; - 매입위탁구분에 따라 배송구분이 달라집니다. 배송구분을 확인해주세요.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="1","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >텐바이텐배송&nbsp;
	  <input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="2","checked","")%>>업체(무료)배송&nbsp;
	  <input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="4","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >텐바이텐무료배송&nbsp;
	  <input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="9","checked","")%>>업체조건배송(개별 배송비부과)&nbsp;
	  <input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <%=chkIIF(oitemdetail.Fdeilverytype="7","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >업체착불배송
	 </td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송방법 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverfixday" value="" onclick="TnCheckFixday(this.form)" <%=chkIIF(Trim(deliverfixday)="" or IsNull(deliverfixday),"checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >택배(일반)&nbsp;
	  <input type="radio" name="deliverfixday" value="X" onclick="TnCheckFixday(this.form)" <%=chkIIF(deliverfixday="X","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >화물&nbsp;
	  <input type="radio" name="deliverfixday" value="C" onclick="TnCheckFixday(this.form)" <%=chkIIF(deliverfixday="C","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >플라워지정일
	  <input type="radio" name="deliverfixday" value="G" onclick="TnCheckFixday(this.form)" <%=chkIIF(deliverfixday="G","checked","")%> <%=chkIIF(mwdiv<>"U" or (deliverytype <> "2" and deliverytype <> "9")," disabled","")%> >해외직구
		<span id="lyrFreightRng" style="display:<%=chkIIF(deliverfixday="X","","none")%>;">
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
	  <input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(deliverarea)="" or IsNull(deliverarea),"checked","")%>>전국배송&nbsp;
	  <input type="radio" name="deliverarea" value="C" <%=chkIIF(deliverarea="C","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >수도권배송&nbsp;
	  <input type="radio" name="deliverarea" value="S" <%=chkIIF(deliverarea="S","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >서울배송&nbsp;
 	  <input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF(oitemdetail.FdeliverOverseas="Y","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> title="해외배송은 상품무게가 입력이 돼야 완료됩니다.">해외배송
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
      - 옵션은 정식등록후 삭제가 불가능합니다. 정확히 입력하세요.
	</td>
</tr>
<tr id="lyDFColor" height="30">
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
		      - 색상별 이미지가 없으면 정식등록이 되지않습니다.(Err:013) 정식등록시에 반드시 등록해주세요.
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
	  <% end if %>
	  <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>필수</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="basic">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품<br />흰배경(누끼)이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (ooimage.Fimgmask <> "") then %>
		<div id="divimgmask" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/mask/<%=imgsubdir%>/<%= ooimage.Fimgmask %>" width="300" height="300">
		</div>
	  <% end if %>
	  <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (선택,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="mask">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 0 then %>
	  <% if (itemaddimage(0) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add1/<%=imgsubdir%>/<%=itemaddimage(0) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 1 then %>
	  <% if (itemaddimage(1) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add2/<%=imgsubdir%>/<%=itemaddimage(1) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 2 then %>
	  <% if (itemaddimage(2) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add3/<%=imgsubdir%>/<%=itemaddimage(2) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 3 then %>
	  <% if (itemaddimage(3) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add4/<%=imgsubdir%>/<%=itemaddimage(3) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 추가이미지5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if ubound(itemaddimage) >= 4 then %>
	  <% if (itemaddimage(4) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/add5/<%=imgsubdir%>/<%=itemaddimage(4) %>" width="300" height="300">
		</div>
	  <% end if %>
		<% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 기존의 제품설명이미지는 사용하지 않고 상품설명이미지를 사용합니다. 기존에 등록된 제품설명이미지는 사용은 하되 추가 수정은 되지않고 삭제만 됩니다.</strong></font>
 	</td>
 </tr>
<%
			'상품설명 이미지
			mainImg(1) = ooimage.Fimgmain
			mainImg(2) = ooimage.Fimgmain2
			mainImg(3) = ooimage.Fimgmain3
			mainImg(4) = ooimage.Fimgmain4
			mainImg(5) = ooimage.Fimgmain5
			mainImg(6) = ooimage.Fimgmain6
			mainImg(7) = ooimage.Fimgmain7
			mainImg(8) = ooimage.Fimgmain8
			mainImg(9) = ooimage.Fimgmain9
			mainImg(10) = ooimage.Fimgmain10

			for i=1 to 7
%>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #<%=i%> :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (mainImg(i) <> "") then %>
		<div id="divimgmain<%=chkIIF(i>1,i,"")%>" style="display:block;">
		<img src="<%=partnerUrl%>/waitimage/main<%=chkIIF(i>1,i,"")%>/<%=imgsubdir%>/<%=mainImg(i) %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain<%=chkIIF(i>1,i,"")%>" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgmain<%=chkIIF(i>1,i,"")%>" onchange="CheckImage(this, <%= CMAIN_IMG_MAXSIZE %>, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>, 'jpg,gif,png', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmain<%=chkIIF(i>1,i,"")%>, 40, <%=CMAIN_IMG_MAXWIDTH%>, <%=CMAIN_IMG_MAXHEIGHT%>)"> (선택, 너비 <%=CMAIN_IMG_MAXWIDTH%>px, Max <%= getFileSize(CMAIN_IMG_MAXSIZE) %>, jpg,gif,png)
		<input type="hidden" name="main<%=chkIIF(i>1,i,"")%>">
	</td>
</tr>
<% next %>
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
		<input type="file" name="mobileimgmain" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain2" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain3" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain4" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain5" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain6" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain7" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain8" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain9" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain10" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain11" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
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
		<input type="file" name="mobileimgmain12" onchange="CheckImage(this, <%= CMOBILE_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif', 40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage2('mobile12', 40, 640, 1200)"> (선택,640X1200, Max <%= CMOBILE_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="mobile12">
	</td>
</tr>

<!--// 20160601추가한부분 -->

</table>


<!-- 11.등록정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>11.등록정보</strong> </td>
    <td align="right">3회 이상 보류시, 반려 처리(재등록불가)되므로 참고 부탁 드립니다.</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	 <tr>
	 	<td width="15%" align="center" bgcolor="#DDDDFF">진행일자</td>
	 	<td bgcolor="#FFFFFF">
	 		<%IF isARray(arrList) THEN%>
	 <% dim count2, strMsg, sMsgCd2, sMsgcd0
	 count2 = 0
	 sMsgCd2 = ""
	 sMsgcd0 = ""
	 For intLoop = 0 To UBound(arrList,2)
	 strMsg = ""
	 		IF arrList(2,intLoop) = "2" THEN
	 			count2 = count2 + 1
	 			strMsg = count2&"차"
	 			sMsgCd2 = sMsgCd2 + "^" + arrList(6,intLoop)
	 		ELSEIF arrList(2,intLoop) = "0" THEN
	 				sMsgCd0 = sMsgCd0 + "^" + arrList(6,intLoop)
	 		END IF
	 %>
	 <div style="padding:3"><font color="<%=GetCurrStateColor(arrList(2,intLoop))%>"><%=strMsg%><%=fnGetCurrStateShortName(arrList(2,intLoop))%></font>: <%=arrList(4,intLoop)%> &nbsp;<%IF arrList(3,intLoop) <> "" THEN%>[<%=replace(arrList(3,intLoop),"^","/")%>]<%END IF%></div>
	 <%Next%>
	  <%ELSEIF isArray(arrold) THEN
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

 <% if oitemdetail.FCurrState="1" or oitemdetail.FCurrState="5" then %>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
	 <tr>
	 	  <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
			<td style="padding:5px">
				<input type="checkbox" name="chkSR" value="Y" onClick="jsChkSellReserve();"> 상품오픈예약:
				<input type="text" name="dSR" value="" size="10" class="input"   onClick="jsPopCal('dSR');">
				<input type="image" name="imgSR" src="/images/admin_calendar.png" onClick="jsPopCal('dSR');"  >
				  상품 승인이 되지 않았거나, 사용안함 상태일 경우 예약된 시간에 오픈이 되지 않습니다.
			   텐바이텐 배송일 경우, 입고 확인 후 오픈예약이 가능합니다.
				</td>
			<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tR>
</table>
<% end if %>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
    	<input type="button" value="임시저장" class="button" onclick="SubmitSave(false);" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %>>
       <%IF count2<2 THEN%>&nbsp;&nbsp;&nbsp;
      <input type="button" value="승인보류 (재등록요청)" class="button" onclick="jsUniWaitState(2,'<%=count2%>');" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %>>
      <%END IF%>
      &nbsp;&nbsp;&nbsp;
      <input type="button" value="승인반려 (재등록불가)" class="button" onclick="jsUniWaitState(0,'<%=count2%>');" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %>>
			&nbsp;&nbsp;&nbsp;
			<input type="button" value="승인" class="button" onclick="SubmitSave(true);" <% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then response.write "disabled" %> >
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
</form>
<!-- 표 하단바 끝-->
<form name="borufrm" method="post" action="doitemregboru.asp">
<input type="hidden" name="hidM" value="U">
<input type="hidden" name="itemid" value="<%= request("itemid") %>">
<input type="hidden" name="sCS" value="">
<input type="hidden" name="sMsgcd" value="">
<input type="hidden" name="sMsg" value="">
<input type="hidden" name="sRU" value="wait_item_modify.asp?itemid=<%=itemid%>&designer=<%=makerid%>">
</form>

<% if application("Svr_Info")	= "Dev" then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="600"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<div id="boxes"></div>
<div id="mask"></div>
<div id="dialog">
<!-- #include virtual="/admin/itemmaster/item_confirm_inc.asp"-->
</div>
<script type="text/javascript">
// 매입위탁구분 및 배송구분세팅
TnCheckUpcheYN(itemreg);
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements[i].name == "deliverytype") {
        if (itemreg.elements[i].value == "<%= oitemdetail.Fdeilverytype %>") {
            itemreg.elements[i].checked = true;
        }
    }
}

// 한정
TnSilentCheckLimitYN(itemreg);
// 세일
// TnCheckSailYN(itemreg);

<% if oitemdetail.FCurrState<>"1" and oitemdetail.FCurrState<>"5" then %>
alert('승인 대기 상태가 아닙니다.');
<% end if %>

	// 안전인증체크. 전안법
	jsSafetyCheck('<%= oitemdetail.FsafetyYn %>','');
</script>
<%
set oitemdetail = Nothing
set oitemreg = Nothing
set ooimage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
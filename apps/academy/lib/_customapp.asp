<script>
<!--
//기존 : callNativeFunction('호출할함수명',Jsonobject);
//변경 : webkit.messageHandlers.호출할함수명.postMessage(jsonobject);

function callNativeFunction(funcname, args) {
    if ( !args ) { args = {} }
    args['funcname'] = funcname;
    registerCallback(funcname, args);
	//alert(JSON.stringify(args));
	<% if flgDevice = "I" or flgDevicePC = "M" then %>
		if(funcname=="uploadImage"){
			webkit.messageHandlers.uploadImage.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="openerJsCallClose"){
			webkit.messageHandlers.openerJsCallClose.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="popupBrowser"){
			webkit.messageHandlers.popupBrowser.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="activateConfirmRightBtns"){
			webkit.messageHandlers.activateConfirmRightBtns.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="activateRegisterRightBtns"){
			webkit.messageHandlers.activateRegisterRightBtns.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else{
			eval("window.webkit.messageHandlers." + funcname + ".postMessage(" + encodeURIComponent(JSON.stringify(args)) + ")");
		}
	<% else %>
		window.location = 'callNativeFunction:' + encodeURIComponent(JSON.stringify(args));
	<% end if %>
}

var _callbacks = [];

function registerCallback(funcname, args) {
    if ( !args['callback'] ) { return; }
    _callbacks[funcname] = args['callback'];
    delete args['callback'];
}

function callback(funcname, jsonString) {
    _callbacks[funcname](JSON.parse(decodeURIComponent(jsonString)));
}

//FROM_LNB 추가 2014/09/20
var OpenType = {
    FROM_RIGHT: "OPEN_TYPE__FROM_RIGHT",
    FROM_BOTTOM: "OPEN_TYPE__FROM_BOTTOM"
}

var BtnType = {
    CONFIRM: "BTN_TYPE__CONFIRM",
    REGISTER: "BTN_TYPE__REGISTER"
}


//현재 팝업 닫기
function fnAPPclosePopup(){
    callNativeFunction('closePopup');
    return false;
}

//팝업
function fnAPPpopupBrowser(openType, leftToolBarBtns, title, rightToolBarBtns, iurl, pageType) {
    if (!pageType) pageType="";
    callNativeFunction('popupBrowser', {
    	"openType": openType,
    	"ltbs": leftToolBarBtns,
    	"title": title,
    	"rtbs": rightToolBarBtns,
    	"url": iurl,
    	"pageType": pageType
    });
    return false;
}

//현재 창 닫으면서 오픈 창 Js호출
function fnAPPopenerJsCallClose(jsfunc){
    callNativeFunction('openerJsCallClose', {"jsfunc": jsfunc});
    return false;    
}

//카테고리 선택 팝업
function fnAPPpopupCategory(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "카테고리 설정", [BtnType.CONFIRM], url, "category");
	return false;
}

//동영상 삽입 팝업
function fnAPPpopupVod(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "동영상 삽입", [BtnType.CONFIRM], url, "vod");
	return false;
}

//제작 특이사항 입력 팝업
function fnAPPpopupReqContents(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "제작 특이사항 입력", [BtnType.CONFIRM], url, "requirecontents");
	return false;
}

//옵션 셋팅 팝업
function fnAPPpopupOptionSet(querystring,optlevel){
	var title;
	if(optlevel==1){
		title="단일 옵션 설정"
		url = "<%=g_AdminURL%>/apps/academy/itemmaster/popOptionSet.asp?"+querystring;
	}else{
		title="이중 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popOptionMulti.asp?"+querystring;
	}
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], title, [BtnType.CONFIRM], url, "optionsetting");
	return false;
}

//검색 키워드 입력 팝업
function fnAPPpopupKeyWord(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "검색 키워드 입력", [BtnType.CONFIRM], url, "keyword");
	return false;
}

//안전인증 대상 입력 팝업
function fnAPPpopupSafeInfo(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "안전인증 대상", [BtnType.CONFIRM], url, "safeinfo");
	return false;
}

//상품정보 제공 고시 팝업
function fnAPPpopupItemInfoDiv(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "상품정보제공고시", [BtnType.CONFIRM], url, "iteminfo");
	return false;
}

//배송비 안내 입력 팝업
function fnAPPpopupDeliveryInfo(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "배송비 안내 입력", [BtnType.CONFIRM], url, "deliveryinfo");
	return false;
}

//상품리스트 필터 팝업
function fnAPPpopupSearchFilter(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "필터", [BtnType.CONFIRM], url, "searchfilter");
	return false;
}

//상품 정보 팝업
function fnAPPpopupItemDetail(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "상품 정보", [BtnType.CONFIRM], url, "itemdetail");
	return false;
}

//상품 가격 변경 요청 팝업
function fnAPPpopupItemPriceEdit(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "가격 변경 요청", [BtnType.CONFIRM], url, "itempriceedit");
	return false;
}

//이중 옵션 창의 입력 팝업
function fnAPPpopupMultiOption(querystring,div){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "이중 옵션 설정", [BtnType.CONFIRM], "<%=g_AdminURL%>/apps/academy/itemmaster/popOptionMultiInput.asp?div="+div+"&"+querystring, "multiOption");
	return false;
}

var _selComp;
var _selTarget;

function fnAPPuploadImage(comp) {
	_selComp = comp;
	var paramname = comp.name;

	var upurl = "<%=uploadUrl2%>/linkweb/academy/items/WaitDIYItemRegister_Process_App.asp?paramname="+paramname;

	if (paramname=="imgbasic"){
		callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadFinish1});
	}else if(paramname=="imgadd1"){
		callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadFinish2});
	}else if(paramname=="imgadd2"){
		callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadFinish3});
	}
	return false;
}
function _appUploadFinish(ret,ino){
    if (_selComp){
        _selComp.value=ret.name;
        $('#imgspan'+ino).empty();
		$('#imgspan'+ino).css("background-image", "url(<%=fingersImgUrl%>/diyItem/waitimage/basic/07/" + ret.name + ")");
    }
}
function appUploadFinish1(ret){
    _appUploadFinish(ret,1);
}
function appUploadFinish2(ret){
    _appUploadFinish(ret,2);
}
function appUploadFinish3(ret){
    _appUploadFinish(ret,3);
}

function fnAPPuploadAddImage(comp,target) {
    _selComp = comp;
	_selTarget = target;
    var paramname = comp.name;
    var upurl = "<%=uploadUrl2%>/linkweb/academy/items/WaitDIYItemRegister_Process_App.asp?paramname="+paramname;
    callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback": appUploadAddFinish});
    return false;
}

function appUploadAddFinish(ret){
    if (_selComp){
		$("#" + _selTarget + " #addimgname").val(ret.name);
        $("#" + _selTarget + " button").remove();
		$("#" + _selTarget + " #imgArea").append("<img src=<%=uploadUrl2%>/diyItem/waitimage/basic/07/" + ret.name + "' />");
    }
}

function fnAPPpopupTest(url,winid) {
	var popwin = window.open(url,winid,"width=500 height=400 scrollbars=yes resizable=yes");
	popwin.focus();
}

//현재 팝업 Right 확인 버튼 활성화
function fnAPPShowRightConfirmBtns(){
    callNativeFunction('activateConfirmRightBtns');
    return false;
}

//상품 등록 팝업 Right 버튼 활성화
function fnAPPShowRightRegisterBtns(){
    callNativeFunction('activateRegisterRightBtns');
    return false;
}

//현재 팝업창 캡션 변경
function fnAPPchangPopCaption(caption){
	callNativeFunction('changPopCaption', {"caption": caption});
    return false;
}

//현재 팝업 타이틀 감추기
function fnAPPhideTitle(){
    callNativeFunction('hideTitle');
    return false;
}

//현재 팝업 타이틀 보이기
function fnAPPshowTitle(){
    callNativeFunction('showTitle');
    return false;
}

//오픈 창 Js호출
function fnAPPopenerJsCall(jsfunc){
    callNativeFunction('openerJsCall', {"jsfunc": jsfunc});
    return false;    
}
//-->
</script>
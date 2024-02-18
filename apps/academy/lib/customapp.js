//기존 : callNativeFunction('호출할함수명',Jsonobject);
//변경 : webkit.messageHandlers.호출할함수명.postMessage(jsonobject);


function callNativeFunction(funcname, args) {
    if ( !args ) { args = {} }
    args['funcname'] = funcname;
    registerCallback(funcname, args);
    window.location = 'webkit.messageHandlers.' + funcname + '.postMessage:' + encodeURIComponent(JSON.stringify(args));
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

//현재 창 닫으면서 오픈 창 Js호출
function fnAPPopenerJsCallClose(jsfunc){
    callNativeFunction('openerJsCallClose', {"jsfunc": jsfunc});
    return false;    
}

var _selComp;

function fnAPPuploadImage(comp) {
    _selComp = comp;
    var paramname = comp.name;

    var upurl = "<%=uploadUrl2%>/linkweb/doevaluatewithimage_android_onlyimageupload.asp?paramname="+paramname;
    if (paramname=="imgbasic"){
        callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback": appUploadFinish1});
    }else if(paramname=="imgadd1"){
        callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback": appUploadFinish2});
    }else if(paramname=="imgadd2"){
        callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback": appUploadFinish3});
    }
    return false;
}
function _appUploadFinish(ret,ino){
   //alert("["+ino+"]");
    if (_selComp){
        _selComp.value=ret.name;
        $('#imgspan'+ino).empty();
		$('#imgspan'+ino).css("background-image", "url(<%=vImgURL%>" + ret.name + ")");
    }
}
function _appUploadFinish2(){
	$('#imgspan4').empty();
	$('#imgspan4').css("background-image", "url(http://image.thefingers.co.kr/diyitem/webimage/icon1/00/S1000007898.jpg)");
}

//동영상 삽입 팝업
function fnAPPpopupVod(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "동영상 삽입", [BtnType.CONFIRM], url, "vod");
	return false;
}

//카테고리 선택 팝업
function fnAPPpopupCategory(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "카테고리 설정", [BtnType.CONFIRM], url, "category");
	return false;
}

//옵션 선택 팝업
function fnAPPpopupOptionSet(url,title){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], title, [BtnType.CONFIRM], url, "option");
	return false;
}

//제작 특이사항 입력 팝업
function fnAPPpopupReqContents(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "제작 특이사항 입력", [BtnType.CONFIRM], url, "requirecontents");
	return false;
}

//검색 키워드 입력 팝업
function fnAPPpopupKeyWord(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "검색 키워드 입력", [BtnType.CONFIRM], url, "keyword");
	return false;
}

//배송비 안내 입력 팝업
function fnAPPpopupDeliveryInfo(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "배송비 안내 입력", [BtnType.CONFIRM], url, "deliveryinfo");
	return false;
}

//안전 인증 대상 팝업
function fnAPPpopupSafeInfo(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "안전인증 대상", [BtnType.CONFIRM], url, "safeinfo");
	return false;
}

//상품정보제공고시 입력 팝업
function fnAPPpopupItemInfoDiv(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "상품정보제공고시", [BtnType.CONFIRM], url, "iteminfo");
	return false;
}

//상품 검색 필터 팝업
function fnAPPpopupSearchFilter(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "필터", [BtnType.CONFIRM], url, "searchfilter");
	return false;
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

//현재 팝업 Right 버튼 활성화
function fnAPPShowRightConfirmBtns(){
    callNativeFunction('activateConfirmRightBtns');
    return false;
}

//현재 팝업 Right 버튼 활성화
function fnAPPShowRightRegisterBtns(){
    callNativeFunction('activateRegisterRightBtns');
    return false;
}

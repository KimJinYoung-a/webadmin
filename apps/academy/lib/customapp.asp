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
		}else if(funcname=="uploadaddImage"){
			webkit.messageHandlers.uploadaddImage.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="reuploadImage"){
			webkit.messageHandlers.reuploadImage.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="reuploadaddImage"){
			webkit.messageHandlers.reuploadaddImage.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="openerJsCallClose"){
			webkit.messageHandlers.openerJsCallClose.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="nativeJsCallReturn"){
			webkit.messageHandlers.nativeJsCallReturn.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="closePopup"){
			webkit.messageHandlers.closePopup.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="popupBrowser"){
			webkit.messageHandlers.popupBrowser.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="activateConfirmRightBtns"){
			webkit.messageHandlers.activateConfirmRightBtns.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="activateRegisterRightBtns"){
			webkit.messageHandlers.activateRegisterRightBtns.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="setPushReceiveYN"){
			webkit.messageHandlers.setPushReceiveYN.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="hideButtonNoneClickLayer"){
			webkit.messageHandlers.hideButtonNoneClickLayer.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="changeBadgeCount"){
			webkit.messageHandlers.changeBadgeCount.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="parentsWinReLoad"){
			webkit.messageHandlers.parentsWinReLoad.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="popupOuterBrowser"){
			webkit.messageHandlers.popupOuterBrowser.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else if(funcname=="parentsWinJsCall"){
			webkit.messageHandlers.parentsWinJsCall.postMessage(encodeURIComponent(JSON.stringify(args)));
		}else{
			eval("window.webkit.messageHandlers." + funcname + ".postMessage(" + encodeURIComponent(JSON.stringify(args)) + ")");
		}
	<% else %>
		window.location = 'callNativeFunction:' + encodeURIComponent(JSON.stringify(args));
	<% end if %>
}

function encSpecialCharNativeFun(orgStr){
    //return(orgStr);
    return orgStr.replace(/\n/g,"&enqt;").replace(/\"/g,"&dbqt;").replace(/\'/g,"&siqt;");
}

function decSpecialCharNativeFun(orgStr){
    //return(orgStr);
    return orgStr.replace(/&enqt;/g,"\n").replace(/&dbqt;/g,"\"").replace(/&siqt;/g,"\'");
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
    callNativeFunction('closePopup',{});
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

//앱의 푸시 알림 상태 값 가져오기
function fnAPPJsCallPushYN(jsfunc){
    callNativeFunction('nativeJsCallReturn', {"jsfunc": jsfunc});
    return false;    
}

//카테고리 선택 팝업
function fnAPPpopupCategory(url){
	//alert(url)
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
function fnAPPpopupOptionWaitSet(querystring,optlevel){
	//alert(querystring +"/"+optlevel);
	var title,url;
	if(optlevel==1){
		title="단일 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionSetWait.asp?"+querystring
	}else{
		title="이중 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiWait.asp?"+querystring
	}
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], title, [BtnType.CONFIRM], url, "optionsetting");
	return false;
}

//옵션 셋팅 팝업
function fnAPPpopupOptionWaitEditSet(querystring,optlevel){
	//alert(querystring +"/"+optlevel);
	var title,url;
	if(optlevel==1){
		title="단일 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionSetWaitEdit.asp?"+querystring
	}else{
		title="이중 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiWaitEdit.asp?"+querystring
	}
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], title, [BtnType.CONFIRM], url, "optionsetting");
	return false;
}

//옵션 수정 셋팅 팝업
function fnAPPpopupOptionEditSet(querystring,optlevel){
	var title,url;
	if(optlevel==1){
		title="단일 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionEditSet.asp?"+querystring
	}else{
		title="이중 옵션 설정"
		url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiEdit.asp?"+querystring
	}
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], title, [BtnType.CONFIRM], url, "optionsetting");
	return false;
}

function fnAPPpopupMultiOptionWait(querystring){
	//alert(querystring);
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiInputWait.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "이중 옵션 설정", [BtnType.CONFIRM], url, "optionsetting2");
	return false;
}

function popOptionMultiCountWaitSet(querystring){
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiCountSetWait.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "옵션별 수량 설정", [BtnType.CONFIRM], url, "optionsetting3");
	return false;
}

function popOptionCountWaitSet(querystring){
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionCountSetWait.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "옵션별 수량 설정", [BtnType.CONFIRM], url, "optionsetting4");
	return false;
}

function popOptionMultiCountSet(querystring){
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiCountSet.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "옵션별 수량 설정", [BtnType.CONFIRM], url, "optionsetting3");
	return false;
}

function popOptionCountSet(querystring){
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionCountSet.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "옵션별 수량 설정", [BtnType.CONFIRM], url, "optionsetting4");
	return false;
}

function fnAPPpopupMultiOption(querystring){
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiInputWaitEdit.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "이중 옵션 설정", [BtnType.CONFIRM], url, "optionsetting2");
	return false;
}

function fnAPPpopupMultiOptionEdit(querystring){
	var url="<%=g_AdminURL%>/apps/academy/itemmaster/popup/popOptionMultiEditInput.asp?"+querystring
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "이중 옵션 설정", [BtnType.CONFIRM], url, "optionsetting2");
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
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "필터", [BtnType.CONFIRM], url, "searchfilter");
	return false;
}

//상품 정보 팝업
function fnAPPpopupItemDetail(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "상품 정보", [BtnType.CONFIRM], url, "itemdetail");
	return false;
}

//등록 대기 상품 수정 팝업
function fnAPPpopupWaitItemEdit(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "작품 등록", [BtnType.REGISTER], url, "waitregwin");
	return false;
}

//등록 대기 리스트 상품 수정 팝업
function fnAPPpopupWaitItemListEdit(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "작품 등록", [BtnType.REGISTER], url, "waiteditwin");
	return false;
}

//상품 가격 변경 요청 팝업
function fnAPPpopupItemPriceEdit(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "가격 변경 요청", [BtnType.CONFIRM], url, "itempriceedit");
	return false;
}

//상품등록 불러오기 팝업
function fnAPPpopupItemListCall(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "불러오기", [], url, "itemlistcall");
	return false;
}

var _selComp;
var _selTarget;
var _selOldFileName='';
var _selOldFileName2='';
var _selOldFileName3='';
var _delFileidx='';
var _delFiletarget='';
var _delFileName='';

function fnImgTempSaveEnd1(waititemid){
	//alert(waititemid);
	$("#tempSaveYn").val("Y");
	$("#waititemid").val(waititemid);
	fnAPPuploadImage(_selComp,_selTarget);
}
function fnImgTempSaveEnd2(waititemid){
	$("#tempSaveYn").val("Y");
	$("#waititemid").val(waititemid);
	fnAPPReUploadImage(_selComp,_selTarget);
}
function fnImgTempSaveEnd3(waititemid){
	$("#tempSaveYn").val("Y");
	$("#waititemid").val(waititemid);
	fnAPPuploadAddImage(_selComp,_selTarget);
}
function fnImgTempSaveEnd4(waititemid){
	$("#tempSaveYn").val("Y");
	$("#waititemid").val(waititemid);
	fnAPPReuploadAddImage(_selComp,_selTarget);
}

function fnAPPuploadImage(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
	var oldfilename;
	_selComp = comp;
	_selTarget = target;
	if($("#waititemid").val()==""){
		document.itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_ItemIDGet_Process_App.asp?target=1";
		document.itemreg.target="FrameCKP";
		document.itemreg.submit();
	}else{
		var paramname = comp;
		if(paramname=="imgbasic"){
			oldfilename=_selOldFileName;
		}else if(paramname=="imgadd1"){
			oldfilename=_selOldFileName2;
		}else if(paramname=="imgadd2"){
			oldfilename=_selOldFileName3;
		}
		var upurl = "<%=uploadUrl2%>/linkweb/academy/items/WaitDIYItemRegister_Process_App.asp?waititemid="+$("#waititemid").val()+"&paramname="+paramname + "&oldfilename=" + oldfilename;
		//alert(upurl);
		if (paramname=="imgbasic"){
			callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadFinish1});
		}else if(paramname=="imgadd1"){
			callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadFinish2});
		}else if(paramname=="imgadd2"){
			callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadFinish3});
		}
		return false;
	}
}

function _appUploadFinish(ret,ino){
    if (_selComp){
        _selComp.value=ret.name;
		var imgname=ret.name.substring(0,ret.name.length-4);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
        $('#imgspan'+ino + " button").remove();
		$('#imgspan'+ino).append("<img src='<%=fingersImgUrl%>/diyItem/waitimage/"+_selTarget+"/"+folername+"/"+ret.name+"' onclick=fnAPPReUploadImage('"+_selComp+"','"+_selTarget+"'); />").find("img").load(function(){fnAPPHideButtonNoneClickLayer();});
		$("#" + _selComp).val(ret.name);
		if(_selComp=="imgbasic"){
			_selOldFileName=ret.name;
		}else if(_selComp=="imgadd1"){
			_selOldFileName2=ret.name;
		}else if(_selComp=="imgadd2"){
			_selOldFileName3=ret.name;
		}
		//alert(ret.name);
    }
}

function fnAPPReUploadImage(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
	var oldfilename;
	_selComp = comp;
	_selTarget = target;
	_delFiletarget = "waitedit";
	if($("#waititemid").val()==""){
		document.itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_ItemIDGet_Process_App.asp?target=2";
		document.itemreg.target="FrameCKP";
		document.itemreg.submit();
	}else{
		var paramname = comp;
		if(paramname=="imgbasic"){
			if(_selOldFileName==""){
				_selOldFileName=$("#imgbasic").val();
			}
			oldfilename=_selOldFileName;
		}else if(paramname=="imgadd1"){
			if(_selOldFileName2==""){
				_selOldFileName2=$("#imgadd1").val();
			}
			oldfilename=_selOldFileName2;
		}else if(paramname=="imgadd2"){
			if(_selOldFileName3==""){
				_selOldFileName3=$("#imgadd2").val();
			}
			oldfilename=_selOldFileName3;
		}

		var upurl = "<%=uploadUrl2%>/linkweb/academy/items/WaitDIYItemRegister_Process_App.asp?waititemid="+$("#waititemid").val()+"&paramname="+paramname + "&oldfilename=" + oldfilename;
		if (paramname=="imgbasic"){
			callNativeFunction('reuploadImage', {"upurl":upurl,"paramname":paramname,"callback":_appUploadFinish1});
		}else if(paramname=="imgadd1"){
			callNativeFunction('reuploadImage', {"upurl":upurl,"paramname":paramname,"callback":_appUploadFinish2});
		}else if(paramname=="imgadd2"){
			callNativeFunction('reuploadImage', {"upurl":upurl,"paramname":paramname,"callback":_appUploadFinish3});
		}
		return false;
	}
}

function __appUploadFinish(ret,ino){
    if (_selComp){
        _selComp.value=ret.name;
		var imgname=ret.name.substring(0,ret.name.length-4);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
		
        $('#imgspan'+ino+" img").remove();
		$('#imgspan'+ino).append("<img src='<%=fingersImgUrl%>/diyItem/waitimage/"+_selTarget+"/"+folername+"/"+ret.name+"' onclick=fnAPPReUploadImage('"+_selComp+"','"+_selTarget+"'); />");
		$("#" + _selComp).val(ret.name);
		if(_selComp=="imgbasic"){
			_selOldFileName=ret.name;
		}else if(_selComp=="imgadd1"){
			_selOldFileName2=ret.name;
		}else if(_selComp=="imgadd2"){
			_selOldFileName3=ret.name;
		}
    }
}

function _appUploadFinish1(ret){
    __appUploadFinish(ret,1);
}
function _appUploadFinish2(ret){
    __appUploadFinish(ret,2);
}
function _appUploadFinish3(ret){
    __appUploadFinish(ret,3);
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

function fnAPPuploadRealImage(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
	var oldfilename;
	_selComp = comp;
	_selTarget = target;
	var paramname = comp;
	if(paramname=="imgbasic"){
		if(_selOldFileName==""){
			_selOldFileName=$("#imgbasic").val();
		}
		oldfilename=_selOldFileName;
	}else if(paramname=="imgadd1"){
		if(_selOldFileName2==""){
			_selOldFileName2=$("#imgadd1").val();
		}
		oldfilename=_selOldFileName2;
	}else if(paramname=="imgadd2"){
		if(_selOldFileName3==""){
			_selOldFileName3=$("#imgadd2").val();
		}
		oldfilename=_selOldFileName3;
	}
	var upurl = "<%=uploadUrl2%>/linkweb/academy/items/DIYItemEdit_Process_App.asp?itemid="+$("#itemid").val()+"&paramname="+paramname + "&oldfilename=" + oldfilename;
	//alert(upurl);
	if (paramname=="imgbasic"){
		callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadRealFinish1});
	}else if(paramname=="imgadd1"){
		callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadRealFinish2});
	}else if(paramname=="imgadd2"){
		callNativeFunction('uploadImage', {"upurl":upurl,"paramname":paramname,"callback":appUploadRealFinish3});
	}
	return false;
}

function _appUploadRealFinish(ret,ino){
    if (_selComp){
        _selComp.value=ret.name;
		var imgname=ret.name.substring(0,ret.name.length-4);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
        $('#imgspan'+ino + " button").remove();
		$('#imgspan'+ino).append("<img src='<%=fingersImgUrl%>/diyItem/webimage/"+_selTarget+"/"+folername+"/"+ret.name+"' onclick=fnAPPReuploadRealImage('"+_selComp+"','"+_selTarget+"'); />");
		$("#" + _selComp).val(ret.name);
		if(_selComp=="imgbasic"){
			_selOldFileName=ret.name;
		}else if(_selComp=="imgadd1"){
			_selOldFileName2=ret.name;
		}else if(_selComp=="imgadd2"){
			_selOldFileName3=ret.name;
		}
		//alert(ret.name);
    }
}

function fnAPPReuploadRealImage(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
	var oldfilename;
	_selComp = comp;
	_selTarget = target;
	_delFiletarget = "edit";
	var paramname = comp;
	if(paramname=="imgbasic"){
		if(_selOldFileName==""){
			_selOldFileName=$("#imgbasic").val();
		}
		oldfilename=_selOldFileName;
	}else if(paramname=="imgadd1"){
		if(_selOldFileName2==""){
			_selOldFileName2=$("#imgadd1").val();
		}
		oldfilename=_selOldFileName2;
	}else if(paramname=="imgadd2"){
		if(_selOldFileName3==""){
			_selOldFileName3=$("#imgadd2").val();
		}
		oldfilename=_selOldFileName3;
	}
	var upurl = "<%=uploadUrl2%>/linkweb/academy/items/DIYItemEdit_Process_App.asp?itemid="+$("#itemid").val()+"&paramname="+paramname + "&oldfilename=" + oldfilename;
	//alert(upurl);
	if (paramname=="imgbasic"){
		callNativeFunction('reuploadImage', {"upurl":upurl,"paramname":paramname,"callback":appReUploadRealFinish1});
	}else if(paramname=="imgadd1"){
		callNativeFunction('reuploadImage', {"upurl":upurl,"paramname":paramname,"callback":appReUploadRealFinish2});
	}else if(paramname=="imgadd2"){
		callNativeFunction('reuploadImage', {"upurl":upurl,"paramname":paramname,"callback":appReUploadRealFinish3});
	}
	return false;
}

function _appReUploadRealFinish(ret,ino){
	//alert(ret.name + "/" + ino);
    if (_selComp){
        _selComp.value=ret.name;
		var imgname=ret.name.substring(0,ret.name.length-4);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
		
        $('#imgspan'+ino+" img").remove();
		$('#imgspan'+ino).append("<img src='<%=fingersImgUrl%>/diyItem/webimage/"+_selTarget+"/"+folername+"/"+ret.name+"' onclick=fnAPPReuploadRealImage('"+_selComp+"','"+_selTarget+"'); />");
		$("#" + _selComp).val(ret.name);
		if(_selComp=="imgbasic"){
			_selOldFileName=ret.name;
		}else if(_selComp=="imgadd1"){
			_selOldFileName2=ret.name;
		}else if(_selComp=="imgadd2"){
			_selOldFileName3=ret.name;
		}
    }
}

function appUploadRealFinish1(ret){
    _appUploadRealFinish(ret,1);
}
function appUploadRealFinish2(ret){
    _appUploadRealFinish(ret,2);
}
function appUploadRealFinish3(ret){
    _appUploadRealFinish(ret,3);
}

function appReUploadRealFinish1(ret){
    _appReUploadRealFinish(ret,1);
}
function appReUploadRealFinish2(ret){
    _appReUploadRealFinish(ret,2);
}
function appReUploadRealFinish3(ret){
    _appReUploadRealFinish(ret,3);
}

function fnAPPuploadAddImage(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
    _selComp = comp;
	_selTarget = target;
	if($("#waititemid").val()==""){
		document.itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_ItemIDGet_Process_App.asp?target=3";
		document.itemreg.target="FrameCKP";
		document.itemreg.submit();
	}else{
		var paramname = comp;
		var upurl = "<%=uploadUrl2%>/linkweb/academy/items/WaitDIYItemRegister_Process_App.asp?waititemid="+$("#waititemid").val()+"&paramname="+paramname + "&oldfilename=" + $("#DetailList"+target+" #addimgname").val();
		callNativeFunction('uploadaddImage', {"upurl":upurl,"paramname":paramname,"callback": appUploadAddFinish});
		return false;
	}
}

function appUploadAddFinish(ret){
    if (_selComp){
		var imgname=ret.name.substring(0,ret.name.length-5);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
		$("#DetailList" + _selTarget + " input[name='addimgname']").val(ret.name);
        $("#DetailList" + _selTarget + " button").remove();
		$("#DetailList" + _selTarget + " #imgArea" + _selTarget).append("<img src='<%=fingersImgUrl%>/diyItem/waitcontentsimage/" + folername + "/" + ret.name + "' onclick=fnAPPReuploadAddImage('addimgname" + _selTarget + "','" + _selTarget + "');>");
		_selOldFileName=ret.name;
    }
}

function fnAPPReuploadAddImage(comp,target) {
	//alert(target + " / " + _selOldFileName + " / " + _selTarget);
    _selComp = comp;
	_selTarget = target;
	_delFiletarget = "waitedit";
	_delFileName = $("#DetailList"+target+" #addimgname").val();
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
	if($("#waititemid").val()==""){
		document.itemreg.action="/apps/academy/itemmaster/WaitDIYItemRegister_ItemIDGet_Process_App.asp?target=4";
		document.itemreg.target="FrameCKP";
		document.itemreg.submit();
	}else{
		var paramname = comp;
		var upurl = "<%=uploadUrl2%>/linkweb/academy/items/WaitDIYItemRegister_Process_App.asp?waititemid="+$("#waititemid").val()+"&paramname="+paramname + "&oldfilename=" + $("#DetailList"+target+" #addimgname").val();
		callNativeFunction('reuploadaddImage', {"upurl":upurl,"paramname":paramname,"callback": appReUploadAddFinish});
		return false;
	}
}

function appReUploadAddFinish(ret){
	//alert(_selTarget + " / " + ret.name);
    if (_selComp){
		var imgname=ret.name.substring(0,ret.name.length-5);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
		$("#DetailList" + _selTarget + " input[name='addimgname']").val(ret.name);
        $("#DetailList" + _selTarget + " img").remove();
		$("#DetailList" + _selTarget + " #imgArea" + _selTarget).append("<img src='<%=fingersImgUrl%>/diyItem/waitcontentsimage/" + folername + "/" + ret.name + "' onclick=fnAPPReuploadAddImage('addimgname" + _selTarget + "','" + _selTarget + "');>");
		_selOldFileName=ret.name;
    }
}

function fnAPPuploadAddImageReal(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
    _selComp = comp;
	_selTarget = target;
	var paramname = comp;
	var upurl = "<%=uploadUrl2%>/linkweb/academy/items/DIYItemRegister_Process_App.asp?itemid="+$("#itemid").val()+"&paramname="+paramname + "&oldfilename=" + $("#DetailList"+target+" #addimgname").val();
	callNativeFunction('uploadaddImage', {"upurl":upurl,"paramname":paramname,"callback": appUploadAddFinishReal});
	return false;
}

function appUploadAddFinishReal(ret){
    if (_selComp){
		var imgname=ret.name.substring(0,ret.name.length-5);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
		$("#DetailList" + _selTarget + " input[name='addimgname']").val(ret.name);
        $("#DetailList" + _selTarget + " button").remove();
		$("#DetailList" + _selTarget + " #imgArea" + _selTarget).append("<img src='<%=fingersImgUrl%>/diyItem/contentsimage/" + folername + "/" + ret.name + "' onclick=fnAPPReuploadAddImageReal('addimgname" + _selTarget + "','" + _selTarget + "');>");
		_selOldFileName=ret.name;
		//alert($("#DetailList" + _selTarget + " #addimgname").val());
    }
}

function fnAPPReuploadAddImageReal(comp,target) {
	//alert("이미지 업로드 경우 실섭을 향하고 있어\n상품등록 이미지 테스트는 실제 서버에 올린 후 테스트 가능합니다.(실섭 이미지와 교체 방지)");
	_selComp = comp;
	_selTarget = target;
	_delFiletarget = "edit";
	_delFileName = $("#DetailList"+target+" input[name='addimgname']").val();
	var paramname = comp;
	var upurl = "<%=uploadUrl2%>/linkweb/academy/items/DIYItemRegister_Process_App.asp?itemid="+$("#itemid").val()+"&paramname="+paramname + "&oldfilename=" + $("#DetailList"+target+" #addimgname").val();
	callNativeFunction('reuploadaddImage', {"upurl":upurl,"paramname":paramname,"callback": appReUploadAddFinishReal});
	return false;
}

function appReUploadAddFinishReal(ret){
    if (_selComp){
		var imgname=ret.name.substring(0,ret.name.length-5);
		var itemid=imgname.substring(1,imgname.length);
		var folername = Num2Str(parseInt(parseInt(itemid) / 10000),2,'0','R');
		$("#DetailList" + _selTarget + " input[name='addimgname']").val(ret.name);
        $("#DetailList" + _selTarget + " img").remove();
		$("#DetailList" + _selTarget + " #imgArea" + _selTarget).append("<img src='<%=fingersImgUrl%>/diyItem/contentsimage/" + folername + "/" + ret.name + "' onclick=fnAPPReuploadAddImageReal('addimgname" + _selTarget + "','" + _selTarget + "');>");
		_selOldFileName=ret.name;
    }
}

function fnImageDeleteSet(parameter){
	//alert(_delFileName);
	if(parameter=="imgbasic"){
		$("#imgspan1 img").remove();
		if(_delFiletarget=="waitedit"){
			$("#imgspan1").append("<button type='button' onclick=fnAPPuploadImage('imgbasic','basic');>이미지 등록1</button>");
		}else{
			$("#imgspan1").append("<button type='button' onclick=fnAPPuploadRealImage('imgbasic','basic');>이미지 등록1</button>");
		}
		
		$("#imgbasic").val("");
	}else if(parameter=="imgadd1"){
		$("#imgspan2 img").remove();
		if(_delFiletarget=="waitedit"){
			$("#imgspan2").append("<button type='button' onclick=fnAPPuploadImage('imgadd1','add1');>이미지 등록2</button>");
		}else{
			$("#imgspan2").append("<button type='button' onclick=fnAPPuploadRealImage('imgadd1','add1');>이미지 등록2</button>");
		}
		
		$("#imgadd1").val("");
	}else if(parameter=="imgadd2"){
		$("#imgspan3 img").remove();
		if(_delFiletarget=="waitedit"){
			$("#imgspan3").append("<button type='button' onclick=fnAPPuploadImage('imgadd2','add2');>이미지 등록3</button>");
		}else{
			$("#imgspan3").append("<button type='button' onclick=fnAPPuploadRealImage('imgadd2','add2');>이미지 등록3</button>");
		}
		
		$("#imgadd2").val("");
	}else{
		if(_delFileName!=""){
			document.itemreg.delmode.value=_delFiletarget;
			document.itemreg.delfilename.value=_delFileName;
			document.itemreg.action="/apps/academy/itemmaster/DIYItem_img_Del_Process_App.asp";
			document.itemreg.target = "FrameCKP";
			document.itemreg.submit();
		}
		var targetnum = parameter.replace('addimgname','');
		//alert(targetnum);
		$("#DetailList" + targetnum + " input[name='addimgname']").val("");
		$("#DetailList" + targetnum + " img").remove();
		if(_delFiletarget=="waitedit"){
			$("#DetailList" + targetnum + " #imgArea" + targetnum).append("<button type='button' class='btnImgRegist' onclick=fnAPPuploadAddImage('addimgname" + targetnum + "','" + targetnum + "');>이미지 등록</button>");
		}else{
			$("#DetailList" + targetnum + " #imgArea" + targetnum).append("<button type='button' class='btnImgRegist' onclick=fnAPPuploadAddImageReal('addimgname" + targetnum + "','" + targetnum + "');>이미지 등록</button>");
		}
		
		_delFiletarget="";
		_delFileName="";
	}
}

//이미지 로드 확인 후 콜 (버튼 클릭 방지 레이어 닫기)
function fnAPPHideButtonNoneClickLayer(){
    callNativeFunction('hideButtonNoneClickLayer');
    return false;
}

//부모창 새로 고침
function fnAPPParentsWinReLoad(){
    callNativeFunction('parentsWinReLoad');
    return false;
}

//NEW 팝업 with Url
function fnAPPpopupItemRegPreview(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "미리보기", [], url, "preview");
	return false;
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

//상품 등록 팝업 Right 버튼 감추기
function fnAPPHideRightRegisterBtns(){
    callNativeFunction('activateRegisterRightBtnsHide');
    return false;
}

//뱃지 카운트 변경 호출
function fnAPPChangeBadgeCount(badge,count){
    callNativeFunction('changeBadgeCount',{"badgename":badge,"count":count});
    return false;
}

//주문 정보 팝업
function fnAPPpopupOrderBasicInfo(url,title){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], title, [], url, "requirecontents");
	return false;
}

//송장 입력 팝업
function fnAPPpopupSongjangInput(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "송장입력", [BtnType.CONFIRM], url, "songjanginput");
	return false;
}

//미출고 사유 입력 팝업
function fnAPPpopupUnDeliverReasonInput(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "미출고 사유 입력", [BtnType.CONFIRM], url, "undeliverreason");
	return false;
}

//송장 일괄 등록 팝업(테스트용)
function fnAPPpopupSongjangInputAll(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "송장 일괄 등록", [BtnType.CONFIRM], url, "songjanginput");
	return false;
}

//CS 처리결과 작성 팝업
function fnAPPpopupCSResultInput(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "CS 처리결과 작성", [BtnType.CONFIRM], url, "csresultwrite");
	return false;
}

//CS 관련 도움말 팝업
function fnAPPpopupCsHelpInfo(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "CS 관련 도움말", [], url, "cshelp");
	return false;
}

//응원톡 팝업
function fnAPPpopupCheerUpTalk(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "응원톡", [], url, "cheeruptalk");
	return false;
}

//응원톡 답글쓰기 팝업
function fnAPPpopupCheerUpWrite(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "답글쓰기", [BtnType.CONFIRM], url, "cheeruptalk");
	return false;
}

//구매후기 팝업
function fnAPPpopupItemReview(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "구매후기", [], url, "itemreview");
	return false;
}

//Q&A 팝업
function fnAPPpopupQna(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "Q&A", [], url, "qnapop");
	return false;
}

//Q&A 답글쓰기 팝업
function fnAPPpopupQnaWrite(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "답글쓰기", [BtnType.CONFIRM], url, "qnawrite");
	return false;
}

//Q&A Detail 팝업
function fnAPPpopupQnaDetail(url){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "Q&A", [], url, "qnadetailpop");
	return false;
}
//외부브라우져
function fnAPPpopupOuterBrowser(url) {
    callNativeFunction('popupOuterBrowser', {"url": url});
    return false;
}

//부모 창 Js호출
function fnAPPParentsWinJsCall(jsfunc){
    callNativeFunction('parentsWinJsCall', {"jsfunc": jsfunc});
    return false;    
}

//문의하기 답글쓰기 팝업
function fnAPPpopupFreeBoardReply(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "답글쓰기", [BtnType.CONFIRM], url, "freereply");
	return false;
}
//문의하기 답글쓰기 팝업
function fnAPPpopupFreeBoardAsk(url){
	fnAPPpopupBrowser(OpenType.FROM_BOTTOM, [], "문의하기", [BtnType.CONFIRM], url, "freeask");
	return false;
}
/*
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
*/
//-->
</script>
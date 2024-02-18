/*
+---------------------------------------------------------------------------------------------------------+
|                               [ 페이지 이동 ]  공통 Script 함수선언                                     |
+------------------------------------------+--------------------------------------------------------------+
|                함 수 명                  |                          기    능                            |
+------------------------------------------+--------------------------------------------------------------+
| TnGotoProduct(v)                         | 상품페이지로 이동                                            |
+------------------------------------------+--------------------------------------------------------------+
| AddEval(OrdSr,itID,OptCd)                | 상품후기 쓰기 페이지 이동                                    |
+------------------------------------------+--------------------------------------------------------------+
| viewVideo(idx)                           | 동영상 강좌 보기 이동                                        |
+------------------------------------------+--------------------------------------------------------------+

+---------------------------------------------------------------------------------------------------------+
|                              [ 레이어 및 팝업 ]  공통 Script 함수선언                                   |
+------------------------------------------+--------------------------------------------------------------+
|                함 수 명                  |                          기    능                            |
+------------------------------------------+--------------------------------------------------------------+
| ZoomItemPop(itemid)                      | 상품 상세 이미지/설명 보기 팝업                              |
+------------------------------------------+--------------------------------------------------------------+
| popSNSPost(svc,tit,link)                 | 쇼셜네트워크로 글보내기 팝업                                 |
+------------------------------------------+--------------------------------------------------------------+
| jsShowMailBox(frm,selVal,strVal)         | Email comboBox 선택표시 처리                                 |
+------------------------------------------+--------------------------------------------------------------+

+---------------------------------------------------------------------------------------------------------+
|                                 [ 폼 실행 ]  공통 Script 함수선언                                       |
+------------------------------------------+--------------------------------------------------------------+
|                함 수 명                  |                          기    능                            |
+------------------------------------------+--------------------------------------------------------------+
| DownloadlecturerCoupon(lecturercouponidx)| 강좌 쿠폰 다운로드 받기 실행                                 |
+------------------------------------------+--------------------------------------------------------------+
| DownloadDiyItemCoupon(itemcouponidx)     | 상품 쿠폰 다운로드 받기 실행                                 |
+------------------------------------------+--------------------------------------------------------------+
| TnAddShoppingBag(bool)                   | 장바구니에 상품을 담기                                       |
+------------------------------------------+--------------------------------------------------------------+
| TnAddFavorite(iitemid)                   | 관심 품목 담기 - 상품 페이지 전용                            |
+------------------------------------------+--------------------------------------------------------------+
| TnAddFavoriteList()                      | 관심 품목 담기 - 복수 상품용                                 |
+------------------------------------------+--------------------------------------------------------------+

+---------------------------------------------------------------------------------------------------------+
|                               [ 기타 기능 ]  공통 Script 함수선언                                       |
+------------------------------------------+--------------------------------------------------------------+
| islogin()                                | 로그인여부 확인 (true or false)                              |
+------------------------------------------+--------------------------------------------------------------+
| getCookie(name)                          | name에 해당하는 쿠키값 접수                                  |
+------------------------------------------+--------------------------------------------------------------+

*/
// PNG를 지원안하는 브라우저에서 출력
function setPng24(obj) {
    obj.width=obj.height=1;
    obj.className=obj.className.replace(/\bpng24\b/i,'');
    obj.style.filter =
    "progid:DXImageTransform.Microsoft.AlphaImageLoader(src='"+ obj.src +"',sizingMethod='image');"
    obj.src=''; 
    return '';
}


//상품후기 쓰기
function AddEval(OrdSr,itID,OptCd){	
	var winEval; 
	winEval = window.open('/myfingers/goodsusing/diyitem/diyitem_goodsUsingWrite.asp?orderserial=' + OrdSr + '&itemid=' + itID + '&optionCD=' + OptCd,"popeval","width=730,height=760,status=no,resizable=yes,scrollbars=yes");
	winEval.focus();
}

// 상품후기 팝업
function popEvaluate(iid,mtd){
	var subwin;
	subwin = window.open("/diyshop/PopItemEvaluate.asp?itemid=" + iid + "&sortMtd=" + mtd,"popeval","width=770,height=600,status=no,resizable=no,scrollbars=yes");
	subwin.focus();
}

//로그인 여부 확인(쿠키)
function islogin() {
	if(getCookie('uinfo')) {
		return "True";
	} else {
		return "False";
	}
}

// 쿠키를 가져온다
function getCookie(name){
 var nameOfCookie = name + "=";
 var x = 0;

 while ( x <= document.cookie.length )
 {
  var y = (x+nameOfCookie.length);
  if ( document.cookie.substring( x, y ) == nameOfCookie ) {
   if ( (endOfCookie=document.cookie.indexOf( ";", y )) == -1 )
   endOfCookie = document.cookie.length;
   return unescape( document.cookie.substring( y, endOfCookie ) );
  }

  x = document.cookie.indexOf( " ", x ) + 1;

  if ( x == 0 )
   break;
 }
 return "";
}

// 강좌 쿠폰 받기
function DownloadlecturerCoupon(lecturercouponidx){
	var popwin=window.open('/myfingers/downloadlecturercoupon.asp?lecturercouponidx=' + lecturercouponidx,'DownloadCoupon','width=550,height=550,scrollbars=no,resizable=no');
	popwin.focus();
}

// 상품 쿠폰 받기
function DownloadDiyItemCoupon(itemcouponidx){
	var popwin=window.open('/myfingers/downloaditemcoupon.asp?itemcouponidx=' + itemcouponidx,'DownloadCoupon','width=550,height=550,scrollbars=no,resizable=no');
	popwin.focus();
}

// 장바구니 담기
function TnAddShoppingBag(bool){
    var frm = document.sbagfrm;
    var optCode = "0000";
    
    
    var MOptPreFixCode="Z";

    if (!frm.item_option){
        //옵션 없는경우

    }else if (!frm.item_option[0].length){
        //단일 옵션
        if (frm.item_option.value.length<1){
            alert('옵션을 선택 하세요.');
            frm.item_option.focus();
            return;
        }

        if (frm.item_option.options[frm.item_option.selectedIndex].id=="S"){
            alert('품절된 옵션은 구매하실 수 없습니다.');
            frm.item_option.focus();
            return;
        }

        optCode = frm.item_option.value;
    }else{
        //이중 옵션 경우

        for (var i=0;i<frm.item_option.length;i++){
            if (frm.item_option[i].value.length<1){
                alert('옵션을 선택 하세요.');
                frm.item_option[i].focus();
                return;
            }

            if (frm.item_option[i].options[frm.item_option[i].selectedIndex].id=="S"){
                alert('품절된 옵션은 구매하실 수 없습니다.');
                frm.item_option[i].focus();
                return;
            }

            if (i==0){
                optCode = MOptPreFixCode + frm.item_option[i].value.substr(1,1);
            }else if (i==1){
                optCode = optCode + frm.item_option[i].value.substr(1,1);
            }else if (i==2){
                optCode = optCode + frm.item_option[i].value.substr(1,1);
            }
        }

        if (optCode.length==2){
            optCode = optCode + "00";
        }

        if (optCode.length==3){
            optCode = optCode + "0";
        }
    }

    frm.itemoption.value = optCode;

    for (var j=0; j < frm.itemea.value.length; j++){
        if (((frm.itemea.value.charAt(j) * 0 == 0) == false)||(frm.itemea.value==0)){
    		alert('수량은 숫자만 가능합니다.');
    		e.focus();
    		return;
    	}
    }

    if (frm.requiredetail){

		if (frm.requiredetail.value.length<1){
			alert('주문 제작 상품 문구를 작성해 주세요.');
			frm.requiredetail.focus();
			return;
		}

		if(GetByteLength(frm.requiredetail.value)>255){
			alert('문구 입력은 한글 최대 120자 까지 가능합니다.');
			frm.requiredetail.focus();
			return;
		}
	}

    if (bool==true){
        frm.action = "/lecpay/DIYShopBag_process.asp?tp=pop";
        var BagWin = window.open('','iiBagWin','width=350,height=310,scrollbars=no,resizable=no');
        BagWin.focus();

        frm.target = "iiBagWin";
        frm.submit();

    }else{
        frm.target = "_self";
    	frm.action="/lecpay/DIYShopBag_process.asp";
    	frm.submit();
    }

}

//이중 옵션 인경우 필요
function CheckMultiOption(comp){
    var frm = comp.form;
    var compid = comp.id;
    var compvalue = comp.value;
    var compname  = comp.name;

    var optSelObj = eval(frm.name + "." + compname);

    var PreSelObj = null;
    var NextSelObj = null;
    var ReDrawObj = null;

    if (!optSelObj.length){
        return;
    }

    if ((compid==0)&&(optSelObj.length>1)) {
        NextSelObj = optSelObj[1];
        if (optSelObj.length>2) {
            ReDrawObj = optSelObj[2];
        }else{
            ReDrawObj = optSelObj[1];
        }
    }

    if ((compid==1)&&(optSelObj.length>2)) {
        PreSelObj  = optSelObj[0];
        NextSelObj = optSelObj[2];
        ReDrawObj = optSelObj[2];
    }

    if (compid==2) {
        PreSelObj  = optSelObj[1];
    }

    if ((PreSelObj!=null)&&(PreSelObj.value.length<1)){
        alert('상위 옵션을 먼저 선택 하세요.');
        comp.value = '';
        PreSelObj.focus();
        return;
    }

    // 최 하위만 품절 세팅
    var found = false;
    var issoldout = false;


    if ( (compvalue.length>0) && (( (ReDrawObj!=null)&&(optSelObj.length-compid==2) )||( (ReDrawObj!=null)&&(optSelObj.length-compid==3)&&(NextSelObj.value.length>0) ))) {
        for (var i=0; i<NextSelObj.length; i++){
            if (NextSelObj.options[i].value.length<1) continue;

            found = false;
            issoldout = false;
            for (var j=0;j<Mopt_Code.length;j++){
                // Box2Ea, Select1-Change
                if ((compid==0)&&(optSelObj.length==2)){
                    if (Mopt_Code[j].substr(1,1)==compvalue.substr(1,1)&&(Mopt_Code[j].substr(2,1)==ReDrawObj.options[i].value.substr(1,1))){
                        found = true;
                        ReDrawObj.options[i].style.color= "#888888";
                        break;
                    }
                }

                // Box3Ea, Select2-Change
                else if ((compid==1)&&(optSelObj.length==3)) {
                    if ((Mopt_Code[j].substr(1,1)==PreSelObj.value.substr(1,1))&&(Mopt_Code[j].substr(2,1)==comp.value.substr(1,1))&&(Mopt_Code[j].substr(3,1)==ReDrawObj.options[i].value.substr(1,1))){
                        found = true;
                        ReDrawObj.options[i].style.color= "#888888";
                        break;
                    }
                }

                // Box3Ea, Select2 Value Exists, Select1-Change
                else if ((compid==0)&&(optSelObj.length==3)&&(NextSelObj.value.length>0)){
                    if ((Mopt_Code[j].substr(1,1)==compvalue.substr(1,1))&&(Mopt_Code[j].substr(2,1)==NextSelObj.value.substr(1,1))&&(Mopt_Code[j].substr(3,1)==ReDrawObj.options[i].value.substr(1,1))){
                        found = true;
                        ReDrawObj.options[i].style.color= "#888888";
                        break;
                    }
                }
            }


            if (!found){
                ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255) + " (품절)";
                ReDrawObj.options[i].id = "S";
                ReDrawObj.options[i].style.color= "#DD8888";
            }else{
                if (Mopt_S[j]==true){
                    ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255) + " (품절)";
                    ReDrawObj.options[i].id = "S";
                    ReDrawObj.options[i].style.color= "#DD8888";
                }else{
                    if ( Mopt_LimitEa[j].length>0){
                        ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255) + " (한정 " + Mopt_LimitEa[j] + " 개)";
                    }else{
                        ReDrawObj.options[i].text = ReDrawObj.options[i].value.substr(2,255);
                    }
                    ReDrawObj.options[i].style.color= "#888888";
                    ReDrawObj.options[i].id = "";
                }
            }
        }
    }
}


// 관심 품목 담기 - 상품 페이지 전용 : 상품 코드로 변경
function TnAddFavorite(iitemid){
	//if (confirm('관심품목에 추가 하시겠습니까?')){
		var params = "mode=add&itemid=" + iitemid ;

        var FavWin = window.open('/myfingers/popMyDIYFavorite.asp?' + params ,'FavWin','width=380,height=300,scrollbars=no,resizable=no');
    	FavWin.focus();
	//}
}

// 관심 품목 담기 -- 다수 선택 가능
function TnAddFavoriteList(){
	var ArrayFavItemID='';
	var chkbx = document.getElementsByName('chkbxFav');

	for (var i=0;i<chkbx.length;i++) {
			if (chkbx[i].checked){
				ArrayFavItemID=ArrayFavItemID  + ',' + chkbx[i].value;
			}
	}

	if (ArrayFavItemID.length < 1){
			alert('하나 이상 상품을 선택해 주세요');
			return
	}	else	 {
			if (confirm('관심품목에 추가하시겠습니까?')){

			var FavWin = window.open('/myfingers/popMyDIYFavorite.asp?mode=AddFavItems&bagarray=' + ArrayFavItemID ,'FavWin','width=380,height=300,scrollbars=no,resizable=no');
			FavWin.focus();
			}

	}

}

//상품 추가 이미지 PopUp
function ZoomItemPop(itemid) {
	var popwin = window.open("/diyshop/PopZoomItem.asp?itemid=" + itemid + '&pop=pop',"win3",'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,width=800,height=520');
    popwin.focus();
}

// 쇼셜네트워크로 글보내기
function popSNSPost(svc,tit,link,pre,tag) {
    // tit 및 link는 반드시 UTF8로 변환하여 호출요망!
    var popwin = window.open("/apps/goSNSposts.asp?svc=" + svc + "&link="+link + "&tit="+tit + "&pre="+pre + "&tag="+tag,'popSNSpost');
    popwin.focus();
}

// Email comboBox 관련
function jsShowMailBox(frm,selVal,strVal) {
	
	if (eval(frm+"."+selVal).value == 'etc') {
		eval(frm+"."+strVal).style.display = '';
		eval(frm+"."+strVal).value = '';
		eval(frm+"."+strVal).readOnly = false;
		eval(frm+"."+strVal).focus();
	}
	else
	{
		eval(frm+"."+strVal).style.display = 'none';
		eval(frm+"."+strVal).value = eval(frm+"."+selVal).value;
	}
}

//동영상 강좌 보기 이동
function viewVideo(idx)
{
	top.document.location.href="/corner/diy_video.asp?idx="+idx+"";
}

function TnGotoProduct(v){
	location.href = '/diyshop/shop_prd.asp?itemid='+v;
}

function FnGotoLecture(v){
	document.location = '/lecture/lecturedetail.asp?lec_idx='+v;
}

function FnGotoLecShoppingBag(){
	document.location = '/lecpay/apply.asp';
}

function FnAddLecShoppingBag(){
    var frm = document.sbagfrm;
    var optCode = "0000";

    if (!frm.lecOption){
        //옵션 없는경우

    }else if (!frm.lecOption[0].length){
        //단일 옵션
        if (frm.lecOption.value.length<1){
            alert('일정을 선택 하세요.');
            frm.lecOption.focus();
            return;
        }

        if (frm.lecOption.options[frm.lecOption.selectedIndex].id=="S"){
            alert('마감된 강좌는 신청하실 수 없습니다.');
            frm.lecOption.focus();
            return;
        }

        optCode = frm.lecOption.value;
    }

    frm.itemoption.value = optCode;

	frm.method="post";
	frm.target="returnleclist";
	frm.action="/lecpay/apply_process.asp";
	frm.submit();
}

function FnAddWaitPersonList(){
    var frm = document.sbagfrm;
    var optCode = "0000";

    if (!frm.lecOption){
        //옵션 없는경우

    }else if (!frm.lecOption[0].length){
        //단일 옵션
        if (frm.lecOption.value.length<1){
            alert('일정을 선택 하세요.');
            frm.lecOption.focus();
            return;
        }

        if (frm.lecOption.options[frm.lecOption.selectedIndex].id!="S"){
            alert('신청가능한 강좌는 대기접수를 하실 수 없습니다.');
            frm.lecOption.focus();
            return;
        }

        optCode = frm.lecOption.value;
    }

    frm.itemoption.value = optCode;

	var waitpop = window.open("","waitpop","width=320,height=200,scrollbars=yes");
	frm.method="post";
	frm.target="waitpop";
	frm.action="/lecture/waitperson.asp";
	frm.submit();
}

function FnCheckLecutureShoppingBag() {
    var frm = document.sbagfrm;
    var optCode = "0000";

    if (!frm.lecOption){
        //옵션 없는경우

    }else if (!frm.lecOption[0].length){
        //단일 옵션
        if (frm.lecOption.value.length<1){
            alert('일정을 선택 하세요.');
            frm.lecOption.focus();
            return;
        }

		optCode = frm.lecOption.value;
        if (frm.lecOption.options[frm.lecOption.selectedIndex].id=="S"){
            // 마감안내 레이어
            centerOpenLayer('iPopSoldOut');
            return;
        }
    }

    // 수강안내 레이어
    //centerOpenLayer('iPopConfirmApply');
    FnAddLecShoppingBag();
    return;
}

function FnCheckWaitPerson() {
    var frm = document.sbagfrm;
    var optCode = "0000";

    if (!frm.lecOption){
        //옵션 없는경우

    }else if (!frm.lecOption[0].length){
        //단일 옵션
        if (frm.lecOption.value.length<1){
            alert('일정을 선택 하세요.');
            frm.lecOption.focus();
            return;
        }

		optCode = frm.lecOption.value;
        if (frm.lecOption.options[frm.lecOption.selectedIndex].id!="S"){
            // 진행중 강좌 안내 레이어
            //centerOpenLayer('iPopIngLecture');
            document.all['iPopIngLecture'].style.visibility="visible";
            return;
        }
    }

    // 수강안내 레이어
    //centerOpenLayer('iPopSoldOut');
    document.all['iPopSoldOut'].style.visibility="visible";
    return;
}

// 레이어 화면 정중앙에 열기
function centerOpenLayer(fm) {
    //화면크기 계산
	var bodyWidth,bodyHeight

	if (/MSIE/.test(navigator.userAgent)) { 
		bodyWidth    = document.body.clientWidth;
		bodyHeight    = document.body.clientHeight; 
	} else {
		bodyWidth    = window.innerWidth;
		bodyHeight    = window.innerHeight; 
	}

	var divWidth    = document.all[fm].offsetWidth; 
	var divHeight    = document.all[fm].offsetHeight; 
    var divLeft = 0, divTop = 0; 
    if(bodyWidth > divWidth) divLeft = Math.ceil((bodyWidth - divWidth) / 2); 
    if(bodyHeight > divHeight) divTop = Math.ceil((bodyHeight - divHeight) / 2);

	document.all[fm].style.left = divLeft; 
	document.all[fm].style.top = divTop; 
	document.all[fm].style.visibility="visible";
}

// 레이어 닫기
function closeLayer(fm) {
	document.all[fm].style.visibility="hidden";
}

function msglogin(){
	FnMustLoginMsg();
}

function FnMustLoginMsg(){
	alert('로그인 후 사용하세요.');
}

function TnFindZip(frmname){
	window.open('/lib/searchzip.asp?target=' + frmname, 'findzipcdode', 'width=460,height=250,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
}

function TnFindZipNew(frmname){
	window.open('/lib/searchzip_new.asp?target=' + frmname, 'findzipcdode', 'width=580,height=690,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
}

function TnTabNumber(thisform,target,num) {
   if (eval("document.frminfo." + thisform + ".value.length") == num) {
	  eval("document.frminfo." + target + ".focus()");
   }
}
function IsDigit(v){
	for (var j=0; j < v.length; j++){
		if ((v.charAt(j) * 0 == 0) == false){
			return false;
		}
	}
	return true;
}

// 플레시 인베드 //
function FlashEmbed(fid,fn,wd,ht,para)
{
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="' + wd + '" height="' + ht + '" id="' + fid + '" align="middle">');
	document.write('<param name="allowScriptAccess" value="sameDomain">');
	document.write('<param name="movie" value="' + fn + para + '">');
	document.write('<param name="menu" value="false">');
	document.write('<param name="quality" value="high">');
	document.write('<param name="wmode" value="transparent">');
	document.write('<embed src="' + fn + para + '" menu="false" quality="high" wmode="transparent" width="' + wd + '" height="' + ht + '" name="' + fid + '" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
	document.write('</object>');
}

// 미디어플레이어 인베드 //
function WMVEmbed(fid,fn,wd,ht)
{
	document.write('<object ID="' + fid + '" WIDTH="' + wd + '" HEIGHT="' + ht + '"  classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" CODEBASE=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab standby="Loading Microsoft?Windows? Media Player components..." type="application/x-oleobject">');
	document.write('<param name="Filename" value="' + fn + '">');
	document.write('<param name="AutoStart" value="false">');
	document.write('<param name="ShowControls" value="true">');
	document.write('<param name="ShowPositionControls" value="false">');
	document.write('<param name="ShowTracker" value="false">');
	document.write('<param name="ShowGotoBar" value="false">');
	document.write('<param name="ShowDisplay" value="false">');
	document.write('<param name="ShowStatusBar" value="false">');
	document.write('<embed type="application/x-mplayer2">');
	document.write('</object>');
}


// 로그아웃 //
function TnLogOut(frm){
	var ret = confirm('로그아웃 하시겠습니까?');

	if (ret){
		frm.action = '/login/dologout.asp';
		frm.submit();
	}
}

// 검색창 리셋
function search_clear(frm){
	frm.rect.value = "";
	frm.extUrl.value = "";
}

// 검색 실행
function TnSearch(frm){
	if (!frm.rect.value.length){
		alert('검색어를 입력하세요.');
		return;
	}

	frm.submit();
}
function NewTnSearch(frm){
	if (!frm.rect.value.length){
		alert('검색어를 입력하세요.');
		return;
	}
	if(frm.extUrl.value =="") {
		frm.submit();
	}else{
		self.location.href=frm.extUrl.value;		
	}
}

// 검색 실행(옵션)
function TnSearchOpt(frm){
	if (frm.rect.value==''){
		alert('검색어를 입력하세요.');
		return;
	}

	frm.submit();
}

// 팝업창 자동 리사이즈
// 팝업창에서 window.onload = popupResize;
// 온로드 이벤트를 이미 사용하고 있으면 
//	window.onload = function() {
//		popupResize();	// 추가
//	}

// 팝업창 자동리사이즈, Width를 지정하면 지정한대로
function popupResize(innerWidth,innerHeight)
{
	var strAgent = navigator.userAgent.toLowerCase();
	var strVersion = strAgent.substr(strAgent.indexOf("msie")+5,1);
    var IE	= strAgent.indexOf("MSIE") ?	true : false;
    
	if (IE)
	{
		var addHeight = (strVersion >=  7) ? 70 : 55;	// 7 이상은 URL창크기만큼 추가

		var innerBody = document.body;
		
		if (!innerHeight)
			var innerHeight = innerBody.scrollHeight + (innerBody.offsetHeight - innerBody.clientHeight);
		if (!innerWidth)
			var innerWidth = innerBody.scrollWidth + (innerBody.offsetWidth - innerBody.clientWidth);

		innerWidth += 10;
		innerHeight += addHeight;
		window.resizeTo(innerWidth,innerHeight);
	}
	else					// FF
	{
		var Dwidth = parseInt(document.body.scrollWidth);
		var Dheight = parseInt(document.body.scrollHeight);
		var divEl = document.createElement("div");
		divEl.style.position = "absolute";
		divEl.style.left = "0px";
		divEl.style.top = "0px";
		divEl.style.width = "100%";
		divEl.style.height = "100%";
	    document.body.appendChild(divEl);
	    window.resizeBy(Dwidth-divEl.offsetWidth, Dheight-divEl.offsetHeight);
		document.body.removeChild(divEl);
	}
}


function GetByteLength(val){
 	var real_byte = val.length;
 	for (var ii=0; ii<val.length; ii++) {
  		var temp = val.substr(ii,1).charCodeAt(0);
  		if (temp > 127) { real_byte++; }
 	}

   return real_byte;
}


// iframe 길이 자동
function resizeIfr(obj, minHeight) {
	minHeight = minHeight || 10;

	try {
		var getHeightByElement = function(body) {
			var last = body.lastChild;
			try {
				while (last && last.nodeType != 1 || !last.offsetTop) last = last.previousSibling;
				return last.offsetTop+last.offsetHeight;
			} catch(e) {
				return 0;
			}
			
		}
				
		var doc = obj.contentDocument || obj.contentWindow.document;
		if (doc.location.href == 'about:blank') {
			obj.style.height = minHeight+'px';
			return;
		}
		
		//var h = Math.max(doc.body.scrollHeight,getHeightByElement(doc.body));
		//var h = doc.body.scrollHeight;
		if (/MSIE/.test(navigator.userAgent)) {
			var h = doc.body.scrollHeight;
		} else {
			var s = doc.body.appendChild(document.createElement('DIV'))
			s.style.clear = 'both';

			var h = s.offsetTop;
			s.parentNode.removeChild(s);
		}
		
		//if (/MSIE/.test(navigator.userAgent)) h += doc.body.offsetHeight - doc.body.clientHeight;
		if (h < minHeight) h = minHeight;
	
		obj.style.height = h + 'px';
		if (typeof resizeIfr.check == 'undefined') resizeIfr.check = 0;
		if (typeof obj._check == 'undefined') obj._check = 0;

//		if (obj._check < 5) {
//			obj._check++;
			setTimeout(function(){ resizeIfr(obj,minHeight) }, 200); // check 5 times for IE bug
//		} else {
			//obj._check = 0;
//		}	
	} catch (e) { 
		//alert(e);
	}
	
}


//문자열의 공백여부 체크
function jsChkBlank(str)
{
    if (str == "" || str.split(" ").join("") == ""){
        return true;
	}
    else{
        return false;
	}
}

// 고객센타 공지사항 팝업
function pop_Notice(nid)
{
	var w;
	if(nid) {
		w = window.open("/cscenter/pop_NoticeList.asp?ntcId="+nid,"popNotice",'width=580,height=768,scrollbars=yes,resizable=yes');
	} else {
		w = window.open("/cscenter/pop_NoticeList.asp","popNotice",'width=580,height=768,scrollbars=yes,resizable=yes');
	}
	w.focus();
}

// 고객센타 찾아오는길 약도/지도 팝업
function pop_fingersmap(areaid,mod)
{
	var pop_fingersmap;

	pop_fingersmap = window.open("/cscenter/fingers_map_pop.asp?areaid="+areaid+"&mode="+mod,"pop_fingersmap",'width=712,height=750,scrollbars=yes,resizable=yes');
	pop_fingersmap.focus();
}
	

// 이미지 사이즈 리사이징
function Resizeimg(limitwidth,fileid)	{
	var frm = document.getElementById(fileid);
	if (frm.width > limitwidth){
		frm.width=limitwidth;
	}
}

// 이미지 상세보기 팝업
function popShowImg(v){
	  var p = (v);
	  w = window.open("/myfingers/showimage.asp?img=" + v, "imageView", "width=10,height=10,status=no,resizable=yes,scrollbars=yes");
      w.focus();
}

// 관심강좌 등록 팝업(강좌페이지용)
function popRegWishBox(v) {
	  var w = window.open("/myfingers/wishlist/popWishList.asp?lec_idx=" + v, "regWishLisp", "width=350,height=200");
      w.focus();
}
// 관심강좌 등록 팝업(복수등록용)
function popRegWishList() {
	var ArrayLecIdx='';
	var chkbx = document.getElementsByName('chkbxWish');

	for (var i=0;i<chkbx.length;i++) {
			if (chkbx[i].checked){
				ArrayLecIdx=ArrayLecIdx  + ',' + chkbx[i].value;
			}
	}

	if (ArrayLecIdx.length < 1){
			alert('하나 이상 강좌를 선택해 주세요');
			return
	}	else	 {
			if (confirm('관심 강좌에 추가하시겠습니까?')){

			var w = window.open('/myfingers/wishlist/popWishList.asp?lec_idx=' + ArrayLecIdx ,'regWishLisp','width=350,height=200');
			w.focus();
			}

	}
}

// 관심강사 등록 팝업(복수등록용)
function popRegTeachWishList() {
	var ArrayLecIdx='';
	var chkbx = document.getElementsByName('chkbxWish');

	for (var i=0;i<chkbx.length;i++) {
			if (chkbx[i].checked){
				ArrayLecIdx=ArrayLecIdx  + ',' + chkbx[i].value;
			}
	}

	if (ArrayLecIdx.length < 1){
			alert('한명 이상의 강사를 선택해 주세요');
			return
	}	else	 {
			if (confirm('관심 강사에 추가하시겠습니까?')){

			var w = window.open('/myfingers/wishlist/popTeachWishList.asp?teach_id=' + ArrayLecIdx + '&mode=add','regWishLisp','width=350,height=200');
			w.focus();
			}

	}
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_changeProp(objName,x,theProp,theValue) { //v6.0
  var obj = MM_findObj(objName);
  if (obj && (theProp.indexOf("style.")==-1 || obj.style)){
    if (theValue == true || theValue == false)
      eval("obj."+theProp+"="+theValue);
    else eval("obj."+theProp+"='"+theValue+"'");
  }
}

// Trim
function jsTrim(v){
	return v.replace(/^(\s+)|(\s+)$/g, "");
}

function jsChkNumber(value) {
	var temp = new String(value)
		
	if(temp.search(/\D/) != -1) {
		return false;
	}
		return true;	
}

// 강좌 상세보기 팝업
function ZoomLecturePop(lecIdx){
	var pZoom = window.open('/lecture/lib/pop_zoomLecture.asp?lec_idx='+ lecIdx,'ZoomLecPop','width=900,height=580');
	pZoom.focus();
}

// 모든 강좌 후기 팝업
function pop_all_vallist(lecturer_id){
	var pval = window.open('/lecture/lib/pop_valuation_list.asp?lecturer_id='+ lecturer_id,'valpop','width=778,height=500,scrollbars=yes,resizable=yes');
	pval.focus();
}

// 창닫기
function selfClose() {
	if (/MSIE/.test(navigator.userAgent)) { 
		if(navigator.appVersion.indexOf("MSIE 8.0")>=0) {
			window.opener='Self';
			window.open('','_parent','');
			window.close();
		} else if(navigator.appVersion.indexOf("MSIE 7.0")>=0) {
			window.open('about:blank','_self').close();
		} else { 
			window.opener = self;
			self.close();
		}
	} else {
		self.close();
	}
}


//로그인 후 로그인 페이지 팝업처리
function jsChklogin(blnLogin){
	if (blnLogin == "True"){
		return true;
	}
	if(confirm("로그인 하시겠습니까?")){
			var winLogin = window.open('/login/popuserloginpage.asp?iframe=o','popLogin','width=400,height=300');
			winLogin.focus();
	}
	return false;

}

//실명확인 여부확인 및 실명확인 페이지 팝업
function jsChkRealname(cRNCheck) {
	if(cRNCheck=='Y') {
		return true;
	} else {
		var winRNCheck = window.open('/member/PopCheckName.asp','popNameCheck','width=515,height=460');
		winRNCheck.focus();
	}
	return false;
}


//링크 점선 전체 없애기
function bluring(){ 
if(event.srcElement.tagName=="A"||event.srcElement.tagName=="IMG") document.body.focus();} 
document.onfocusin=bluring; 

function myqnawrite(){
	var popwin;
	popwin = window.open("/myfingers/qna/myqnawrite.asp","myqnawrite","width=700,height=580,scrollbars=yes,resizabled=yes");
	popwin.focus();
}

function myqnawriteWithParam(iorderserial,iqadiv,iitemid){
	var popwin;
	popwin = window.open("/myfingers/qna/myqnawrite.asp?orderserial="+iorderserial+"&qadiv="+iqadiv+"&itemid="+iitemid,"myqnawrite","width=700,height=580,scrollbars=yes,resizabled=yes");
	popwin.focus();
}

// for radio button checked Index
function getCheckedIndex(comp){
    var i =0;
    for( var i = 0 ; i <comp.length;  i++){
        if(comp[i].checked) return i;
    }
    return -1;
}

// 비회원 메일링 팝업
function popMailling_InMain()
{
	var popMailling = window.open('/member/mailzine/notmember_pop.asp','popMailling','width=520,height=732');
	popMailling.focus();
}

// 패스워드 복잡도 검사
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

    if((numAlpha>0&&numNums>0)||(numAlpha>0&&numSpecials>0)||(numNums>0&&numSpecials>0)) {
    	sRst=true;
    } else {
    	sRst=false;
    }
    return sRst;
}
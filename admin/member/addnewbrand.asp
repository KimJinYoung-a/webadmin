<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 브랜드
' History : 2015.05.27 서동석 생성
'			2022.02.09 한용민 수정(전시카테고리 담당MD 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%

dim i
dim designer
dim groupid

dim pcuserdiv : pcuserdiv = requestCheckVar(request("pcuserdiv"),16)

'/2013.12.02 한용민 추가
if not(C_ADMIN_AUTH or C_AUTH) then
	if pcuserdiv="999_50" or pcuserdiv="501_21" or pcuserdiv="502_21" or pcuserdiv="503_21" or pcuserdiv="903_21" then	' 900_21 출고처(기타)
		response.write "<script language='javascript'>"
		response.write "	alert('[권한없음] 매입처만 등록 가능 합니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
end if
%>
<script type='text/javascript'>

function CopyZip(flag,post1,post2,add,dong){
	if (flag=="s"){
		frmbrand.company_zipcode.value= post1 + "-" + post2;
		frmbrand.company_address.value= add;
		frmbrand.company_address2.value= dong;
	}else if(flag=="m"){
		frmbrand.return_zipcode.value= post1 + "-" + post2;
		frmbrand.return_address.value= add;
		frmbrand.return_address2.value= dong;
	}else if(flag=="p"){
		frmbrand.p_return_zipcode.value= post1 + "-" + post2;
		frmbrand.p_return_address.value= add;
		frmbrand.p_return_address2.value= dong;
	}
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SameReturnAddr(bool){
    var frm = document.frmbrand;
	if (bool){
		frm.return_zipcode.value = frm.company_zipcode.value;
		frm.return_address.value = frm.company_address.value;
		frm.return_address2.value = frm.company_address2.value;
	}else{
		frm.return_zipcode.value = "";
		frm.return_address.value = "";
		frm.return_address2.value = "";
	}
}

function SameReturnAddr2(bool){
    var frm = document.frmbrand;
	if (bool){
		frm.p_return_zipcode.value = frm.return_zipcode.value;
		frm.p_return_address.value = frm.return_address.value;
		frm.p_return_address2.value = frm.return_address2.value;
	}else{
		frm.p_return_zipcode.value = "";
		frm.p_return_address.value = "";
		frm.p_return_address2.value = "";
	}
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

function precheck(frm){
	if (frm.company_name.value.length<1){
		alert('사업자 등록상의 회사명을 입력하세요.');
		frm.company_name.focus();
		return;
	}

	if (frm.ceoname.value.length<1){
		alert('사업자 등록상의 대표자명을 입력하세요.');
		frm.ceoname.focus();
		return;
	}

	if (frm.company_no.value.length<1){
		alert('사업자 등록 번호를 입력하세요.');
		frm.company_no.focus();
		return;
	}

	if (frm.jungsan_gubun.value.length<1){
		alert('과세구분을 선택하세요.');
		frm.jungsan_gubun.focus();
		return;
	}

	var errMsg = chkIsValidJungsanGubun(frm.company_no.value, frm.jungsan_gubun.value);
	if (errMsg != "OK") {
		alert(errMsg);
		retutn;
	}

	if (frm.company_zipcode.value.length<1){
		alert('우편번호를 선택하세요.');
		frm.company_zipcode.focus();
		return;
	}

	if (frm.company_address.value.length<1){
		alert('사업자 등록상의 주소1을 입력하세요.');
		frm.company_address.focus();
		return;
	}

	if (frm.company_address2.value.length<1){
		alert('사업자 등록상의 주소2를 입력하세요.');
		frm.company_address2.focus();
		return;
	}

	if (frm.company_uptae.value.length<1){
		alert('사업자 등록상의 업태를 입력하세요.');
		frm.company_uptae.focus();
		return;
	}

	if (frm.company_upjong.value.length<1){
		alert('사업자 등록상의 업종을 입력하세요.');
		frm.company_upjong.focus();
		return;
	}

	if (frm.company_tel.value.length<1){
		alert('업체 전화번호를 입력하세요.');
		frm.company_tel.focus();
		return;
	}

	if (frm.manager_name.value.length<1){
		alert('담당자 성함을 입력하세요.');
		frm.manager_name.focus();
		return;
	}

	if (frm.manager_phone.value.length<1){
		alert('담당자 전화번호를 입력하세요.');
		frm.manager_phone.focus();
		return;
	}

	if (frm.manager_email.value.length<1){
		alert('담당자 이메일을 입력하세요.');
		frm.manager_email.focus();
		return;
	}

	if (frm.manager_hp.value.length<1){
		alert('담당자 핸드폰을 입력하세요.');
		frm.manager_hp.focus();
		return;
	}

    if (frm.jungsan_date.value.length<1){
		alert('정산일을 선택하세요.');
		frm.jungsan_date.focus();
		return;
	}

    if (frm.jungsan_date_off.value.length<1){
		alert('오프 정산일을 선택하세요. - 기본은 온라인과 동일합니다.');
		frm.jungsan_date_off.focus();
		return;
	}

    var partnerCnt = frm.partnerCnt.value;
    if (partnerCnt=='') partnerCnt=0;

    // 최초 브랜드가 아닐경우
    if (partnerCnt>0){
        if (frm.jungsan_date.value!=''){
            if (frm.jungsan_date.value!='말일'){
                if (!confirm('온라인 정산일의 기본값은 말일 입니다. 계속 진행 하시겠습니까?')){
                    return;
                }
            }
        }
        if (frm.jungsan_date_off.value!=''){
            if (frm.jungsan_date_off.value!='말일'){
                if (!confirm('오프라인 정산일의 기본값은 말일 입니다. 계속 진행 하시겠습니까?')){
                    return;
                }
            }
        }
    }

// -----------------	//


/*
	if (frm.userdiv.value.length<1){
		alert('업체 구분을 선택하세요.');
		frm.userdiv.focus();
		return;
	}
*/
    var pcuserdiv = getFieldValue(frm.pcuserdiv);
    // 9999_02:매입처, 9999_14:아카데미, 999_50:제휴사(온라인) , 501_21:직영매장, 502_21:가맹점, 503_21:도매 ,9999_21:출고처(기타)

    if (pcuserdiv.length<1){
		alert('브랜드 구분을 선택하세요.');
		frm.pcuserdiv[0].focus();
		return;
	}

	if (frm.uid.value.length<2){
		alert('브랜드 아이디를 입력하세요.');
		frm.uid.focus();
		return;
	}

    var regex = "^[a-zA-Z0-9_]+$";
	if(frm.uid.value.match(regex) == null){
		alert("브랜드 아이디에는 영문, 숫자, 밑줄(_) 만 입력할 수 있습니다.");
		va.focus();
	}

	if (frm.password.value.length<1){
		alert('브랜드 패스워드를 입력하세요.');
		frm.password.focus();
		return;
	}


	if (frm.password.value.length < 8 || frm.password.value.length > 16){
			alert("패스워드는 공백없이 8~16자입니다.");
			frm.password.focus();
			return ;
		 }

	if (!fnChkComplexPassword(frm.password.value)) {
			alert('패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
			frm.password.focus();
			return;
		}

	if(frm.password.value==frm.uid.value) {
			alert("아이디와 다른 비밀번호를 사용해주세요.");
			frm.password.focus();
			return  ;
		}

	//if (frm.passwordS.value.length<1){
	//	alert('브랜드 2차 패스워드를 입력하세요.');
	//	frm.passwordS.focus();
	//	return;
	//}

	//if (frm.passwordS.value.length < 8 || frm.passwordS.value.length > 16){
	//	alert("2차 패스워드는 공백없이 8~16자입니다.");
	//	frm.passwordS.focus();
	//	return ;
	//}

	//if (!fnChkComplexPassword(frm.passwordS.value)) {
	//	alert('2차 패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
	//	frm.passwordS.focus();
	//	return;
	//}

	//if(frm.passwordS.value==frm.uid.value) {
	//	alert("아이디와 다른 비밀번호를 사용해주세요.");
	//	frm.passwordS.focus();
	//	return  ;
	//}

	//if(frm.passwordS.value==frm.password.value) {
	//	alert("비밀번호와  다른 비밀번호를 사용해주세요.");
	//	frm.passwordS.focus();
	//	return  ;
	//}

	if (frm.socname_kor.value.length<1){
		alert('브랜드명-한글을 입력하세요.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('브랜드명-영문을 입력하세요.');
		frm.socname.focus();
		return;
	}

    if ((frm.p_return_zipcode.value.length<1)||(frm.p_return_address.value.length<1)){
        alert('물류 반품주소를 입력하세요.');
		frm.p_return_address.focus();
		return;
    }



    //일반 매입처.
    if (pcuserdiv=="9999_02"){
        /*
        var selltype=getFieldValue(frm.selltype);
        if (selltype.length<1){
    		alert('브랜드 판매채널을 선택하세요.');
    		frm.selltype[0].focus();
    		return;
    	}
	    */

	    if ((!frm.isusing[0].checked)&&(!frm.isusing[1].checked)){
    		alert('사용여부를 선택하세요.');
    		frm.isusing[0].focus();
    		return;
    	}

		/*
		// 항상 Y 로 생성한다. 등록 이후에 수정가능(skyer9)
    	if ((!frm.isextusing[0].checked)&&(!frm.isextusing[1].checked)){
    		alert('제휴몰 사용여부를 선택하세요.');
    		frm.isextusing[0].focus();
    		return;
    	}

        //제휴사 브랜드 설정 confirm 텐바이텐Y,제휴N인경우 (생각없이 N로 설정하므로 컨펌)
        if ((frm.isusing[0].checked)&&(frm.isextusing[1].checked)){
            if (!confirm('제휴몰 브랜드 사용여부 N인경우 InterPark,Lotte 등 제휴몰에 판매하지 않습니다. 계속하시겠습니까?')) {
                frm.isextusing[0].focus();
                return;
            }
        }
		*/

    	if ((!frm.streetusing[0].checked)&&(!frm.streetusing[1].checked)){
    		alert('스트리트 사용여부를 선택하세요.');
    		frm.streetusing[0].focus();
    		return;
    	}

		/*
    	if ((!frm.extstreetusing[0].checked)&&(!frm.extstreetusing[1].checked)){
    		alert('제휴몰 스트리트 사용여부를 선택하세요.');
    		frm.extstreetusing[0].focus();
    		return;
    	}
    	*/

    	if ((!frm.specialbrand[0].checked)&&(!frm.specialbrand[1].checked)){
    		alert('커뮤니티 사용여부를 선택하세요.');
    		frm.specialbrand[0].focus();
    		return;
    	}

    	if ((frm.catecode.value.length<1)&&(frm.offcatecode.value.length<1)){
    		alert('온라인 또는 오프라인 카테고리 구분을 선택하세요. \n- 둘 중 하나는 필수 사항입니다.');
    		//frm.catecode.focus();
    		return;
    	}

    	if (frm.standardmdcatecode.value.length<1){
    		alert('전시카테고리 담당MD를 선택하세요.');
    		frm.standardmdcatecode.focus();
    		return;
    	}

    	if ((frm.mduserid.value.length<1)&&(frm.offmduserid.value.length<1)){
    		alert('온라인 또는 오프라인 담당MD 구분을 선택하세요. \n- 둘 중 하나는 필수 사항입니다.');
    		frm.mduserid.focus();
    		return;
    	}

    	if (frm.maeipdiv.value.length<1){
    		alert('매입 구분을 선택하세요.');
    		frm.maeipdiv.focus();
    		return;
    	}

    	if (!IsDouble(frm.defaultmargine.value)){
    		alert('기본마진을 입력하세요. - 실수만 가능합니다.');
    		frm.defaultmargine.focus();
    		return;
    	}

    	if(frm.defaultdeliverytype.options[1].selected == true){
    		if(frm.defaultFreeBeasongLimit.value == ""){
    			alert('조건 배송의 경우 무료배송기준금액을 입력해주세요.');
    			frm.defaultFreeBeasongLimit.focus();
    			return;
    		}
    		if(frm.defaultDeliverPay.value == ""){
    			alert('조건 배송의 경우 배송비를 입력해주세요.');
    			frm.defaultDeliverPay.focus();
    			return;
    		}
    		if(isNaN(frm.defaultFreeBeasongLimit.value)){
    			alert('금액은 숫자로 입력해주세요.');
    			frm.defaultFreeBeasongLimit.value = "";
    			frm.defaultFreeBeasongLimit.focus();
    			return;
    		}
    		if(isNaN(frm.defaultDeliverPay.value)){
    			alert('배송비는 숫자로 입력해주세요.');
    			frm.defaultDeliverPay.value = "";
    			frm.defaultDeliverPay.focus();
    			return;
    		}
            if (frm.defaultFreeBeasongLimit.value*1<=0){
                alert('조건 배송의 경우 무료배송기준금액은 0원 이상이어야 합니다.');
                frm.defaultFreeBeasongLimit.focus();
                return;
            }
            if (frm.defaultDeliverPay.value*1<=2000){
                alert('조건 배송의 경우 배송비는 2000원 이상만 입력가능입니다.');
                frm.defaultDeliverPay.focus();
                return;
            }

    	}
	}

    //매입처(아카데미)
    if (pcuserdiv=="9999_14"){
        var selltype=frm.selltype.value;

        if (frm.mduserid.value.length<1){
    		alert('담당MD 구분을 선택하세요.');
    		frm.mduserid.focus();
    		return;
    	}

        var lec_yn = getFieldValue(frm.lec_yn);
        var diy_yn = getFieldValue(frm.diy_yn);

        if ((lec_yn=="N")&&(diy_yn=="N")){
            alert('강좌/DIY 둘중 하나는 사용으로 설정하셔야 합니다.');
            frm.lec_yn[0].focus();
            return;
        }

        if ((lec_yn=="Y")&&(frm.lec_margin.value.length<1)){
            alert('강좌 기본 마진을 입력하세요.');
            frm.lec_margin.focus();
            return;
        }

        if ((lec_yn=="Y")&&(frm.mat_margin.value.length<1)){
            alert('재료비 기본 마진을 입력하세요.');
            frm.mat_margin.focus();
            return;
        }

        if ((diy_yn=="Y")&&(frm.diy_margin.value.length<1)){
            alert('DIY상품 기본 마진을 입력하세요.');
            frm.diy_margin.focus();
            return;
        }

        if ((frm.diy_yn[0].checked)&&(frm.diy_dlv_gubun.value.length<1)){
            alert('DIY 배송구분을 선택하세요.');
            frm.diy_dlv_gubun.focus();
            return;
        }

        if (frm.diy_dlv_gubun.value=="9"){
            if (!IsDigit(frm.DefaultFreebeasongLimit.value)){
                alert('배송비 기준 숫자만 가능합니다.');
                frm.DefaultFreebeasongLimit.focus();
                return;
            }

            if (!IsDigit(frm.DefaultDeliverPay.value)){
                alert('배송비  숫자만 가능합니다.');
                frm.DefaultDeliverPay.focus();
                return;
            }

            if (frm.DefaultFreebeasongLimit.value*1<=0){
                alert('금액을 0원 이상 입력하세요.');
                frm.DefaultFreebeasongLimit.focus();
                return;
            }

            if (frm.DefaultDeliverPay.value*1<=0){
                alert('금액을 0원 이상 입력하세요.');
                frm.DefaultDeliverPay.focus();
                return;
            }

        }

        if ((lec_yn=="Y")&&(diy_yn=="N")){
            frm.selltype.value="10";
            frm.maeipdiv.value="M";
            frm.defaultmargine.value=frm.lec_margin.value;
        }

        if ((lec_yn=="N")&&(diy_yn=="Y")){
            frm.selltype.value="20";
            frm.maeipdiv.value="U";
            frm.defaultmargine.value=frm.diy_margin.value;

        }

        if ((lec_yn=="Y")&&(diy_yn=="Y")){
            frm.selltype.value="30";

        }

        if (diy_yn=="Y"){
            frm.maeipdiv.value="U";
            frm.defaultmargine.value=frm.diy_margin.value;
            frm.defaultdeliverytype.value = frm.diy_dlv_gubun.value;
        }
	}

	// 온라인제휴사, etc출고처
	if ((pcuserdiv=="999_50") || (pcuserdiv=="900_21") || (pcuserdiv=="902_21") || (pcuserdiv=="903_21")){
	    if (frm.purchasetype.value.length<1){
	        alert('정산 방식을 선택 하세요. 필수 값입니다.');
	        frm.purchasetype.focus()
	        return;
	    }
	}

	//온라인제휴사
	if (pcuserdiv=="999_50"){
	    if (frm.commission.value.length<1){
	        alert('수수료를 입력 하세요.');
	        frm.commission.focus()
	        return;
	    }
	}

	if ((pcuserdiv!="999_50") && (pcuserdiv!="900_21") && (pcuserdiv!="902_21") && (pcuserdiv!="903_21") && (pcuserdiv!="501_21") && (pcuserdiv!="502_21") && (pcuserdiv!="503_21")){
		//스트리트 표시여부 제휴몰을 텐바이텐과 통일.
		if(frm.streetusing[0].checked){
			frm.extstreetusing.value = "Y";
		}else if(frm.streetusing[1].checked){
			frm.extstreetusing.value = "N";
		}
	}

	var ret = confirm('브랜드 정보를 저장 하시겠습니까?');

	if (ret){
		if (frm.groupid.value.length<1) {
			icheckframe.location.href="icheckframe.asp?mode=CheckSocnoOnSave&socno=" + frm.company_no.value+"&pcuserdiv="+pcuserdiv;
		}else{
		    icheckframe.location.href="icheckframe.asp?uid=" + frm.uid.value + "&password=" + frm.password.value+"&pcuserdiv="+pcuserdiv;
		}
	}
}

function AddProc(mode){
    var frm = document.frmbrand;

	try {
		if (mode == "checkidpassword") {
			frm.submit();
			return;
		}

		if (mode == "CheckSocnoOnSave") {
			icheckframe.location.href="icheckframe.asp?uid=" + frm.uid.value + "&password=" + frm.password.value;
			return;
		}

		if (mode == "CheckSocno") {
			alert('등록가능한 사업자번호입니다.');
			return;
		}
	} catch (err) {
		alert(err.message);
		return;
	}
}

function chkIsValidJungsanGubun(company_no, jungsan_gubun) {
	// 000-00-00000
	//
	// 가운데 두글자 : 구분코드
	// =========================================================================
	// 01-79 : 개인사업자+과세사업자
	// 90-99 : 개인사업자+면세사업자
	// 기타 : 과세 면세 모두 가능
	//
	// 앞자리 세글자 : 지역(1-6) + 세무서일련번호
	// =========================================================================
	// 108 = 1(서울) + 08(동작)
	//
	// 앞자리 888 = 영세(해외), 간이과세
	// =========================================================================

	if (company_no.length != 12) {
		// return "잘못된 사업자번호입니다.";
		return "OK";
	}

	var soc_gubun = company_no.substring(4, 6)*1;
	var IsForeign = (company_no.substring(0, 3) == "888");

	if (IsForeign) {
		if ((jungsan_gubun != "영세(해외)") && (jungsan_gubun != "간이과세")) {
			return "영세(해외), 간이과세 사업자만 가능한 사업자번호입니다.";
		}

		return "OK";
	} else {
		if (jungsan_gubun == "영세(해외)") {
			return "영세(해외) 사업자로 변경 불가능한 사업자번호입니다.";
		}

		/*
		if ((soc_gubun >= 1) && (soc_gubun <= 79)) {
			if (jungsan_gubun == "면세") {
				return "면세로 등록할 수 없는 사업자번호입니다.";
			}

			return "OK";
		}
		*/

		if ((soc_gubun >= 90) && (soc_gubun <= 99)) {
			if (jungsan_gubun != "면세") {
				return "면세로만 등록가능한 사업자번호입니다.";
			}

			return "OK";
		}

		return "OK";
	}
}

function SearchSocno(frm){

	if (frm.company_no.value.length<1){
		alert('사업자 등록 번호를 입력하세요.');
		frm.company_no.focus();
		return;
	}

	if (frm.company_no.value.length != 12){
		alert('사업자 등록 번호는 000-00-00000 형식으로 입력해야 합니다.');
		frm.company_no.focus();
		return;
	}

	if (frm.groupid.value.length<1){
		icheckframe.location.href="icheckframe.asp?mode=CheckSocno&socno=" + frm.company_no.value;
	}else{
		alert('사업자번호를 변경할 경우 기존 정보가 변경됩니다.');
	}

}

function ModiInfo(frm){
	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		//frm.submit();
	}

}

function DisableSocInfo(frm){
<% if (C_ADMIN_AUTH <> true) then %>
	frm.company_name.readOnly = true;
	frm.company_no.readOnly = true;
	frm.ceoname.readOnly = true;
	frm.jungsan_gubun.readOnly = true;

	frm.company_name.style.background = "#EEEEEE";
	frm.company_no.style.background = "#EEEEEE";
	frm.ceoname.style.background = "#EEEEEE";
	frm.jungsan_gubun.style.background = "#EEEEEE";
<% end if %>
}

function CopyFromBrandInfo(){
	frmupche.company_name.value = frmbuf.company_name.value;
	frmupche.ceoname.value = frmbuf.ceoname.value;
	frmupche.company_no.value = frmbuf.company_no.value;
	frmupche.jungsan_gubun.value = frmbuf.jungsan_gubun.value;
	frmupche.company_zipcode.value = frmbuf.company_zipcode.value;
	frmupche.company_address.value = frmbuf.company_address.value;
	frmupche.company_address2.value = frmbuf.company_address2.value;
	frmupche.company_uptae.value = frmbuf.company_uptae.value;
	frmupche.company_upjong.value = frmbuf.company_upjong.value;
	frmupche.company_tel.value = frmbuf.company_tel.value;
	frmupche.company_fax.value = frmbuf.company_fax.value;
	frmupche.jungsan_bank.value = frmbuf.jungsan_bank.value;
	frmupche.jungsan_acctno.value = frmbuf.jungsan_acctno.value;
	frmupche.jungsan_acctname.value = frmbuf.jungsan_acctname.value;


	frmupche.manager_name.value = frmbuf.manager_name.value;
	frmupche.manager_phone.value = frmbuf.manager_phone.value;
	frmupche.manager_email.value = frmbuf.manager_email.value;
	frmupche.manager_hp.value = frmbuf.manager_hp.value;

	frmupche.deliver_name.value = frmbuf.deliver_name.value;
	frmupche.deliver_phone.value = frmbuf.deliver_phone.value;
	frmupche.deliver_email.value = frmbuf.deliver_email.value;
	frmupche.deliver_hp.value = frmbuf.deliver_hp.value;

	frmupche.jungsan_name.value = frmbuf.jungsan_name.value;
	frmupche.jungsan_phone.value = frmbuf.jungsan_phone.value;
	frmupche.jungsan_email.value = frmbuf.jungsan_email.value;
	frmupche.jungsan_hp.value = frmbuf.jungsan_hp.value;

}

function inputDeliveryType(ddt)
{
	if(ddt == "U")
	{
		document.getElementById("ddtdiv").style.display = "block";
	}
	else
	{
		document.frmbrand.defaultdeliverytype.options[0].selected = true;
		document.frmbrand.defaultFreeBeasongLimit.value = "";
		document.frmbrand.defaultDeliverPay.value = "";
		document.getElementById("ddtdiv").style.display = "none";
		document.getElementById("paydiv").style.display = "none";
	}
}

function inputDeliveryPay(pay)
{
	if(pay == "9")
	{
		document.getElementById("paydiv").style.display = "block";
	}
	else
	{
		document.frmbrand.defaultFreeBeasongLimit.value = "";
		document.frmbrand.defaultDeliverPay.value = "";
		document.getElementById("paydiv").style.display = "none";
	}
}



function clickLec(comp){

}

function clickDiy(comp){
    if (comp.value=="Y"){
        iDiyDlv.style.display="inline";
    }else{
        iDiyDlv.style.display="none";
    }
}

function stepNext(){
    var frm = document.frmNext;

    var pcuserdiv=getFieldValue(frm.pcuserdiv);

    if (pcuserdiv.length<1){
        alert('브랜드 구분을 먼저 선택하세요.');
        frm.pcuserdiv[0].focus();
        return;
    }

    frm.submit();
}

function chkCompdiygbn(comp){
    var frm = comp.form;
    if (comp.value=="9"){
        frm.DefaultFreebeasongLimit.style.background = '#FFFFFF';
        frm.DefaultDeliverPay.style.background  = '#FFFFFF';

        frm.DefaultFreebeasongLimit.readOnly = false;
        frm.DefaultDeliverPay.readOnly = false;

        frm.DefaultFreebeasongLimit.value=frm.pDFL.value;
        frm.DefaultDeliverPay.value=frm.pDDP.value;


    }else{
        frm.DefaultFreebeasongLimit.style.background = '#BBBBBB';
        frm.DefaultDeliverPay.style.background  = '#BBBBBB';

        frm.DefaultFreebeasongLimit.readOnly = true;
        frm.DefaultDeliverPay.readOnly = true;

        frm.DefaultFreebeasongLimit.value=0;
        frm.DefaultDeliverPay.value=0;
    }
}

function delcomRow(){
    //매장은 이곳에서 등록 불가
    var f = document.frmNext;

    for (i=0;i<f.pcuserdiv.length;i++){
    	if (f.pcuserdiv[i].value=="501_21" || f.pcuserdiv[i].value=="502_21" || f.pcuserdiv[i].value=="503_21" ){
    		f.pcuserdiv.remove(i);
    		i--;
    	}
    }
}

var orgjungsan_gubun = "일반과세";
function fnJungsanGubunChanged() {
	var frm = document.frmbrand;
	var company_no = document.getElementById("company_no");

	if ((orgjungsan_gubun != "영세(해외)") && (frm.jungsan_gubun.value != "영세(해외)")) {
		orgjungsan_gubun = frm.jungsan_gubun.value;
		return;
	}
	orgjungsan_gubun = frm.jungsan_gubun.value;

	if (frm.jungsan_gubun.value == "영세(해외)") {
		// 해외는 사업자번호 자동설정된다(888-00-00000)

		company_no.className = "text_ro";
		frm.company_no.readOnly = true;
		frm.company_no.value = "888-00-00000";

		frm.btnSearchSocno.disabled = true;

		frm.checksocnoyn.value = "Y";
	} else {
		company_no.className = "text";
		frm.company_no.readOnly = false;
		frm.company_no.value = "";

		frm.btnSearchSocno.disabled = false;

		frm.checksocnoyn.value = "N";
	}
}

</script>

<% if (pcuserdiv="") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext">
    <tr height="30" bgcolor="#FFFFFF">
    <td width="150" bgcolor="<%= adminColor("pink") %>">1. 브랜드 구분 선택</td>
    <td >
        <% drawPartnerCommCodeBox false,"pcuserdiv","pcuserdiv","9999_02","" %>

        <%'<script>delcomRow();</script>%>
    </td>
</tr>
<tr>
    <td colspan="2" height="30" bgcolor="#FFFFFF" align="center"><input type="button" value="다음" onClick="stepNext();"></td>
</tr>
</form>
</table>
<% else %>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbrand" method="post" action="/admin/member/doupcheedit.asp" target="FrameCKP">
<input type="hidden" name="mode" value="addnewupchebrand">
<input type="hidden" name="partnerCnt" value="">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>1.업체관련정보</b></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" class="text" name="groupid" value="" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
		<input type="button" class="button" value="업체선택" onClick="PopUpcheSelect('frmbrand'); DisableSocInfo(frmbrand);">
		</td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("tabletop") %>">입점브랜드ID</td>
		<td height="25" colspan="3" bgcolor="#FFFFFF"></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**사업자등록정보**&nbsp;&nbsp;&nbsp;(중복된 사업자번호는 등록할 수 없습니다.)</td>
	</tr>

	<tr>
		<td width="120" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 회사명(상호)</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_name" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td width="90" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 대표자</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="ceoname" value="" size="16" maxlength="16" style="background-color:#EEEEEE;" readonly></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 사업자번호</td>
		<input type="hidden" name="checksocnoyn" value="N">
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" id="company_no" name="company_no" value="" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
			<!--<input type="button" class="button" name="btnSearchSocno" value="검색" onClick="SearchSocno(frmbrand)">//-->
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 과세구분</td>
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()" readonly>
			<option value="일반과세" >일반과세</option>
			<option value="간이과세" >간이과세</option>
			<option value="원천징수" >원천징수</option>
			<option value="면세" >면세</option>
			<option value="영세(해외)" >영세(해외)</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 사업장소재지</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="company_zipcode" value="" size="7" maxlength="7" style="background-color:#EEEEEE;" readonly>
			<% '<input type="button" class="button_s" value="검색" onClick="FnFindZipNew('frmbrand','C')"> %>
			<% '<input type="button" class="button_s" value="검색(구)" onClick="TnFindZipNew('frmbrand','C')"> %>
			<% '<input type="button" class="button" value="검색(구)" onClick="popZip('s');"> %>
		    <br>
			<input type="text" class="text" name="company_address" value="" size="30" maxlength="64" style="background-color:#EEEEEE;" readonly>&nbsp;
			<input type="text" class="text" name="company_address2" value="" size="46" maxlength="64" style="background-color:#EEEEEE;" readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 업태</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 업종</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**업체기본정보**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 대표전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_tel)"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_fax)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">사무실주소</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="return_zipcode" value="" size="7" maxlength="7">
			<input type="button" class="button_s" value="검색" onClick="FnFindZipNew('frmbrand','D')">
			<input type="button" class="button_s" value="검색(구)" onClick="TnFindZipNew('frmbrand','D')">
			<% '<input type="button" class="button" value="검색(구)" onClick="popZip('m');"> %>

			<input type="checkbox" name="samezip2" onclick="SameReturnAddr(this.checked)">상동
		<br>
		<input type="text" class="text" name="return_address" value="" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="" size="46" maxlength="64">
		</td>
	</tr>
	<!-- 브랜드 정보로 변경
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">반품 주소</td>
		<td colspan="3" height="25" bgcolor="#FFFFFF">초기 반품주소는 사무실 주소와 동일하게 설정됩니다.</td>
	</tr>
	-->
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**결제계좌정보**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<% DrawBankCombo "jungsan_bank", "" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctno" value="" size="24" maxlength="32" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctname" value="" size="24" maxlength="16" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 정산일</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		온라인 : <% DrawJungsanDateCombo "jungsan_date", "" %>
		&nbsp;
		오프라인 : <% DrawJungsanDateCombo "jungsan_date_off", "" %>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**업체 담당자정보**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_phone)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_hp)"></td>
	</tr>


	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_phone)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_hp)"></td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="30" >
    	<td bgcolor="<%= adminColor("pink") %>" colspan="6"><b>2.브랜드관련정보</b></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td height="25" colspan="4">**브랜드 기본정보**</td>
    </tr>
    <tr height="30" bgcolor="#FFFFFF">
        <td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 브랜드 구분</td>
        <td colspan="3">
            <%= getPartnerCommCodeName("pcuserdiv",pcuserdiv) %>
            <input type="hidden" name="pcuserdiv" value="<%= pcuserdiv %>">
        </td>
    </tr>
	<tr height="50">
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 브랜드ID</td>
		<td bgcolor="#FFFFFF" >
    		<input type="text" class="text" name="uid" value="" size="24" maxlength="24">
    		<p>(영문, 숫자, 밑줄(_) 만 가능 특수문자 금지)</p>
			<% if (pcuserdiv="501_21") or (pcuserdiv="502_21") or (pcuserdiv="503_21") then %>
			<ul>
				<li>텐바이텐 매장 - streetshopxxx</li>
				<li>아이띵소 매장 - ithinksoxxxxx, 3pl_its_xxxxx</li>
				<li>도매 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;- wholesale1xxx</li>
				<li>대행 매장 &nbsp; &nbsp; &nbsp; - ygentshop1xxx, 3pl_xxx_xxxxx</li>
				<li>아이띵소 해외출고처 &nbsp; &nbsp; &nbsp; - its_exp_xxxxx</li>
			</ul>
			<% end if %>
		</td>
		<td width="100"  bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 패스워드</td>
		<td bgcolor="#FFFFFF" >
		    <input type="password" class="text" name="password" value="" size="16" maxlength="24">
		    <%'<input type="password" class="text" name="passwordS" value="" size="16" maxlength="24">%>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 브랜드명(KR)</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="socname_kor" value="" size="30" maxlength="32">
		</td>
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 브랜드명(EN)</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="socname" value="" size="30" maxlength="32">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">브랜드명 표시</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select name="socname_use" class="select">
				<option value="K">브랜드명(KR)</option>
				<option value="E" selected>브랜드명(EN)</option>
			</select>
		</td>
	</tr>
	<!--
	<tr >
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 업체구분</td>
		<td bgcolor="#FFFFFF" colspan=2>
		<% DrawBrandGubunCombo "userdiv", "" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", "" %></td>
	</tr>
	-->
	<tr bgcolor="#FFFFFF">
    	<td height="25" colspan="4">**브랜드 물류정보**</td>
    </tr>
	<tr>
		<td height="25"  bgcolor="<%= adminColor("pink") %>">기본택배사</td>
		<td bgcolor="#FFFFFF" ><% drawSelectBoxDeliverCompany "defaultsongjangdiv","" %></td>
		<td width="90" bgcolor="<%= adminColor("pink") %>" >랙번호(물류)</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="prtidx" value="9999" size="4" maxlength="4">
		(기본값 : 9999)</td>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" >물류(반품)주소</td>
		<td bgcolor="#FFFFFF" colspan=5 >
			<input type="text" class="text" name="p_return_zipcode" value="" size="7" maxlength="7">
			<input type="button" class="button_s" value="검색" onClick="FnFindZipNew('frmbrand','I')">
			<input type="button" class="button_s" value="검색(구)" onClick="TnFindZipNew('frmbrand','I')">
			<% '<input type="button" class="button" value="검색(구)" onClick="popZip('p');"> %>

			<input type="checkbox" name="p_samezip" onclick="SameReturnAddr2(this.checked)">상동(사무실주소와)
			<br>
			<input type="text" class="text" name="p_return_address" value="" size="30" maxlength="64">&nbsp;
			<input type="text" class="text" name="p_return_address2" value="" size="46" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">배송담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("pink") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.deliver_phone)"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("pink") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.deliver_hp)"></td>
	</tr>
</table>

<p>
<% ''' 9999_15 추가 2016/05/16 %>
<% if (pcuserdiv="9999_02") or (pcuserdiv="9999_15") then %>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
	    <% if (pcuserdiv="9999_15") then %>
	    <td height="25" colspan="6">**매입처(핑거스상품) 추가정보**</td>
	    <% else %>
		<td height="25" colspan="6">**매입처(일반) 추가정보**</td>
	    <% end if %>

	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><!--판매채널--></td>
		<td bgcolor="#FFFFFF" colspan=2>
		<input type="hidden" name="selltype" value="0">
		<!--
		<input type="radio" name="selltype" value="0"> 온/OFF 전체 <input type="radio" name="selltype" value="9"> 오프라인전용
		-->
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>">구매유형</td>
		<td bgcolor="#FFFFFF" colspan=2>
			<% drawPartnerCommCodeBox false,"purchasetype","purchasetype","1","" %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">온라인대표<br>관리카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode","" %></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >온라인 담당MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "mduserid", session("ssBctId") , "on" %></td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">전시카테고리<br>담당MD</td>
		<td bgcolor="#FFFFFF" colspan=5><%= fnStandardDispCateSelectBox(1,"", "standardmdcatecode", "", "")%></td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>">오프라인 카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "offcatecode", ""  %></td>
		<td bgcolor="<%= adminColor("pink") %>" >오프라인 담당MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "offmduserid", "" , "off" %></td>
	</tr>
	<tr>
		<td rowspan="3" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 브랜드 사용여부<br>&nbsp;&nbsp;(카테고리노출)</td>
		<td bgcolor="#FFFFFF">텐바이텐</td>
		<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" checked >사용 <input type=radio name="isusing" value="N" >사용안함</td>
		<td rowspan="3" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 스트리트 표시여부<br>&nbsp;&nbsp;(브랜드운영관련)</td>
		<td bgcolor="#FFFFFF">텐바이텐</td>
		<td bgcolor="#FFFFFF"><input type=radio name="streetusing" value="Y" checked >사용 <input type=radio name="streetusing" value="N" >사용안함</td>
	</tr>
	<tr >
		<td bgcolor="#FFFFFF">제휴몰</td>
		<td bgcolor="#FFFFFF">
			Y (등록후 수정가능)
			<input type="hidden" name="isextusing" value="Y">

			<!-- 스트리트 표시여부 제휴몰 히든처리. 스트리트 표시여부 텐바이텐 과 값 동일. //-->
			<input type="hidden" name="extstreetusing" value="">
			<!--
			<input type=radio name="isextusing" value="Y" >사용 <input type=radio name="isextusing" value="N" checked >사용안함.

			<td bgcolor="#FFFFFF">제휴몰</td>
			<td bgcolor="#FFFFFF"><input type=radio name="extstreetusing" value="Y" >사용 <input type=radio name="extstreetusing" value="N" checked >사용안함	</td>
			-->
		</td>
		<td bgcolor="#FFFFFF">커뮤니티(상품Q/A)</td>
		<td bgcolor="#FFFFFF"><input type=radio name="specialbrand" value="Y" checked >사용 <input type=radio name="specialbrand" value="N" >사용안함</td>
	</tr>

	<tr >
		<td bgcolor="#FFFFFF" height="24">텐바이텐 OFF</td>
		<td bgcolor="#FFFFFF">
			N (등록후 수정가능)
			<input type="hidden" name="isoffusing" value="N">
		</td>
		<td bgcolor="#FFFFFF"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("pink") %>">Only 노출</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="onlyflg" value="Y"  >Y <input type=radio name="onlyflg" value="N" checked >N</td>
		<td bgcolor="<%= adminColor("pink") %>">Artist 노출</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="artistflg" value="Y"  >Y <input type=radio name="artistflg" value="N" checked >N</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">K-Design 노출</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="kdesignflg" value="Y"  >Y <input type=radio name="kdesignflg" value="N" checked >N</td>
		<td bgcolor="#FFFFFF" colspan="3"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**계약관련사항**</td>
	</tr>
	<!--
	<tr bgcolor="#FFFFFF">
		<td colspan=6>
		* 온라인 매입일경우 -> 오프라인도 매입으로 설정.<br>
		* 온라인 위탁일경우 -> 오프라인도 위탁으로 설정.
		</td>
	</td>
	-->
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> 브랜드 기본마진</td>
		<td bgcolor="#FFFFFF" colspan=5>
			<table cellpadding="1" cellspacing="1" border="0" class="a">
			<tr>
				<td>
					<% DrawBrandMWUCombo_2011 "maeipdiv","" %>
					<input type="text" class="text" name="defaultmargine" value="" size="4" style="text-align:right"> %
				</td>
			</tr>
			<tr id="ddtdiv" style="display:none;">
				<td>
					업체배송조건설정:
					<select class='select' name="defaultdeliverytype" onchange="inputDeliveryPay(this.value)">
						<option value="null" selected>업체무료배송</option>
						<option value="9">업체조건배송</option>
						<option value="7">업체착불배송</option>
					</select>
				</td>
			</tr>
			<tr id="paydiv" style="display:none;">
				<td>
					<input type="text" name="defaultFreeBeasongLimit" value="" size="7" maxlength="7">원 미만 구매시 배송료 <input type="text" name="defaultDeliverPay" value="" size="7" maxlength="7">원
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<!--
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> 브랜드 기본마진</td>
		<td bgcolor="#FFFFFF" colspan="5">
			매입 <input type="text" class="text" name="" value="" size="4"> %  /
			위탁 <input type="text" class="text" name="" value="" size="4"> %  /
			업체배송 <input type="text" class="text" name="" value="" size="4"> %
			(향후, 이 기준값으로 변경예정)
		</td>
	</tr>
    -->
	<tr height="40">
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>

</table>
<% end if %>

<% if (pcuserdiv="9999_14") then %>
<!-- 아카데미 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**매입처(아카데미) 추가정보**
		<input type="hidden" name="selltype" value="10">        <!-- 자동설정 -->
		<input type="hidden" name="purchasetype" value="0">
		<input type="hidden" name="catecode" value="999">       <!-- 전시안함 -->
		<input type="hidden" name="offcatecode" value="999">    <!-- 전시안함 -->
		<input type="hidden" name="offmduserid" value="">       <!-- OFFMD 없음 -->

		<input type="hidden" name="isextusing" value="N">       <!-- 제휴몰 사용안함 -->
		<input type="hidden" name="extstreetusing" value="N">   <!-- 제휴몰 Street 사용안함 -->
		<input type="hidden" name="isoffusing" value="N">       <!-- OFF 사용안함 -->
		<input type="hidden" name="specialbrand" value="N">     <!-- specialbrand 커뮤니티 사용안함 -->
		<input type="hidden" name="onlyflg" value="N">          <!-- onlyflg 사용안함 -->
		<input type="hidden" name="artistflg" value="N">        <!-- artistflg 사용안함 -->
		<input type="hidden" name="kdesignflg" value="N">       <!-- kdesignflg 사용안함 -->

		<input type="hidden" name="maeipdiv" value="M">         <!-- 매입구분 -->
		<input type="hidden" name="defaultmargine" value="50">  <!-- 기본마진(유동) -->
		<input type="hidden" name="defaultdeliverytype" value="">

		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"></td>
		<td bgcolor="#FFFFFF" ></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >담당MD</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "" , "fingers" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">브랜드<br>사용여부</td>
		<td bgcolor="#FFFFFF">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N</td>
		<td bgcolor="<%= adminColor("pink") %>">스트리트<br>표시여부<br>(브랜드운영관련)</td>
		<td bgcolor="#FFFFFF">
		    <input type=radio name="streetusing" value="Y" checked >Y
		    <input type=radio name="streetusing" value="N" >N</td>
	</tr>

	<tr >
		<td width="120" bgcolor="#DDDDFF" rowspan="2">강좌 진행 여부</td>
		<td bgcolor="#FFFFFF" rowspan="2">
		<input type="radio" name="lec_yn" value="Y" checked onClick="clickLec(this)"> Y
		<input type="radio" name="lec_yn" value="N" onClick="clickLec(this)"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">강좌기본마진</td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="lec_margin" value="" size="4" maxlength="3"> (%)
		</td>
	</tr>
	<tr>
	    <td width="120" bgcolor="#DDDDFF">재료기본마진</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="mat_margin" value="" size="4" maxlength="3"> (%)
		</td>
	</tr>
	<tr >
		<td width="120" bgcolor="#DDDDFF" >DIY 진행 여부</td>
		<td  bgcolor="#FFFFFF" width="200" >
		<input type="radio" name="diy_yn" value="Y"  onClick="clickDiy(this);"> Y
		<input type="radio" name="diy_yn" value="N"  checked  onClick="clickDiy(this);"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">기본마진</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="diy_margin" value="" size="4" maxlength="3"> (%)
		</td>
	</tr>

	<tr id="iDiyDlv" style="display:none">
		<td width="120" bgcolor="#DDDDFF">DIY배송구분</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<select name="diy_dlv_gubun" onChange="chkCompdiygbn(this);">
		<option value="0" >기본(업체무료배송)
		<option value="9" selected >업체 조건배송
		</select>
		<br>
		<input type="hidden" name="pDFL" value="">
		<input type="hidden" name="pDDP" value="">
		<input type="text" name="DefaultFreebeasongLimit" value="" size="9" maxlength="9">원 이상 무료배송
		/미만 배송비 <input type="text" name="DefaultDeliverPay" value="" size="9" maxlength="9">원
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>

<%
	'///// 추가정보 - 999(제휴사) /////
	if (pcuserdiv="999_50") then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> 추가정보**</td>
		<input type="hidden" name="catecode" value="999"> <!-- 전시안함 -->
		<input type="hidden" name="offcatecode" value="999"> <!-- 전시안함 -->

		<input type="hidden" name="isextusing" value="N"> <!-- 제휴몰 사용안함 -->
		<input type="hidden" name="streetusing" value="N"> <!--  Street 사용안함 -->
		<input type="hidden" name="extstreetusing" value="N"> <!-- 제휴몰 Street 사용안함 -->
		<input type="hidden" name="isoffusing" value="N">
		<input type="hidden" name="specialbrand" value="N"> <!-- specialbrand 커뮤니티 사용안함 -->
		<input type="hidden" name="onlyflg" value="N"> <!-- onlyflg 사용안함 -->
		<input type="hidden" name="artistflg" value="N"> <!-- artistflg 사용안함 -->
		<input type="hidden" name="kdesignflg" value="N"> <!-- kdesignflg 사용안함 -->

		<input type="hidden" name="maeipdiv" value="M">   <!-- 매입구분 -->
		<input type="hidden" name="defaultmargine" value="20">  <!-- 기본마진(유동) -->

		<input type="hidden" name="M_margin" value="0">
		<input type="hidden" name="W_margin" value="0">
		<input type="hidden" name="U_margin" value="0">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">브랜드 사용여부</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >담당자(영업)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "", "sell" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 매출계정</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"sellacccd","selltype","","" %>

		</td>
		<td bgcolor="<%= adminColor("pink") %>">기본 정산방식</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","","" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 매출부서</td>
		<td bgcolor="#FFFFFF">
		   <%= fndrawSaleBizSecCombo(true,"sellBizCd","","") %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">수수료</td>
		<td bgcolor="#FFFFFF">
		    <input type="text" name="commission" value="" size="4">%
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">계산서발행방식</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"taxevaltype","taxevaltype","","" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">(기타매출)정산방법</td>
		<td bgcolor="#FFFFFF">
        <% drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype","","" %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> 제휴사 입점 정보**</td>

	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">제휴형태</td>
		<td bgcolor="#FFFFFF">
		   <% drawPartnerCommCodeBox true,"mallSellType","pmallSellType","","" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">연동방식</td>
		<td bgcolor="#FFFFFF">
		    <% drawPartnerCommCodeBox true,"pcomType","pcomType","","" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">제휴어드민URL</td>
		<td bgcolor="#FFFFFF" colspan="3">
		   <input type="text" name="padminUrl" value="" size="60" maxlength="120">
		</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("pink") %>">제휴어드민계정</td>
		<td bgcolor="#FFFFFF" colspan="">
		   ID <input type="text" name="padminId" value="" size="10" maxlength="32">
		   PW <input type="password" name="padminPwd" value="" size="10" maxlength="32">
		</td>
		<td bgcolor="<%= adminColor("pink") %>">주문처리담당</td>
		<td bgcolor="#FFFFFF">
            <% drawSelectBoxCoWorker_OnOff "offmduserid", "", "sell" %>
		</td>
	</tr>
    <tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 배송비 조건</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="hidden" name="defaultdeliverytype" value="9">
			<input type="text" name="defaultFreeBeasongLimit" value="" size="8" maxlength="7">원 미만 구매시
			배송료 <input type="text" name="defaultDeliverPay" value="" size="7" maxlength="7">원
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>


<%
	'///// 추가정보 - 900(출고처), 902(협력업체), 903(3PL대표) /////
	if (pcuserdiv="900_21") or (pcuserdiv="902_21") or (pcuserdiv="903_21") then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> 추가정보**</td>
		<input type="hidden" name="catecode" value="999"> <!-- 전시안함 -->
		<input type="hidden" name="offcatecode" value="999"> <!-- 전시안함 -->

		<input type="hidden" name="isextusing" value="N"> <!-- 제휴몰 사용안함 -->
		<input type="hidden" name="streetusing" value="N"> <!--  Street 사용안함 -->
		<input type="hidden" name="extstreetusing" value="N"> <!-- 제휴몰 Street 사용안함 -->
		<input type="hidden" name="isoffusing" value="N">
		<input type="hidden" name="specialbrand" value="N"> <!-- specialbrand 커뮤니티 사용안함 -->
		<input type="hidden" name="onlyflg" value="N"> <!-- onlyflg 사용안함 -->
		<input type="hidden" name="artistflg" value="N"> <!-- artistflg 사용안함 -->
		<input type="hidden" name="kdesignflg" value="N"> <!-- kdesignflg 사용안함 -->

		<input type="hidden" name="maeipdiv" value="M">   <!-- 매입구분 -->
		<input type="hidden" name="defaultmargine" value="20">  <!-- 기본마진(유동) -->

		<input type="hidden" name="M_margin" value="0">
		<input type="hidden" name="W_margin" value="0">
		<input type="hidden" name="U_margin" value="0">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">브랜드 사용여부</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >담당자(영업)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "", "sell" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 정산방식</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","","" %>
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>


<%
	'///// 추가정보 - 501(직영점), 502(가맹점), 503(도매처) /////
	if (pcuserdiv="501_21") or (pcuserdiv="502_21") or (pcuserdiv="503_21") then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> 추가정보**</td>
		<input type="hidden" name="catecode" value="999"> <!-- 전시안함 -->
		<input type="hidden" name="offcatecode" value="999"> <!-- 전시안함 -->

		<input type="hidden" name="isextusing" value="N"> <!-- 제휴몰 사용안함 -->
		<input type="hidden" name="streetusing" value="N"> <!--  Street 사용안함 -->
		<input type="hidden" name="extstreetusing" value="N"> <!-- 제휴몰 Street 사용안함 -->
		<input type="hidden" name="isoffusing" value="N">
		<input type="hidden" name="specialbrand" value="N"> <!-- specialbrand 커뮤니티 사용안함 -->
		<input type="hidden" name="onlyflg" value="N"> <!-- onlyflg 사용안함 -->
		<input type="hidden" name="artistflg" value="N"> <!-- artistflg 사용안함 -->
		<input type="hidden" name="kdesignflg" value="N"> <!-- kdesignflg 사용안함 -->

		<input type="hidden" name="maeipdiv" value="M">   <!-- 매입구분 -->
		<input type="hidden" name="defaultmargine" value="20">  <!-- 기본마진(유동) -->

		<input type="hidden" name="M_margin" value="0">
		<input type="hidden" name="W_margin" value="0">
		<input type="hidden" name="U_margin" value="0">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">브랜드 사용여부</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" checked >Y
		    <input type=radio name="isusing" value="N" >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >담당자(영업)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", "", "sell" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 매출계정</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"sellacccd","selltype","","" %>

		</td>
		<td bgcolor="<%= adminColor("pink") %>">기본 정산방식</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","0","" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 매출부서</td>
		<td bgcolor="#FFFFFF">
		   <%= fndrawSaleBizSecCombo(true,"sellBizCd","","") %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">수수료</td>
		<td bgcolor="#FFFFFF">
		    <input type="text" name="commission" value="" size="4">%
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">계산서발행방식</td>
		<td bgcolor="#FFFFFF">
		<% drawPartnerCommCodeBox true,"taxevaltype","taxevaltype","","" %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">(기타매출)정산방법</td>
		<td bgcolor="#FFFFFF">
        <% drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype","","" %>
		</td>
	</tr>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>
</table>
<% end if %>

<% if (pcuserdiv="") then %>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="30"><td align="center">[먼저 브랜드 구분을 선택 하세요.]</td></tr>
</table>
<% end if %>

</form>

<!--
!!!!!! icheckframe.asp 두번 돌린다. (AddProc() 참조)
-->
<iframe src="" name="icheckframe" width="200" height="0" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
<iframe name="FrameCKP" src="" frameborder="0" width="600"  height="400"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

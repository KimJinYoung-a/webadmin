<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 브랜드
' History : 2018.08.03 정태훈 생성
'			2022.02.24 한용민 수정(일반(간이)사업자, 원천징수, 해외사업자 체크생성후 저장하는 로직 생성)
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
dim groupid, categorylarge

dim pcuserdiv : pcuserdiv = requestCheckVar(request("pcuserdiv"),16)
dim hp : hp = requestCheckVar(request("hp"),16)
dim email : email = requestCheckVar(request("email"),128)
dim cd1 : cd1 = requestCheckVar(request("cd1"),3)
dim cate1 : cate1 = requestCheckVar(request("cate1"),3)
dim companyno : companyno = requestCheckVar(request("companyno"),16)

if cate1 <> "" then
	categorylarge=cate1
else
	categorylarge=cd1
end if

'/2013.12.02 한용민 추가
if not(C_ADMIN_AUTH or C_MngPart or C_partnership_part) then
	if pcuserdiv="999_50" or pcuserdiv="501_21" or pcuserdiv="502_21" or pcuserdiv="900_21" then
		response.write "<script language='javascript'>"
		response.write "	alert('[권한없음] 매입처만 등록 가능 합니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
end if
%>
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
<script type='text/javascript'>

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

function precheck(frm){

    var pcuserdiv = getFieldValue(frm.pcuserdiv);
    // 9999_02:매입처, 9999_14:아카데미, 999_50:제휴사(온라인) , 501_21:직영매장 503_21:기타매장 ,9999_21:출고처(기타)

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

	if(frm.uid.value.search(/\W|\s/g) > -1){
		alert("브랜드 아이디에는 특수문자 또는 공백을 입력할 수 없습니다.");
		va.focus();
	}

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

	if (pcuserdiv=="9999_02"){
		if (frm.jungsan_date.value=="" && frm.jungsan_date_off.value==""){
			alert('정산일을 선택해 주세요.');
			return;
		}

		if (frm.catecode.value.length<1){
			alert('온라인 카테고리 구분을 선택하세요.');
			//frm.catecode.focus();
			return;
		}

		if (frm.standardmdcatecode.value.length<1){
			alert('전시카테고리 담당MD를 선택하세요.');
			frm.standardmdcatecode.focus();
			return;
		}
		
		if (frm.mduserid.value.length<1){
			alert('담당MD 구분을 선택하세요.');
			frm.mduserid.focus();
			return;
		}
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

    	if (frm.catecode.value.length<1){
    		alert('온라인 카테고리 구분을 선택하세요.');
    		//frm.catecode.focus();
    		return;
    	}
		if (frm.standardmdcatecode.value.length<1){
			alert('전시카테고리 담당MD를 선택하세요.');
			frm.standardmdcatecode.focus();
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
            if (frm.defaultDeliverPay.value*1<2000){
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
	if ((pcuserdiv=="999_50")||(pcuserdiv=="900_21")||(pcuserdiv=="902_21")||(pcuserdiv=="503_21")){
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

	if ((pcuserdiv!="999_50")&&(pcuserdiv!="900_21")&&(pcuserdiv!="902_21")&&(pcuserdiv!="503_21")){
		//스트리트 표시여부 제휴몰을 텐바이텐과 통일.
		if(frm.streetusing[0].checked){
			frm.extstreetusing.value = "Y";
		}else if(frm.streetusing[1].checked){
			frm.extstreetusing.value = "N";
		}
	}

	if (pcuserdiv=="9999_02" || pcuserdiv=="902_21" || pcuserdiv=="503_21"){
		if (frm.email.value==""){
			alert('업체 담당자 이메일을 입력해 주세요.');
			frm.email.focus();
			return;
		}

		if (frm.hp.value==""){
			alert('업체 담당자 핸드폰 번호를 입력해 주세요.');
			frm.hp.focus();
			return;
		}
	}

	if (frm.signtype.value==""){
		alert('신규 계약 등록 구분을 선택해주세요.');
		frm.signtype.focus();
		return;
	}

	// 해외사업자
	if ( $("input[name=businessgubun]:radio:checked").val()=="5" ){
		if (frm.company_no.value==""){
			alert('해외 사업자번호를 확인해주세요.');
			frm.company_no.focus();
			return;
		}
		if (frm.partcheck.value==""){
			alert('해외 사업자번호를 확인해주세요.');
			frm.partcheck.focus();
			return;
		}
	// 원천징수
	}else if ( $("input[name=businessgubun]:radio:checked").val()=="3" ){
		if (frm.company_no.value==""){
			alert('주민등록번호를 입력해주세요.');
			frm.company_no.focus();
			return;
		}
		if (frm.partcheck.value==""){
			alert('주민등록번호를 입력하고 확인해주세요.');
			frm.partcheck.focus();
			return;
		}
	// 일반(간이)사업자
	}else{
		if (frm.company_no.value==""){
			alert('사업자번호를 입력해주세요.');
			frm.company_no.focus();
			return;
		}
		if (frm.partcheck.value==""){
			alert('사업자번호를 입력하고 확인해주세요.');
			frm.partcheck.focus();
			return;
		}
	}

	if (frm.partcheck.value==""){
		alert('사업자번호를 입력하고 확인해주세요.');
		frm.partcheck.focus();
		return;
    }

	var ret = confirm('브랜드 정보를 저장 하시겠습니까?');

	if (ret){
		frm.target="FrameCKP";
		frm.action="doupchebrand.asp";
		frm.submit();
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

function PopUpcheSelectCustom(frmname){
	document.frmbrand.mode.value = "addnewupchebrand2";
	var popwin = window.open("/admin/member/popupcheselect.asp?mode=newbrand&frmname=" + frmname,"popupcheselect","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function viewtable(){
	document.getElementById("upchediv1").style.display = "";
	document.getElementById("upchediv2").style.display = "";
	document.getElementById("upchediv3").style.display = "";
	document.getElementById("upchediv4").style.display = "";
	document.getElementById("upchediv5").style.display = "";
	document.getElementById("upchediv6").style.display = "";
	document.getElementById("upchediv7").style.display = "";
	document.getElementById("upchediv8").style.display = "";
	document.getElementById("upchediv9").style.display = "";
	document.getElementById("upchediv10").style.display = "";
	document.getElementById("upchediv11").style.display = "";
	document.getElementById("upchediv12").style.display = "";
	document.getElementById("upchediv13").style.display = "";
	//document.getElementById("upchediv14").style.display = "";
	document.getElementById("upchediv15").style.display = "";
	document.getElementById("upchediv16").style.display = "";
	document.getElementById("upchediv17").style.display = "";
	document.getElementById("upchediv18").style.display = "";
	document.getElementById("upchediv19").style.display = "";
}

function DisableSocInfo(){
	var frm = document.frmbrand;
<% if (C_ADMIN_AUTH <> true) then %>
	frm.company_name.readOnly = true;
	frm.company_no.readOnly = true;
	frm.ceoname.readOnly = true;
	frm.jungsan_gubun.readOnly = true;

	$("input[name='company_name']").css("background-color","#EEEEEE");
	$("input[name='company_no']").css("background-color","#EEEEEE");
	$("input[name='ceoname']").css("background-color","#EEEEEE");
	$("input[name='jungsan_gubun']").css("background-color","#EEEEEE");
<% end if %>
}

function businessgubun_change(){
	var businessgubun = $("input[name=businessgubun]:radio:checked").val();
	var businessgubun3 = document.getElementById("businessgubun3");

	// 해외사업자
	if (businessgubun=="5"){
		document.getElementById("businessgubun1").style.display = "none";
		document.getElementById("businessgubun2").style.display = "none";
		document.getElementById("businessgubun3").style.display = "";

		// 해외는 사업자번호 자동설정된다(888-00-00000)
		$("#company_no").val("888-00-00000");
		$("#company_no3").val("888-00-00000");
		//$("#company_no3").attr("readonly",true);

		//if (frm.coSearchBtn) {
		//	frm.coSearchBtn.disabled = true;
		//}

	// 원천징수
	} else if (businessgubun=="3"){
		document.getElementById("businessgubun1").style.display = "none";
		document.getElementById("businessgubun2").style.display = "";
		document.getElementById("businessgubun3").style.display = "none";

	// 일반(간이)사업자
	} else {
		document.getElementById("businessgubun1").style.display = "";
		document.getElementById("businessgubun2").style.display = "none";
		document.getElementById("businessgubun3").style.display = "none";
	}
}

function fnCheckUpcheNo(frm){
	var businessgubun = $("input[name=businessgubun]:radio:checked").val();

	// 해외사업자
	if (businessgubun=="5"){
		$("#company_no").val($("#company_no3").val())
		frm.target="FrameCKP";
		frm.action="checkUpcheSelect.asp";
		frm.submit();

	// 원천징수
	} else if (businessgubun=="3"){
		$("#company_no").val($("#company_no2").val())
		var company_no=$("#company_no").val().replace("-", "");

		if (!jsChkSocialNum1(company_no)){
			alert('주민등록번호를 다시 입력해 주세요.');
			return;
		}
		frm.target="FrameCKP";
		frm.action="checkUpcheSelect.asp";
		frm.submit();
		
	// 일반(간이)사업자
	} else {
		var bizNo = frm.company_no1.value;
		bizNo = bizNo.replace(/-/gi,"");
		if(frm.company_no1.value==""){
			alert("사업자번호를 입력해주세요.");
		}
		else{
			var sumMod=0;
			sumMod += parseInt(bizNo.substring(0,1));
			sumMod += parseInt(bizNo.substring(1,2)) * 3 % 10;
			sumMod += parseInt(bizNo.substring(2,3)) * 7 % 10;
			sumMod += parseInt(bizNo.substring(3,4)) * 1 % 10;
			sumMod += parseInt(bizNo.substring(4,5)) * 3 % 10;
			sumMod += parseInt(bizNo.substring(5,6)) * 7 % 10;
			sumMod += parseInt(bizNo.substring(6,7)) * 1 % 10;
			sumMod += parseInt(bizNo.substring(7,8)) * 3 % 10;
			sumMod += Math.floor(parseInt(bizNo.substring(8,9)) * 5 / 10);
			sumMod += parseInt(bizNo.substring(8,9)) * 5 % 10;
			sumMod += parseInt(bizNo.substring(9,10));

			if(sumMod % 10 != 0){
				alert("사업자 등록번호가 잘 못 되었습니다.");
				return false;
			}else if ($("#company_no1").val().length != 12){
				alert('사업자 등록 번호는 000-00-00000 형식으로 입력해야 합니다..');
				$("#company_no1").focus();
				return;
			}else{
				$("#company_no").val($("#company_no1").val())
				frm.target="FrameCKP";
				frm.action="checkUpcheSelect.asp";
				frm.submit();
			}
		}
	}
}

function fnbizNoHyphen(num) {
     num = num.replace(/-/g, "");
     var num_str = num.toString();
     var result = '';
 
      for(var i=0; i<num_str.length; i++) {
            var tmp = num_str.length-(i+1);
            if(i==5){
				result = '-' + result;
			}
			else if(i==7){
				result = '-' + result;
			}
            result = num_str.charAt(tmp) + result;
       }
       return result;
}

function fnjuminNoHyphen(num) {
     num = num.replace(/-/g, "");
     var num_str = num.toString();
     var result = '';
 
      for(var i=0; i<num_str.length; i++) {
            var tmp = num_str.length-(i+1);
            if(i==7){
				result = '-' + result;
			}
            result = num_str.charAt(tmp) + result;
       }
       return result;
}

</script>
<% if (pcuserdiv="") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext">
    <tr height="30" bgcolor="#FFFFFF">
    <td width="150" bgcolor="<%= adminColor("pink") %>">1. 브랜드 구분 선택</td>
    <td >
        <% drawPartnerCommCodeBox false,"pcuserdiv","pcuserdiv","9999_02","" %>

        <script>delcomRow();</script>
    </td>
</tr>
<tr>
    <td colspan="2" height="30" bgcolor="#FFFFFF" align="center"><input type="button" value="다음" onClick="stepNext();"></td>
</tr>
</form>
</table>
<% else %>
<form name="frmbrand" method="post" action="/admin/member/doupchebrand.asp" target="FrameCKP">
<input type="hidden" name="mode" value="addnewupchebrand">
<input type="hidden" name="partnerCnt" value="">
<input type="hidden" name="partcheck" value="">
<input type="hidden" name="defaultsongjangdiv" value="">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="30" >
    	<td bgcolor="<%= adminColor("pink") %>" colspan="6"><b>브랜드관련정보</b></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td height="25" colspan="4">**브랜드 기본정보**</td>
    </tr>
	<tr height="50">
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font> 브랜드ID</td>
		<td bgcolor="#FFFFFF" >
    		<input type="text" class="text" name="uid" value="" size="24" maxlength="24">
    		<div>(영문, 숫자만 가능 특수문자 금지)</div>
		</td>
		<td width="100"  bgcolor="<%= adminColor("pink") %>"><font color=red>*</font>브랜드 구분</td>
		<td bgcolor="#FFFFFF" >
			<%= getPartnerCommCodeName("pcuserdiv",pcuserdiv) %>
            <input type="hidden" name="pcuserdiv" value="<%= pcuserdiv %>">
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
		<td width="100" bgcolor="<%= adminColor("pink") %>" >구매유형</td>
		<td bgcolor="#FFFFFF" colspan=2>
		<input type="hidden" name="selltype" value="0">
			<% drawPartnerCommCodeBox false,"purchasetype","purchasetype","1","" %>
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>">정산일</td>
		<td bgcolor="#FFFFFF" colspan=2>
			온라인 : <% DrawJungsanDateCombo "jungsan_date", "" %>
			&nbsp;
			오프라인 : <% DrawJungsanDateCombo "jungsan_date_off", "" %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font>온라인대표<br>관리카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", categorylarge %></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font>온라인 담당MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "mduserid", session("ssBctId") , "on" %></td>
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"><font color=red>*</font>전시카테고리<br>담당MD</td>
		<td bgcolor="#FFFFFF" colspan=5><%= fnStandardDispCateSelectBox(1,"", "standardmdcatecode", "", "")%></td>
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
					<select class="select" name="defaultdeliverytype" onchange="inputDeliveryPay(this.value)">
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
	<tr>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> 업체 담당자 E-Mail</td>
		<td bgcolor="#FFFFFF" colspan="2">
			<input type="text" class="text" name="email" size="30" value="<%=email%>">
		</td>
		<td bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> 업체 담당자 핸드폰</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type="text" class="text" name="hp"  value="<%=hp%>" size="15"></td>
	</tr>


</table>
<% end if %>

<% if (pcuserdiv="902_21") or (pcuserdiv="503_21") then %>
<input type="hidden" name="jungsan_date" value="">
<input type="hidden" name="jungsan_date_off" value="">
<input type="hidden" name="defaultmargine" value="20">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> 업체 담당자 E-Mail</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="email" size="30" value="<%=email%>">
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" ><font color=red>*</font> 업체 담당자 핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="hp"  value="<%=hp%>" size="15"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 정산방식</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype","","" %>
		</td>
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

<% if (pcuserdiv="999_50") or (pcuserdiv="900_21") then %>
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
	<% if (pcuserdiv="999_50") then %>
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
	<% end if %>
	<tr height="40">
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>
</table>

<% end if %>
<p>
	<input type="radio" name="signtype" value="1">신규계약등록(수기계약)
	<input type="radio" name="signtype" value="2">신규계약등록(U+전자서명)
	<input type="radio" name="signtype" value="3">신규계약등록(DocuSign)
<p>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4"><b>업체관련정보</b></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" colspan="4">
			<input type="hidden" id="company_no" name="company_no" value="">
			<input type="hidden" name="groupid" value="">
			<input type="radio" name="businessgubun" value="1" checked onclick="businessgubun_change();" >일반(간이)사업자
			<input type="radio" name="businessgubun" value="3" onclick="businessgubun_change();" >원천징수
			<input type="radio" name="businessgubun" value="5" onclick="businessgubun_change();" >해외사업자
			&nbsp;&nbsp;
			<span id="bizCheck"></span>
		</td>
	</tr>
	<tr id="businessgubun1" style="display:">
		<td bgcolor="<%= adminColor("tabletop") %>">사업자 번호</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" class="text" id="company_no1" name="company_no1" value="<%=companyno%>" size="15" onkeyup="this.value=fnbizNoHyphen(this.value)">
			<input type="button" class="button" value="확인" onClick="fnCheckUpcheNo(this.form);">&nbsp;
		</td>
	</tr>
	<tr id="businessgubun2" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">주민번호</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" class="text" id="company_no2" name="company_no2" value="<%=companyno%>" size="15" onkeyup="this.value=fnjuminNoHyphen(this.value)">
			<input type="button" class="button" value="확인" onClick="fnCheckUpcheNo(this.form);">&nbsp;
		</td>
	</tr>
	<tr id="businessgubun3" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">해외사업자 번호</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" class="text" id="company_no3" name="company_no3" value="<%=companyno%>" size="15" >
			<input type="button" class="button" value="확인" onClick="fnCheckUpcheNo(this.form);">&nbsp;
			신규업체의 경우 해외사업자 번호는 888-00-00000 으로 입력해 주세요. 자동생성 됩니다.
		</td>
	</tr>
	<tr id="upchediv1" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">입점브랜드ID</td>
		<td height="25" colspan="3" bgcolor="#FFFFFF"></td>
	</tr>

	<tr id="upchediv2" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**사업자등록정보**&nbsp;&nbsp;&nbsp;(중복된 사업자번호는 등록할 수 없습니다.)</td>
	</tr>

	<tr id="upchediv3" style="display:none">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 회사명(상호)</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_name" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td width="90" bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 대표자</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="ceoname" value="" size="16" maxlength="16" style="background-color:#EEEEEE;" readonly></td>
	</tr>
	<tr id="upchediv4" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 과세구분</td>
		<input type="hidden" name="checksocnoyn" value="N">
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()" readonly>
			<option value="일반과세" >일반과세</option>
			<option value="간이과세" >간이과세</option>
			<option value="원천징수" >원천징수</option>
			<option value="면세" >면세</option>
			<option value="영세(해외)" >영세(해외)</option>
			</select>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font></td>
		<td bgcolor="#FFFFFF">			
		</td>
	</tr>
	<tr id="upchediv5" style="display:none">
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
	<tr id="upchediv6" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 업태</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 업종</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="" size="30" maxlength="32" style="background-color:#EEEEEE;" readonly></td>
	</tr>

	<tr id="upchediv7" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**업체기본정보**</td>
	</tr>

	<tr id="upchediv8" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 대표전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_tel)"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.company_fax)"></td>
	</tr>
	<tr id="upchediv9" style="display:none">
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
	<tr id="upchediv10" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**결제계좌정보**</td>
	</tr>

	<tr id="upchediv11" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<% DrawBankCombo "jungsan_bank", "" %>
		</td>
	</tr>
	<tr id="upchediv12" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctno" value="" size="24" maxlength="32" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
		</td>
	</tr>
	<tr id="upchediv13" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" class="text" name="jungsan_acctname" value="" size="24" maxlength="16" style="background-color:#EEEEEE;" readonly>
		&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
		</td>
	</tr>
	<tr id="upchediv15" style="display:none">
		<td colspan="4" bgcolor="#FFFFFF" height="25">**업체 담당자정보**</td>
	</tr>

	<tr id="upchediv16" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_phone)"></td>
	</tr>
	<tr id="upchediv17" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>"><font color=red>*</font> 핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.manager_hp)"></td>
	</tr>


	<tr id="upchediv18" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_phone)"></td>
	</tr>
	<tr id="upchediv19" style="display:none">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="" size="16" maxlength="16" onFocusOut="phone_format(frmbrand.jungsan_hp)"></td>
	</tr>
	<tr height="40">
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value=" 정보 저장 " onclick="precheck(frmbrand);"></td>
	</tr>
</table>

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
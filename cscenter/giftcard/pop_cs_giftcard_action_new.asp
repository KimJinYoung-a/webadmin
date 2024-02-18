<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/CSFunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%

'[코드정리]
'------------------------------------------------------------------------------
'A008			주문취소
'
'[변수정리]
'------------------------------------------------------------------------------
'CSFunction.asp
'
'dim IsStatusRegister			'접수
'dim IsStatusEdit				'수정
'dim IsStatusFinishing			'처리완료 시도
'dim IsStatusFinished			'처리완료

'dim IsDisplayPreviousCSList	'이전 CS 내역
'dim IsDisplayCSMaster			'CS 마스터정보
'dim IsDisplayItemList			'상품목록
'dim IsDisplayRefundInfo		'환불정보
'dim IsDisplayButton			'버튼
'
'dim IsPossibleModifyCSMaster
'dim IsPossibleModifyItemList
'dim IsPossibleModifyRefundInfo



dim i, id, mode, divcd, giftorderserial, ckAll, iPgGubun

id			= request("id")
divcd		= request("divcd")
giftorderserial	= request("giftorderserial")
mode		= request("mode")
ckAll		= request("ckAll")

dim IsOrderCanceled
dim OrderMasterState
dim IsTicketOrder



'==============================================================================
'CS접수마스터 가져오기
dim ocsaslist

set ocsaslist = New CCSASList

ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if



'==============================================================================
'CS접수마스터 정보가 없을경우 신규 접수
if (ocsaslist.FResultCount<1) then
	set ocsaslist.FOneItem = new CCSASMasterItem

	ocsaslist.FOneItem.FId = 0
	ocsaslist.FOneItem.Fdivcd = divcd

	mode = "regcsas"
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    giftorderserial = ocsaslist.FOneItem.Forderserial

    if (ocsaslist.FOneItem.FCurrState = "B007") then
		mode = "finished"
    else
    	if (mode = "finishreginfo") then
    		'
    	else
    		mode = "editreginfo"
    	end if
    end if
end if

Call SetCSVariable(mode, divcd)



'==============================================================================
''환불정보
dim orefund

set orefund = New CCSASList

orefund.FRectCsAsID = ocsaslist.FOneItem.FId

orefund.GetOneRefundInfo

if (orefund.FOneItem.Fencmethod = "TBT") then
	orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
elseif (orefund.FOneItem.Fencmethod = "PH1") then
	orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
end if

function Decrypt(encstr)
	if (Not IsNull(encstr)) and (encstr <> "") then
		Decrypt = TBTDecrypt(encstr)
		exit function
	end if
	Decrypt = ""
end function



'==============================================================================
''주문 마스타
dim ogiftcardordermaster

set ogiftcardordermaster = new cGiftCardOrder

ogiftcardordermaster.FRectgiftorderserial = giftorderserial

if Left(giftorderserial,1)<>"G" then
	response.write "잘못된 접속입니다."
	response.end
else
    ogiftcardordermaster.getCSGiftcardOrderDetail
end if

IsOrderCanceled = (ogiftcardordermaster.FOneItem.Fcancelyn = "Y")
OrderMasterState = ogiftcardordermaster.FOneItem.FIpkumDiv
'iPgGubun = (ogiftcardordermaster.FOneItem.Fpggubun)    ' 기프트카드는 pg구분 없음? 전역변수가 없어서 오류가 나서 우선 변수만 추가해둠.

'=============================================================================='==============================================================================
'원주문 상품금액
dim orgitemcostsum, orgpercentcouponpricesum

'접수상품 합계금액
dim regitemcostsum, regpercentcouponpricesum



'==============================================================================
''접수 불가시 메세지
dim JupsuInValidMsg

if (Left(giftorderserial,1)<>"A") and (ogiftcardordermaster.FResultCount<1) then
    response.write "<br><br>!!! 과거 주문내역이거나 주문 내역이 없습니다. - 관리자 문의 요망"
    dbget.close()	:	response.End
end if

''접수 가능 여부
dim IsJupsuProcessAvail

if (ogiftcardordermaster.FResultCount>0) then
	if (ogiftcardordermaster.FOneItem.FCancelyn <> "N") then
		JupsuInValidMsg = "정상 주문건만 취소 가능합니다."
		IsJupsuProcessAvail = false
	else
		JupsuInValidMsg = ""
		IsJupsuProcessAvail = true
	end if

	if (ogiftcardordermaster.FOneItem.Fjumundiv = "7") then
		JupsuInValidMsg = "등록된 Gift카드주문은 취소할 수 없습니다. 등록이전 상태로 전환하세요." & ogiftcardordermaster.FOneItem.Fjumundiv
		IsJupsuProcessAvail = false
	end if
else
    IsJupsuProcessAvail = false
end if

if (ogiftcardordermaster.FOneItem.Fsubtotalprice < 0) and (IsJupsuProcessAvail = true) then
	IsJupsuProcessAvail = false
	JupsuInValidMsg = "마이너스주문에 대해 CS접수할 수 없습니다."
end if

%>

<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript'>
var IsCsPowerUser               = <%= LCase(C_CSPowerUser) %>;

var IsStatusRegister 			= <%= LCase(IsStatusRegister) %>;
var IsStatusEdit 				= <%= LCase(IsStatusEdit) %>;
var IsStatusFinishing 			= <%= LCase(IsStatusFinishing) %>;
var IsStatusFinished 			= <%= LCase(IsStatusFinished) %>;

var IsDisplayPreviousCSList 	= <%= LCase(IsDisplayPreviousCSList) %>;
var IsDisplayCSMaster 			= <%= LCase(IsDisplayCSMaster) %>;
var IsDisplayItemList 			= <%= LCase(IsDisplayItemList) %>;
var IsDisplayRefundInfo 		= <%= LCase(IsDisplayRefundInfo) %>;
var IsDisplayButton 			= <%= LCase(IsDisplayButton) %>;

var IsCSCancelInfoNeeded		= <%= LCase(IsCSCancelInfoNeeded(divcd)) %>;
var IsCSRefundNeeded			= <%= LCase(IsCSRefundNeeded(divcd, OrderMasterState)) %>;

var IsPossibleModifyCSMaster	= <%= LCase(IsPossibleModifyCSMaster) %>;
var IsPossibleModifyItemList	= <%= LCase(IsPossibleModifyItemList) %>;
var IsPossibleModifyRefundInfo	= <%= LCase(IsPossibleModifyRefundInfo) %>;

var IsCSCancelProcess			= <%= LCase(IsCSCancelProcess(divcd)) %>;
var IsCSReturnProcess			= <%= LCase(IsCSReturnProcess(divcd)) %>;
var IsCSServiceProcess			= <%= LCase(IsCSServiceProcess(divcd)) %>;

var IsDeletedCS 				= <%= LCase(ocsaslist.FOneITem.FDeleteyn = "Y") %>;

var ERROR_MSG_TRY_MODIFY		= "<%= ERROR_MSG_TRY_MODIFY %>";

var CDEFAULTBEASONGPAY 		= <%=Cint(getDefaultBeasongPayByDate(now())) %>; // 텐바이텐 기본 배송비
var divcd 					= "<%= divcd %>";
var mode 					= "<%= mode %>";
var giftorderserial 			= "<%= giftorderserial %>";

var IsAdminLogin 			= IsCsPowerUser; ///<%= LCase((session("ssBctId") = "icommang") or (session("ssBctId") = "iroo4") or (session("ssBctId") = "bseo")) %>;
var IsOrderFound 			= <%= LCase(ogiftcardordermaster.FResultCount > 0) %>;
var IsRefundInfoFound 		= <%= LCase(orefund.FResultCount > 0) %>;

<% if (ogiftcardordermaster.FResultCount > 0) then %>
var IsThisMonthJumun 		= <%= LCase(datediff("m", ogiftcardordermaster.FOneItem.FRegdate, now()) <= 0) %>;
<% else %>
var IsThisMonthJumun 		= false;
<% end if %>

// ============================================================================
// 마스터 사유구분 설정
// ============================================================================
function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv){

	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
		return;
	}

    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// CS 사유구분 표시 (AJAX)
// ============================================================================
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv) {

	if ((IsPossibleModifyCSMaster != true) || (IsPossibleModifyItemList != true)) {
		alert(ERROR_MSG_TRY_MODIFY);
		return;
	}

    var params = "?gubun01=" + value_gubun01 + "&gubun02=" + value_gubun02 + "&name_gubun01=" + name_gubun01 + "&name_gubun02=" + name_gubun02 + "&name_gubun01name=" + name_gubun01name + "&name_gubun02name=" + name_gubun02name +"&name_frm=" + name_frm + "&targetDiv=" + targetDiv;
    initializeURL("/cscenter/action/ajax_cs_gubun_select.asp" + params);
    initializeReturnFunction("processAjaxCSGubunSelect(" + targetDiv + ")");
    initializeErrorFunction("onErrorAjaxCSGubunSelect()");
    startRequest();
}

function processAjaxCSGubunSelect(targetDiv) {
    eval(targetDiv).innerHTML = xmlHttp.responseText;
}

function onErrorAjaxCSGubunSelect() {
    alert("데이터를 읽는 도중에 에러가 발생했습니다. 잠시후 다시 시도해보시기 바랍니다.[CODE:" + xmlHttp.status + "]");
}

function colseCausepop(targetDiv){
    eval(targetDiv).innerHTML = "";
}

// ============================================================================
// CS 접수
// ============================================================================
function CsRegProc(frm) {

	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// 취소, 반품, 환불
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// 취소, 반품
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// 환불정보가 정상인지 체크
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// 환불수단이 옳바른지
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

    if (IsCSCancelProcess){
        if(confirm("취소 접수 하시겠습니까?")){
			if (frm.returnmethod) {
				if (frm.returnmethod.value == "R000") {
					frm.refundrequire.value = "0";
				}
			}

            frm.submit();
        }
    }else if (IsCSRefundNeeded) {
        if(confirm("환불 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if(confirm("접수 하시겠습니까?")){
        frm.submit();
    }
}

function CheckCSMasterForSave(frm) {
    if (frm.divcd.value.length<1){
        alert("접수 구분을 선택하세요.");
        frm.divcd.focus();
        return false;
    }

    if (frm.title.value.length<1) {
        alert("제목을 입력하세요.");
        frm.title.focus();
        return false;
    }

    if (frm.gubun01.value.length<1) {
        alert("사유 구분을 입력하세요.");
        return false;
    }

    return true;
}

// ============================================================================
// 환불 관련 체크 Form
// ============================================================================
function CheckReturnForm(frm) {
    if (!frm.returnmethod) { return true; }
    if (!frm.refundrequire) { return true; }

    if (frm.returnmethod.value.length < 1) {
        alert('환불 방식을 선택해 주세요.');
        frm.returnmethod.focus();
        return false;
    }

	if (frm.returnmethod.value == "R000") {
		// 환불없음 이면 체크 안한다.
		return true;
	}

	if (frm.refundrequire.value*0 != 0) {
        alert('환불 금액을 입력하세요.');
        frm.refundrequire.focus();
        return;
	}

	if ((frm.refundrequire.value*1 <= 0) && (frm.returnmethod.value != "R000")) {
		alert('환불 금액은 0 보다 커야합니다. 또는 환불없음을 선택하세요.');
        return false;
	}


	// ====================================================================
	if (frm.returnmethod.value=="R007") {
		// 무통장
        var mooconfirm = false;
        if ((frm.rebankaccount) && (frm.rebankaccount.value.length < 1)) {
            mooconfirm = true;
        }

        if ((frm.rebankownername) && (frm.rebankownername.value.length < 1)) {
            mooconfirm = true;
        }

        if ((frm.rebankname) && (frm.rebankname.value.length < 1)) {
            mooconfirm = true;
        }

        if (mooconfirm == true) {
        	// 무통장 계좌정보는 나중에 별도로 입력할 수 있다.
            if (!confirm('환불 계좌가 없습니다. \n\n환불 계좌 없이 등록 하시겠습니까?')) {
                if ((IsStatusRegister == true) || (IsStatusEdit == true)) {
                	frm.rebankaccount.focus();
                }
                return false;
            }
        }
	}

	if (frm.returnmethod.value == "R900") {
    	if (confirm("CS서비스가 아닌경우(결제금액환불) 마일리지 대신 예치금으로 환불하세요.\n\n마일리지 환불 하시겠습니까?") != true) {
    		return false;
    	}
	}

	// ====================================================================
	if ((frm.returnmethod.value=="R900") || (frm.returnmethod.value=="R910")) {
		// 마일리지, 예치금환불
        if ((frm.refund_userid) && (frm.refund_userid.value.length<1)) {
            alert('비회원에게 가능하지 않은 환불방식입니다. 다른 환불 방식을 선택하세요.');
            return false;
        }
	}

    return true;
}

function ChangeReturnMethod(comp){
    if (comp==undefined) return;

    var returnmethod = comp.value;

    document.all.refundinfo_R007.style.display = "none";
    document.all.refundinfo_R050.style.display = "none";
    document.all.refundinfo_R100.style.display = "none";
    document.all.refundinfo_R900.style.display = "none";

    if (comp.value=="R007"){
        //무통장 환불
        document.all.refundinfo_R007.style.display = "";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R120")||(comp.value=="R400")){
        //실시간 이체 취소//ALL@ 결제 취소 //신용카드 결제 취소 //신용카드 부분취소//휴대폰
        document.all.refundinfo_R100.style.display = "";
    }else if(comp.value=="R050"){
        //입점몰 결제 취소
        document.all.refundinfo_R050.style.display = "";
    }else if ((comp.value=="R900") || (comp.value=="R910")) {
        //마일리지 환불, 예치금 환불
        document.all.refundinfo_R900.style.display = "";
    }

}

// ============================================================================
// 환불 금액이 적절한지
// ============================================================================
function IsRefundInfoOK(frm) {

	if ((IsCSCancelProcess != true) && (IsCSReturnProcess != true)) {
		return true;
	}

	// 취소, 반품시 : 환불금액과 비교
	if (frm.orgsubtotalprice && frm.refundsubtotalprice) {
	    if (frm.orgsubtotalprice.value*1 < frm.refundsubtotalprice.value*1) {
	        alert('결제금액 이상으로 환불할 수 없습니다.\n\n마일리지, 쿠폰 등을 환원체크하세요.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.returnmethod) {
		if (frm.returnmethod.value == "R000") {
			// 환불없음 이면 체크 안한다.
			return true;
		}
	}

	if (frm.refundrequire && frm.returnmethod) {
	    if ((frm.refundsubtotalprice.value*1 < 1) && ((frm.returnmethod.value != "R000"))) {
	        alert('환불대상 금액이 없습니다.\n\n환불없음 또는 쿠폰, 마일리지 등을 환원체크 해제하세요');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	if (frm.remainsubtotalprice) {
	    if (frm.remainsubtotalprice.value*1 < 0) {
	        alert('취소 후 결제 금액이 마이너스가 될 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');
	        if (IsAdminLogin != true) {
	        	return false;
	        }
	    }
	}

	return true;
}

function CheckReturnMethod(frm) {
	if (!frm.returnmethod) { return true; }
	if (!frm.refundsubtotalprice) { return true; }

	if ((frm.returnmethod.value != "R100") && (frm.returnmethod.value != "R007") && (frm.returnmethod.value != "R020") && (frm.returnmethod.value != "R000")) {
        alert('적용불가 환불수단입니다.\n\n적용가능 환불수단 : 환불없음, 신용카드/실시간이체 취소, 무통장환불');
        return false;
	}

	if ((frm.accountdiv.value == "7") && ((frm.returnmethod.value != "R007") && (frm.returnmethod.value != "R000"))) {
        alert('무통장환불 또는 환불없음을 선택하세요');
        return false;
	}

	if ((frm.accountdiv.value == "100") && ((frm.returnmethod.value != "R100") && (frm.returnmethod.value != "R000"))) {
        alert('신용카드취소 또는 환불없음을 선택하세요');
        return false;
	}

	if ((frm.accountdiv.value == "20") && ((frm.returnmethod.value != "R020") && (frm.returnmethod.value != "R000"))) {
        alert('실시간이체취소 또는 환불없음을 선택하세요');
        return false;
	}

    return true;
}

// ============================================================================
// 삭제
// ============================================================================
function CsRegCancelProc(frm) {
    if (confirm('등록된 접수 내역을 삭제 하시겠습니까?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

// ============================================================================
// 수정
// ============================================================================
function CsRegEditProc(frm) {

	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	// 취소, 반품, 환불
	if ((IsCSCancelProcess == true) || (IsCSReturnProcess == true) || (divcd == "A003")) {
        if (CheckReturnForm(frm) != true) {
            return;
        }
	}

	// 취소, 반품
    if ((IsCSCancelProcess) || (IsCSReturnProcess)) {
		// 환불정보가 정상인지 체크
		if (IsRefundInfoOK(frm) != true) {
			return;
		}

		// 환불수단이 옳바른지
		if (CheckReturnMethod(frm) != true) {
			return;
		}
    }

    if (confirm('수정 하시겠습니까?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

// ============================================================================
// 완료처리
// ============================================================================
function CsRegFinishProc(frm) {
	var btn = document.getElementById("btnFinishReturn");

	// 마스터 체크
    if (!CheckCSMasterForSave(frm)) {
        return;
    }

	if (IsStatusFinishing && (divcd == "A007" || ((divcd == "A003") && (frm.returnmethod.value=="R007")))) {
		if (IsAdminLogin) {
			alert('이곳에서 완료처리 하여도 신용카드 승인취소/무통장 환불처리는 이루어 지지 않습니다.[어드민권한]');
		} else {
			alert('이곳에서 완료처리 하여도 신용카드 승인취소/무통장 환불처리는 이루어 지지 않습니다.\n\n완료 처리 할 수 없습니다.');
			return;
		}
	}

    //환불요청 , 신용카드 취소요청
    if ((divcd == "A003") || (divcd == "A007")) {
        if (frm.contents_finish.value.length<1){
            alert('처리 내용을 입력하세요.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '완료처리 진행 하시겠습니까?';

	if (btn) {
		btn.disabled = true;
	}

    if (confirm(confirmMsg )) {
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }

	if (btn) {
		btn.disabled = false;
	}
}
</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">
<form name="frmaction" method="post" action="pop_cs_giftcard_action_new_process.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="modeflag2" value="<%= mode %>">
<input type="hidden" name="giftorderserial" value="<%= giftorderserial %>" >
<input type="hidden" name="id" value="<%= ocsaslist.FOneItem.Fid %>">
<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
<input type="hidden" name="ipkumdiv" value="<%= ogiftcardordermaster.FOneItem.Fipkumdiv %>">
<input type="hidden" name="accountdiv" value="<%= ogiftcardordermaster.FOneItem.Faccountdiv %>">
<input type="hidden" name="orgitemcostsum" value="<%= ogiftcardordermaster.FOneItem.Fsubtotalprice %>">





<!-- ====================================================================== -->
<!-- 1. 이전 CS 내역                                                        -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_prev_cslist.asp" -->



<!-- ====================================================================== -->
<!-- 2. CS 마스터 정보                                                      -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_master_info.asp" -->



<!-- ====================================================================== -->
<!-- 3. 상품정보                                                            -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_item_list.asp" -->


</table>



<!-- ====================================================================== -->
<!-- 4. 취소/환불/업체정산 정보                                             -->
<!-- ====================================================================== -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" width="500" valign="top">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="5" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>취소관련 정보</b>
            </td>
        </tr>
		<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_cancel_info.asp" -->
      </table>
    </td>
    <td bgcolor="#FFFFFF" valign="top" align="left">
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>환불관련 정보</b>
            </td>
        </tr>
        <!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_refund_info.asp" -->
        </table>

        <p>

    </td>
</tr>
</table>
<!-- ====================================================================== -->
<!-- 4. 취소/환불/업체정산 정보                                             -->
<!-- ====================================================================== -->



<!-- ====================================================================== -->
<!-- 5. 버튼                                                                -->
<!-- ====================================================================== -->
<!-- #include virtual="/cscenter/giftcard/include/inc_cs_giftcard_action_button.asp" -->
<!-- ====================================================================== -->
<!-- 5. 버튼                                                                -->
<!-- ====================================================================== -->

</form>

<script>

// 페이지 시작시 작동하는 스크립트
function getOnload(){

	if (IsStatusFinishing && (divcd == "A007" || divcd == "A003")) {
		if ((divcd == "A003") && (!frmaction.returnmethod)) {
			alert("결제완료 이전 주문에 대해 환불할 수 없습니다.");
			frmaction.finishbutton.disabled = true;
		} else {
			if (divcd == "A007" || ((divcd == "A003") && (frmaction.returnmethod.value=="R007"))) {
				alert('이곳에서 완료처리 하여도 \n\n\n신용카드 승인취소/무통장 환불처리는 이루어 지지 않으니 유의하시기 바랍니다.!\n\n\n\n\n\n');
			}
		}
	}

}

window.onload = getOnload;

</script>

<%

set ogiftcardordermaster = Nothing

%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

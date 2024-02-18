<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
function IsMaySocialNo(icompanyno)
    IsMaySocialNo = false
    if isNULL(icompanyno) then Exit function
    IsMaySocialNo = LEN(trim(replace(icompanyno,"-","")))=13
end function

dim i
dim makerid, yyyy1,mm1, onoffGubun, jidx, isauto, nextjidx
makerid 		= requestCheckvar(request("makerid"),32)
yyyy1   		= requestCheckvar(request("yyyy1"),10)
mm1     		= requestCheckvar(request("mm1"),10)
onoffGubun     	= requestCheckvar(request("onoffGubun"),10)
jidx            = requestCheckvar(request("jidx"),10)
isauto          = requestCheckvar(request("isauto"),10)
nextjidx        = requestCheckvar(request("nextjidx"),10)
dim groupid
groupid = getPartnerId2GroupID(makerid)


'// ============================================================================
dim ojungsanTaxCC
set ojungsanTaxCC = new CUpcheJungsanTax
ojungsanTaxCC.FRectMakerid = makerid
ojungsanTaxCC.FRectTargetGbn = onoffGubun
ojungsanTaxCC.FRectJjungsanIdx = jidx
ojungsanTaxCC.getOneUpcheJungsanTax


dim PrdCommissionSum : PrdCommissionSum = 0

if (ojungsanTaxCC.FresultCount>0) then
	if (ojungsanTaxCC.FOneItem.IsCommissionTax) then
	    PrdCommissionSum = ojungsanTaxCC.FOneItem.Ftotalcommission
	end if
end if
rw makerid
rw onoffGubun
rw jidx
if PrdCommissionSum = 0 then
    if (request("autotype")="V2") then
    response.write "<script>"&vbCRLF
    response.write "opener.addResultLog('"&request("jidx")&"','수수료0');"&vbCRLF
    response.write "opener.fnNextEvalProc();"&vbCRLF
    response.write "</script>"
    else
	response.write "<script>alert('수수료 매출정보가 없습니다.');</script>"
	response.write "수수료 매출정보가 없습니다"
    end if
	dbget.close()	:	response.End
end if

if ojungsanTaxCC.FOneItem.IsEvaledTax then
    if (request("autotype")="V2") then
    response.write "<script>"&vbCRLF
    response.write "opener.addResultLog('"&request("jidx")&"','기정산확정');"&vbCRLF
    response.write "opener.fnNextEvalProc();"&vbCRLF
    response.write "</script>"
    else
    response.write "<script>alert('이미 정산 확정된 내역입니다.');</script>"
	response.write "이미 정산 확정된 내역입니다."
    end if
	dbget.close()	:	response.End
end if


'// ============================================================================
dim opartner, ogroup
dim stypename

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid
opartner.FPageSize = 1
opartner.GetOnePartnerNUser


set ogroup = new CPartnerGroup
ogroup.FRectGroupid = ojungsanTaxCC.FOneItem.Fgroupid
ogroup.GetOneGroupInfo

if ogroup.FResultCount<1 then
    response.write "<script>alert('그룹 코드가 지정되지 않았거나, 정산정보가 없습니다.');</script>"
	response.write "그룹 코드가 지정되지 않았거나, 정산정보가 없습니다"
	dbget.close()	:	response.End
end if

dim MaySocialNo : MaySocialNo=FALSE ''주민번호로 발급
if IsMaySocialNo(ogroup.FOneItem.Fcompany_no) then
    MaySocialNo = true
    ogroup.FOneItem.Fcompany_no = ogroup.FOneItem.FdecCompNo
end if

if (NOT MaySocialNo) then
    if LEN(replace(ogroup.FOneItem.Fcompany_no,"-",""))<>10 then
        response.write "<script>alert('사업자 번호가 올바르지 않습니다.');</script>"
    	response.write "사업자 번호가 올바르지 않습니다."& replace(ogroup.FOneItem.Fcompany_no,"-","") & "::" & LEN(replace(ogroup.FOneItem.Fcompany_no,"-",""))
    	dbget.close()	:	response.End
    end if
end if


stypename = "세금계산서"

dim jungsan_hpall, jungsan_hp1,jungsan_hp2,jungsan_hp3
jungsan_hpall = Trim(ogroup.FOneItem.Fjungsan_hp)
jungsan_hpall = split(jungsan_hpall,"-")

if UBound(jungsan_hpall)>=0 then
	jungsan_hp1 = jungsan_hpall(0)
end if

if UBound(jungsan_hpall)>=1 then
	jungsan_hp2 = jungsan_hpall(1)
end if

if UBound(jungsan_hpall)>=2 then
	jungsan_hp3 = jungsan_hpall(2)
end if

if (jungsan_hp2="") and (jungsan_hp3="") and (Len(jungsan_hp1)=11) then
    jungsan_hp3 = MID(jungsan_hp1,8,4)
    jungsan_hp2 = MID(jungsan_hp1,4,4)
    jungsan_hp1 = LEFT(jungsan_hp1,3)
end if

dim Bill365URL : Bill365URL = "http://www.bill36524.com"  '' :8090: test, 80: real
dim swfName    : swfName = "DzEBankFlexAPI" ''"dZAmfApp"
dim swfURL     : swfURL = "/designer/jungsan/"


Dim EVAL_CompanyNo  : EVAL_CompanyNo = "2118700620"

if (replace(ogroup.FOneItem.Fcompany_no,"-","")=EVAL_CompanyNo) then
    response.write "<script>alert('텐바이텐 사업자 발행 불가.');</script>"
	response.write "텐바이텐 사업자 발행 불가."
	''if (session("ssBctID")<>"icommang") then ''TEST
	dbget.close()	:	response.End
    ''end if
end if


%>
<script src="/designer/jungsan/AC_OETags.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript">
	AC_FL_RunContent(
    	"src", "<%= swfURL&swfName %>",
    	"width", "300",
    	"height", "10",
    	"align", "middle",
    	"id", "<%= swfName %>",
    	"quality", "high",
    	"bgcolor", "#869ca7",
    	"name", "<%= swfName %>",
    	"allowScriptAccess","always",
    	"type", "application/x-shockwave-flash",
    	"pluginspage", "http://www.adobe.com/go/getflashplayer"
    );
</script>
<script language='javascript'>
var pLogIdx = 0;
var fxStarted = false;


function getMatchStr(stre,pt){

    var pat = "[<]"+pt+"[>](.*?)[<]\/"+pt+"[>]";

    var re = new RegExp(pat,"g");

    var resultArray = re.exec(stre);

    if (resultArray==null){
        return "";
    }else{
        return (resultArray[1])
    }

}



function thisMovie(movieName){
    if(navigator.appName.indexOf("Microsoft") != -1){
        return window[movieName];
    }else {
        return document[movieName];
    }
}

function AddNew(key, value)
{
 var obj = new Object();
 obj.key = key;
 obj.value = value;
 return obj;
}


//01.로그인
function FxLogin(iid,ipwd){

    if (fxStarted) return;
    fxStarted = true;

    pLogIdx = 0;

    var obj = AddNew("ID", iid);
    var obj1 = AddNew("PASSWD", ipwd);
    var obj2 = AddNew("USER_IP", "<%= request.ServerVariables("REMOTE_ADDR") %>");

    var arr = new Array(obj, obj1, obj2);
    try {
        thisMovie("<%= swfName %>").Login(arr);
    } catch (e) {
        alert('플래시 파일 로드 오류 - 문의 요망(070-7515-5403 서동석)\n\n'+e.message);
	}
    document.all.txtMsg.innerHTML = "bill36524.com 에 로그인중입니다. 잠시 기다려주세요..";
    //alert('startedlogin');
}




//01.로그인 결과
function FxLoginResult(retObj){
    //alert(retObj);
    var result = retObj.RESULT;
    var company_no = "<%= EVAL_CompanyNo %>";


    document.all.txtMsg.innerHTML = "";
    if (result=="00000"){
        //사업자번호 체크
        if (retObj.NO_ID!=company_no){
            hideLogin();
            alert('bill36524 사이트에 가입된 사업자번호와 텐바이텐에 등록된 사업자번호가 일치하지 않습니다.\n\nbill36524에 등록된 사업자번호:' + retObj.NO_ID + '\n텐바이텐에 등록된 사업자번호:'+company_no);
            return;
        }
//alert('TEST 중')
//return;
        preSaveLog();
    }else{
        hideLogin();
        alert(retObj.RESULT_MSG);
    }

}

//발행전 저장
function preSaveLog(){
    var frm = document.frm;
    <% if (jungsan_hp1="") or (jungsan_hp2="") or (jungsan_hp3="") or (Len(jungsan_hp1)>3) or (Len(jungsan_hp2)>4) or (Len(jungsan_hp3)>4) then %>
        <% if (request("autotype")="V2") then %>
        opener.addResultLog('<%=request("jidx")%>','<strong>휴대폰번호<strong>');
        opener.fnNextEvalProc();
        <% else %>
        alert('정산 담당자 핸드폰 번호가 올바르지 않습니다. \n업체정보수정에서 정산담당자 핸드폰을 000-000-0000 대시 형태로 수정후 사용하세요.');
        hideLogin();
        <% end if %>
        return;
    <% end if %>


    frm.action="dotaxregAdm.asp";
	frm.target = "ipreSave";
	frm.submit();

}

//발행후 저장
function saveTaxEvalResult(result,no_tax,result_msg,no_iss){
    var frm = taxSaveFrm;
    frm.action="saveTaxResultAdm.asp";
    frm.idx.value = pLogIdx;
    frm.result.value = result;
    frm.no_tax.value = no_tax;
    frm.no_iss.value = no_iss;
    frm.result_msg.value = result_msg;

	frm.target = "ipreSave";
	frm.submit();

	fxStarted = false;
}


function billTaxEvalFlexApi(pidx){

    pLogIdx = pidx;
    <%
    dim FG_VAT : FG_VAT = ojungsanTaxCC.FOneItem.getBill_FG_VAT
    %>
    var obj1 = AddNew("FG_BILL","<%= ojungsanTaxCC.FOneItem.getBill_FG_BILL %>");   //청구1 영수2
    var obj2 = AddNew("YN_TURN","Y");   //Y정발행 N역발행  :: 역발행시 발행요청 , 정발행시 승인요청
    var obj3 = AddNew("FG_IO","1");     //1매출 2매입
    <% if (MaySocialNo) then %>
    var obj4 = AddNew("FG_PC","2");     // 2016/09/29 핑거스 작품 개인
    <% else %>
    var obj4 = AddNew("FG_PC","1");     //1기업 2개인
    <% end if %>
    var obj5 = AddNew("FG_FINAL","1");  //0저장 1 발송 2승인 3반려 4승인취소요청
    var obj6 = AddNew("YN_CSMT","N"); // 확인
    var obj7 = AddNew("FG_VAT","<%= FG_VAT %>");    // 1과세,2영세,3면세
    var obj8 = AddNew("AM","<%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>");
    var obj9 = AddNew("AM_VAT","<%= ojungsanTaxCC.FOneItem.getJungsanTaxVat %>");
    var obj10 = AddNew("AMT","<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>");

    var obj11 = AddNew("AMT_CASH","0");
    var obj12 = AddNew("AMT_CHECK","0");
    var obj13 = AddNew("AMT_NOTE","0");
    var obj14 = AddNew("YMD_WRITE","<%= Replace(ojungsanTaxCC.FOneItem.GetPreFixSegumil,"-","") %>");

    var obj15 = AddNew("BUY_NO_BIZ","<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>");
    var obj16 = AddNew("BUY_NM_CORP","<%= ogroup.FOneItem.FCompany_name %>");
    var obj17 = AddNew("BUY_NM_CEO","<%= ogroup.FOneItem.Fceoname %>");
    var obj18 = AddNew("BUY_BIZ_STATUS","<%= ogroup.FOneItem.Fcompany_uptae %>");
    var obj19 = AddNew("BUY_BIZ_TYPE","<%= ogroup.FOneItem.Fcompany_upjong %>");

    var obj20 = AddNew("BUY_ADDR1","<%= ogroup.FOneItem.Fcompany_address %>");
    var obj21 = AddNew("BUY_ADDR2","<%= ogroup.FOneItem.Fcompany_address2 %>");
    var obj22 = AddNew("BUY_DAM_DEPT","");
    var obj23 = AddNew("BUY_DAM_NM","<%= ogroup.FOneItem.Fjungsan_name %>");
    var obj24 = AddNew("BUY_DAM_EMAIL","<%= ogroup.FOneItem.Fjungsan_email %>");

    var obj25 = AddNew("BUY_DAM_MOBIL1","<%= jungsan_hp1 %>");
    var obj26 = AddNew("BUY_DAM_MOBIL2","<%= jungsan_hp2 %>");
    var obj27 = AddNew("BUY_DAM_MOBIL3","<%= jungsan_hp3 %>");

    var obj28 = AddNew("BUY_DAM_TEL1","<%= jungsan_hp1 %>");
    var obj29 = AddNew("BUY_DAM_TEL2","<%= jungsan_hp2 %>");
    var obj30 = AddNew("BUY_DAM_TEL3","<%= jungsan_hp3 %>");


    var obj31 = AddNew("SELL_NO_BIZ","2118700620");
    var obj32 = AddNew("SELL_NM_CEO","최은희");
    var obj33 = AddNew("SELL_NM_CORP","(주)텐바이텐");

    var obj34 = AddNew("SELL_DAM_NM","계산서담당자");  //2017/06/01
    var obj35 = AddNew("SELL_DAM_EMAIL","accounts@10x10.co.kr");

    var obj36 = AddNew("SELL_DAM_MOBIL1","02");
    var obj37 = AddNew("SELL_DAM_MOBIL2","554");
    var obj38 = AddNew("SELL_DAM_MOBIL3","2033");

    var obj39 = AddNew("SELL_DAM_TEL1","02");
    var obj40 = AddNew("SELL_DAM_TEL2","554");
    var obj41 = AddNew("SELL_DAM_TEL3","2033");

    var obj42 = AddNew("SELL_ADDR1","서울시 종로구 대학로 57");
    var obj43 = AddNew("SELL_ADDR2","홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐");
    var obj44 = AddNew("SELL_BIZ_STATUS","도소매외");
    var obj45 = AddNew("SELL_BIZ_TYPE","전자상거래외");

    var obj46 = AddNew("SELL_DAM_DEPT","재무회계팀");


    var obj47 = AddNew("AMT_AR","0");   //외상미수금
    //var obj47 = AddNew("AMT_AR","<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>");   //외상미수금
    //var obj48 = AddNew("CD_SVC","<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>");   //CD_SVC ??
    var obj48 = AddNew("NO_SERIAL",pidx);   //일련번호

    var obj49 = AddNew("DC_RMK2", "[10x10 scm 연동발행 : ID" + pidx + "]");

    //201002버전.
    var obj50 = AddNew("YN_FX","N"); // 수정 세금계산서 여부  Y:수정 세금 계산서, N: 정상 발행 <== 필수 입력 입니다
    var obj51 = AddNew("NO_SENDER_PK","<%= ojungsanTaxCC.FOneItem.getBill_NO_SENDER_PK %>");
    
    //2016/04/18 추가
    var obj52 = AddNew("YN_ISS","0");  //FG_VAT 가 3(면세) 일경우 YN_ISS : NULL 일경우 전송제외 YN_ISS : 0 일경우 국세청 전송요청
    
    <% if (TRUE) or (FG_VAT="3") then %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45, obj46, obj47, obj48, obj49, obj50, obj51, obj52);
    <% else %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45,obj46,obj47,obj48, obj49, obj50, obj51);
    <% end if %>

    var objline1 = AddNew("ITEM_STD", "<%= Right(Replace(ojungsanTaxCC.FOneItem.Fyyyymm,"-",""),4) %>");
    var objline2 = AddNew("NM_ITEM", "<%= ojungsanTaxCC.FOneItem.getBill_NM_ITEM %>");
    var objline3 = AddNew("NO_ITEM", "1");
    var objline4 = AddNew("AM", "<%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>");
    var objline5 = AddNew("AM_VAT", "<%= ojungsanTaxCC.FOneItem.getJungsanTaxVat %>");
    var objline6 = AddNew("AMT", "<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>");
    var objline7 = AddNew("DD_WRITE", "<%= Mid(ojungsanTaxCC.FOneItem.GetPreFixSegumil,9,2) %>");
    var objline8 = AddNew("MM_WRITE", "<%= Mid(ojungsanTaxCC.FOneItem.GetPreFixSegumil,6,2) %>");
    //var objline9 = AddNew("QTY", "1");      //수량
    //var objline10 = AddNew("UM", "<%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>");      //단가

    var arrline1 = new Array(objline1, objline2,objline3, objline4, objline5, objline6, objline7, objline8);

    var arrlineArr = new Array(arrline1);

    thisMovie("<%= swfName %>").SendTaxMuch(1);

    thisMovie("<%= swfName %>").SendTaxAccount("", arr, arrlineArr);
    //thisMovie("<%= swfName %>").SendTaxAccount("", arr, arrlineArr, null, "");
    document.all.txtMsg.innerHTML = "계산서 발행중입니다. 잠시 기다려주세요..";
}

function closeMe(){
    opener.location.reload();
    window.close();
}

//02.세금계산서 발행 결과
function FxSendTaxAccountResult(retObj){
    var result = retObj.RESULT;
    var result_msg  = retObj.RESULT_MSG;
    var tb_tax = retObj.OBJ_TBTAX;
    if (tb_tax!=null){
        var no_tax = tb_tax.NO_TAX;
        var no_iss = tb_tax.NO_ISS; //국세청승인번호
    }else{
        var no_tax = "";
    }

    saveTaxEvalResult(result,no_tax,result_msg,no_iss);
    document.all.txtMsg.innerHTML = "";
    hideLogin();


    if (result!="00000"){
        if (result=="10000"){
            if (result_msg=="API 기발행 세금계산서") {
                <% if (request("autotype")="V2") then %>
                opener.addResultLog("<%=jidx%>",result_msg)
                opener.fnNextEvalProc()
                return;
                <% else %>
                alert("오류 : " + result_msg + "");
                <% end if %>
            }else{
                alert("오류 : " + result_msg + "\n\nbill36524.com 로그인 하신후 \n사용자환경설정 => 인증서 등록에서 인증서 등록후 사용하시기 바랍니다.");
            }
        }else{
            alert(result_msg);
        }
        location.reload();  //재로딩 안하면 내부적으로 오류발생(중복발행을 막는듯)
    }


    /*
    else{
        //popTax :: 정상발행

        //FxShowTaxAccount(no_tax,compNo);

        alert("계산서가 발행 되었습니다. \n텐바이텐에서 승인후 (익일)출력가능합니다.");
        opener.location.reload();
        window.close();
    }
    */
}

//popupMove 관련
var bdown = false;
var x, y;
var sElem;

function mdown(evt)
{
	evt = (evt) ? evt : ((window.event) ? window.event : "");
	sElem = evt.target ? evt.target : evt.srcElement;
	if (evt.stopPropagation)
	{
		evt.stopPropagation();
		evt.preventDefault();
	}
	evt.returnValue  = false;
	evt.cancelBubble = true;

	if(sElem.className == "drag")
	{
		bdown = true;
		x = evt.clientX;
		y = evt.clientY;
	}
}

function mup()
{
	bdown = false;
}

document.onmousemove = function moveimg(event)
{
	event = (event) ? event : ((window.event) ? window.event : "");
	if(bdown)
	{
		var distX = event.clientX - x;
		var distY = event.clientY - y;
		var targetImg = document.getElementById('POPBillLogin');
		targetImg.style.left = (parseInt(targetImg.style.left) + distX) + 'px';
		targetImg.style.top = (parseInt(targetImg.style.top) + distY) + 'px';
		x = event.clientX;
		y = event.clientY;
		return false;
	}
}


function hideLogin(){
    document.all["POPBillLogin"].style.visibility='hidden';
    document.frm.evalButton.disabled=false;
}

function showLogin(){

    var frm = document.billfrm;
    frm.billid.value = '';
    frm.billpass.value = '';
    hideDoing();
    document.all["POPBillLogin"].style.visibility='visible';

    document.frm.evalButton.disabled=true;
    fxStarted = false;
}

function showDoing(){
    var frm = document.billfrm;
    document.all.ievalBtn.style.display='none';
    document.all.idoingMsg.style.display='inline';
    document.all.popcloseId.style.display='none';
    frm.billid.disabled = true;
    frm.billpass.disabled = true;
}

function hideDoing(){
    var frm = document.billfrm;
    document.all.ievalBtn.style.display='inline';
    document.all.idoingMsg.style.display='none';
    document.all.popcloseId.style.display='inline';
    frm.billid.disabled = false;
    frm.billpass.disabled = false;
}

function billTaxEval(frm){

    if (frm.billid.value.length<1){
        alert('Bill36524 아이디를 입력하세요.');
        frm.billid.focus();
        return;
    }

    if (frm.billpass.value.length<1){
        alert('Bill36524 패스워드를 입력하세요.');
        frm.billpass.focus();
        return;
    }

    showDoing();
    FxLogin(frm.billid.value,frm.billpass.value);
    // 발행이 완료된다음 숨김.. hideLogin();
}

//05. 세금계산서 확인

function FxShowTaxAccount(no_tax, no_biz_no){
    var url = "<%= Bill365URL %>/popupBillTax.jsp?";
    url += "NO_TAX=" + no_tax;
    url += "&NO_BIZ_NO=" + no_biz_no;


    var popwin = window.open(url, "taxwin", "height=700,width=660, menubar=no, location=no, resizeable=no, status=no, scrollbars=no, top=200, left=300");
    popwin.focus();
}

// FLESH 내부에서 기타 예외 발생시 오류 리턴
/*
//한번 발행실패후 계속해서 오류남 : 계산서 발행중이라는 메세지 201002버전..
function FxErrorResult(retObj) {

    alert("ERR:" + retObj + "\n관리자 문의 요망.");

    if (pLogIdx!=0){
        var frm = taxSaveFrm;
        frm.action="saveTaxResult.asp";
        frm.idx.value = pLogIdx;
        frm.result.value = "999";
        frm.no_tax.value = "";
        frm.result_msg.value = retObj;

    	frm.target = "ipreSave";
    	frm.submit();

    	fxStarted = false;
	}
	hideLogin();
}
*/

//세금계산서 발행시 에러 처리:네트웍오류 및 처리하지 못한 예외 발생
function DzErrorEvent(faultEvent){
    var errinfo = "";

    errinfo = "faultEvent.message:" + faultEvent.message + "\n";
    errinfo += "faultEvent.errorID:" + faultEvent.errorID + "\n";
    errinfo += "faultEvent.faultCode:" + faultEvent.faultCode + "\n";
    errinfo += "faultEvent.faultDetail:" + faultEvent.faultDetail + "\n";
    errinfo += "faultEvent.faultString:" + faultEvent.faultString + "\n";

    //form1.fxlog.value = errinfo;
    alert("ERR:" + errinfo + "\n관리자 문의 요망");

    hideLogin();
}


function ActTaxReg(frm){
//alert('죄송합니다. \n bill36524.com 사이트 접속이 원활하지 않아 잠시 계산서 발행을 중지합니다.');
//return;

    <% if (MaySocialNo) then %>
        if (frm.biz_no.value.length!=13){
    		alert('사업자 등록 번호가 올바르지 않거나 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
    		return;
    	}
    <% else %>
    	if (frm.biz_no.value.length!=10){
    		alert('사업자 등록 번호가 올바르지 않거나 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
    		return;
    	}
    <% end if %>
    
	if (frm.corp_nm.value.length<1){
		alert('사업자 명이 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.ceo_nm.value.length<1){
		alert('대표자 명이 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.biz_status.value.length<1){
		alert('업태가 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.biz_type.value.length<1){
		alert('업종이 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.addr.value.length<1){
		alert('사업장 주소가 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.dam_nm.value.length<1){
		alert('담당자 성명이 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.email.value.length<1){
		alert('담당자 이메일이 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

	if (frm.write_date.value.length<1){
		alert('계산서 발행일 입력 후 사용하세요.');
		return;
	}

    if (!thisMovie("<%= swfName %>")){
        alert('swf 파일이 로딩 되지 않았습니다.');
        return;
    }

    if (frm.billSite[1].checked){
        if (confirm('발행 하시겠습니까?')){
            FxLogin('tenbyten','cube1010!!');  
        }
        return;
    }

/*
    if (frm.billSite[1].checked){
        if (confirm('팝업창에서 bill36524.com 아이디와 패스워드를 입력하신후 발행하시면 됩니다. 계속 하시겠습니까?')){
            showLogin();

        }
        return;
    }

    if (confirm('<%= stypename %> 를 발행 하시겠습니까?')){
	    frm.action="dotaxreg.asp";
	    frm.target = "";
		frm.submit();
	}
*/
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="25" valign="top">
        <td>
        	<img src="/images/icon_star.gif" width="16" height="16" align="absbottom">
        	<strong>전자 <%= stypename %> 발행</strong>
        </td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="dotaxreg.asp">
	<input type=hidden name=jungsanid value="<%= ojungsanTaxCC.FOneItem.FId %>">
	<input type=hidden name=jungsanname value="<%= ojungsanTaxCC.FOneItem.Ftitle %>">
	<input type=hidden name=jungsangubun value="<%= ojungsanTaxCC.FOneItem.FtargetGbn %>">
	<input type=hidden name=makerid value="<%= makerid %>">
	<input type=hidden name=jgubun value="<%= ojungsanTaxCC.FOneItem.Fjgubun %>">

	<input type=hidden name=biz_no value="<%= replace(replace(socialnoReplace(ogroup.FOneItem.Fcompany_no),"-","")," ","") %>" >
	<input type=hidden name=corp_nm value="<%= ogroup.FOneItem.FCompany_name %>">
	<input type=hidden name=ceo_nm value="<%= ogroup.FOneItem.Fceoname %>">
	<input type=hidden name=biz_status value="<%= ogroup.FOneItem.Fcompany_uptae %>">
	<input type=hidden name=biz_type value="<%= ogroup.FOneItem.Fcompany_upjong %>">


	<input type=hidden name=addr value="<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %>">
	<input type=hidden name=dam_nm value="<%= ogroup.FOneItem.Fjungsan_name %>">
	<input type=hidden name=email value="<%= ogroup.FOneItem.Fjungsan_email %>">
	<input type=hidden name=hp_no1 value="<%= jungsan_hp1 %>">
	<input type=hidden name=hp_no2 value="<%= jungsan_hp2 %>">
	<input type=hidden name=hp_no3 value="<%= jungsan_hp3 %>">

	<input type=hidden name=sb_type value="01"> <!-- 매출 01 매입 02 -->
	<input type=hidden name=tax_type value="<%= ojungsanTaxCC.FOneItem.Ftaxtype %>">
	<input type=hidden name=bill_type value="01"> <!-- 영수 01 청구 18 -->
	<input type=hidden name=pc_gbn value="C"> <!-- 개인 P 기업 C -->

	<input type=hidden name=item_count value="1">
	<input type=hidden name=item_nm value="<%= ojungsanTaxCC.FOneItem.getBill_NM_ITEM %>">
	<input type=hidden name=item_qty value="1">
	<input type=hidden name=item_price value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>">
	<input type=hidden name=item_amt value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxSuply %>">
	<input type=hidden name=item_vat value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxVat %>">
	<input type=hidden name=item_remark value="">

	<input type=hidden name=credit_amt value="<%= ojungsanTaxCC.FOneItem.getJungsanTaxSum %>">

	<input type=hidden name=cur_u_user_no value="261744"> <!-- DEV 1000394, REAL 244730, ON 261744 -->
	<input type=hidden name=cur_dam_nm value="김민환">
	<input type=hidden name=cur_email value="accounts@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="02">
	<input type=hidden name=cur_hp_no2 value="554">
	<input type=hidden name=cur_hp_no3 value="2033">

    <input type=hidden name=autotype value="<%=request("autotype")%>">
    

    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td height="20" colspan="2">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>전자 <%= stypename %> 발행방법</strong>
   		</td>
 	</tr>

    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td colspan="2" height="20" valign="middle">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>등록된 사업자정보 확인</strong>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF" width="30%">사업자명</td>
		<td><%= ogroup.FOneItem.FCompany_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">대표자명</td>
		<td><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">사업자번호</td>
		<td><%= socialnoReplace(ogroup.FOneItem.Fcompany_no) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">과세구분</td>
		<td><%= ogroup.FOneItem.Fjungsan_gubun %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">사업장소재지</td>
		<td><%= ogroup.FOneItem.Fcompany_address %>&nbsp;<%= ogroup.FOneItem.Fcompany_address2 %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">업태</td>
		<td><%= ogroup.FOneItem.Fcompany_uptae %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">업종</td>
		<td><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">계산서발행일</td>
		<% if False and (ojungsanTaxCC.FOneItem.Fdifferencekey>0) then %>
		<td><input type=text name=write_date value="" size="10" maxlength=10 readonly ><a href="javascript:calendarOpen(frm.write_date);"><img src="/images/calicon.gif" border=0 align=absmiddle></a></td>
		<% else %>
		<td><input type=text name=write_date value="<%= ojungsanTaxCC.FOneItem.GetPreFixSegumil %>" size="10" maxlength=10 readonly style="border:0"></td>
		<% end if %>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">발행금액</td>
		<td><b><%= FormatNumber(ojungsanTaxCC.FOneItem.getJungsanTaxSum,0) %></b> (공급가 : <%= FormatNumber(ojungsanTaxCC.FOneItem.getJungsanTaxSuply,0) %> 부가세: <%= FormatNumber(ojungsanTaxCC.FOneItem.getJungsanTaxVat,0) %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">품목명</td>
		<td><%= ojungsanTaxCC.FOneItem.getBill_NM_ITEM %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">NO_SENDER_PK</td>
		<td><%= ojungsanTaxCC.FOneItem.getBill_NO_SENDER_PK %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;<b>* 매월 10일 까지 발행시 : 정상발행</b><br>
			&nbsp;&nbsp;<b>* 매월 11일 이후 발행시 : 이월발행(입금처리도 이월(15일)됩니다.)</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">정산담당자명</td>
		<td><%= ogroup.FOneItem.Fjungsan_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">정산담당자E-mail</td>
		<td><%= ogroup.FOneItem.Fjungsan_email %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">정산담당자 핸드폰번호</td>
		<td><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;* 업체정보를 확인하시고, 미입력된 정보는 어드민 업체정보수정에서 수정후 진행하시기 바랍니다.<br>
			&nbsp;&nbsp;* 정산담당자의 정보를 입력하시면, 세금계산서의 발행상황을 E-mail과 문자서비스로 알려드립니다.
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" align="center">
	    <% if (FALSE) then %>
	    <input type="radio" name="billSite" value="N" checked ><strong>네오포트 </strong>
	    <input type="radio" name="billSite" value="B" ><font color=red><strong>bill36524.com (2010년부터)</strong></font>
	    <% else %>
	    <input type="radio" name="billSite" value="N" disabled ><font color=gray><strong>네오포트 (사용불가<!--네오포트기존회원-->)</strong></font>
	    <input type="radio" name="billSite" value="B" checked><font color=red><strong>bill36524.com (2010년부터)</strong></font>
    <% end if %>
	    </td>
	</tr>

</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">

    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<input type=button name="evalButton" value="전자 <%= stypename %> 발행" onClick="ActTaxReg(frm)">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- 표 하단바 끝-->
<div id='POPBillLogin' style='position:absolute; left:100px; top:240px; width:140; height:100; z-index:2; visibility: hidden'>
<table width="420" height="260" border="0" cellpadding="0" cellspacing="2" bgcolor="#000000" class="a">
  <form name="billfrm">
  <tr >
    <td height="20" onMouseDown="mdown(event);" onMouseUp="mup();"  class="drag" bgcolor="#333399">
    &nbsp;<font color="#ffffff"><strong>bill36524 계산서발행</strong></font>
    </td>
  </tr>
  <tr>
    <td height="210" colspan="2" valign="top" bgcolor="#FFFFFF" align="center">
        <table border=0 width="100%" class="a">
        <tr>
            <td>
            <table border=0 width="90%" class="a">
                <tr>
                    <td>1. http://www.bill36524.com 에 회원가입을하세요.</td>
                </tr>
                <tr>
                    <td>2. 인증서등록 및 사용포인트를 충전하세요.</td>
                </tr>
                <tr>
                    <td>3. 아래 페이지에서 http://www.bill36524.com 의 아이디와 패스워드를 입력하신후 계산서발행 버튼을 클릭하세요.</td>
                </tr>
                <tr>
                    <td>4. 계산서 발행시 시간이 오래 소요될 수 있으니 기다려주시기 바랍니다.(최장 1분)</td>
                </tr>
            </table>
            </td>
        </tr>
        <tr height="120">
            <td align="center">
                <table border="0" cellspacing="2" cellpadding="2" width="330" height="100" class="a" bgcolor="#CCCCCC" >
                    <tr bgcolor="#FFFFFF">
                        <td width="130">bill36524 아이디</td>
                        <td><input type="text" name="billid" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) billfrm.billpass.focus();"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                        <td width="130">bill36524 패스워드</td>
                        <td><input type="password" name="billpass" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) billTaxEval(billfrm);"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF" id="ievalBtn">
                        <td colspan="2" align="center"><input type="button" value="계산서발행" onclick="billTaxEval(billfrm);"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF" id="idoingMsg" style="display:none">
                        <td colspan="2" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/loading.gif" width="269" height="14">
                        <br><div id="txtMsg" name="txtMsg"><!-- 처리중입니다.잠시만 기다려주세요... --></div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        </table>
    </td>
  </tr>
  <tr id="popcloseId" ><td bgcolor="#FFFFFF" align="right"><a href="javascript:hideLogin();">close</a></td></tr>
  </form>
</table>
</div>

<form name="taxSaveFrm" method="post">
<input type="hidden" name="idx" value="">
<input type="hidden" name="result" value="">
<input type="hidden" name="no_tax" value="">
<input type="hidden" name="no_iss" value="">
<input type="hidden" name="billsiteCode" value="B"> <!-- 더존B, 웹캐시W -->
<input type="hidden" name="result_msg" value="">
<input type="hidden" name="jungsangubun" value="<%= ojungsanTaxCC.FOneItem.FtargetGbn %>">
<input type="hidden" name="write_date" value="<%= ojungsanTaxCC.FOneItem.GetPreFixSegumil %>">
<input type="hidden" name="jungsanid" value="<%= ojungsanTaxCC.FOneItem.FId %>">
<input type="hidden" name="isauto" value="<%= isauto %>">
</form>
<iframe name="ipreSave" id="ipreSave" width="500" height="50"></iframe>
<%
set ojungsanTaxCC = Nothing
set opartner = Nothing
set ogroup = Nothing
%>

<script language=javascript>
function reActEval(){
    <% if (nextjidx<>"") then %>
        <% if (jidx<>nextjidx) then %>
        opener.evalOneTax(<%=nextjidx%>)
        <% end if %>
    <% elseif (request("autotype")="V2") then %>
        opener.addResultLog("<%=jidx%>","v")
        opener.fnNextEvalProc()
    <% end if %>
}

function getOnload(){
    <% if (isauto<>"") then %>
    setTimeout("FxLogin('tenbyten','cube1010!!')",2000);
    <% end if %>
}
window.onload = getOnload;
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

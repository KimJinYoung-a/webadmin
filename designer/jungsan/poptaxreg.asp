<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim id
id = requestCheckvar(request("id"),10)

dim makerid, groupid
makerid = session("ssBctId")
groupid = getPartnerId2GroupID(makerid)

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
''ojungsan.FRectdesigner = session("ssBctID")
ojungsan.FRectGroupID = groupid
if (groupid<>"") then
ojungsan.JungsanMasterList
end if

if ojungsan.FResultCount<1 then
	response.write "<script>alert('정산정보가 없습니다.');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

''수수료 발행 불가
if ojungsan.FItemList(0).IsCommissionTax then
    response.write "<script>alert('수수료 계산서 발행 불가');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

makerid = ojungsan.FItemList(0).Fdesignerid

dim opartner, ogroup
dim stypename

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid ''session("ssBctId")
opartner.FPageSize = 1
opartner.GetOnePartnerNUser



set ogroup = new CPartnerGroup
ogroup.FRectGroupid = ojungsan.FItemList(0).Fgroupid
ogroup.GetOneGroupInfo

if ogroup.FResultCount<1 then
    response.write "<script>alert('그룹 코드가 지정되지 않았거나, 정산정보가 없습니다.');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

if (ojungsan.FItemList(0).IsElecTaxCase) then
	stypename = "세금계산서"
elseif (ojungsan.FItemList(0).IsElecFreeTaxCase) then
	stypename = "계산서"
else
	response.write "<script>alert('세금계산서 혹은 계산서만 발행 가능합니다. - 이미 발행 하였거나 발행할 정보가 없습니다.');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

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



dim Bill365URL : Bill365URL = "http://www.bill36524.com"  '' :8090: test, 80: real
dim swfName    : swfName = "DzEBankFlexAPI" ''"dZAmfApp"

IF application("Svr_Info")="Dev" THEN
    ''아놔..
    ''Bill365URL = "http://www.bill36524.com:8090"
    ''swfName = "DzEBankFlexAPI_test"
end if

%>



<script language="VBScript">
<!--
Sub <%= swfName %>_FSCommand(ByVal command, ByVal args)
    call <%= swfName %>_DoFSCommand(command, args)
end sub
//-->
</script>

<script src="AC_OETags.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript">
	AC_FL_RunContent(
    	"src", "<%= swfName %>",
    	"width", "0",
    	"height", "0",
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

//결과. :: 다른버전. dZAmfApp.swf
function <%= swfName %>_DoFSCommand(cmd, args) {
    var result, result_msg, no_tax, no_iss ;

    var compNo = '<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>';

    //alert(args);
    switch (cmd) {
        case "Login" :
            result = getMatchStr(args,"RESULT");
            result_msg = getMatchStr(args,"RESULT_MSG");
            if (result!="00000"){
                alert(result_msg);

            }else{
                var no_id = getMatchStr(args,"NO_ID");
                //사업자번호 체크
                if (no_id!=compNo){
                    alert('bill36524 사이트에 가입된 사업자번호와 텐바이텐에 등록된 사업자번호가 일치하지 않습니다.\n\nbill36524에 등록된 사업자번호:' + no_id + '\n텐바이텐에 등록된 사업자번호:' + compNo);
                    return;
                }

                preSaveLog();
            }
            break;
        case "SendTaxAccount" :
            result  = getMatchStr(args,"RESULT");
            result_msg  = getMatchStr(args,"RESULT_MSG");
            no_tax = getMatchStr(args,"NO_TAX");
            no_iss = getMatchStr(args,"NO_ISS"); //국세청승인번호

            saveTaxEvalResult(result,no_tax,result_msg,no_iss);
            if (result!="00000"){
                if (result=="10000"){
                    alert(result + result_msg + "\n\nbill36524.com에서 \n사용자환경설정 => 인증서 등록에서 인증서 등록후 사용하시기 바랍니다.");
                }else{
                    alert(result + result_msg);
                }
            }else{
                //popTax
                FxShowTaxAccount(no_tax,compNo);
            }
            break;
        default :
            alert(cmd);
            alert(args);
            break;
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
        alert('플래시 파일 로드 오류 - 문의 요망(070-7515-5403 서동석)');
	}
    document.all.txtMsg.innerHTML = "bill36524.com 에 로그인중입니다. 잠시 기다려주세요..";
    //alert('startedlogin');
}




//01.로그인 결과
function FxLoginResult(retObj){
    //alert(retObj);
    var result = retObj.RESULT;
    var company_no = "<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>";

    document.all.txtMsg.innerHTML = "";
    if (result=="00000"){
        //사업자번호 체크
        if (retObj.NO_ID!=company_no){
            hideLogin();
            alert('bill36524 사이트에 가입된 사업자번호와 텐바이텐에 등록된 사업자번호가 일치하지 않습니다.\n\nbill36524에 등록된 사업자번호:' + retObj.NO_ID + '\n텐바이텐에 등록된 사업자번호:'+company_no);
            return;
        }

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

        alert('정산 담당자 핸드폰 번호가 올바르지 않습니다. \n업체정보수정에서 정산담당자 핸드폰을 000-000-0000 대시 형태로 수정후 사용하세요.');
        hideLogin();
        return;

    <% end if %>


    frm.action="dotaxreg.asp";
	frm.target = "ipreSave";
	frm.submit();

}

//발행후 저장
function saveTaxEvalResult(result,no_tax,result_msg,no_iss){
    var frm = taxSaveFrm;
    frm.action="saveTaxResult.asp";
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
    dim FG_VAT
    if (ojungsan.FItemList(0).Ftaxtype="03") then
        FG_VAT = "2"
    elseif (ojungsan.FItemList(0).Ftaxtype="02") then
        FG_VAT = "3"
    else
        FG_VAT = "1"
    end if
    %>
    var obj1 = AddNew("FG_BILL","1");   //청구1 영수2
    var obj2 = AddNew("YN_TURN","Y");   //Y정발행 N역발행  :: 역발행시 발행요청 , 정발행시 승인요청
    var obj3 = AddNew("FG_IO","1");     //1매출 2매입
    var obj4 = AddNew("FG_PC","1");     //1기업 2개인
    var obj5 = AddNew("FG_FINAL","1");  //0저장 1 발송 2승인 3반려 4승인취소요청
    var obj6 = AddNew("YN_CSMT","N");
    var obj7 = AddNew("FG_VAT","<%= FG_VAT %>");    // 1과세,2영세,3면세
    var obj8 = AddNew("AM","<%= ojungsan.FItemList(0).GetTotalTaxSuply %>");
    var obj9 = AddNew("AM_VAT","<%= ojungsan.FItemList(0).GetTotalTaxvat %>");
    var obj10 = AddNew("AMT","<%= ojungsan.FItemList(0).GetTotalSuplycash %>");

    var obj11 = AddNew("AMT_CASH","0");
    var obj12 = AddNew("AMT_CHECK","0");
    var obj13 = AddNew("AMT_NOTE","0");
    var obj14 = AddNew("YMD_WRITE","<%= Replace(ojungsan.FItemList(0).GetPreFixSegumil,"-","") %>");

    var obj15 = AddNew("SELL_NO_BIZ","<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>");
    var obj16 = AddNew("SELL_NM_CORP","<%= ogroup.FOneItem.FCompany_name %>");
    var obj17 = AddNew("SELL_NM_CEO","<%= ogroup.FOneItem.Fceoname %>");
    var obj18 = AddNew("SELL_BIZ_STATUS","<%= ogroup.FOneItem.Fcompany_uptae %>");
    var obj19 = AddNew("SELL_BIZ_TYPE","<%= ogroup.FOneItem.Fcompany_upjong %>");

    var obj20 = AddNew("SELL_ADDR1","<%= ogroup.FOneItem.Fcompany_address %>");
    var obj21 = AddNew("SELL_ADDR2","<%= ogroup.FOneItem.Fcompany_address2 %>");
    var obj22 = AddNew("SELL_DAM_DEPT","");
    var obj23 = AddNew("SELL_DAM_NM","<%= ogroup.FOneItem.Fjungsan_name %>");
    var obj24 = AddNew("SELL_DAM_EMAIL","<%= ogroup.FOneItem.Fjungsan_email %>");

    var obj25 = AddNew("SELL_DAM_MOBIL1","<%= jungsan_hp1 %>");
    var obj26 = AddNew("SELL_DAM_MOBIL2","<%= jungsan_hp2 %>");
    var obj27 = AddNew("SELL_DAM_MOBIL3","<%= jungsan_hp3 %>");

    var obj28 = AddNew("SELL_DAM_TEL1","<%= jungsan_hp1 %>");
    var obj29 = AddNew("SELL_DAM_TEL2","<%= jungsan_hp2 %>");
    var obj30 = AddNew("SELL_DAM_TEL3","<%= jungsan_hp3 %>");


    var obj31 = AddNew("BUY_NO_BIZ","2118700620");
    var obj32 = AddNew("BUY_NM_CEO","최은희");
    var obj33 = AddNew("BUY_NM_CORP","(주)텐바이텐");

    var obj34 = AddNew("BUY_DAM_NM","이은주");
    var obj35 = AddNew("BUY_DAM_EMAIL","accounts@10x10.co.kr");

    var obj36 = AddNew("BUY_DAM_MOBIL1","02");
    var obj37 = AddNew("BUY_DAM_MOBIL2","554");
    var obj38 = AddNew("BUY_DAM_MOBIL3","2033");

    var obj39 = AddNew("BUY_DAM_TEL1","02");
    var obj40 = AddNew("BUY_DAM_TEL2","554");
    var obj41 = AddNew("BUY_DAM_TEL3","2033");

    var obj42 = AddNew("BUY_ADDR1","서울시 종로구 동숭동");
    var obj43 = AddNew("BUY_ADDR2","1-45 자유빌딩2층");
    var obj44 = AddNew("BUY_BIZ_STATUS","도소매외");
    var obj45 = AddNew("BUY_BIZ_TYPE","전자상거래외");

    var obj46 = AddNew("BUY_DAM_DEPT","온라인");

    var obj47 = AddNew("AMT_AR","<%= ojungsan.FItemList(0).GetTotalSuplycash %>");   //외상미수금
    //var obj48 = AddNew("CD_SVC","<%= ojungsan.FItemList(0).GetTotalSuplycash %>");   //CD_SVC ??
    var obj48 = AddNew("NO_SERIAL",pidx);   //일련번호

    var obj49 = AddNew("DC_RMK2", "[10x10 scm 연동발행 : ID" + pidx + "]");

    //201002버전.
    var obj50 = AddNew("YN_FX","N"); // 수정 세금계산서 여부  Y:수정 세금 계산서, N: 정상 발행 <== 필수 입력 입니다
    var obj51 = AddNew("NO_SENDER_PK","DZ_TEN_ON_<%= ojungsan.FRectId %>_<%= ojungsan.FItemList(0).Fdifferencekey %>_<%= ojungsan.FItemList(0).GetTotalSuplycash %>");
    
    //2016/04/18 추가
    var obj52 = AddNew("YN_ISS","0");  //FG_VAT 가 3(면세) 일경우 YN_ISS : NULL 일경우 전송제외 YN_ISS : 0 일경우 국세청 전송요청
    
    <% if (TRUE) or (FG_VAT="3") then %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45, obj46, obj47, obj48, obj49, obj50, obj51, obj52);
    <% else %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45, obj46, obj47, obj48, obj49, obj50, obj51);
    <% end if %>

    var objline1 = AddNew("ITEM_STD", "<%= Right(Replace(ojungsan.FItemList(0).Fyyyymm,"-",""),4) %>");
    var objline2 = AddNew("NM_ITEM", "<%= ojungsan.FItemList(0).getBillItemName %>");
    var objline3 = AddNew("NO_ITEM", "1");
    var objline4 = AddNew("AM", "<%= ojungsan.FItemList(0).GetTotalTaxSuply %>");
    var objline5 = AddNew("AM_VAT", "<%= ojungsan.FItemList(0).GetTotalTaxvat %>");
    var objline6 = AddNew("AMT", "<%= ojungsan.FItemList(0).GetTotalSuplycash %>");
    var objline7 = AddNew("DD_WRITE", "<%= Mid(ojungsan.FItemList(0).GetPreFixSegumil,9,2) %>");
    var objline8 = AddNew("MM_WRITE", "<%= Mid(ojungsan.FItemList(0).GetPreFixSegumil,6,2) %>");
    //var objline9 = AddNew("QTY", "1");      //수량
    //var objline10 = AddNew("UM", "<%= ojungsan.FItemList(0).GetTotalTaxSuply %>");      //단가

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
                alert("오류 : " + result_msg + "");
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
	if (frm.biz_no.value.length!=10){
		alert('사업자 등록 번호가 올바르지 않거나 등록되어 있지 않습니다. - 업체정보 수정후 사용하세요.');
		return;
	}

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
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" width="16" height="16" align="absbottom">
        	<strong>전자 <%= stypename %> 발행</strong>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="dotaxreg.asp">
	<input type=hidden name=jungsanid value="<%= ojungsan.FItemList(0).FId %>">
	<input type=hidden name=jungsanname value="<%= ojungsan.FItemList(0).Ftitle %>">
	<input type=hidden name=jungsangubun value="ON">
	<input type=hidden name=makerid value="<%= makerid %>">

	<input type=hidden name=biz_no value="<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>" >
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

	<input type=hidden name=sb_type value="02"> <!-- 매출 01 매입 02 -->
	<input type=hidden name=tax_type value="<%= ojungsan.FItemList(0).Ftaxtype %>">
	<input type=hidden name=bill_type value="18"> <!-- 영수 01 청구 18 -->
	<input type=hidden name=pc_gbn value="C"> <!-- 개인 P 기업 C -->

	<input type=hidden name=item_count value="1">
	<input type=hidden name=item_nm value="<%= ojungsan.FItemList(0).getBillItemName %>">
	<input type=hidden name=item_qty value="1">
	<input type=hidden name=item_price value="<%= ojungsan.FItemList(0).GetTotalSuplycash %>">
	<input type=hidden name=item_amt value="<%= ojungsan.FItemList(0).GetTotalTaxSuply %>">
	<input type=hidden name=item_vat value="<%= ojungsan.FItemList(0).GetTotalTaxvat %>">
	<input type=hidden name=item_remark value="">

	<input type=hidden name=credit_amt value="<%= ojungsan.FItemList(0).GetTotalTaxSuply + ojungsan.FItemList(0).GetTotalTaxvat %>">

	<input type=hidden name=cur_u_user_no value="261744"> <!-- DEV 1000394, REAL 244730, ON 261744 -->
	<input type=hidden name=cur_dam_nm value="최보연">
	<input type=hidden name=cur_email value="accounts@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="000">
	<input type=hidden name=cur_hp_no2 value="000">
	<input type=hidden name=cur_hp_no3 value="0000">


    <!--
    <tr align="center" bgcolor="#FFFFFF">
		<td colspan="2">
		* 2005년 3월분 정산분(발행일 3월 31일)부터는 전자 <%= stypename %> 발행을 사용하셔야 합니다.
		</td>
	</tr>
	-->
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td height="20" colspan="2">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>전자 <%= stypename %> 발행방법</strong>
	   		&nbsp;&nbsp;&nbsp;&nbsp;
	   		<a href="http://www.bill36524.com" target="_blank"><font color="blue">>>bill36524 회원가입하기</font></a>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			<img src="/images/icon_num01.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>기존에 네오포트에서 발행을 하셨고, bill36524로 이관을 하신경우</b></font><br>
				&nbsp;&nbsp;1.세금계산서 발행창에서 [bill36524]를 선택하시고 발행하시면 됩니다.<br>
				&nbsp;&nbsp;2.bill36524에 건수충전이 되어 있어야 하며, 인증서가 있어야 합니다.<br>
			<img src="/images/icon_num02.gif" width="16" height="16" align="absbottom">
			<!--
			<font color="red"><b>네오포트를 이용중이고, 아직 bill36524로 이관을 안하신 경우</b></font><br>
				&nbsp;&nbsp;1.세금계산서 발행창에서 [네오포트]를 선택하시고 발행하시면 됩니다.<br>
				&nbsp;&nbsp;2.이번달에 한해서 한시적으로 텐바이텐에서 발행수수료를 지불합니다.<br>
			<img src="/images/icon_num03.gif" width="16" height="16" align="absbottom">
			-->
			<font color="red"><b>신규입점업체의 경우</b></font><br>
				<!-- &nbsp;&nbsp;1.현재 네오포트(www.neoport.net)에 사업자회원으로 회원가입이 불가능합니다.<br>
				&nbsp;&nbsp;2.(네오포트 서비스는 12월부터 서비스가 중단되었습니다.)<br>
				-->
				&nbsp;&nbsp;1.bill36524에 회원 가입 후에 위 1번과 같이 발행창에서 [bill36524]를 선택하시고 발행하시면 됩니다.

			<br>
			<img src="/images/icon_num02.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>발행오류 대처방법</b></font><br>
			[인증서 데이터 값이 없습니다.] 또는 [국세청서명데이터 저장오류] => bill36524로그인 후 왼쪽 세로메뉴에 [사용자환경설정] 클릭 인증서등록 탭에 인증서 재등록후 사용<br>
			[포인트가 부족합니다.] =>  bill36524로그인 후 포인트 충전 후 사용<br>
			&nbsp;&nbsp;<a href="/designer/jungsan/popTaxHelp.asp" target="_taxHelp"><font color="#0000FF"><strong>[세금계산서 발행방법 안내 자세히 보기 ☞]</strong></font></a>

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
		<td><%= ogroup.FOneItem.Fcompany_no %></td>
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
		<% if False and (ojungsan.FItemList(0).Fdifferencekey>0) then %>
		<td><input type=text name=write_date value="" size="10" maxlength=10 readonly ><a href="javascript:calendarOpen(frm.write_date);"><img src="/images/calicon.gif" border=0 align=absmiddle></a></td>
		<% else %>
		<td><input type=text name=write_date value="<%= ojungsan.FItemList(0).GetPreFixSegumil %>" size="10" maxlength=10 readonly style="border:0"></td>
		<% end if %>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">발행금액</td>
		<td><b><%= FormatNumber(ojungsan.FItemList(0).GetTotalSuplycash,0) %></b> (공급가 : <%= FormatNumber(ojungsan.FItemList(0).GetTotalTaxSuply,0) %> 부가세: <%= FormatNumber(ojungsan.FItemList(0).GetTotalTaxvat,0) %>)</td>
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
	    <% if (ojungsan.FItemList(0).Favailneo=1) then %>
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
<input type="hidden" name="jungsangubun" value="ON">
<input type="hidden" name="write_date" value="<%= ojungsan.FItemList(0).GetPreFixSegumil %>">
<input type="hidden" name="jungsanid" value="<%= ojungsan.FItemList(0).FId %>">
</form>
<iframe name="ipreSave" id="ipreSave" width="0" height="0"></iframe>
<%
set ojungsan = Nothing
set opartner = Nothing
set ogroup = Nothing
%>

<script language=javascript>
function SvcErrMsg(){
    //alert('잠시 서비스 점검중입니다.');
    var alertStr;
    //alertStr = "이번달부터 전자세금계산서 연동발행시 변경사항이 있습니다.";
    //alertStr += "\n\nbill36524에 로그인 하신 후, 왼쪽 세로메뉴에 [사용자환경설정]에서";
    //alertStr += "\n4번째 항목에 있는 인증서 등록을 해주셔야 정상적으로 세금계산서가 연동발행됩니다.";
    //alertStr += "\n인증서 등록이 안되어 있을 경우, 텐바이텐SCM과의 연동발행이 되지 않습니다.\n\n";

    //alertStr = "!!! 현재 bill36524.com 접속이 원활하지 않습니다.";
    //alertStr += "\n발행시 시간이 오래걸릴경우 잠시 후 시도하시기 바랍니다."
    //alertStr = "";

    if (alertStr!="") {
        //alert(alertStr);
    }

}
window.onload = SvcErrMsg;
</script>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

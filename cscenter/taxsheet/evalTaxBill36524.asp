<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

dim taxIdx  : taxIdx  = requestCheckVar(request("taxIdx"),32)
''dim taxType : taxType = requestCheckVar(request("taxType"),16)

function Get3PLUpcheInfoByTPLCompanyid(tplcompanyid, byRef tplcompanyname, byRef tplgroupid, byRef tplbillUserID, byRef tplbillUserPass)
	dim sqlStr

	sqlStr = " select top 1 t.tplcompanyid, t.tplcompanyname, t.groupid as tplgroupid, billUserID as tplbillUserID, billUserPass as tplbillUserPass "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_partner.dbo.tbl_partner_tpl t "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.tplcompanyid = '" + CStr(tplcompanyid) + "' "
	'' response.write sqlStr
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF then
		tplcompanyname = db2html(rsget("tplcompanyname"))
		tplgroupid = rsget("tplgroupid")
		tplbillUserID = rsget("tplbillUserID")
		tplbillUserPass = rsget("tplbillUserPass")
	end if
	rsget.close
end function


'// ============================================================================
dim oTax
set oTax = new CTax
oTax.FRecttaxIdx = taxIdx
oTax.GetTaxRead

dim Bill365URL : Bill365URL = "http://www.bill36524.com"  '' :8090: test, 80: real
dim swfName    : swfName = "DzEBankFlexAPI" ''"dZAmfApp"
dim swfURL     : swfURL = "/designer/jungsan/"


'// ============================================================================
dim sell_hp, sell_hp1, sell_hp2, sell_hp3
dim buy_hp, buy_hp1, buy_hp2, buy_hp3

sell_hp = Split(oTax.FOneItem.FsupplyRepTel, "-")
buy_hp = Split(oTax.FOneItem.FrepTel, "-")

if (UBound(sell_hp) >= 0) then
	sell_hp1 = sell_hp(0)
end if

if (UBound(sell_hp) >= 1) then
	sell_hp2 = sell_hp(1)
end if

if (UBound(sell_hp) >= 2) then
	sell_hp3 = sell_hp(2)
end if

if (UBound(buy_hp) >= 0) then
	buy_hp1 = buy_hp(0)
end if

if (UBound(buy_hp) >= 1) then
	buy_hp2 = buy_hp(1)
end if

if (UBound(buy_hp) >= 2) then
	buy_hp3 = buy_hp(2)
end if


'// ============================================================================
if (oTax.FOneItem.Fbilldiv = "52") or (oTax.FOneItem.Fbilldiv = "55") then
	response.write "텐바이텐 이외 사업자 발행불가"
	response.end
end if


'// ============================================================================
dim reg_socno
dim reg_subsocno
dim reg_socname
dim reg_ceoname
dim reg_socaddr
dim reg_socstatus
dim reg_socevent
dim reg_managername
dim reg_managerphone
dim reg_managermail

dim tplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass

reg_socno			= oTax.FOneItem.FsupplyBusiNo
reg_subsocno		= oTax.FOneItem.FsupplyBusiSubNo
reg_socname			= oTax.FOneItem.FsupplyBusiName
reg_ceoname			= oTax.FOneItem.FsupplyBusiCEOName
reg_socaddr			= oTax.FOneItem.FsupplyBusiAddr
reg_socstatus		= oTax.FOneItem.FsupplyBusiType
reg_socevent		= oTax.FOneItem.FsupplyBusiItem
reg_managername		= oTax.FOneItem.FsupplyRepName
reg_managerphone	= oTax.FOneItem.FsupplyRepTel
reg_managermail		= oTax.FOneItem.FsupplyRepEmail


'// ============================================================================
dim FG_VAT : FG_VAT = "1"			'// 1과세, 3면세, 2영세(잘못된것 아님 : 빌365)

if IsNull(oTax.FOneItem.Ftaxtype) then
	oTax.FOneItem.Ftaxtype = ""
end if

'// Y : 과세 / N : 면세 / 0 : 영세
Select Case oTax.FOneItem.Ftaxtype
	Case "Y"
		FG_VAT = "1"
	Case "N"
		FG_VAT = "3"
	Case "0"
		FG_VAT = "2"
	Case Else
		response.write "과세구분 설정 에러"
		response.end
End Select


'// ============================================================================
dim isueDate

if IsNull(oTax.FOneItem.FisueDate) then
	oTax.FOneItem.FisueDate = ""
end if

if (oTax.FOneItem.FisueDate = "") then
	response.write "발행일자 설정 에러"
	response.end
else
	isueDate = oTax.FOneItem.FisueDate
end if


'// ============================================================================
dim ipkumdate : ipkumdate = ""

if IsNull(oTax.FOneItem.Fipkumdate) then
	oTax.FOneItem.Fipkumdate = ""
end if

'// 고객 주문의 경우 입금일자
ipkumdate = oTax.FOneItem.Fipkumdate


'// ============================================================================
dim consignYN

if IsNull(oTax.FOneItem.FconsignYN) then
	oTax.FOneItem.FconsignYN = ""
end if

if (oTax.FOneItem.FconsignYN = "") then
	response.write "위수탁구분 설정 에러"
	response.end
else
	consignYN = oTax.FOneItem.FconsignYN
end if


'// ============================================================================
if (oTax.FOneItem.Fbilldiv = "99") then
	Call Get3PLUpcheInfoByTPLCompanyid(oTax.FOneItem.Ftplcompanyid, tplcompanyname, tplgroupid, tplbillUserID, tplbillUserPass)
end if

%>
<script src="/designer/jungsan/AC_OETags.js" language="javascript"></script>
<script language='javascript'>

var fxStarted = false;

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


    var obj = AddNew("ID", iid);
    var obj1 = AddNew("PASSWD", ipwd);
    var obj2 = AddNew("USER_IP", "<%= request.ServerVariables("REMOTE_ADDR") %>");

    var arr = new Array(obj, obj1, obj2);

    if (!(thisMovie("<%= swfName %>"))) {
        alert('플래시 객체 생성오류');
    }

    thisMovie("<%= swfName %>").Login(arr);
   // alert('startedlogin');
}




//01.로그인 결과
function FxLoginResult(retObj){
//alert(retObj);
    var result = retObj.RESULT;
    var company_no = "<%= replace(replace(oTax.FOneItem.FbusiNo,"-","")," ","") %>";

    if (result=="00000"){
        billTaxEvalFlexApi();
    }else{
        alert(retObj.RESULT_MSG);
    }

}

//발행전 저장

//발행후 저장
function saveTaxEvalResult(result,no_tax,result_msg,no_iss){
    var frm = document.taxSaveFrm;

    frm.action="saveTaxResult.asp";
    frm.result.value = result;
    frm.no_tax.value = no_tax;
    frm.result_msg.value = result_msg;
    frm.no_iss.value = no_iss;

	frm.target = "ipreSave";
	frm.submit();

	fxStarted = false;
}


function billTaxEvalFlexApi(){

	<% if (ipkumdate <> "") then %>
	var obj1 = AddNew("FG_BILL","2");   //청구1 영수2
	<% else %>
	var obj1 = AddNew("FG_BILL","1");   //청구1 영수2
	<% end if %>

    var obj2 = AddNew("YN_TURN","Y");   //Y정발행 N역발행  :: 역발행시 발행요청 , 정발행시 승인요청

    var obj3 = AddNew("FG_IO","1");     //1매출 2매입
    var obj4 = AddNew("FG_PC","1");     //1기업 2개인
    var obj5 = AddNew("FG_FINAL","1");  //0저장 1 발송 2승인 3반려 4승인취소요청

	var obj6 = AddNew("YN_CSMT","<%= consignYN %>");	// 위수탁구분 Y 위수탁 N 정상

	var obj7 = AddNew("FG_VAT","<%= FG_VAT %>");    // 1과세, 3면세, 2영세(잘못된 것 아님)

	var obj8 = AddNew("AM","<%= oTax.FOneItem.FtotalPrice-oTax.FOneItem.FtotalTax %>");
    var obj9 = AddNew("AM_VAT","<%= oTax.FOneItem.FtotalTax %>");
    var obj10 = AddNew("AMT","<%= oTax.FOneItem.FtotalPrice %>");

    <% if (oTax.FOneItem.Fbilldiv = "01" or oTax.FOneItem.Fbilldiv = "11") then %>
    var obj11 = AddNew("AMT_CASH","<%= oTax.FOneItem.FtotalPrice %>");			// 현금
    <% else %>
    var obj11 = AddNew("AMT_AR","<%= oTax.FOneItem.FtotalPrice %>");			// 외상미수금
    <% end if %>

    var obj12 = AddNew("AMT_CHECK","0");
    var obj13 = AddNew("AMT_NOTE","0");
    var obj14 = AddNew("YMD_WRITE","<%= Replace(isueDate,"-","") %>");

	// 공급자
    var obj15 = AddNew("SELL_NO_BIZ","<%= Replace(reg_socno, "-", "") %>");
    var obj16 = AddNew("SELL_NM_CORP","<%= reg_socname %>");
    var obj17 = AddNew("SELL_NM_CEO","<%= reg_ceoname %>");
    var obj18 = AddNew("SELL_BIZ_STATUS","<%= reg_socstatus %>");
    var obj19 = AddNew("SELL_BIZ_TYPE","<%= reg_socevent %>");
    var obj20 = AddNew("SELL_ADDR1","<%= reg_socaddr %>");
    var obj21 = AddNew("SELL_ADDR2","");
    var obj22 = AddNew("SELL_DAM_DEPT","");
    var obj23 = AddNew("SELL_DAM_NM","<%= reg_managername %>");
    var obj24 = AddNew("SELL_DAM_EMAIL","<%= reg_managermail %>");
    var obj25 = AddNew("SELL_DAM_MOBIL1","<%= sell_hp1 %>");
    var obj26 = AddNew("SELL_DAM_MOBIL2","<%= sell_hp2 %>");
    var obj27 = AddNew("SELL_DAM_MOBIL3","<%= sell_hp3 %>");
    var obj28 = AddNew("SELL_DAM_TEL1","<%= sell_hp1 %>");
    var obj29 = AddNew("SELL_DAM_TEL2","<%= sell_hp2 %>");
    var obj30 = AddNew("SELL_DAM_TEL3","<%= sell_hp3 %>");

	// 공급받는자
    var obj31 = AddNew("BUY_NO_BIZ","<%= replace(replace(oTax.FOneItem.FbusiNo,"-","")," ","") %>");
    var obj32 = AddNew("BUY_NM_CEO","<%= oTax.FOneItem.FbusiCEOName %>");
    var obj33 = AddNew("BUY_NM_CORP","<%= oTax.FOneItem.FbusiName %>");
    var obj34 = AddNew("BUY_DAM_NM","<%= db2html(oTax.FOneItem.FrepName) %>");
    var obj35 = AddNew("BUY_DAM_EMAIL","<%= db2html(oTax.FOneItem.FrepEmail) %>");
    var obj36 = AddNew("BUY_DAM_MOBIL1","<%= buy_hp1 %>");
    var obj37 = AddNew("BUY_DAM_MOBIL2","<%= buy_hp2 %>");
    var obj38 = AddNew("BUY_DAM_MOBIL3","<%= buy_hp3 %>");
    var obj39 = AddNew("BUY_DAM_TEL1","<%= buy_hp1 %>");
    var obj40 = AddNew("BUY_DAM_TEL2","<%= buy_hp2 %>");
    var obj41 = AddNew("BUY_DAM_TEL3","<%= buy_hp3 %>");
    var obj42 = AddNew("BUY_ADDR1","<%= oTax.FOneItem.FbusiAddr %>");
    var obj43 = AddNew("BUY_ADDR2","");
    var obj44 = AddNew("BUY_BIZ_STATUS","<%= oTax.FOneItem.FbusiType %>");
    var obj45 = AddNew("BUY_BIZ_TYPE","<%= oTax.FOneItem.FbusiItem %>");
    var obj46 = AddNew("BUY_DAM_DEPT","");

    var obj47 = AddNew("YN_FX","N"); // 수정 세금계산서 여부  Y:수정 세금 계산서, N: 정상 발행 <== 필수 입력 입니다

<% if (Trim(oTax.FOneItem.Forderserial) <> "") then %>
    var obj48 = AddNew("DC_RMK2","주문번호/출고코드 : <%= oTax.FOneItem.Forderserial %>");
<% else %>
    var obj48 = AddNew("DC_RMK2","인덱스코드 : <%= oTax.FOneItem.Forderidx %>");
<% end if %>
    var today = new Date() ;

	// alert( today.getYear() + "" + (today.getMonth()+1) + "" +today.getDate() + "" +today.getHours() + "" +today.getMinutes() + "" +today.getSeconds());
    // var obj49 = AddNew("NO_SENDER_PK","DZ_PK_" +today.getYear() + "" + (today.getMonth()+1) + "" +today.getDate() + "" +today.getHours() + "" +today.getMinutes() + "" +today.getSeconds());

<%

' 1. 비고에 값이 있고 첫 두글자가 SO 로 되어 있으면 출고분계산서, 아니면 주문번호로를 PK 로 한다. (SO_주문번호, CUST_주문번호)
' 2. 비고에 값이 없고 orderidx 에 0 이 아닌 값이 있으면 가맹점계산서(FRAN_orderidx)
' 3. 비고에 값이 없고 orderidx 에 0 이면 추가발행계산서(TAX_taxIdx)

%>
<% if (Trim(oTax.FOneItem.Forderserial) <> "") and (Left(oTax.FOneItem.Forderserial, 2) = "SO") then %>
	// 출고코드
	var obj49 = AddNew("NO_SENDER_PK","SO_" + "<%= Trim(oTax.FOneItem.Forderserial) %>");
<% elseif (Trim(oTax.FOneItem.Forderserial) <> "") and (Left(oTax.FOneItem.Forderserial, 2) <> "SO") then %>
    <%
    dim osePK
    ''osePK = getOrderSerialPK(oTax.FOneItem.Forderserial)
    ''if (osePK="") then
    ''    response.write "alert('이미 발행 되었거나 올바른 주문번호가 아닙니다. - 관리자문의요망');return;"
	''end if
	osePK = oTax.FOneItem.Forderserial & "_" & reg_socno
    %>
	// 주문번호
	var obj49 = AddNew("NO_SENDER_PK","CUST_" + "<%= Trim(osePK) %>");

<% else %>

	// 기타
	var obj49 = AddNew("NO_SENDER_PK","TAX_" + "<%= Trim(CStr(oTax.FOneItem.FtaxIdx)) %>");

<% end if %>

	// 종사업장번호
	var obj50 = AddNew("SELL_REG_ID","<%= reg_subsocno %>");
	var obj51 = AddNew("BUY_REG_ID","<%= Trim(CStr(NULL2Blank(oTax.FOneItem.FbusiSubNo))) %>");

    //2016/04/18 추가
    var obj52 = AddNew("YN_ISS","0");  //FG_VAT 가 3(면세) 일경우 YN_ISS : NULL 일경우 전송제외 YN_ISS : 0 일경우 국세청 전송요청
    
    <% if (TRUE) or (FG_VAT="3") then %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45, obj46, obj47, obj48, obj49, obj50, obj51, obj52);
    <% else %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45,obj46,obj47,obj48, obj49, obj50, obj51);
    <% end if %>

    var objline1 = AddNew("ITEM_STD", "");
    var objline2 = AddNew("NM_ITEM", "<%= oTax.FOneItem.Fitemname %>");
    <% if (oTax.FOneItem.Fbilldiv = "01") or (oTax.FOneItem.Fbilldiv = "11") then %>
    var objline3 = AddNew("NO_ITEM", "");
    <% else %>
    var objline3 = AddNew("NO_ITEM", "1");
    <% end if %>
    var objline4 = AddNew("AM", "<%= oTax.FOneItem.FtotalPrice-oTax.FOneItem.FtotalTax %>");
    var objline5 = AddNew("AM_VAT", "<%= oTax.FOneItem.FtotalTax %>");
    var objline6 = AddNew("AMT", "<%= oTax.FOneItem.FtotalPrice %>");
    var objline7 = AddNew("DD_WRITE", "<%= Mid(isueDate,9,2) %>");
    var objline8 = AddNew("MM_WRITE", "<%= Mid(isueDate,6,2) %>");

    var arrline1 = new Array(objline1, objline2,objline3, objline4, objline5, objline6, objline7, objline8);

    var arrlineArr = new Array(arrline1);
    showDoing();

    thisMovie("<%= swfName %>").SendTaxMuch(1);
    thisMovie("<%= swfName %>").SendTaxAccount("", arr, arrlineArr);

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

    hideDoing();
    if (result!="00000"){
        alert(result_msg);
    }else{
        //popTax :: 정상발행

        //FxShowTaxAccount(no_tax,compNo);

        alert("계산서가 발행 되었습니다. ");
        // aaaaaaa
        opener.location.reload();
        // aaaaaaaa
        window.close();
    }
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
    return;
    //document.all["POPBillLogin"].style.visibility='hidden';
    //document.frm.evalButton.disabled=false;
}

function showLogin(){
    return;
    /*
    var frm = document.billfrm;
    frm.billid.value = '';
    frm.billpass.value = '';
    hideDoing();
    document.all["POPBillLogin"].style.visibility='visible';

    document.frm.evalButton.disabled=true;
    fxStarted = false;
    */
}

function showDoing(){
    document.all.idoingMsg.style.display='inline';
}

function hideDoing(){
    document.all.idoingMsg.style.display='none';
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

function getOnLoad(){
    setTimeout("evalTx()",1000)


}

function evalTx(){
    if (confirm('발행일 : <%= isueDate %>\n발행 하시겠습니까?')){
        
		<%
		Select Case oTax.FOneItem.Fbilldiv
			Case "01"
				'// 고객 - 공급자 텐바이텐
				response.write "FxLogin('customer','20011010');"
			Case "11"
				'// 고객 - 공급자(업체별)
				response.write "FxLogin('customer','20011010');"
			Case "02"
				'// 가맹점 - 공급자 텐바이텐
				response.write "FxLogin('accounts','20011010');"
			Case "03"
				'// 프로모션 - 공급자 텐바이텐
				response.write "FxLogin('promotion','20011010');"
			Case "51"
				'// 기타 - 공급자 텐바이텐
				response.write "FxLogin('accounts','20011010');"
			Case "99"
				'// 3PL업체
				response.write "FxLogin('" & tplbillUserID & "','" & tplbillUserPass & "');"
			Case Else
				response.write "FxLogin('customer','20011010');"
		End Select
		%>
    }
}

function closeMe(){
    window.close();
}

// FLESH 내부에서 기타 예외 발생시 오류 리턴
function FxErrorResult(retObj) {
    alert("ERR:" + retObj + "\n관리자 문의 요망");
    hideDoing();
}

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
    hideDoing();
}
</script>
<table border="0" cellspacing="0" cellpadding="0" width="500">
<tr>
    <td>
    <script language="JavaScript" type="text/javascript">
    	AC_FL_RunContent(
        	"src", "<%= swfURL & swfName %>",
        	"width", "400",
        	"height", "100",
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
    </td>
</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="500">
<tr bgcolor="#FFFFFF" id="idoingMsg" style="display:none">
    <td colspan="2" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/loading.gif" width="269" height="14"><br>처리중입니다.잠시만 기다려주세요...</td>
</tr>
</table>
<form name="taxSaveFrm" method="post">
<input type="hidden" name="taxIdx" value="<%= taxIdx %>">
<input type="hidden" name="result" value="">
<input type="hidden" name="no_tax" value="">
<input type="hidden" name="result_msg" value="">
<input type="hidden" name="no_iss" value="">
<input type="hidden" name="write_date" value="<%= isueDate %>">
</form>
<iframe name="ipreSave" id="ipreSave" width="400" height="110"></iframe>
<%
set oTax = Nothing
%>
<script language=javascript>
//IE8에서 플래시가 나중에 표시됨..
window.onload=getOnLoad;

</script>
<input type="button" value="발행" onclick="evalTx();">
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

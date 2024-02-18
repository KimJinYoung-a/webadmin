<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payRequestCls.asp"-->
<%
Dim taxKey      : taxKey = requestCheckvar(request("taxKey"),24)
Dim taxSellType : taxSellType = requestCheckvar(request("taxSellType"),10)
Dim clsEsero

if taxSellType="" then taxSellType="0" ''기본 매입



Dim appDate
Dim sellCorpNo, sellJongNo, sellCorpName, sellCeoName, sellEmail
Dim buyCorpNo, buyJongNo, BuyCorpName, BuyCeoName, buyEmail

Dim suplySum,taxSum,totSum
Dim bigo, DtlName
Dim taxModiType, evalTypeNm, taxType, recreqGubunNm, DtlDate, DtlSuplysum,DtltaxSum, DtlBigo,reqDate,sendDate,regdate
Dim matchType,matchKey,matchState,bizSecCD,erpLinkType,erpLinkKey,addCnt
Dim cust_cd, cust_nm, arap_cd,arap_nm, prod_cd, prod_nm
Dim clsPay, spayrequestTitle
Dim mayErrType, jacctcd

IF (taxSellType="0") then
    sellCorpNo      = ""
    sellJongNo      = ""
    sellCorpName    = ""
    sellCeoName     = ""
    sellEmail       = ""

    buyCorpNo       = "2118700620"
    buyJongNo       = ""
    BuyCorpName     = "(주)텐바이텐"
    BuyCeoName      = "최은희"
    buyEmail        = ""
ELSE
    sellCorpNo      = "2118700620"
    sellJongNo      = ""
    sellCorpName    = "(주)텐바이텐"
    sellCeoName     = "최은희"
    sellEmail       = ""

    buyCorpNo       = ""
    buyJongNo       = ""
    BuyCorpName     = ""
    BuyCeoName      = ""
    buyEmail        = ""
ENd IF

Dim ArrVal, IsEditMode
IF (taxKey<>"") then  ''수정 모드.
    set clsEsero = new CEsero
    clsEsero.FtaxKey = taxKey
    ArrVal = clsEsero.fnGetEseroOneTax
    set clsEsero = Nothing
end if

IF IsArray(ArrVal) then
    ''T.taxKey,T.appDate,T.sellCorpNo,T.sellJongNo,T.sellCorpName,T.sellCeoName,T.sellEmail,T.buyCorpNo,T.buyJongNo
	'',T.BuyCorpName,T.BuyCeoName,T.buyEmail,T.totSum,T.suplySum,T.taxSum,T.taxSellType,T.taxModiType,T.taxType,T.evalTypeNm
	'',T.Bigo,T.recreqGubunNm,T.DtlDate,T.DtlName,T.DtlSuplysum,T.DtltaxSum,T.DtlBigo,T.reqDate,T.sendDate,T.regdate
	'',M.matchType, M.matchKey, M.matchState, M.bizSecCD, M.erpLinkType, M.erpLinkKey
	'',(select Count(*) from db_Partner.dbo.tbl_Esero_TaxMatch C where C.taxKey=T.taxKey and C.matchSeq>0) as addCnt
	IsEditMode = true
	appDate = ArrVal(1,0)
	sellCorpNo = ArrVal(2,0)
	sellJongNo = ArrVal(3,0)
	sellCorpName = ArrVal(4,0)
	sellCeoName  = ArrVal(5,0)
	sellEmail    = ArrVal(6,0)

	buyCorpNo  = ArrVal(7,0)
	buyJongNo  = ArrVal(8,0)
	BuyCorpName  = ArrVal(9,0)
	BuyCeoName  = ArrVal(10,0)
	buyEmail  = ArrVal(11,0)

	totSum      = ArrVal(12,0)
	suplySum    = ArrVal(13,0)
	taxSum      = ArrVal(14,0)

	taxSellType = ArrVal(15,0)
	taxModiType = ArrVal(16,0)
	taxType     = ArrVal(17,0)
	evalTypeNm  = ArrVal(18,0)
	Bigo        = ArrVal(19,0)
	recreqGubunNm = ArrVal(20,0)
	DtlDate     = ArrVal(21,0)
	DtlName     = ArrVal(22,0)
	DtlSuplysum = ArrVal(23,0)
	DtltaxSum   = ArrVal(24,0)
	DtlBigo   = ArrVal(25,0)
	reqDate   = ArrVal(26,0)
	sendDate  = ArrVal(27,0)
	regdate  = ArrVal(28,0)

	matchType   = ArrVal(29,0)
	matchKey    = ArrVal(30,0)
	matchState  = ArrVal(31,0)
	bizSecCD    = ArrVal(32,0)
	erpLinkType = ArrVal(33,0)
	erpLinkKey  = ArrVal(34,0)
	addCnt      = ArrVal(35,0)

	cust_cd     = ArrVal(36,0)
	cust_nm     = ArrVal(37,0)
	arap_cd     = ArrVal(38,0)
	arap_nm     = ArrVal(39,0)
	prod_cd     = ArrVal(40,0)
	prod_nm     = ArrVal(41,0)

    mayErrType  = ArrVal(47,0)
    jacctcd = ArrVal(48,0)

    ''품목정보 => 자금용도로 매칭하기 위한
    if (matchType=9) then

        set clsPay = new CPayRequest
    	clsPay.FpayrequestIdx = matchKey
    	clsPay.fnGetPayRequestReceiveData
	    spayrequestTitle	= clsPay.FpayRequestTitle
	    SET clsPay=Nothing

    end if


End if

Dim IsElecTax
IsElecTax = (Not (taxModiType="9")) and (Not (taxModiType=""))

Dim inValidStr

Dim sTotCnt, sArr, intLoop

IF (IsEditMode) then
    ''' 매핑 상세내역 리스트
    set clsEsero = new CEsero
    clsEsero.FtaxKey = taxKey
    sArr = clsEsero.fnGetMappingList
    set clsEsero = Nothing

    If IsArray(sArr) then
        sTotCnt = UBound(sArr,2)
    end if
ENd IF

%>


<script language='javascript'>
//거래처 정보 보기
function jsGetCust(){
	var Strparm="";
	var cust_cd = "";
	var rdoCgbn = "2"; //매입
	if (cust_cd!=""){
		Strparm = "?selSTp=1&sSTx="+ cust_cd;
	}else{
	    Strparm = "?rdoCgbn="+rdoCgbn;
	}
	Strparm = Strparm + "&opnType=eTax";
	var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1200, height=600,resizable=yes, scrollbars=yes");
	winC.focus();
}

//거래처 선택
function jsSetCust(custcd, custnm, ceonm, custno ){
    var frm = document.frmEtax;
    frm.hidcustcd.value = custcd;
    var currSellType = "<%= taxSellType %>";

    if (currSellType=="0"){
        frm.sellCorpName.value = custnm;
        frm.sellCeoName.value = ceonm;
        frm.sellCorpNo.value = custno;
    }else{
        frm.buyCorpName.value = custnm;
        frm.buyCeoName.value = ceonm;
        frm.buyCorpNo.value = custno;
    }
}

function changeFrm(comp){
    var currSellType = "<%= taxSellType %>";

    if (currSellType!=comp.value){
        document.frm.submit();
    }
}

function saveHandTax(isEdit){
    var frm = document.frmEtax;
    var preFrm =document.frmPreCheck;
    //if (frm.hidcustcd.value.length<1){
    //    alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
    //    return;
    //}

    if (frm.sellCorpNo.value.length<1){
        alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
        return;
    }

    if (frm.sellCorpName.value.length<1){
        alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
        return;
    }

    if (frm.sellCeoName.value.length<1){
        alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
        return;
    }

    if (frm.buyCorpNo.value.length<1){
        alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
        return;
    }

    if (frm.buyCorpName.value.length<1){
        alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
        return;
    }

    if (frm.buyCeoName.value.length<1){
        alert('거래처 정보 오류 - 사업자정보 선택 후 사용요망');
        return;
    }

    if (frm.DtlName.value.length<1){
        alert('품목 정보 오류 - 품목 입력 후 사용요망');
        return;
    }

    if (frm.taxType.value.length<1){
        alert('과세 구분 오류 - 과세구분 입력 후 사용요망');
        return;
    }

    if (frm.suplySum.value.length<1){
        alert('공급가 정보 오류 - 공급가 입력 후 사용요망');
        return;
    }

    if (frm.taxSum.value.length<1){
        alert('세금 정보 오류 - 세금 입력 후 사용요망');
        return;
    }

    if (frm.totSum.value.length<1){
        alert('총금액 정보 오류 - 총금액 입력 후 사용요망');
        return;
    }

    if ((frm.totSum.value*1)!=(frm.taxSum.value*1+frm.suplySum.value*1)){
        alert('총금액<>공급가+세금 정보가 일치하지 않습니다.' + frm.totSum.value + '<>' + (frm.taxType.value*1+frm.suplySum.value*1));
        return;
    }

    if ((frm.taxType.value=="1")&&(frm.taxSum.value*1==0)){
        alert('과세이나 세금액 없음');
        return;
    }

    if (((frm.taxType.value=="2")||(frm.taxType.value=="3"))&&(frm.taxSum.value*1!=0)){
        alert('면세/영세이나 세금액 있음');
        return;
    }

    var buf='입력';
    if (isEdit) buf='수정';

    if (confirm('수기 세금 계산서 정보를 ['+buf+'] 하시겠습니까?\n\n전자 세금계산서인경우 엑셀 또는 XML로 입력하셔야 합니다.')){
        frm.target = "ifrm_PreCheck";
        frm.submit();

    }
}

function confirmedSubmit(){
    var frm = document.frmEtax;
    frm.target = "";
    if (confirm('동일 내역의 세금계산서가 존재 합니다.\n\n(사업자, 발행일, 금액 동일) 계속 하시겠습니까?')){
        frm.duppConfirm.value="on";
        frm.submit();
    }
}

function reCalcuFillSum(comp){
    var frm = document.frmEtax;


    if (frm.appDate.value.length==10){
        frm.DtlDate1.value = frm.appDate.value.substr(5,2);
        frm.DtlDate2.value = frm.appDate.value.substr(8,2);
    }else{
        frm.DtlDate1.value = "";
        frm.DtlDate2.value = "";
    }

    if (frm.taxType.value==""){
        alert('과세 구분을 선택 하세요.');
        frm.taxType.focus();
        return;
    }

    if (frm.taxType.value=="1"){
        if ((comp.name!="DtltaxSum")){
            frm.DtltaxSum.value = parseInt(frm.DtlsuplySum.value*0.1);
        }

        frm.suplySum.value=frm.DtlsuplySum.value;
        frm.taxSum.value=frm.DtltaxSum.value;
        frm.totSum.value=frm.suplySum.value*1+frm.taxSum.value*1;
    }else{
        frm.suplySum.value=frm.DtlsuplySum.value*1;
        if (frm.DtltaxSum.value.length<1) frm.DtltaxSum.value=0;
        frm.taxSum.value=frm.DtltaxSum.value*1;
        frm.totSum.value=frm.suplySum.value*1+frm.taxSum.value*1;
    }



}

function delHandTax(itaxkey){
    var frm = document.frmAct;
    frm.taxKey.value=itaxkey;

    if (itaxkey.length!=24){
        alert('승인번호가 올바르지 않습니다.');
        return;
    }

    if (confirm('수기 계산서를 삭제 하시겠습니까?')){
        frm.mode.value="delHandTax";
        frm.submit();
    }
}

function popTargetDetail(itargetGb,iidx,iridx){
    var popURL ='';
    if (itargetGb=="1"){
        popURL = "/admin/upchejungsan/nowjungsanmasteredit.asp?id="+iidx;
    }else if (itargetGb=="2"){
        popURL = "/admin/offupchejungsan/off_jungsanstateedit.asp?idx="+iidx;

    // 원가 매칭
    }else if (itargetGb=="4"){
        popURL = "/admin/newstorage/PurchasedProductSheetModify.asp?idx="+iidx;
    }else if (itargetGb=="9"){
        popURL = "/admin/approval/eapp/modeappPayDoc.asp?ipridx="+iidx+"&iridx="+iridx;
    }else if (itargetGb=="11"){
        popURL = "/cscenter/taxsheet/Tax_view.asp?taxIdx="+iidx;
    }else{
        popURL = "";
    }

    var popWin = window.open(popURL,'popTargetDetail','width=1400,height=800,scrollbars=yes,resizable=yes');
    popWin.focus();
}

//자금관리부서 선택
var G_matchSeq = 0;
function jsGetPart(imatchseq, t){
    if (!confirm('상품매입금 또는 결제요청내역과 매핑할 수 있는 자료는 수기 매핑 않하는것이 원칙입니다.\n\n결제요청 자료와 매핑할 자료가 없는경우에만 사용.\n\n그래도 계속 진행 하시겠습니까?')){
        return;
    }

    G_matchSeq = imatchseq;
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne2.asp?taxKey='+t,'popP','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//editDtlName
function editDtlName(){
    var frm = document.frmAct;
    if (confirm('품목명을 수정하시겠습니까?')){
        frm.mode.value="modiDtlName"
        frm.DtlName.value=document.frmEtax.DtlName.value;
        frm.taxKey.value = '<%= taxKey %>';
        frm.submit();
    }
}

//자금관리부서 등록
function jsSetPart(bizSecCd, sPNM){
    var frm = document.frmAct;
    if (confirm('사업 부문을 ' + sPNM + '로 수정하시겠습니까?')){
        frm.mode.value="modiBizSec"
	    frm.bizSecCd.value = bizSecCd;
	    frm.taxKey.value = '<%= taxKey %>';
	    frm.matchSeq.value = G_matchSeq;
	    frm.submit();
	}
}

//수지항목 불러오기
function jsGetARAP(imatchseq, t){
    G_matchSeq = imatchseq;
    rdoGB = "<%= CHKIIF(taxSellType="0","2","1") %>";
	var winARAP = window.open("/admin/linkedERP/arap/popGetARAP2.asp?rdoGB="+rdoGB+"&taxKey="+t,"popARAP1","width=800,height=600,resizable=yes, scrollbars=yes");
	winARAP.focus();
}

//선택 수지항목 가져오기
function jsSetARAP(dAC, sANM,sACC,sACCNM){
    var frm = document.frmAct;
    if (confirm('수지 항목을 ' + sANM + '로 수정하시겠습니까?')){
        frm.mode.value="modiArapCD"
	    frm.arap_cd.value = dAC;
	    frm.taxKey.value = '<%= taxKey %>';
	    frm.matchSeq.value = G_matchSeq;
	    frm.submit();
	}

}

// 매핑정보 삭제
function delMapDtl(matchSeq){
    var frm = document.frmAct;
    if (confirm('매핑 정보를 삭제 하시겠습니까?')){
        frm.mode.value="delMapDTL"
	    frm.taxKey.value = '<%= taxKey %>';
	    frm.matchSeq.value = matchSeq;
	    frm.submit();
	}
}
// 수기 매핑으로 변환
function chgHandMap(matchSeq){
    var frm = document.frmAct;
    if (confirm('매핑 링크를 삭제 하시겠습니까?')){
        frm.mode.value="chgHandMap"
	    frm.taxKey.value = '<%= taxKey %>';
	    frm.matchSeq.value = matchSeq;
	    frm.submit();
	}
}

function sendERP(iTaxKey){
    var frm = document.frmAct;
    if (confirm('증빙서류를 ERP로 전송하시겠습니까?')){
        frm.mode.value="sendDocErp"
        frm.taxKey.value = iTaxKey;
        //alert(document.frmEtax.chkPLANDATE.checked);
        //if (document.frmEtax.chkPLANDATE.checked==true){
        //    frm.chkPLANDATE.value = "on";
        //}else{
            frm.chkPLANDATE.value = "";
        //}
        frm.action="eTax_sERP_process.asp"; //2016/05/10
        frm.submit();
    }
}

function sendERPHand(iTaxKey){
    var frm = document.frmAct;
    if (confirm('증빙서류를 수기 입력 완료 처리 하시겠습니까?')){
        frm.mode.value="finishDocHand"
        frm.taxKey.value = iTaxKey;
        frm.submit();
    }
}

function delSendErpLinkKey(iTaxKey){
    var frm = document.frmAct;
    if (confirm('증빙서류를 전송 내역을 삭제 하시겠습니까?')){
        frm.mode.value="delErpLinkKey"
        frm.taxKey.value = iTaxKey;
        frm.submit();
    }
}

function mayErrEvalSave(){

    var frm = document.frmAct;
    if (confirm('오발행 예상건으로 지정 하시겠습니까?')){
        frm.mode.value="mayErrStat"
	    frm.taxKey.value = '<%= taxKey %>';
	    frm.submit();
	}
}

function mayErrEvalDel(){

    var frm = document.frmAct;
    if (confirm('오발행 예상건 삭제 하시겠습니까?')){
        frm.mode.value="mayErrStatDel"
	    frm.taxKey.value = '<%= taxKey %>';
	    frm.submit();
	}
}

</script>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<form name="frm">
<tr>
    <td>
        <% if (IsElecTax) then %>
        <input type="hidden" name="taxSellType" value="<%= taxSellType %>">

        <b><%= getSellTypeName(taxSellType) %></b>
        <%= gettaxModiTypeName(taxModiType) %>
        <%=gettaxTypeName(taxType) %>
        세금 계산서 <input type="text" name="taxKey" value="<%=taxKey%>" size="30" class="text_ro" readonly >

        <% else %>
        수기 (세금)계산서 입력
    &nbsp;&nbsp;
        <input type="radio" name="taxSellType" value="0" <%= CHKIIF(taxSellType="0","checked","") %> onClick="changeFrm(this)">매입
        <input type="radio" name="taxSellType" value="1" <%= CHKIIF(taxSellType="1","checked","") %> onClick="changeFrm(this)">매출
        <% end if %>
    </td>
</tr>
</form>
</table>
<p>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top">
	<form name="frmEtax" method="POST" action="eTax_process.asp">
	<input type="hidden" name="taxKey" value="<%=taxKey%>">
	<input type="hidden" name="hidcustcd" value="<%=cust_cd%>">
	<input type="hidden" name="taxSellType" value="<%=taxSellType%>">
	<input type="hidden" name="mode" value="handTaxInput">
	<input type="hidden" name="duppConfirm" value="">

        <td width="49%">
        	<!-- 공급자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>공급자 정보</b>
        			</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">사업자번호</td>
        			<td ><input type=text name="sellCorpNo" size=16 value="<%= sellCorpNo %>" class="<%= CHKIIF(sellCorpNo<>"","text_ro","text") %>" <%= CHKIIF(sellCorpNo<>"","readonly","") %> >
        			<% if (taxSellType="0") and (Not IsElecTax) then %>
        			<input type="button" class="button" value="선택" onClick="jsGetCust()">
        			<% end if %>
        			</td>
        			<td bgcolor="#F0F0FD" height="25">종번호</td>
        			<td ><input type=text name="sellJongNo" size=8 value="<%= sellJongNo %>" class="<%= CHKIIF(sellJongNo<>"","text_ro","text") %>" <%= CHKIIF(sellJongNo<>"","readonly","") %> ></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
        			<td><input type=text name="sellCorpName" size=16 value="<%= sellCorpName %>" border=0 class="<%= CHKIIF(sellCorpName<>"","text_ro","text") %>" <%= CHKIIF(sellCorpName<>"","readonly","") %> ></td>
        			<td width="70" bgcolor="#F0F0FD">대표자</td>
        			<td><input type=text name="sellCeoName" size=16 value="<%= sellCeoName %>" class="<%= CHKIIF(sellCeoName<>"","text_ro","text") %>" <%= CHKIIF(sellCeoName<>"","readonly","") %> ></td>
        		</tr>

        		<% if (taxSellType="0") then %>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">거래처코드</td>
        			<td ><%= cust_cd %> </td>
        			<td colspan="2"><%= cust_nm %></td>
        	    </tr>
        		<% end if %>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">담당자</td>
        			<td > </td>
        			<td colspan="2"><%= sellEmail %></td>
        	    </tr>

        		<!--
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">이메일</td>
        			<td colspan=3><input type=text name="sellEmail" size=20 value="<%= sellEmail %>" class="text" ></td>
        		</tr>
        		-->
        	</table>
        	<!-- 공급자정보 끝 -->
        </td>
        <td>&nbsp;</td>
        <td width="49%">
        	<!-- 공급받는자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>공급받는자 정보</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        		    <td bgcolor="#F0F0FD" height="25">사업자번호</td>
        			<td ><input type=text name="buyCorpNo" size=16 value="<%= buyCorpNo %>" class="<%= CHKIIF(buyCorpNo<>"","text_ro","text") %>" <%= CHKIIF(buyCorpNo<>"","readonly","") %> >
        			<% if (taxSellType<>"0") and (Not IsElecTax) then %>
        			<input type="button" class="button" value="선택" onClick="jsGetCust()">
        			<% end if %>
        			</td>
        			<td bgcolor="#F0F0FD" height="25">종번호</td>
        			<td ><input type=text name="buyJongNo" size=8 value="<%= buyJongNo %>" class="<%= CHKIIF(buyJongNo<>"","text_ro","text") %>" <%= CHKIIF(buyJongNo<>"","readonly","") %>></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
        			<td><input type=text name="buyCorpName" size=16 value="<%= buyCorpName %>" border=0 class="<%= CHKIIF(buyCorpName<>"","text_ro","text") %>" <%= CHKIIF(BuyCorpName<>"","readonly","") %> ></td>
        			<td width="70" bgcolor="#F0F0FD">대표자</td>
        			<td><input type=text name="buyCeoName" size=16 value="<%= buyCeoName %>" class="<%= CHKIIF(buyCeoName<>"","text_ro","text") %>" <%= CHKIIF(BuyCeoName<>"","readonly","") %> ></td>
        		</tr>
        		<% if (taxSellType<>"0") then %>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">거래처코드</td>
        			<td ><%= cust_cd %> </td>
        			<td colspan="2"><%= cust_nm %></td>
        	    </tr>
        		<% end if %>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">담당자</td>
        			<td > </td>
        			<td colspan="2"><%= buyEmail %></td>
        	    </tr>

        	</table>
        	<!-- 공급받는자정보 끝 -->
        </td>
	</tr>
</table>
<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="120" height="25">작성일</td>
		<td width="100">과세구분</td>
		<td width="100">공급가액</td>
		<td width="100">세액</td>
		<td width="100">합계금액</td>
		<td>비고</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25">
			<input type="text" size="10" name="appDate" value="<%=appDate%>" onClick="calendarOpen(frmEtax.appDate);" style="cursor:hand;" class="writebox" onChange="reCalcuFillSum(this)">
			<a href="javascript:calendarOpen(frmEtax.appDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		</td>
		<td>
		<select name="taxType" onChange="reCalcuFillSum(this);">
		<option value="">선택
		<option value="1" <%= CHKIIF(taxType="1","selected","") %> >과세
		<option value="2" <%= CHKIIF(taxType="2","selected","") %> >영세
		<option value="3" <%= CHKIIF(taxType="3","selected","") %> >면세
		</select>
		</td>
		<td><input type="text" name="suplySum" value="<%= (suplySum) %>" size=9 class="text_ro" readonly style="text-align=right"></td>
		<td><input type="text" name="taxSum" value="<%= (taxSum) %>" size=9 class="text_ro" readonly style="text-align=right"></td>
		<td><input type="text" name="totSum" value="<%= (totSum) %>" size=9 class="text_ro" readonly style="text-align=right"></td>
		<td><input type="text" size="30" name="bigo" class="text" value="<%= bigo %>" maxlength="60"></td>
	</tr>
</table>

<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="30" height="25">월</td>
		<td width="30">일</td>
		<td>품목</td>
		<td width="100">공급가액</td>
		<td width="100">세액</td>
		<td width="100">영수/청구</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"><input type="text" name="DtlDate1" value="<%= Mid(DtlDate,6,2) %>" size="2" class="text_ro" readonly ></td>
		<td><input type="text" name="DtlDate2" value="<%= Right(DtlDate,2) %>" size="2" class="text_ro" readonly ></td>
		<td><input type=text name="DtlName" size=40 value="<%=DtlName%>" class="text">
		<% if (spayrequestTitle<>"") and (spayrequestTitle<>DtlName) then %>
		<input type="button" value="수정" class="button" onClick="editDtlName();"><br><font color=red onDblClick="document.frmEtax.DtlName.value='<%=spayrequestTitle%>';"><%=spayrequestTitle%></font></td>
		<% else %>
		<input type="button" value="수정" class="button" onClick="editDtlName();">
		</td>
		<% end if %>

		<td><input type=text name="DtlsuplySum" size=10 value="<%= (DtlSuplysum) %>"  class="text" style="text-align=right" onKeyUp="reCalcuFillSum(this)"> </td>
		<td><input type=text name="DtltaxSum" size=10 value="<%= (DtltaxSum) %>"  class="text" style="text-align=right" onKeyUp="reCalcuFillSum(this)"> </td>
		<td>
		<select name="recreqGubunNm">
		<option value="청구" <%= CHKIIF(recreqGubunNm="청구","selected","") %> >청구
		<option value="영수" <%= CHKIIF(recreqGubunNm="영수","selected","") %> >영수
		</select>
		</td>

	</tr>

</table>

<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if (Not IsElecTax) then %>
<tr>
    <td bgcolor="#FFFFFF" align="center">
        <% if (IsEditMode) then %>
        <input type="button" value="수기계산서 수정" onClick="saveHandTax(<%= LCASE(IsEditMode) %>);">
        &nbsp;
        <input type="button" value="수기계산서 삭제" onClick="delHandTax('<%= taxKey %>');">
        <% else %>
        <input type="button" value="수기계산서 입력" onClick="saveHandTax(<%= LCASE(IsEditMode) %>);">
        <% end if %>
    </td>
</tr>
<% end if %>
</table>
<p>
<% IF (IsEditMode) then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" colspan="8" >* 매핑 / erp전송상태 </td>
</tr>
<tr bgcolor="#F0F0FD" align="center">
    <td>매핑상태</td>
    <td>매핑구분</td>
    <td>매핑IDX</td>
    <td>결제단계</td>
    <td>사업부문</td>
    <td>수지항목</td>
    <td>거래종류</td>
    <td>ERP전송상태</td>
</tr>
<%
Dim iMatchSeq
Dim imatchType
Dim imatchKey
Dim imatchState
Dim ibizSecCD
Dim ierpLinkType
Dim ierpLinkKey
Dim jstatus
Dim ireportidx

%>
<% IF IsArray(sArr) then %>
    <%  For intLoop = 0 To UBound(sArr,2) %>
    <%
        ''0 M.taxKey ,M.matchSeq, M.matchType, M.matchKey, M.matchState, M.bizSecCD
        ''6, M.erpLinkType, M.erpLinkKey,jstatus, ireportidx

        iMatchSeq    = sArr(1,intLoop)
        imatchType  = sArr(2,intLoop)
        imatchKey   = sArr(3,intLoop)
		imatchState = sArr(4,intLoop)
		ibizSecCD    = sArr(5,intLoop)
		ierpLinkType  = sArr(6,intLoop)
		ierpLinkKey    = sArr(7,intLoop)
		jstatus      = sArr(8,intLoop)
		ireportidx  = sArr(9,intLoop)

    %>
    <tr bgcolor="#FFFFFF" align="center">
        <td><%= getMatchStateName(imatchState) %></td>
        <td><%= getMatchTypeName(imatchType) %></td>
        <td>
            <%= imatchKey %>
            <% if Not IsNULL(imatchKey) and (imatchKey<>0) then %>
            <img src="/images/icon_arrow_link.gif" onClick="popTargetDetail('<%= imatchType %>','<%= imatchKey %>','<%=ireportidx%>')" style="cursor:pointer">
            <% end if %>
        </td>
        <td>
            <%= getCommonTargetStatus(imatchType,jstatus) %>
        </td>
        <td><%= getbizSecCDName(ibizSecCD) %>
            <% IF IsNULL(ibizSecCD) or (imatchKey=0) then %>
            <img src="/images/icon_search.jpg" onClick="jsGetPart(0, '<%= taxKey %>');" style="cursor:pointer">
            <% end if %>
        </td>
        <td><%= arap_nm %>
            <% If Not IsNULL(arap_cd) then %>
            <br>[<%= arap_cd %>]
            <% end if %>
            <% IF IsNULL(arap_cd) or (imatchKey=0) or (jacctcd=8340003) then %>
            <img src="/images/icon_search.jpg" onClick="jsGetARAP(0, '<%= taxKey %>');" style="cursor:pointer">
            <% end if %>
        </td>
        <td><%= prod_nm %></td>
        <td>
            <% if Not IsNULL(ierpLinkType) then %><!-- 전송완료 -->
                [<%= ierpLinkType %>]
                <%= ierpLinkKey %>

                <% if session("ssBctID")="icommang" or session("ssBctID")="coolhas" then %>
                <% if matchType="11" then %>
                    <a href="javascript:chgHandMap(<%= iMatchSeq %>);">[수기변환]</a>
                <% end if %>
                <% end if %>
            <% else %>
                <% if ierpLinkType="H" then %>
                <img src="/images/i_delete.gif" onClick="delMapDtl(<%= iMatchSeq %>);" style="cursor:pointer">
                <% elseif (matchType="1") or (matchType="2") or (matchType="3") then %>

                <% else %>
                <img src="/images/i_delete.gif" onClick="delMapDtl(<%= iMatchSeq %>);" style="cursor:pointer">
                <% end if %>
            <% end if %>
        </td>
    </tr>
    <% next %>
<% else %>
    <tr bgcolor="#FFFFFF" align="center" height="30">
        <td colspan="4" align="center"> 매핑 내역 없음 =&gt;</td>
        <td><input type="button" value="사업부문수기지정" class="button" onClick="jsGetPart(0, '<%= taxKey %>')"></td>
        <td><%= arap_nm %>
            <% IF IsNULL(arap_cd) or (imatchKey=0) then %>
            <img src="/images/icon_search.jpg" onClick="jsGetARAP(0, '<%= taxKey %>');" style="cursor:pointer">
            <% end if %>
        </td>
        <td><%= prod_nm %></td>
        <td></td>
    </tr>
<% end if %>
<tr>
    <td colspan="8" align="center" bgcolor="#FFFFFF" height="50">
        <%
            Dim iTargetState : iTargetState =-1
            IF IsArray(sArr) then
                iTargetState = sArr(8,0)   ''첫번째거.
            END IF
         %>
            <% if IsERPSendAvail(matchState, matchType, erpLinkType, erpLinkKey, bizSecCD,iTargetState , arap_cd, inValidStr) then %>
                <% if isPLAN_DATEDefaultSend(imatchType, taxSellType, arap_cd) then %>
                <input type="checkbox" name="chkPLANDATE" value="" checked >(수입/지출)예정정보입력
                <% else %>
                <input type="checkbox" name="chkPLANDATE" value=""  >(수입/지출)예정정보입력
                <% end if %>
                <p>
                <input type="button" value="  ERP 전송  " onClick="sendERP('<%= taxKey %>');" class="button" >
            <% else %>
                <b><%= inValidStr %></b>
                <% end if %>
                &nbsp;&nbsp;&nbsp;
                <% if IsERPHandInpuAvail(matchState, matchType, erpLinkType, erpLinkKey, bizSecCD, iTargetState, arap_cd, inValidStr) then %>
                <input type="button" value="ERP 수기 처리" onClick="sendERPHand('<%= taxKey %>');" class="button" >
                <% else %>
                <b><%= inValidStr %></b>
                <% if (C_ADMIN_AUTH) or (C_MngPart) then %>
                    (관리자메뉴 : <input type="button" value="전송내역삭제" onClick="delSendErpLinkKey('<%= taxKey %>');" class="button" >)
                <% end if %>
            <% end if %>
    </td>
</tr>
<% if (IsElecTax) then %>
<tr>
    <td colspan="8" align="center" bgcolor="#FFFFFF" height="50">
    <% if isNULL(mayErrType) then %>
    <input type="button" value="오발행 예상 저장" onClick="mayErrEvalSave();" class="button" >

    <% else %>
    <input type="button" value="오발행 예상 삭제" onClick="mayErrEvalDel();" class="button" >
    <% end if %>
    </td>
</tr>
<% end if %>
</table>
<p>
<% end if %>
</form>

<form name="frmAct" method="post" action="eTax_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="taxKey" value="">
<input type="hidden" name="bizSecCd" value="">
<input type="hidden" name="arap_cd" value="">
<input type="hidden" name="matchSeq" value="">
<input type="hidden" name="chkPLANDATE" value="">
<input type="hidden" name="DtlName" value="">
</form>
<iframe src="" name="ifrm_PreCheck" id="ifrm_PreCheck" width="220" height="325" frameborder="0" scrolling="no"></iframe>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteMifinCls.asp"-->

<%

dim sitename, research, searchType, page
dim Matchorderserial, OutMallOrderSerial
dim excNoOrderSerial, shppDivDtl
sitename = RequestCheckVar(Request("sitename"),32)
research = RequestCheckVar(Request("research"),32)
searchType = RequestCheckVar(Request("searchType"),32)
page     = RequestCheckVar(Request("page"),10)
Matchorderserial     = RequestCheckVar(Request("Matchorderserial"),32)
OutMallOrderSerial   = RequestCheckVar(Request("OutMallOrderSerial"),32)
excNoOrderSerial     = RequestCheckVar(Request("excNoOrderSerial"),32)
shppDivDtl           = RequestCheckVar(Request("shppDivDtl"),10)

if (page="") then page=1
if (searchType="") and research="" then searchType="1"
'if (research="") then excNoOrderSerial = "Y"
if (shppDivDtl="") and research="" then shppDivDtl="N"

dim i
Dim sqlStr

Dim iOutMallDlvCode
Dim oMiFin
set oMiFin = new CxSiteMifinCls
oMiFin.FCurrPage = page
oMiFin.FPageSize = 50
oMiFin.FRectSellsite = sitename
oMiFin.FRectSearchType = searchType
oMiFin.FRectMatchorderserial = Matchorderserial
oMiFin.FRectOutMallOrderSerial = OutMallOrderSerial
oMiFin.FRectExcNoOrderSerial = excNoOrderSerial
oMiFin.FRectshppDivDtl = shppDivDtl

oMiFin.getXSiteMifinLIST


'''/// /admin/apps/outMallAutoJob.asp 동일 함수 존재 동시수정요망
function N_TenDlvCode2CommonDlvCode(imallname,itenCode)
    if (LCASE(imallname)="lottecom") or (LCASE(imallname)="lottecomm") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteDlvCode(itenCode)
    elseif (LCASE(imallname)="lotteimall") then
		If (Now() > #09/01/2015 00:00:00#) Then
			N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteiMallNewDlvCode(itenCode)
		Else
			N_TenDlvCode2CommonDlvCode = TenDlvCode2LotteiMallDlvCode(itenCode)
		End If
    elseif (LCASE(imallname)="interpark") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2InterParkDlvCode(itenCode)
    elseif (LCASE(imallname)="cjmall") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2cjMallDlvCode(itenCode)
    elseif (LCASE(imallname)="gseshop") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2GSShopDlvCode(itenCode)
    elseif (LCASE(imallname)="homeplus") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HomeplusDlvCode(itenCode)
    elseif (LCASE(imallname)="ezwel") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2EzwelDlvCode(itenCode)
    elseif (LCASE(imallname)="auction1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2AuctionDlvCode(itenCode)
    elseif (LCASE(imallname)="gmarket1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2GmarketDlvCode(itenCode)
    elseif (LCASE(imallname)="nvstorefarm") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2NvstorefarmDlvCode(itenCode)
    elseif (LCASE(imallname)="11st1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode211stDlvCode(itenCode)
    elseif (LCASE(imallname)="ssg") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2SsgDlvCode(itenCode)
    elseif (LCASE(imallname)="halfclub") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HalfClubDlvCode(itenCode)
    elseif (LCASE(imallname)="coupang") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2CoupangDlvCode(itenCode)
    elseif (LCASE(imallname)="hmall1010") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2HmallDlvCode(itenCode)
    elseif (LCASE(imallname)="wmp") or (LCASE(imallname)="wmpfashion") then
        N_TenDlvCode2CommonDlvCode = TenDlvCode2WMPDlvCode(itenCode)
    elseif (LCASE(imallname)="gsisuper") or (LCASE(imallname)="lfmall") then
        '// gsisuper => 사방넷API 이용
        N_TenDlvCode2CommonDlvCode = TenDlvCode2SabangNetDlvCode(itenCode)
    end if
end function

%>


<script language='javascript'>
function goPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}


function X_sendSongJangCJMALL(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('제휴사 택배사코드 미지정');
        return;
    }

    if ((OutMallSDiv=="99")&&(songjangNo=="")){
        songjangNo="기타"
    }

    if (songjangNo==""){
        alert('송장번호 미지정..');
        return;
    }


    //proc_gubun=sfin:발송완료 //dfin:배송완료

    var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
	//var popwin=window.open('/admin/etc/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
	var popwin=window.open('https://wapi.10x10.co.kr/outmall/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangGSShop(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('제휴사 택배사코드 미지정');
        return;
    }

    if ((OutMallSDiv=="99")&&(songjangNo=="")){
        songjangNo="기타"
    }

    if (songjangNo==""){
        alert('송장번호 미지정..');
        return;
    }

    var params = "ten_ord_no="+tenorderserial+"&ordclmNo="+OutMallOrderSerial+"&ordSeq="+OrgDetailKey+"&delvEntrNo="+OutMallSDiv+"&invoNo="+songjangNo;
	var popwin=window.open('https://wapi.10x10.co.kr/outmall/gsshop/actGSShopSongjangInputProc.asp?' + params,'sendSongJangGSShop','width=600,height=400,scrollbars=yes,resizable=yes');
    //var popwin=window.open('/admin/etc/gsshop/actGSShopSongjangInputProc.asp?' + params,'sendSongJangGSShop','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendRecvStateCJMALL(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName, errTakBae){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "ten_ord_no="+tenorderserial+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&rcv_nm="+receiveName;
	var popwin;

    /*
	if(errTakBae == '2' ){
		popwin=window.open("/admin/etc/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
	    popwin.focus();
	    comp.disabled=true;
	}else{
		if (confirm('경동택배 홈페이지에서 인수확인 하셨나요?')){
			popwin=window.open("/admin/etc/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
		    popwin.focus();
		    comp.disabled=true;
		}
	}
    */
	var popwin=window.open("https://wapi.10x10.co.kr/outmall/cjmall/actCjmallRecvStateInputProc.asp?" + params,"sendRecvStateCJMALL","width=600,height=400,scrollbars=yes,resizable=yes");
    popwin.focus();
	comp.disabled=true;
}

function sendSongJangHomeplus(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
//alert('Wapi로 pop')
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    //var popwin=window.open("/admin/etc/homeplus/actHomeplusSongjangInputProc.asp?" + params,"sendRecvStateHomeplus","width=600,height=400,scrollbars=yes,resizable=yes");
    var popwin=window.open("https://wapi.10x10.co.kr/outmall/proc/Homeplus_SongjangProc.asp?" + params,"sendRecvStateHomeplus","width=600,height=400,scrollbars=yes,resizable=yes"); //방화벽확인필요
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangEzwel(OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
	//var popwin=window.open("/admin/etc/ezwel/actEzwelSongjangInputProc.asp?" + params,"sendRecvStateEzwel","width=600,height=400,scrollbars=yes,resizable=yes");
	var popwin=window.open('https://wapi.10x10.co.kr/outmall/proc/Ezwel_SongjangProc.asp?' + params,'sendRecvStateEzwel','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangAuction(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('https://wapi.10x10.co.kr/outmall/proc/Auction_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}


function sendSongJangGmarket(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }

	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('https://wapi.10x10.co.kr/outmall/proc/Gmarket_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangStorefarm(comp,tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, receiveName){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+receiveName;
    var popwin=window.open('https://wapi.10x10.co.kr/outmall/proc/Nvstorefarm_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangHalfclub(comp, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmallOptionCode, outmallOptionName, itemno, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmallOptionCode+"&outmallOptionName="+outmallOptionName+"&itemno="+itemno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('https://wapi.10x10.co.kr/outmall/proc/halfclub_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangSabangNet(comp, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmallOptionCode, outmallOptionName, itemno, songjangDiv, songjangNo, shoplinkerorderid){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmallOptionCode+"&outmallOptionName="+outmallOptionName+"&itemno="+itemno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&shoplinkerorderid=" + shoplinkerorderid;
    var popwin=window.open('<%=apiURL%>/outmall/proc/sabangnet_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}


function sendSongJangHmall(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, beasongNum, reserve01){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&beasongNum="+beasongNum+"&reserve01="+reserve01;
    var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangHmallCS(prctp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo, beasongNum, reserve01){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&beasongNum="+beasongNum+"&reserve01="+reserve01+"&prctp="+prctp;
    var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangWMP(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/wmp_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangWMPfashion(comp, OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/proc/wmpfashion_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function jsChulgoTargetMake(imallname){
    if (imallname=="ssg"){
        var params = "prctp=999";
        var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }
}

function sendSongJang(comp,OutMallOrderSerial,OrgDetailKey,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('제휴사 택배사코드 미지정');
        return;
    }

    if ((OutMallSDiv=="99")&&(songjangNo=="")){
        songjangNo="기타"
    }

    if (songjangNo==""){
        alert('송장번호 미지정..');
        return;
    }


    //proc_gubun=sfin:발송완료 //dfin:배송완료
//alert('Wapi로 pop')
    var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
    //var popwin=window.open('/admin/etc/lotte/actLotteSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('<%=apiURL%>/outmall/proc/LotteCom_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');


    popwin.focus();
    comp.disabled=true;
}

function sendSongJangiMall(comp,OutMallOrderSerial,OrgDetailKey,sendQnt,sendDate,outmallGoodsID,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert('제휴사 택배사코드 미지정');
        return;
    }

    if (songjangNo==""){
        alert('송장번호 미지정');
        return;
    }

    //proc_gubun=sfin:발송완료 //dfin:배송완료

    var params = "cmdparam=songjangip&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&sendQnt="+sendQnt+"&sendDate="+sendDate+"&outmallGoodsID="+outmallGoodsID+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
    //var popwin=window.open('/admin/etc/lotteimall/actLotteiMallReq.asp?' + params,'xSiteSongjangInputProciMall','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('<%=apiURL%>/outmall/proc/Lotteimall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

// 신규 시스템
function sendSongJangiMallNew(comp,OutMallOrderSerial,OrgDetailKey,sendQnt,sendDate,outmallGoodsID,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정");
        return;
    }

    var params = "mode=sendsongjang&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&sendQnt="+sendQnt+"&sendDate="+sendDate+"&outmallGoodsID="+outmallGoodsID+"&hdc_cd="+OutMallSDiv+"&inv_no="+songjangNo;
	//var popwin=window.open("/admin/etc/orderInput/xSiteCSOrder_lotteimall_Process.asp?" + params,"sendSongJangiMallNew","width=600,height=400,scrollbars=yes,resizable=yes");
	var popwin=window.open('<%=apiURL%>/outmall/proc/Lotteimall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJangInterpark(comp,OutMallOrderSerial,OrgDetailKey,yyyymmdd,OutMallSDiv,songjangNo){
    if (OutMallSDiv==""){
        //alert('제휴사 택배사코드 미지정');
        //return;

        OutMallSDiv="169167";
    }

    if ((OutMallSDiv=="169167")&&(songjangNo=="")){
        songjangNo="기타"
    }

    if (songjangNo==""){
        alert('송장번호 미지정..');
        return;
    }



    var params = "ordclmNo="+OutMallOrderSerial+"&ordSeq="+OrgDetailKey;
    params=params+"&delvDt="+yyyymmdd+"&delvEntrNo="+OutMallSDiv+"&invoNo="+songjangNo;
    params=params+"&optPrdTp=01&optOrdSeqList="+OrgDetailKey
    //alert(params)
    //var popwin=window.open('/admin/etc/interparkXML/actInterparkSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('<%=apiURL%>/outmall/interpark/actInterparkSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function sendSongJang11st(tenorderserial,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, beasongNum){

    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo+"&songjangDiv="+beasongNum;
    var popwin=window.open('<%=apiURL%>/outmall/proc/11st_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
    comp.disabled=true;
}

function pop_mayDelOrder(sitename){
    var popwin=window.open('/admin/etc/orderInput/pop_etcSiteSongjangInput.asp?sitename=' + sitename,'pop_mayDelOrder','width=1500,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function pop_etcSiteCsMatch(sitename, outmallorderserial, outmallorderseq, shppDivDtlNm) {
    var popwin=window.open('/admin/etc/orderInput/pop_etcSiteCsMatch.asp?mode=matchcs&sitename=' + sitename + '&outmallorderserial=' + outmallorderserial + '&outmallorderseq=' + outmallorderseq + '&shppDivDtlNm=' + shppDivDtlNm,'pop_etcSiteCsMatch','width=300,height=100,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//======================================================================================================
function receiveXSiteMifinList(isitename){
    var params = ""
    if (isitename=="ssg"){
        var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="ezwel"){
        //var popwin=window.open('<%=apiURL%>/outmall/ezwel/ezwel_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        var popwin=window.open('/admin/etc/ezwel/ezwel_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="hmall1010"){
        var popwin=window.open('<%=apiURL%>/outmall/hmall/hmall_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="coupang"){
        var popwin=window.open('<%=apiURL%>/outmall/coupang/coupang_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="wmp"){
        var popwin=window.open('<%=apiURL%>/outmall/wmp/wmp_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="wmpfashion"){
        var popwin=window.open('<%=apiURL%>/outmall/wmpfashion/wmpfashion_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="cjmall"){
        var popwin=window.open('<%=apiURL%>/outmall/cjmall/cjmall_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="interpark"){
        var popwin=window.open('<%=apiURL%>/outmall/interpark/interpark_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else if(isitename=="11st1010"){
        var popwin=window.open('<%=apiURL%>/outmall/11st/11st_getMichulgoList.asp?' + params,'XSiteGetMichulgoList','width=600,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
    }else{
        alert(isitename + " 정의되지 않았습니다.");
    }

}

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sitename%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function ssgDlvFinishSend(outmallorderserial,tenorderserial,tenitemid,tenitemoption,dlvfinishdt){
	var params = "prctp=3&outmallorderserial="+outmallorderserial+"&tenorderserial="+tenorderserial+"&tenitemid="+tenitemid+"&tenitemoption="+tenitemoption+"&dlvfinishdt="+dlvfinishdt
 	var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function ssgDlvFinishSendCS(outmallorderserial,tenorderserial,tenitemid,tenitemoption,dlvfinishdt){
	var params = "prctp=33&outmallorderserial="+outmallorderserial+"&tenorderserial="+tenorderserial+"&tenitemid="+tenitemid+"&tenitemoption="+tenitemoption+"&dlvfinishdt="+dlvfinishdt
 	var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();

}

function exwelDlvFinishSend(outmallorderserial,outmallorderseq){
	var params = "prctp=3&ord_no="+outmallorderserial+"&ord_dtl_sn="+outmallorderseq;
 	//var popwin=window.open('<%=apiURL%>/outmall/proc/Ezwel_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    var popwin=window.open('/admin/etc/ezwel/Ezwel_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function hmallDlvFinishSend(outmallorderserial,outmallorderseq,shppNo,shppSeq,delicoVenId,wblNo){
	var params = "prctp=3&OutMallOrderSerial="+outmallorderserial+"&OrgDetailKey="+outmallorderseq+"&hdc_cd="+delicoVenId+"&songjangNo="+wblNo+"&beasongNum="+shppNo+"&reserve01="+shppSeq;
 	var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function hmallDlvFinishSendCS(outmallorderserial,outmallorderseq,shppNo,shppSeq,delicoVenId,wblNo){
	var params = "prctp=33&OutMallOrderSerial="+outmallorderserial+"&OrgDetailKey="+outmallorderseq+"&hdc_cd="+delicoVenId+"&songjangNo="+wblNo+"&beasongNum="+shppNo+"&reserve01="+shppSeq;
 	var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function hmallCancelFinishSend(outmallorderserial,outmallorderseq,shppNo,shppSeq,delicoVenId,wblNo){
	var params = "prctp=7&OutMallOrderSerial="+outmallorderserial+"&OrgDetailKey="+outmallorderseq+"&hdc_cd="+delicoVenId+"&songjangNo="+wblNo+"&beasongNum="+shppNo+"&reserve01="+shppSeq;
 	var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function cjmallDlvFinishSend(outmallorderseq){
    var params = "ord_dtl_sn="+outmallorderseq+"&rcv_nm=.";
 	var popwin=window.open('<%=apiURL%>/outmall/cjmall/actCjmallRecvStateInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function sendSongJangSSG(ichultype,OutMallOrderSerial,OrgDetailKey,songjangDiv,songjangNo, shppNo, shppSeq, itemno, dlvfinishdt){
    if (ichultype != "6") {
        if (songjangDiv==""){
            alert("제휴사 택배사코드 미지정");
            // return;
        }

        if ((songjangNo == "") && (songjangDiv != "98")) {
            alert("송장번호 미지정!!\n\n송장번호를 모르는 경우, 택배사를 퀵서비스로 지정하세요.");
            return;
        }
    }

    var params = "shppNo="+shppNo+"&shppSeq="+shppSeq+"&delicoVenId="+songjangDiv+"&wblno="+songjangNo+"&itemno="+itemno+"&outmallorderserial="+OutMallOrderSerial+"&orgdetailKey="+OrgDetailKey+"&dlvfinishdt="+dlvfinishdt;

    if ((ichultype=="2") || (ichultype=="6")) {
        params = params + "&prctp=" + ichultype;
    }

   //alert(params);
   //return;
    var popwin=window.open('<%=apiURL%>/outmall/ssg/ssg_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function sendOrderConfirmHMall(outmallorderserial,outmallorderseq,shppNo,shppSeq,delicoVenId,wblNo) {
    var params = "prctp=6&OutMallOrderSerial="+outmallorderserial+"&OrgDetailKey="+outmallorderseq+"&hdc_cd="+delicoVenId+"&songjangNo="+wblNo+"&beasongNum="+shppNo+"&reserve01="+shppSeq;
 	var popwin=window.open('<%=apiURL%>/outmall/proc/hmall_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function sendSongJangCoupang(ichultype, OutMallOrderSerial, OrgDetailKey, outmallGoodNo, outmalloptionno, songjangDiv, songjangNo, beasongNum, splitrequire){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "OutMallOrderSerial="+OutMallOrderSerial+"&OrgDetailKey="+OrgDetailKey+"&outmallGoodNo="+outmallGoodNo+"&outmallOptionCode="+outmalloptionno+"&hdc_cd="+songjangDiv+"&songjangNo="+songjangNo+"&beasongNum="+beasongNum+"&splitrequire="+splitrequire;
    var popwin=window.open('<%=apiURL%>/outmall/proc/coupang_SongjangProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function sendSongJangCjmall(ten_ord_no,OutMallOrderSerial, OrgDetailKey, songjangDiv, songjangNo){
    if (songjangDiv==""){
        alert("제휴사 택배사코드 미지정");
        return;
    }

    if (songjangNo==""){
        alert("송장번호 미지정..");
        return;
    }
	var params = "ten_ord_no="+ten_ord_no+"&ord_no="+OutMallOrderSerial+"&ord_dtl_sn="+OrgDetailKey+"&hdc_cd="+songjangDiv+"&inv_no="+songjangNo;
    var popwin=window.open('<%=apiURL%>/outmall/cjmall/actCJmallSongjangInputProc.asp?' + params,'xSiteSongjangInputProc','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}




</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- 검색 시작 -->


<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">

	<tr align="center">
		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br" rowspan="2">검색<br>조건</td>
		<td align="left" class="td_br">
		    제휴몰 선택 :
            <select name="sitename" class="select">
            <option value="">선택
            <option value="interpark"  <%=CHKIIF(sitename="interpark","selected","") %> >인터파크
            <option value="11st1010"  <%=CHKIIF(sitename="11st1010","selected","") %> >11번가
            <option value="ezwel"  <%=CHKIIF(sitename="ezwel","selected","") %> >이지웰
            <option value="ssg"  <%=CHKIIF(sitename="ssg","selected","") %> >SSG
            <option value="hmall1010"  <%=CHKIIF(sitename="hmall1010","selected","") %> >Hmall
            <option value="coupang"  <%=CHKIIF(sitename="coupang","selected","") %> >Coupang
            <option value="wmp"  <%=CHKIIF(sitename="wmp","selected","") %> >WMP
            <option value="wmpfashion"  <%=CHKIIF(sitename="wmpfashion","selected","") %> >WMPW패션
            <option value="cjmall"  <%=CHKIIF(sitename="cjmall","selected","") %> >CJmall
            </select>


		    &nbsp;&nbsp;
		    검색구분 :
			<select class="select" name="searchType">
	     	<option value="0" <%= chkIIF(searchType="0","selected","") %> >전체</option>
	     	<option value="1" <%= chkIIF(searchType="1","selected","") %> >검토필요</option>
	     	</select>

            &nbsp;&nbsp;
            출고구분 :
            <input type="radio" name="shppDivDtl" value="N" <%=CHKIIF(shppDivDtl="N","checked","")%> >일반출고
            <input type="radio" name="shppDivDtl" value="" <%=CHKIIF(shppDivDtl="","checked","")%> >전체
            <input type="radio" name="shppDivDtl" value="E" <%=CHKIIF(shppDivDtl="E","checked","")%> >교환/반품등

		</td>

		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br" rowspan="2">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center">
		<td align="left" class="td_br">
		    주문번호 :
            <input type="text" class="text" name="Matchorderserial" value="<%= Matchorderserial %>">
            &nbsp;&nbsp;
            제휴 주문번호 :
            <input type="text" class="text" name="OutMallOrderSerial" value="<%= OutMallOrderSerial %>">
            &nbsp;&nbsp;
            <input type="checkbox" name="excNoOrderSerial" value="Y" <%= CHKIIF(excNoOrderSerial="Y", "checked", "") %>> 주문번호없음 제외
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr height="16">
		<td align="left">
			검색결과 : <b><%= FormatNumber(oMiFin.FTotalCount,0) %></b>
            &nbsp;
            페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMiFin.FTotalPage,0) %></b>
            &nbsp;
            최종업데이트 : <%= oMiFin.getLastUpDt %>
		</td>
        <td align="right">
            <% if (sitename<>"") then %>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="<%=sitename%> 미완료목록 재수신" onClick="receiveXSiteMifinList('<%=sitename%>')">
            <% end if %>
        </td>

	</tr>
</table>
<!-- 액션 끝 -->
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl" >
	<tr align="center" class="tr_tablebar">

	    <td width="60" class="td_br">제휴사</td>
	    <td width="70" class="td_br">TEN 매핑<br>주문번호</td>
    	<td width="120" class="td_br">제휴주문번호<br />(원주문번호)</td>
      	<td width="60" class="td_br">제휴<br>상품코드</td>
      	<td width="50" class="td_br">TEN 매핑<br>상품코드</td>
      	<td width="50" class="td_br">TEN 매핑<br>옵션코드</td>
      	<td  class="td_br">상품명 <font color="blue">[옵션명]</font></td>
        <td width="30" class="td_br">수량<br>(TEN/제휴)</td>
        <td width="100" class="td_br">출고구분</td>
        <td width="100" class="td_br">출고일<br />(완료일)</td>
        <td width="100" class="td_br">배송완료일</td>
        <td width="80" class="td_br">정산확정일</td>
        <td width="40" class="td_br">배송<br>구분</td>
      	<td width="100" class="td_br">자사입력송장</td>
      	<!-- td width="80" class="td_br">제휴사<br>택배코드</td -->
        <td width="100" class="td_br">제휴입력송장</td>
      	<td width="80" class="td_br">제휴상태</td>
        <td width="80" class="td_br">TEN상태</td>
        <td width="80" class="td_br">CsId</td>
      	<td width="80" class="td_br">비고</td>
    </tr>
<% if oMiFin.FResultCount>0 then %>
<% for i=0 to oMiFin.FResultcount-1 %>
<%
if isNULL(oMiFin.FItemList(i).Fsongjangdiv) then
    iOutMallDlvCode = ""
else
    iOutMallDlvCode = N_TenDlvCode2CommonDlvCode(oMiFin.FItemList(i).FSellSite,oMiFin.FItemList(i).Fsongjangdiv)
end if
%>
<tr>
    <td class="td_br"><%= oMiFin.FItemList(i).FSellSite %></td>
    <td class="td_br"><a href="#" onClick="popDeliveryTrackingSummaryOne('<%=oMiFin.FItemList(i).FMatchorderserial %>','<%=oMiFin.FItemList(i).Fsongjangno %>','<%=oMiFin.FItemList(i).Fsongjangdiv %>');return false;"><%= oMiFin.FItemList(i).FMatchorderserial %></a></td>
    <td class="td_br">
        <a href="#" onClick="popByExtorderserial('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>');return false;"><%= oMiFin.FItemList(i).FOutMallOrderSerial %></a>
        <% if Not IsNull(oMiFin.FItemList(i).FOrgOutMallOrderSerial) then %>
        <br />(<a href="#" onClick="popByExtorderserial('<%= oMiFin.FItemList(i).FOrgOutMallOrderSerial %>');return false;"><%= oMiFin.FItemList(i).FOrgOutMallOrderSerial %></a>)
        <% end if %>
    </td>
    <td class="td_br"><%= oMiFin.FItemList(i).FoutmallGoodsNo %></td>
    <td class="td_br"><%= oMiFin.FItemList(i).FMatchitemid %></td>
    <td class="td_br"><%= oMiFin.FItemList(i).FMatchitemoption %></td>
    <td class="td_br">
        <%= oMiFin.FItemList(i).Fitemname %>
        <% if oMiFin.FItemList(i).Fitemoptionname<>"" then %>
        <font color="blue">[<%= oMiFin.FItemList(i).Fitemoptionname %>]</font>
        <% end if %>
    </td>
    <td class="td_br" align="center"><%= oMiFin.FItemList(i).Fitemno %>  / <%= oMiFin.FItemList(i).FordQty %></td>
    <td class="td_br" align="center"><%= oMiFin.FItemList(i).getshppDivDtlNm %></td>
    <td class="td_br" align="center">
    <% if IsDate(oMiFin.FItemList(i).Fbeasongdate) then %>
        <% if CDate(oMiFin.FItemList(i).Fbeasongdate)<=dateadd("d",-14,now()) then %>
        <strong><%= oMiFin.FItemList(i).Fbeasongdate %></strong>
        <% else %>
        <%= oMiFin.FItemList(i).Fbeasongdate %>
        <% end if %>
    <% else %>
        <%= oMiFin.FItemList(i).Fbeasongdate %>
    <% end if %>
    </td>
    <td class="td_br" align="center">
        <% if NOT isNULL(oMiFin.FItemList(i).Fdlvfinishdt) then %>
            <% if isDate(LEFT(oMiFin.FItemList(i).Fdlvfinishdt,10)) then %>
                <% if datediff("d",LEFT(oMiFin.FItemList(i).Fdlvfinishdt,10),now())>1 then %>
                    <strong><%= oMiFin.FItemList(i).Fdlvfinishdt %></strong>
                <% else %>
                    <%= oMiFin.FItemList(i).Fdlvfinishdt %>
                <% end if %>
            <% else %>
                <%= oMiFin.FItemList(i).Fdlvfinishdt %>
            <% end if %>
        <% end if %>
    </td>
    <td class="td_br" align="center"><%= oMiFin.FItemList(i).FjungsanFixDate %></td>
    <td class="td_br" align="center">
        <% IF oMiFin.FItemList(i).Fisupchebeasong="Y" THEN %>
        <font color="blue">업체</font>
        <% End IF %>
    </td>
    <td class="td_br"><%= oMiFin.FItemList(i).Fdivname %><br><%= oMiFin.FItemList(i).Fsongjangno %></td>
    <!-- td class="td_br" >
        <%= iOutMallDlvCode %>
        <% if NOT isNULL(oMiFin.FItemList(i).FsongjangDiv) then %>
        (<%=oMiFin.FItemList(i).FsongjangDiv%>)
        <% end if %>
    </td -->
    <td class="td_br" align="center">
		<%= oMiFin.FItemList(i).getOutDlvInputedStr %>
    </td>
    <td class="td_br" align="center">
		<%= oMiFin.FItemList(i).getOutorderStatusNm %>
    </td>
    <td class="td_br" align="center">
        <%= oMiFin.FItemList(i).getTenStatusNm %>
    </td>
    <td class="td_br" align="center">
        <%= oMiFin.FItemList(i).Fasid %>
    </td>
    <td class="td_br">
    <%
    ' if (oMiFin.FItemList(i).isStatusSendReqConfirm) then ''주문확인
    '     Select Case sitename
    '         Case "coupang"

    '         Case Else
    '     End Select
    ' end if

    if (oMiFin.FItemList(i).isStatusSendReqSongjang) then
        Select Case sitename
            Case "ezwel"
        %>
            <input type="button" value="전송" onClick="sendSongJangEzwel('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>');return false;">
        <%
			Case "ssg"
		%>
				<input type="button" value="송장입력" onClick="sendSongJangSSG(1,'<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>','<%= oMiFin.FItemList(i).FshppNo %>','<%= oMiFin.FItemList(i).FshppSeq %>', '<%= oMiFin.FItemList(i).Fitemno %>', '<%= LEFT(oMiFin.FItemList(i).Fdlvfinishdt,10) %>');return false;">
		<%
			Case "coupang"
		%>
				<input type="button" value="송장입력" onClick="sendSongJangCoupang(1,'<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%=oMiFin.FItemList(i).FoutmallGoodsNo%>','<%=oMiFin.FItemList(i).FoutmalloptionNo%>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>','<%= oMiFin.FItemList(i).FshppNo %>','N');return false;">
		<%
			Case "cjmall"
		%>
				<input type="button" value="송장입력" onClick="sendSongJangCjmall('<%=oMiFin.FItemList(i).FMatchorderserial %>_<%= oMiFin.FItemList(i).FMatchitemid %>_<%= oMiFin.FItemList(i).FMatchitemoption %>','<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>');return false;">
		<%
			Case "11st1010"
		%>
				<input type="button" value="송장입력" onClick="sendSongJang11st('<%=oMiFin.FItemList(i).FMatchorderserial %>','<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>','<%= oMiFin.FItemList(i).FshppNo %>');return false;">
		<%
            Case Else

        End Select
    end if
    %>

    <%
    if (oMiFin.FItemList(i).isStatusSendReqChulgo) then
        Select Case sitename
			Case "ssg"
		%>
                <input type="button" value="출고전송" onClick="sendSongJangSSG(2,'<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>','<%= oMiFin.FItemList(i).FshppNo %>','<%= oMiFin.FItemList(i).FshppSeq %>', '<%= oMiFin.FItemList(i).Fitemno %>', '<%= LEFT(oMiFin.FItemList(i).Fdlvfinishdt,10) %>');return false;">
		<%
			Case "hmall1010"
                if (oMiFin.FItemList(i).FshppDivDtlNm="교환출고") then
		%>
                <input type="button" value="출고전송" onClick="sendSongJangHmallCS(22,'<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>','<%= oMiFin.FItemList(i).FshppNo %>','<%= oMiFin.FItemList(i).FshppSeq %>', '<%= oMiFin.FItemList(i).Fitemno %>', '<%= LEFT(oMiFin.FItemList(i).Fdlvfinishdt,10) %>');return false;">
		<%
                end if
            Case Else

        End Select
    end if
    %>

    <%
    if (oMiFin.FItemList(i).isStatusSendReqOrderConfirm) then
        Select Case sitename
			Case "ssg"
		%>
                <input type="button" value="주문확인" onClick="sendSongJangSSG(6,'<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%= oMiFin.FItemList(i).FOrgDetailKey %>','<%= iOutMallDlvCode %>','<%= oMiFin.FItemList(i).FsongjangNo %>','<%= oMiFin.FItemList(i).FshppNo %>','<%= oMiFin.FItemList(i).FshppSeq %>', '<%= oMiFin.FItemList(i).Fitemno %>', '<%= LEFT(oMiFin.FItemList(i).Fdlvfinishdt,10) %>');return false;">
		<%
            Case "hmall1010"
		%>
                <input type="button" value="주문확인" onClick="sendOrderConfirmHMall('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FOrgDetailKey%>','<%=oMiFin.FItemList(i).FshppNo %>','<%=oMiFin.FItemList(i).FshppSeq %>','<%=oMiFin.FItemList(i).FdelicoVenId %>','<%=oMiFin.FItemList(i).FwblNo %>');return false;">
		<%
            Case Else

        End Select
    end if
    %>

    <%
    if (oMiFin.FItemList(i).isStatusSendDliverFinish) then
        Select Case sitename
			Case "ssg"
                if (oMiFin.FItemList(i).FshppDivDtlNm="일반출고") then
		%>
                <input type="button" value="배송완료전송" onClick="ssgDlvFinishSend('<%= CHKIIF(Not IsNull(oMiFin.FItemList(i).FOrgOutMallOrderSerial), oMiFin.FItemList(i).FOrgOutMallOrderSerial, oMiFin.FItemList(i).FOutMallOrderSerial) %>','<%=oMiFin.FItemList(i).FMatchorderserial%>','<%=oMiFin.FItemList(i).FMatchitemid%>','<%=oMiFin.FItemList(i).FMatchitemoption%>','<%= oMiFin.FItemList(i).Fdlvfinishdt %>');return false;">
		<%
                else
		%>
                <input type="button" value="배송완료전송" onClick="ssgDlvFinishSendCS('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FMatchorderserial%>','<%=oMiFin.FItemList(i).FMatchitemid%>','<%=oMiFin.FItemList(i).FMatchitemoption%>','<%= oMiFin.FItemList(i).Fdlvfinishdt %>');return false;">
		<%
                end if
            Case "ezwel"
		%>
                <input type="button" value="배송완료전송" onClick="exwelDlvFinishSend('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FOrgDetailKey%>');return false;">
		<%
        	Case "hmall1010"
                if (oMiFin.FItemList(i).FshppDivDtlNm="교환출고") then
		%>
                <input type="button" value="배송완료전송" onClick="hmallDlvFinishSendCS('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FOrgDetailKey%>','<%=oMiFin.FItemList(i).FshppNo %>','<%=oMiFin.FItemList(i).FshppSeq %>','<%= iOutMallDlvCode %>','<%=oMiFin.FItemList(i).FsongjangNo %>');return false;">
		<%
                else
		%>
                <input type="button" value="배송완료전송" onClick="hmallDlvFinishSend('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FOrgDetailKey%>','<%=oMiFin.FItemList(i).FshppNo %>','<%=oMiFin.FItemList(i).FshppSeq %>','<%=oMiFin.FItemList(i).FdelicoVenId %>','<%=oMiFin.FItemList(i).FwblNo %>');return false;">
		<%
                end if
            Case "cjmall"
		%>
                <input type="button" value="인수전송" onClick="cjmallDlvFinishSend('<%=oMiFin.FItemList(i).FOrgDetailKey%>');return false;">
		<%
            Case Else

        End Select
    end if
    %>

    <%
    if (oMiFin.FItemList(i).isStatusSendCancelFinish) then
        Select Case sitename
            Case "hmall1010"
		%>
                <input type="button" value="주문확인취소" onClick="hmallCancelFinishSend('<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FOrgDetailKey%>','<%=oMiFin.FItemList(i).FshppNo %>','<%=oMiFin.FItemList(i).FshppSeq %>','<%=oMiFin.FItemList(i).FdelicoVenId %>','<%=oMiFin.FItemList(i).FwblNo %>');return false;">
		<%
            Case Else

        End Select
    end if
    %>

    <%
    if (oMiFin.FItemList(i).FshppDivDtlNm="교환출고") and IsNull(oMiFin.FItemList(i).Fasid) then
        %>
        		<input type="button" value="매칭" onClick="pop_etcSiteCsMatch('<%= oMiFin.FItemList(i).FSellSite %>','<%= oMiFin.FItemList(i).FOutMallOrderSerial %>','<%=oMiFin.FItemList(i).FOrgDetailKey%>','<%=oMiFin.FItemList(i).FshppDivDtlNm %>');return false;">
        <%
    end if
    %>

    <% if (FALSE) Then %>
		<!-- 송장전송 -->
		<%
		Select Case ArrList(18,i)
			Case "lotteCom", "lotteComM"
		%>
				<input type="button" value="전송" onClick="sendSongJang(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
			Case "lotteimall"
				if InStr(ArrList(4,i),"-")>0 then
		%>
					<input type="button" value="전송OLD" onClick="sendSongJangiMall(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(8,i) %>','<%= replace(Left(ArrList(15,i),10),"-","") %>','<%= ArrList(19,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
				else
		%>
					<input type="button" value="전송" onClick="sendSongJangiMallNew(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(8,i) %>','<%= replace(Left(ArrList(15,i),10),"-","") %>','<%= ArrList(19,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
				end if
			Case "interpark"
				%>
				<input type="button" value="전송" onClick="sendSongJangInterpark(this,'<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= replace(Left(ArrList(15,i),10),"-","") %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
			Case "cjmall"
			ArrList(10,i) = Replace(ArrList(10,i), "&nbsp;", Chr(32))
			ArrList(10,i) = trim(ArrList(10,i))
		%>
				<input type="button" value="전송" onClick="sendSongJangCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
			Case "gseshop"
		%>
				<input type="button" value="전송" onClick="sendSongJangGSShop(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
			Case "homeplus"
		%>
				<input type="button" value="전송" onClick="sendSongJangHomeplus(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
			Case "ezwel"
		%>
				<input type="button" value="전송" onClick="sendSongJangEzwel(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>');">
		<%
			Case "auction1010"
		%>
				<input type="button" value="전송" onClick="sendSongJangAuction(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case "gmarket1010"
		%>
				<input type="button" value="전송" onClick="sendSongJangGmarket(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">

		<%
			Case "nvstorefarm"
		%>
				<input type="button" value="전송" onClick="sendSongJangStorefarm(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>', '<%=ArrList(9,i)%>');">
		<%
			Case "11st1010"
		%>
				<input type="button" value="전송" onClick="sendSongJang11st(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>', '<%=ArrList(23,i)%>');">
		<%
			Case "ssg"
		%>
				<input type="button" value="전송" onClick="sendSongJangSSG(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>', '<%=ArrList(23,i)%>', '<%=ArrList(24,i)%>', '<%=ArrList(8,i)%>', '<%=ArrList(31,i)%>');">
		<%
			Case "halfclub"
		%>
				<input type="button" value="전송" onClick="sendSongJangHalfclub(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(25,i) %>', '<%= ArrList(26,i) %>', '<%= ArrList(30,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case "gsisuper", "LFmall"
                '// 사방넷API 로 전송
		%>
				<input type="button" value="전송" onClick="sendSongJangSabangNet(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(25,i) %>', '<%= ArrList(26,i) %>', '<%= ArrList(8,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(27,i) %>');" >
		<%
			Case "coupang"
		%>
				<input type="button" value="전송" onClick="sendSongJangCoupang(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= ArrList(19,i) %>', '<%= ArrList(28,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(23,i) %>', '<%= ArrList(29,i) %>');" >
		<%
			Case "hmall1010"
		%>
				<input type="button" value="전송" onClick="sendSongJangHmall(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>', '<%= ArrList(23,i) %>', '<%= ArrList(24,i) %>');" >
		<%
			Case "WMP"
		%>
				<input type="button" value="전송" onClick="sendSongJangWMP(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case "wmpfashion"
		%>
				<input type="button" value="전송" onClick="sendSongJangWMPfashion(this,'<%= ArrList(1,i) %>', '<%= ArrList(4,i) %>', '<%= iOutMallDlvCode %>', '<%= ArrList(10,i) %>');" >
		<%
			Case Else
				response.write "ERR"
		End Select
		%>
	<% elseif ((searchType = "sendRecvState")) then %>
		<!-- 고객인수 확인 및 전송 -->
        <%
        	if ArrList(18,i)="cjmall" then
        		If (FALSE) and ArrList(9,i) = "21" Then		'경동택배//홈페이지 엄청느림..다른 방법으로 해야 될 듯..2015-02-27 19:04 김진영
		%>
		<input type="button" class="button_s" value="경동확인" onClick="sendRecvStateCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(9,i) %>','<%= ArrList(10,i) %>','<%= ArrList(22,i) %>', '1');">
		<%
        		Else
        %>
        <input type="button" value="인수전송" onClick="sendRecvStateCJMALL(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= ArrList(9,i) %>','<%= ArrList(10,i) %>','.', '2');">
        <%
        		End If
        	end if
        %>
    <% elseif ((searchType = "sendChulgo")) then %>
        <%
        	if ArrList(18,i)="ssg" then
		%>
				<input type="button" value="출고처리" onClick="sendSongJangSSG(this,'<%= ArrList(0,i) %>_<%= ArrList(20,i) %>','<%= ArrList(1,i) %>','<%= ArrList(4,i) %>','<%= iOutMallDlvCode %>','<%= CHKIIF(ArrList(9,i)="98" and ArrList(10,i)="","퀵서비스",ArrList(10,i)) %>', '<%=ArrList(23,i)%>', '<%=ArrList(24,i)%>', '<%=ArrList(8,i)%>', '<%=ArrList(31,i)%>');">
		<%
        	end if
        %>
	<% end if %>
    </td>

</tr>
<% next %>
<tr>
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oMiFin.HasPreScroll then %>
        <a href="javascript:goPage('<%= oMiFin.StartScrollPage-1 %>');">[pre]</a>
        <% else %>
            [pre]
        <% end if %>

        <% for i=0 + oMiFin.StartScrollPage to oMiFin.FScrollCount + oMiFin.StartScrollPage - 1 %>
            <% if i>oMiFin.FTotalpage then Exit for %>
            <% if CStr(page)=CStr(i) then %>
            <font color="red">[<%= i %>]</font>
            <% else %>
            <a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
            <% end if %>
        <% next %>

        <% if oMiFin.HasNextScroll then %>
            <a href="javascript:goPage('<%= i %>');">[next]</a>
        <% else %>
            [next]
        <% end if %>
    </td>
</tr>
<% ELSE %>
<tr>
    <td colspan="19" align="center">
    <% if sitename="" then %>
    [제휴몰을 선택하세요.]
    <% else %>
    [검색 결과가 없습니다.]
    <% end if %>
    </td>
</tr>
<% end if %>
</table>
<% if (sitename="") then %>
<script>//alert('쇼핑몰을 선택 하세요.');</script>
<% end if %>
<%
SET oMiFin = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

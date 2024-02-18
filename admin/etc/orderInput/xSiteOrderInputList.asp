<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim sellsite, matchState, research, page, orderserial, outmallorderserial, notilog, isNotiList, csViewYn, apiOrder, overseaViewYn
Dim i, pOrderSerial, isNewOrderLine
sellsite			= requestCheckvar(request("sellsite"),32)
matchState			= requestCheckvar(request("matchState"),10)
csViewYn			= requestCheckvar(request("csViewYn"),10)
overseaViewYn		= requestCheckvar(request("overseaViewYn"),10)
research			= requestCheckvar(request("research"),10)
page				= requestCheckvar(request("page"),10)
orderserial			= requestCheckvar(request("orderserial"),20)
outmallorderserial	= requestCheckvar(request("outmallorderserial"),30)
notilog				= requestCheckvar(request("notilog"),10)
isNotiList			=(notilog = "on")
Dim tplinc : tplinc = requestCheckvar(request("tplinc"),10)


Dim regyyyymmdd : regyyyymmdd = requestCheckvar(request("regyyyymmdd"),10)

If (research="") then
	matchState = "I"
	csViewYn = "Y"
	overseaViewYn = "Y"
	tplinc="0"

	if (session("ssAdminPsn")="17") then ''관계사인경우
		tplinc="1"
	end if
End If
If (page = "") Then page = 1
Dim optLeft2FF, otmpOrder, kakaoGiftOptNmDiff, shopifyPriceDiff
Set otmpOrder = new CxSiteTempOrder
	'2018-08-31 17:32 김진영 하단 pagesize 10 -> 50으로 수정
	otmpOrder.FPageSize					= 50					'배열입력의 문제로 페이지 사이즈 제한 있음(CallDBSendRequestModifyOnlineSellAfterMulti 참조) ?/
	otmpOrder.FCurrPage					= page
	''otmpOrder.FRectCompanyID			= CCOMPID
	otmpOrder.FRectSellSite				= sellsite
	otmpOrder.FRectMatchState			= matchState
	otmpOrder.FRectCsViewYn				= csViewYn
	otmpOrder.FRectOverseaViewYn		= overseaViewYn
	otmpOrder.FRectorderserial			= orderserial
	otmpOrder.FRectoutmallorderserial	= outmallorderserial
	otmpOrder.FRectregYYYYMMDD 			= regyyyymmdd
	otmpOrder.FRectInc3pl				= tplinc
	If (isNotiList) Then
		otmpOrder.getOrderNotiLogList
	Else
		otmpOrder.getOnlineTmpOrderList(true)
	End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script language='javascript'>
function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function xlOnlineOrderUpload(){
	var winFile = window.open("/admin/etc/orderInput/popRegFile.asp","popFile","width=600, height=600 ,scrollbars=yes,resizable=yes");
	winFile.focus();
}

function popMatchItem(){
	alert('Not Using');
	return;

	var params = "";
	var popWin = window.open("/company/partnercompany/partneritemlink_modify.asp" + params,"popitemLink","width=800, height=600 ,scrollbars=yes,resizable=yes");
	popWin.focus();
}

function chkThis(comp){
	AnCheckClick(comp);
}

function chkValidAll(){
	var frm = document.frmArr;
}

// ============================================================================
function CheckProduct(o) {
	var frm;
	if (o.checked) {
		hL(o);
	} else {
		dL(o);
	}
}

function CheckTop(o) {
	var frm;

	if (o.checked) {
		SelectAll();
	} else {
		DeselectAll();
	}
}

function DeselectAll() {
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.chk.disabled == false) {
				frm.chk.checked = false;
				CheckProduct(frm.chk);
			}
		}
	}
}

function SelectAll() {
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.chk.disabled == false) {
				frm.chk.checked = true;
				CheckProduct(frm.chk);
			}
		}
	}
}
// ============================================================================

function fnCheckValidAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        if (comp.disabled==false){
            comp.checked = bool;
            AnCheckClick(comp);
        }
    }else{
        for (var i=0;i<comp.length;i++){
            if (comp[i].disabled==false){
                comp[i].checked = bool;
                AnCheckClick(comp[i]);
            }
        }
    }
}

function xlOnlineOrderLotteiMall(){
    var frm = document.frmSvArr;
    frm.mode.value="ltimallreg";
    frm.submit();
}

function xlOnlineOrderCjMall(){
    var frm = document.frmTmp;
    frm.cmdparam.value="cjmallOrdreg";
    frm.submit();
}

function xlOnlineOrderUpCjMall(){
    var frm = document.frmTmp;
    frm.cmdparam.value="cjmallOrdUp";
    frm.submit();
}
function excelSongjang(v){
    var frm = document.frmXl;
    frm.mallid.value= v;
    frm.submit();
}

// 제휴몰 주문 가져오기
function xSiteOrderInput(sellsite) {
    var frm = document.frmXSiteOrder;

	frm.mode.value = "getxsiteorderlist";
	frm.sellsite.value = sellsite;

	if (sellsite=="lotteimall") {
		frm.action = "xSiteOrder_lotteimall_Process.asp";
	//} else if(sellsite=="lotteCom") {
	//	alert('잠시중지.');
	//	return;
	//	frm.action = "xSiteOrder_lotteCom_Process.asp";
	//} else if(sellsite=="interpark") {
	//    frm.action = "xSiteOrder_interpark_Process.asp";
	}else if(sellsite=="gmarket1010"){
		frm.action = "<%=apiURL%>/outmall/gmarket/xSiteOrder_gmarket_Process.asp"
	}else if(sellsite=="11st1010"){
		frm.action = "<%=apiURL%>/outmall/11st/xSiteOrder_11st1010_Process.asp"
	}else if(sellsite=="interpark"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=interpark"
	}else if(sellsite=="auction1010"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=auction1010"
	}else if(sellsite=="nvstorefarm"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=nvstorefarm"
	}else if(sellsite=="nvstorefarmclass"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=nvstorefarmclass"
	}else if(sellsite=="nvstoremoonbangu"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=nvstoremoonbangu"
	}else if(sellsite=="Mylittlewhoopee"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=Mylittlewhoopee"
	}else if(sellsite=="nvstoregift"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=nvstoregift"
	}else if(sellsite=="ezwel"){
		frm.action = "/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=ezwel"
	}else if(sellsite=="kakaostore"){
		frm.action = "/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=kakaostore&gubunCode=ShippingRequest"
	}else if(sellsite=="boribori1010"){
		frm.action = "/admin/etc/order/xSiteOrder_Ins_Process.asp?sellsite=boribori1010&gubunCode=c"
	}else if(sellsite=="lotteCom"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=lotteCom"
	}else if(sellsite=="gseshop"){
		////alert("주문입력에 시간이 많이 걸립니다.(최대 3분)\n\n에러가 발생하면 한번더 주문입력하시기 바랍니다.");
		////frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=gseshop"

		//var popwin = window.open("popGsShopReceveCall.asp","popGsShopReceveCall","width=800,height=800,scrollbars=yes,resizable=yes");
		//popwin.focus();
		//API에 배송비가 없음..
		alert("XL로 입력해 주세요~~");
		return;
	}else if(sellsite=="ssg"){
		frm.action = "<%=apiURL%>/outmall/ssg/xSiteOrder_ssg_Process.asp"
	}else if(sellsite=="halfclub"){
		frm.action = "<%=apiURL%>/outmall/halfclub/xSiteOrder_halfclub_Process.asp"
	}else if(sellsite=="sabangnet"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=sabangnet"
	}else if(sellsite=="coupang"){
		frm.action = "<%=apiURL%>/outmall/coupang/xSiteOrder_coupang_Process.asp"
	}else if(sellsite=="hmall1010"){
		frm.action = "<%=apiURL%>/outmall/hmall/xSiteOrder_hmall_new_Process.asp"
	}else if(sellsite=="WMP"){
		frm.action = "<%=apiURL%>/outmall/wmp/xSiteOrder_wmp_Process.asp"
	}else if(sellsite=="wmpfashion"){
		frm.action = "<%=apiURL%>/outmall/wmpfashion/xSiteOrder_wmpfashion_Process.asp"
	}else if(sellsite=="LFmall"){
		frm.action = "<%=apiURL%>/outmall/lfmall/xSiteOrder_lfmall_Process.asp"
	}else if(sellsite=="shopify"){
		frm.action = "/admin/etc/shopify/xSiteOrder_shopify_Process.asp"
	}else if(sellsite=="lotteon"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=lotteon"
	}else if(sellsite=="shintvshopping"){
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=shintvshopping&gubunCode=3"
	}else if(sellsite=="wconcept1010"){
		var popwin = window.open("/admin/etc/order/popXSiteOrderInput.asp?sellsite="+sellsite,"xSiteOrderInput","width=1200 height=900 scrollbars=yes resizable=yes");
		popwin.focus()
		return false;	
	}
    frm.submit();
}

function popBatchOrderInput(){
	var popwin = window.open("xSiteOrderInputBatch.asp?sellsite=<%=sellsite%>","xSiteOrderInputBatch","width=1200 height=900 scrollbars=yes resizable=yes");
	popwin.focus()

}

function SubmitInputOrder(frm){
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('선택 주문이 없습니다.');
        return;
    }

    if (confirm('주문을 입력 하시겠습니까?')){
        frm.mode.value="add";
        frm.submit();
    }
}

function AddNewPartnerItemLinkWithOrder(SellSite, orderItemID, orderItemName, orderItemOption, orderItemOptionName) {
	var popwin = window.open("/company/partnercompany/partneritemlink_modify_frame.asp?SellSite=" + SellSite + "&orderItemID=" + orderItemID + "&orderItemName=" + orderItemName + "&orderItemOption=" + orderItemOption + "&orderItemOptionName=" + orderItemOptionName,"AddNewPartnerItemLinkWithOrder","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popMatchItemIDEdit(outMallorderSeq,orderitemid,matchitemoption){
    //alert('수정중');

    var retval = window.showModalDialog("/lib/inputpop.html","selItem", "dialogwidth:450px;dialogheight:550px;center:yes;scroll:no;resizable:no;status:no;help:no;");

    if (retval!=""){
        if (IsDigit(retval)){
            if (confirm('상품코드 :' + retval + ' 수정하시겠습니까 ? ')){
                var popwin = window.open("OrderInput_Process.asp?mode=MatchItemSeqChg&outMallorderSeq="+outMallorderSeq+"&orderItemID="+orderitemid+"&chgItemID="+retval,"OrderInput_Process","width=100 height=100 scrollbars=yes resizable=yes");
                popwin.focus();
            }
        }else{
            alert('숫자만 가능합니다.');
        }
    }
}

function popMatchItemOptionEdit(outMallorderSeq,Matchitemid,matchitemoption){
    var popwin = window.open("popMatchItemOptionEdit.asp?outMallorderSeq="+outMallorderSeq+"&Matchitemid="+Matchitemid+"&matchitemoption="+matchitemoption,"popMatchItemOptionEdit","width=900 height=580 scrollbars=yes resizable=yes");
    popwin.focus();
}

function delInputOrder(outMallorderSeq,OutMallOrderSerial,orderItemID,orderItemOption){
    if (!confirm('삭제 하시겠습니까?')){
        return;
    }
    var popwin = window.open("OrderInput_Process.asp?mode=delpInputOrder&outMallorderSeq="+outMallorderSeq+"&OutMallOrderSerial="+OutMallOrderSerial+"&orderItemID="+orderItemID+"&orderItemOption="+orderItemOption,"OrderInput_Process","width=100 height=100 scrollbars=yes resizable=yes");
    popwin.focus();
}

function chgComp(comp){
    var frm = comp.form;

    //frm.sellsite.disabled = (comp.checked);
    //frm.matchState.disabled = (comp.checked);
    //frm.orderserial.disabled = (comp.checked);
    //frm.outmallorderserial.disabled = (comp.checked);
}

function orderHandModi(ooseq){
    var popwin = window.open("popOrderHandModi.asp?ooseq="+ooseq,"popOrderHandModi","width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function delNotiList(orderno,orderseq){
    if (confirm('주문 수신후/입력전 취소하는 기능입니다. 계속하시겠습니까?')){
        var popwin = window.open("OrderInput_Process.asp?mode=ltimalldel&outMallorderSeq="+orderseq+"&OutMallOrderSerial="+orderno,"OrderInput_Process","width=100 height=100 scrollbars=yes resizable=yes");
        popwin.focus();
    }
}

function updateZipCode(outMallorderSerial) {
	var popzip = window.open("popZipCodeEdit.asp?outMallorderSerial=" + outMallorderSerial,"updateZipCode","width=900 height=580 scrollbars=yes resizable=yes");
	popzip.focus();
}

function updateMemo(outMallorderSerial){
	var popmemo = window.open("popMemoEdit.asp?outMallorderSerial=" + outMallorderSerial,"updateMemo","width=900 height=200 scrollbars=yes resizable=yes");
	popmemo.focus();
}

function GSUpdateOrderserial(outMallorderSerial, sitename, OutMallOrderSeq){
/*
	if(sitename == 'auction'){
		alert('당분간 운영기획팀(김진영 대리)으로 문의 바랍니다.');
		return;
	}
*/
    var iURI = "OrderInput_Process.asp?mode=gsshopupdate&outMallorderSerial="+outMallorderSerial+"&sitename="+sitename+"&OutMallOrderSeq="+OutMallOrderSeq;
    if ((sitename=="ssg")||(sitename=="SSG")){
        iURI = "OrderInput_Process.asp?mode=ssgupdate&outMallorderSerial="+outMallorderSerial+"&sitename="+sitename+"&OutMallOrderSeq="+OutMallOrderSeq;
    }

    if (confirm('중복된 상품을 합치는 기능입니다. 계속하시겠습니까?')){
        var popwin = window.open(iURI,"GSOrderInput_Process","width=800 height=500 scrollbars=yes resizable=yes");
        popwin.focus();
    }
}

function poporderedit(outmallorderseq){
	var poporderedit = window.open('/admin/etc/orderinput/xSiteOrderedit.asp?outmallorderseq='+outmallorderseq,'poporderedit','width=600,height=200,scrollbars=yes,resizable=yes');
	poporderedit.focus();
}
function reqdetailInsert(s){
	var popreqedit = window.open('/admin/etc/orderinput/xSiteReqDetailedit.asp?outmallorderseq='+s,'poporderedit','width=600,height=200,scrollbars=yes,resizable=yes');
	popreqedit.focus();
}
function apiOrderProcess(){
	var v = document.getElementById("apiOrder").value;
    if (confirm(""+v+"몰의 주문 연동 등록 하시겠습니까?")){
		if (v == "cjmall") {
			xlOnlineOrderCjMall();
		}else if(v =="lotteimall"){
			xSiteOrderInput('lotteimall');
		}else if(v == "11st"){
			xSiteOrderInput('11st1010');
		}else if(v == "gmarket1010"){
			xSiteOrderInput('gmarket1010');
		}else if(v == "auction1010"){
			xSiteOrderInput('auction1010');
		}else if(v == "nvstorefarm"){
			xSiteOrderInput('nvstorefarm');
		}else if(v == "nvstorefarmclass"){
			xSiteOrderInput('nvstorefarmclass');
		}else if(v == "nvstoremoonbangu"){
			xSiteOrderInput('nvstoremoonbangu');
		}else if(v == "Mylittlewhoopee"){
			xSiteOrderInput('Mylittlewhoopee');
		}else if(v == "nvstoregift"){
			xSiteOrderInput('nvstoregift');
		}else if(v == "ezwel"){
			xSiteOrderInput('ezwel');
		}else if(v == "kakaostore"){
			xSiteOrderInput('kakaostore');
		}else if(v == "boribori1010"){
			xSiteOrderInput('boribori1010');
		}else if(v == "lotteCom"){
			xSiteOrderInput('lotteCom');
		}else if(v == "interpark"){
			xSiteOrderInput('interpark');
		}else if(v == "gseshop"){
			xSiteOrderInput('gseshop');
		}else if(v == "ssg"){
			xSiteOrderInput('ssg');
		}else if(v == "halfclub"){
			xSiteOrderInput('halfclub');
		}else if(v == "sabangnet"){
			xSiteOrderInput('sabangnet');
		}else if(v == "coupang"){
			xSiteOrderInput('coupang');
		}else if(v == "hmall1010"){
			xSiteOrderInput('hmall1010');
		}else if(v == "WMP"){
			xSiteOrderInput('WMP');
		}else if(v == "wmpfashion"){
			xSiteOrderInput('wmpfashion');
		}else if(v == "LFmall"){
			xSiteOrderInput('LFmall');
		}else if(v == "lotteon"){
			xSiteOrderInput('lotteon');
		}else if(v == "shintvshopping"){
			xSiteOrderInput('shintvshopping');
		}else if(v == "wconcept1010"){
			xSiteOrderInput('wconcept1010');
		}else if(v == "shopify"){
			xSiteOrderInput('shopify');
		}
    }
}
function SubmitDelOrder(frm){
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('선택 주문이 없습니다.');
        return;
    }

    if (confirm('주문을 삭제 하시겠습니까?')){
        frm.mode.value="realDel";
        frm.submit();
    }
}
function UpdateOrderRealPrice(frm){
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('선택 주문이 없습니다.');
        return;
    }

    if (confirm('원판매가를 수정 하시겠습니까?')){
        frm.mode.value="realPriceUpd";
        frm.submit();
    }
}

function popDayRate(yyyy, mm){
	var winRate = window.open("/admin/etc/dayRate/dayRateList.asp?yyyy="+yyyy+"&mm="+mm.replace(/(^0+)/, ""),"popFile","width=600, height=600 ,scrollbars=yes,resizable=yes");
	winRate.focus();
}

function rateCal(outMallorderSerial, sitename, OutMallOrderSeq, paydate){
	var iURI = "OrderInput_Process.asp?mode=rateCal&outMallorderSerial="+outMallorderSerial+"&sitename="+sitename+"&OutMallOrderSeq="+OutMallOrderSeq+"&paydate="+paydate;
    if (confirm('일별 환율로 재계산 하시겠습니까?')){
        var popwin = window.open(iURI,"GSOrderInput_Process","width=800 height=500 scrollbars=yes resizable=yes");
        popwin.focus();
    }
}

$(function() {
	var CAL_Start = new Calendar({
		inputField : "regyyyymmdd", trigger    : "regyyyymmdd_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    * 쇼핑몰 선택 :
	    <% call drawSelectBoxXSiteOrderInputPartner("sellsite", sellsite) %>
	    &nbsp;&nbsp;
		<input type="radio" name="tplinc" value="" <%=CHKIIF(tplinc="","checked","")%> >ALL
		<input type="radio" name="tplinc" value="0" <%=CHKIIF(tplinc="0","checked","")%> >Ten
		<input type="radio" name="tplinc" value="1" <%=CHKIIF(tplinc="1","checked","")%> >3pl
		&nbsp;&nbsp;
	    * 처리상태 :
		<select class="select" name="matchState"  >
			<option value='' <%= chkIIF(matchState="","selected","") %> >전체</option>
	     	<option value='I' <%= chkIIF(matchState="I","selected","") %> >엑셀등록</option>
	     	<!-- option value='P' <%= chkIIF(matchState="P","selected","") %> >상품매칭완료</option -->
	     	<option value='O' <%= chkIIF(matchState="O","selected","") %> >주문입력완료</option>
	     	<option value='D' <%= chkIIF(matchState="D","selected","") %> >기입력삭제</option>
     	</select>
     	&nbsp;&nbsp;
		<input type="checkbox" name="csViewYn" value="Y" <%= chkiif(csViewYn="Y", "checked", "") %> >CS건제외&nbsp;
		<input type="checkbox" name="overseaViewYn" value="Y" <%= chkiif(overseaViewYn="Y", "checked", "") %> >해외몰제외
		&nbsp;&nbsp;
     	* 주문번호:<input type="text" name="orderserial" value="<%=orderserial%>" size="14" maxlength="11"  >
     	&nbsp;&nbsp;
     	* 제휴주문번호:<input type="text" name="outmallorderserial" value="<%= outmallorderserial %>" size="20" maxlength="20" >
		&nbsp;&nbsp;
		* 주문입력일 :
		<input id="regyyyymmdd" name="regyyyymmdd" value="<%=regyyyymmdd%>" class="text" size="10" maxlength="10" />
		<img src="http://scm.10x10.co.kr/images/calicon.gif" id="regyyyymmdd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	<!--
		<input type="checkbox" name="notilog" <%= CHKIIF(notilog="on","checked","") %> onClick="chgComp(this);"> 롯데iMall 주문 통신 내역 보기
	-->
    </td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<% if (isNotiList) then %>
	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="롯데iMall주문 임시등록" onClick="xlOnlineOrderLotteiMall();">
		</td>
		<td align="right">
		</td>
	</tr>
	</table>
	<!-- 액션 끝 -->
<% else %>
	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="1. 엑셀 등록" onClick="xlOnlineOrderUpload();">
			<!--
            <input type="button" class="button" value="롯데iMall주문 임시등록" onClick="xlOnlineOrderLotteiMall();">
            -->
            &nbsp;&nbsp;&nbsp;
			*API주문연동선택 :
			<select class="select" name="apiOrder" id="apiOrder">
		     	<option value='cjmall' <%= chkIIF(apiOrder="cjmall","selected","") %> >CJMALL</option>
		     	<option value='lotteimall' <%= chkIIF(apiOrder="lotteimall","selected","") %> >롯데iMall</option>
		     	<option value='11st' <%= chkIIF(apiOrder="11st","selected","") %> >11번가</option>
		     	<option value='gmarket1010' <%= chkIIF(apiOrder="gmarket1010","selected","") %> >지마켓</option>
				<option value ="">----------</option>
				<option value='interpark' <%= chkIIF(apiOrder="interpark","selected","") %> >인터파크</option>
				<option value='auction1010' <%= chkIIF(apiOrder="auction1010","selected","") %> >옥션</option>
				<option value='nvstorefarm' <%= chkIIF(apiOrder="nvstorefarm","selected","") %> >스토어팜</option>
				<!-- <option value='nvstorefarmclass' <%= chkIIF(apiOrder="nvstorefarmclass","selected","") %> >스토어팜 클래스</option> -->
				<!-- <option value='nvstoremoonbangu' <%= chkIIF(apiOrder="nvstoremoonbangu","selected","") %> >스토어팜 문방구</option> -->
				<option value='nvstoregift' <%= chkIIF(apiOrder="nvstoregift","selected","") %> >스토어팜 선물하기</option>
				<option value='Mylittlewhoopee' <%= chkIIF(apiOrder="Mylittlewhoopee","selected","") %> >스토어팜 캣앤독</option>
				<option value='ezwel' <%= chkIIF(apiOrder="ezwel","selected","") %> >이지웰</option>
				<option value='boribori1010' <%= chkIIF(apiOrder="boribori1010","selected","") %> >보리보리</option>
				<option value='lotteon' <%= chkIIF(apiOrder="lotteon","selected","") %> >롯데On</option>
				<option value='shintvshopping' <%= chkIIF(apiOrder="shintvshopping","selected","") %> >신세계TV쇼핑</option>
				<!-- <option value='gseshop' <%= chkIIF(apiOrder="gseshop","selected","") %> >gseshop</option> -->
				<option value ="">----------</option>
				<option value='ssg' <%= chkIIF(apiOrder="ssg","selected","") %> >신세계몰(SSG)</option>
				<option value='halfclub' <%= chkIIF(apiOrder="halfclub","selected","") %> >하프클럽</option>
				<option value='sabangnet' <%= chkIIF(apiOrder="sabangnet","selected","") %> >사방넷</option>
				<option value='coupang' <%= chkIIF(apiOrder="coupang","selected","") %> >쿠팡</option>
				<option value='hmall1010' <%= chkIIF(apiOrder="hmall1010","selected","") %> >HMall</option>
				<option value='WMP' <%= chkIIF(apiOrder="WMP","selected","") %> >위메프</option>
				<option value='wmpfashion' <%= chkIIF(apiOrder="wmpfashion","selected","") %> >위메프W패션</option>
				<option value='LFmall' <%= chkIIF(apiOrder="LFmall","selected","") %> >LFmall</option>
				<option value ="">----------</option>
				<option value='wconcept1010' <%= chkIIF(apiOrder="wconcept1010","selected","") %> >W컨셉</option>
				<option value ="">----------</option>
				<option value='shopify' <%= chkIIF(apiOrder="shopify","selected","") %> >shopify</option>
	     	</select>
	     	<input type="button" class="button" value="API연동등록" onClick="apiOrderProcess();">
<!-- 2017-06-21 김진영 주석
            <input type="button" class="button" value="CjMall주문 API연동등록" onClick="xlOnlineOrderCjMall();">
-->
            <% if session("ssBctID")="icommang" or session("ssBctID")="kjy8517" then %>
            &nbsp;&nbsp;&nbsp;
            <input type="button" class="button" value="CjMall주문 실판매가업데이트" onClick="xlOnlineOrderUpCjMall();">
            <% end if %>
			&nbsp;
<!-- 2017-06-21 김진영 주석
			<input type="button" class="button" value="롯데iMall 주문 API연동등록" onClick="xSiteOrderInput('lotteimall');">
			&nbsp;
            <input type="button" class="button" value="롯데닷컴 주문 API연동등록" onClick="xSiteOrderInput('lotteCom');" disabled>
            &nbsp;
			<input type="button" class="button" value="11st주문 API연동등록" onClick="xSiteOrderInput('11st1010');">&nbsp;
-->
            <%' if session("ssBctID")="icommang" or session("ssBctID")="kjy8517" then %>
<!--            <input type="button" class="button" value="인터파크 정산 API연동등록" onClick="xSiteOrderInput('interpark');"> -->
            <%' end if %>

			<% If sellsite = "cookatmall" or sellsite = "aboutpet" Then %>
			<input type="button" class="button" value="송장EXCEL" onClick="excelSongjang('<%= sellsite %>');">
			<% End If %>
		</td>
		<td align="right">
			<input type="button" class="button" value="주문일괄입력" onClick="popBatchOrderInput()">
			&nbsp;&nbsp;
			<input type="button" class="button" value="2. 선택내역주문입력" onClick="SubmitInputOrder(frmSvArr)">
		</td>
	</tr>
	</table>
	<!-- 액션 끝 -->
<% end if %>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= otmpOrder.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= otmpOrder.FTotalPage %></b>
	<% If session("ssBctID")="sj100" or session("ssBctID")="nys1006" or session("ssBctID")="hrkang97" or session("ssBctID")="kjy8517" then %>
		<input type="button" class="button" value="삭제" onClick="SubmitDelOrder(frmSvArr);" style=color:red;font-weight:bold>
	<% End If %>

	<% If (sellsite = "interpark") or (sellsite = "gseshop") or (sellsite = "alphamall") or (sellsite = "aboutpet") or (sellsite = "shintvshopping") or (sellsite = "goodwearmall10") Then %>
		<input type="button" class="button" value="0원->1원수정" onClick="UpdateOrderRealPrice(frmSvArr);" style=color:red;font-weight:bold>
	<% End If %>
	</td>
</tr>
<form name="frmSvArr" method="post" action="OrderInput_Process.asp">
<input type="hidden" name="mode" value="add">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><% if (Not isNotiList) then %><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked,frmSvArr.cksel);"><% end if %></td>
    <td width="30">주문<br>구분</td>
    <td width="60">쇼핑몰</td>
    <td width="60">주문자<br>수령인</td>
	<td width="100">제휴주문번호</td>
	<td width="60">제휴주문<br>상세번호</td>
	<td width="100">주문제작문구</td>
	<td width="100">주문제작여부</td>
	<td width="60">우편번호</td>
	<td width="100">배송요청사항</td>
	<td width="60">판매가</td>
	<td width="50">수량</td>
	<td width="60">배송비</td>
  	<td >엑셀상품코드<br>상품명</td>
  	<td width="80">엑셀옵션코드<br>옵션명</td>
  	<td >연결상품코드<br>상품명</td>
  	<td width="80">연결옵션코드<br>옵션명</td>

  	<td width="80">Ten<br>주문번호</td>

  	<td>상품<br>매칭</td>
  	<td>옵션<br>수정</td>
  	<% if isNotiList then %>
  	<td>삭제</td>
  	<% end if %>
</tr>

<%
	Dim availableBrand, isCheckBoxDisable
	for i=0 to otmpOrder.FresultCount -1
		Select Case otmpOrder.FItemList(i).FSellSite
			Case "wconcept"		availableBrand = "o"
			Case "dnshop"		availableBrand = "o"
			Case "wizwid"		availableBrand = "o"
			Case Else			availableBrand = "x"
		End Select

		If Left(otmpOrder.FItemList(i).FmatchItemOption,2) = "FF" OR LEFT(otmpOrder.FItemList(i).ForderItemOption, 2) = "FF" Then
			optLeft2FF = "Y"
		Else
			optLeft2FF = "N"
		End If

		If otmpOrder.FItemList(i).FSellSite = "kakaogift" Then
			If (otmpOrder.FItemList(i).ForderItemOptionName <> otmpOrder.FItemList(i).FmatchItemOptionName) Then
				kakaoGiftOptNmDiff = "Y"
			Else
				kakaoGiftOptNmDiff = "N"
			End If
		Else
			kakaoGiftOptNmDiff = "N"
		End If

		If otmpOrder.FItemList(i).FSellSiteName = "shopify" AND (otmpOrder.FItemList(i).FOverseasPrice = otmpOrder.FItemList(i).FSellPrice) Then
			shopifyPriceDiff = "Y"
		Else
			shopifyPriceDiff = "N"
		End If

		If (otmpOrder.FItemList(i).IsItemMatched <> true) or (otmpOrder.FItemList(i).IsCjMallStarCASE) or (Left(otmpOrder.FItemList(i).FmatchItemOption,2) = "FF") _
		or (otmpOrder.FItemList(i).FDuppExists>0) or (otmpOrder.FItemList(i).FaddDlvExists>0) or (otmpOrder.FItemList(i).isCancelOrder) _
		or (((otmpOrder.FItemList(i).FSellPrice < 1) or ((otmpOrder.FItemList(i).FRealSellPrice < 1) AND (otmpOrder.FItemList(i).FSellSite <> "interpark") AND (otmpOrder.FItemList(i).FSellSite <> "gseshop") AND (otmpOrder.FItemList(i).FSellSite <> "skstoa") AND (otmpOrder.FItemList(i).FSellSite <> "kakaostore") AND (otmpOrder.FItemList(i).FSellSite <> "wconcept1010") AND (otmpOrder.FItemList(i).FSellSite <> "withnature1010") AND (otmpOrder.FItemList(i).FSellSite <> "LFmall") AND (otmpOrder.FItemList(i).FSellSite <> "alphamall") AND (otmpOrder.FItemList(i).FSellSite <> "aboutpet") AND (otmpOrder.FItemList(i).FSellSite <> "shintvshopping") AND (otmpOrder.FItemList(i).FSellSite <> "goodwearmall10") )) and (availableBrand = "x")) _
		or (otmpOrder.FItemList(i).FordercsGbn = "3") or (otmpOrder.FItemList(i).FFFExists >= 1) _
		or (shopifyPriceDiff= "Y") _
		or ((otmpOrder.FItemList(i).FoptionCnt>0) and ((otmpOrder.FItemList(i).FmatchItemOption="0000") AND (otmpOrder.FItemList(i).ForderItemOption="0000"))) then
	 		isCheckBoxDisable = "Y"
	 	Else
	 		isCheckBoxDisable = "N"
	 	End If

'rw "sellprice : " & otmpOrder.FItemList(i).FSellPrice
'rw "realprice : " & otmpOrder.FItemList(i).FRealSellPrice
'rw "availableBrand : " & availableBrand
'rw "optLeft2FF : " & optLeft2FF
%>

<% If ((otmpOrder.FItemList(i).FSellPrice < 1) or (otmpOrder.FItemList(i).FRealSellPrice < 1) AND (otmpOrder.FItemList(i).FSellSite <> "interpark") AND (otmpOrder.FItemList(i).FSellSite <> "gseshop") AND (otmpOrder.FItemList(i).FSellSite <> "skstoa") AND (otmpOrder.FItemList(i).FSellSite <> "kakaostore") AND (otmpOrder.FItemList(i).FSellSite <> "wconcept1010") AND (otmpOrder.FItemList(i).FSellSite <> "withnature1010") AND (otmpOrder.FItemList(i).FSellSite <> "LFmall") AND (otmpOrder.FItemList(i).FSellSite <> "alphamall") AND (otmpOrder.FItemList(i).FSellSite <> "aboutpet") AND (otmpOrder.FItemList(i).FSellSite <> "shintvshopping") AND (otmpOrder.FItemList(i).FSellSite <> "goodwearmall10") ) and (availableBrand = "x") or (optLeft2FF = "Y") or (kakaoGiftOptNmDiff = "Y") Then %>
<tr align="center" bgcolor="RED">
<% Else %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
<% End If %>
	<td>
	<%
		if pOrderSerial=otmpOrder.FItemList(i).FOutMallOrderSerial then
	%>
		  =
	<%
		else
			if (Not isNotiList) then
				if (FALSE) and (C_ADMIN_AUTH) and (otmpOrder.FItemList(i).FSellSiteName="cjmall") then
	%>
	        <input type="checkbox" name="cksel" value="<%= otmpOrder.FItemList(i).FOutMallOrderSerial %>" onclick="CheckProduct(this);" disabled >
	<%
				else
	%>
	    	<input type="checkbox" name="cksel" value="<%= otmpOrder.FItemList(i).FOutMallOrderSerial %>" onclick="CheckProduct(this);" <%= Chkiif(isCheckBoxDisable = "Y", "disabled" ,"")%> >
	<%
	    		end if
	    	end if
	   	end if
	%>
	</td>
	<!--
	<td><input type="checkbox" name="chk" onclick="CheckProduct(this);"></td>
	-->
	<td><%= otmpOrder.FItemList(i).getOrderCsGbnName %>
	<%
		If (otmpOrder.FItemList(i).FDuppExists) Then
			'If ((otmpOrder.FItemList(i).FSellSiteName="gsshop") OR (otmpOrder.FItemList(i).FSellSiteName="nvstorefarm") OR (otmpOrder.FItemList(i).FSellSiteName="auction") OR (otmpOrder.FItemList(i).FSellSiteName="SSG") OR (otmpOrder.FItemList(i).FSellSiteName="gmarket1010")) AND (session("ssAdminPsn") = "14" OR session("ssAdminPsn") = "30" OR session("ssAdminPsn")="7" ) Then
			If ((otmpOrder.FItemList(i).FSellSiteName="gsshop") OR (otmpOrder.FItemList(i).FSellSiteName="LFmall") OR (otmpOrder.FItemList(i).FSellSiteName="skstoa") OR (otmpOrder.FItemList(i).FSellSiteName="kakaostore") OR (otmpOrder.FItemList(i).FSellSite="wconcept1010") OR (otmpOrder.FItemList(i).FSellSite="withnature1010") OR (otmpOrder.FItemList(i).FSellSite="interpark") OR (otmpOrder.FItemList(i).FSellSite="shintvshopping") OR (otmpOrder.FItemList(i).FSellSite="cjmall") OR (otmpOrder.FItemList(i).FSellSite="lotteon") OR (otmpOrder.FItemList(i).FSellSite="nvstorefarm") OR (otmpOrder.FItemList(i).FSellSite="nvstoremoonbangu") OR (otmpOrder.FItemList(i).FSellSite="nvstoregift") OR (otmpOrder.FItemList(i).FSellSite="Mylittlewhoopee") OR (otmpOrder.FItemList(i).FSellSite="lotteimall") OR (otmpOrder.FItemList(i).FSellSite="ezwel") OR (otmpOrder.FItemList(i).FSellSite="WMP") OR (otmpOrder.FItemList(i).FSellSite="wmpfashion") OR (otmpOrder.FItemList(i).FSellSite="hmall1010") OR (otmpOrder.FItemList(i).FSellSite="auction1010") OR (otmpOrder.FItemList(i).FSellSite="boribori1010") OR (otmpOrder.FItemList(i).FSellSite="ssg") OR (otmpOrder.FItemList(i).FSellSite="gmarket1010")) AND (session("ssAdminPsn") = "14" OR session("ssAdminPsn") = "11" OR session("ssAdminPsn") = "22" OR session("ssAdminPsn") = "30" OR session("ssAdminPsn")="7" ) Then
				If (otmpOrder.FItemList(i).FoptionCnt = 0) OR ((otmpOrder.FItemList(i).FoptionCnt>0) and ((otmpOrder.FItemList(i).FmatchItemOption<>"0000") and (otmpOrder.FItemList(i).FmatchItemOption<>"FF00"))) Then
	%>
			<br><input type="button" value="상품중복" class="button" onclick="GSUpdateOrderserial('<%= otmpOrder.FItemList(i).FOutMallOrderSerial %>', '<%=otmpOrder.FItemList(i).FSellSite %>','<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');">
	<%
				End If
			Else
	%>
			<br>상품중복
	<%
			End If
		End If
		If (otmpOrder.FItemList(i).FaddDlvExists) Then
	%>
			<br>다수령지
	<%
		End If
	%>
	</td>
	<td><%= otmpOrder.FItemList(i).FSellSiteName %></td>
	<td>
	<% If otmpOrder.FItemList(i).FSellSite = "cjmall" OR otmpOrder.FItemList(i).FSellSite = "lotteimall" Then %>
		<a href="" onclick="poporderedit('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>'); return false;" /><%= otmpOrder.FItemList(i).FOrderName %><br><%= otmpOrder.FItemList(i).fReceiveName %></a>
	<% Else %>
		<%= otmpOrder.FItemList(i).FOrderName %><br><%= otmpOrder.FItemList(i).fReceiveName %>
	<% End If %>
	</td>
  	<td>
	<% If (otmpOrder.FItemList(i).getOrderCsGbnName<>"") or (otmpOrder.FItemList(i).FSellSite = "gseshop") or (otmpOrder.FItemList(i).FDuppExists>0) or (otmpOrder.FItemList(i).FaddDlvExists>0) Then %>
  	    <a href="javascript:orderHandModi('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>')"><%= otmpOrder.FItemList(i).FOutMallOrderSerial %></a>
	<% Else %>
  	    <%= otmpOrder.FItemList(i).FOutMallOrderSerial %>
	<% End If %>
	<%
		If otmpOrder.FItemList(i).FSellSiteName = "shopify" Then
			response.write "<br /><strong><font color='BLUE'>" & otmpOrder.FItemList(i).FShopifyOrderName & "</font></strong>"
		End If 
	%>
  	</td>
    <td>
	<%
		if otmpOrder.FItemList(i).FSellSite="lotteimall" then
			if ( isNotiList) then
				response.write otmpOrder.FItemList(i).FOrgDetailKey
            else
				response.write Mid(otmpOrder.FItemList(i).FOrgDetailKey,16,11)
			end if
		else
			response.write otmpOrder.FItemList(i).FOrgDetailKey
		end if
	%>
    </td>
    <td>
   	<%
    	If otmpOrder.FItemList(i).FSellSite = "nvstorefarm" or otmpOrder.FItemList(i).FSellSite = "nvstoremoonbangu" or otmpOrder.FItemList(i).FSellSite = "nvstoregift" or otmpOrder.FItemList(i).FSellSite = "Mylittlewhoopee" Then
    		If (Instr(otmpOrder.FItemList(i).ForderItemOptionName, "직접입력:") >= 1) AND (otmpOrder.FItemList(i).FrequireDetail = "") AND (otmpOrder.FItemList(i).FItemdiv = "06") Then
	%>
		<input type="button" value="입력" class="button" onclick="reqdetailInsert('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');">
	<% 		Else %>
		<span onclick="reqdetailInsert('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');" style="cursor:pointer;"><%= otmpOrder.FItemList(i).FrequireDetail %></span>
	<%
    		End If
		ElseIf otmpOrder.FItemList(i).FSellSite = "11st1010" Then
    		If (Instr(otmpOrder.FItemList(i).ForderItemOptionName, "텍스트를 입력하세요:") >= 1) AND (otmpOrder.FItemList(i).FrequireDetail = "") AND (otmpOrder.FItemList(i).FItemdiv = "06") Then
	%>
		<input type="button" value="입력" class="button" onclick="reqdetailInsert('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');">
	<% 		Else %>
		<span onclick="reqdetailInsert('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');" style="cursor:pointer;"><%= otmpOrder.FItemList(i).FrequireDetail %></span>
	<%
    		End If
    	Else
    		response.write otmpOrder.FItemList(i).FrequireDetail
			If Len(otmpOrder.FItemList(i).FrequireDetail) > 2 Then
	%>
				<br / ><input type="button" value="수정" class="button" onclick="reqdetailInsert('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');">
	<%
			End If
    	End If
   	%>
    </td>
    <td>
   	<%
    	If otmpOrder.FItemList(i).FItemdiv = "06" Then
			response.write "<font color='red'>제작문구 필요</font>"
		ElseIf otmpOrder.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='blue'>주문제작 상품</font>"
		Else
			response.write "N"
		End If
   	%>
    </td>
    <td>
	<%
'		If Len(REPLACE(otmpOrder.FItemList(i).FReceiveZipCode,"-","")) < 6 Then
			rw otmpOrder.FItemList(i).FReceiveZipCode
			response.write "<input type='button' class='button' value='수정' onclick=updateZipCode('"&otmpOrder.FItemList(i).FOutMallOrderSerial&"'); >"
'		Else
'			rw otmpOrder.FItemList(i).FReceiveZipCode
'		End If
	%>
    </td>
    <td>
		<%
			rw otmpOrder.FItemList(i).Fdeliverymemo
			response.write "<input type='button' class='button' value='수정' onclick=updateMemo('"&otmpOrder.FItemList(i).FOutMallOrderSerial&"'); >"
		%>
	</td>
    <td align="right"><%= FormatNumber(otmpOrder.FItemList(i).FSellPrice,0) %>
    <% if (otmpOrder.FItemList(i).FSellPrice<>otmpOrder.FItemList(i).FRealSellPrice) then %>
        <% if otmpOrder.FItemList(i).FRealSellPrice>otmpOrder.FItemList(i).FSellPrice then %>
        <br>(<b><font color=red><%=FormatNumber(otmpOrder.FItemList(i).FRealSellPrice,0)%></font></b>)
        <% else %>
        <br>(<%=FormatNumber(otmpOrder.FItemList(i).FRealSellPrice,0)%>)
        <% end if %>
    <% end if %>
    <% if otmpOrder.FItemList(i).isCurDiffPrice then %>
    <br><font color=red><%= otmpOrder.FItemList(i).getCurDiffPriceHtml %></font>
    <% end if %>

	<% If otmpOrder.FItemList(i).FSellSiteName = "shopify" AND (otmpOrder.FItemList(i).FOverseasPrice = otmpOrder.FItemList(i).FSellPrice) Then %>
		<br />
		<input type="button" class="button" value="<%= LEFT(otmpOrder.FItemList(i).FPaydate, 10) %>" onclick="popDayRate('<%= LEFT(otmpOrder.FItemList(i).FPaydate, 4) %>', '<%= Split(otmpOrder.FItemList(i).FPaydate, "-")(1) %>');">
		<input type="button" class="button" value="계산" onclick="rateCal('<%= otmpOrder.FItemList(i).FOutMallOrderSerial %>', '<%=otmpOrder.FItemList(i).FSellSite %>','<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>', '<%= LEFT(otmpOrder.FItemList(i).FPaydate, 10) %>');" >
	<% End If %>
  	<td ><%= otmpOrder.FItemList(i).FItemOrderCount %></td>
  	<td ><%= FormatNumber(otmpOrder.FItemList(i).ForderDlvPay,0) %></td>
  	<td><%= otmpOrder.FItemList(i).ForderItemID %><br>
  	<% IF (Right(otmpOrder.FItemList(i).getorderItemName,Len(NULL2Blank(otmpOrder.FItemList(i).FmatchItemName)))<>otmpOrder.FItemList(i).FmatchItemName) then %>
  	<font color="red"><%= otmpOrder.FItemList(i).getorderItemName %></font>
  	<% else %>
  	<%= otmpOrder.FItemList(i).getorderItemName %>
  	<% end if %>
  	</td>
  	<td><% if (IsNull(otmpOrder.FItemList(i).ForderItemOption)) then response.write "----" else response.write otmpOrder.FItemList(i).ForderItemOption end if %><br><%= otmpOrder.FItemList(i).ForderItemOptionName %></td>

  	<td>
  	<% if IsNull(otmpOrder.FItemList(i).FmatchItemID)  then %>
  	<input type="button" class="button" value="상품연결 등록" onclick="AddNewPartnerItemLinkWithOrder('<%= otmpOrder.FItemList(i).FSellSite %>', '<%= otmpOrder.FItemList(i).ForderItemID %>', '<%= Server.URLEncode(otmpOrder.FItemList(i).getorderItemName) %>', '<%= otmpOrder.FItemList(i).ForderItemOption %>', '<%= Server.URLEncode(otmpOrder.FItemList(i).ForderItemOptionName) %>');">
  	<% else %>
  	    <% if (otmpOrder.FItemList(i).FmatchItemID=0) and (otmpOrder.FItemList(i).FmatchState="I") then %>
  	    <input type="button" value="수정.." onClick="popMatchItemIDEdit('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>','<%= otmpOrder.FItemList(i).FmatchItemID %>','<%= otmpOrder.FItemList(i).FmatchItemOption %>');">
  	    <% else %>
  	    <%= otmpOrder.FItemList(i).FmatchItemID %>
  	    <%= CHKIIF(otmpOrder.FItemList(i).isCurItemSoldOut,"<font color='red'><b>[품절]</b></font>","")%>
  	    <br><%= otmpOrder.FItemList(i).FmatchItemName %>
  	    <% end if %>
  	<% end if %>
  	</td>
  	<td>
  	<% if (otmpOrder.FItemList(i).IsItemOptionNotMatched) then %>
  	    <%= otmpOrder.FItemList(i).FmatchItemOption %>
  	    <%= CHKIIF(otmpOrder.FItemList(i).isCurItemOptionSoldOut,"<font color='red'><b>[품절]</b></font>","")%>
  	    <br><%= otmpOrder.FItemList(i).FmatchItemOptionName %>

      	<% if  ((Not isNotiList) and (otmpOrder.FItemList(i).FmatchState<>"O")) then '' ((isNotiList) and (otmpOrder.FItemList(i).FmatchState="0")) or %>
      	<input type="button" value="수정.." onClick="popMatchItemOptionEdit('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>','<%= otmpOrder.FItemList(i).ForderItemID %>','<%= otmpOrder.FItemList(i).FmatchItemOption %>');">
      	<% end if %>

  	<% else %>
  	    <%= otmpOrder.FItemList(i).FmatchItemOption %><br>
      	<% if (otmpOrder.FItemList(i).IsItemOptionNameNotMatched) then %>
      	<b><font color="red"><%= otmpOrder.FItemList(i).FmatchItemOptionName %></font></b>
      	<% else %>
      	<%= otmpOrder.FItemList(i).FmatchItemOptionName %>
      	<% end if %>
    <% end if%>
    </td>
  	<td><%= otmpOrder.FItemList(i).Forderserial %></td>

    <% if ( isNotiList) then %>
    <td><%= otmpOrder.FItemList(i).getNotiStateString %></td>
    <% else %>
  	<td><%= otmpOrder.FItemList(i).getmatchStateString %></td>
  	<% end if %>

  	<td>
  	<% if (otmpOrder.FItemList(i).FoptionCnt>0) and ((otmpOrder.FItemList(i).FmatchItemOption="0000") or (otmpOrder.FItemList(i).ForderItemOption="0000") or (optLeft2FF = "Y") OR (kakaoGiftOptNmDiff = "Y")) then %>
  	    <% if isNULL(otmpOrder.FItemList(i).Forderserial) then %>
  	    <input type="button" value="수정." onClick="popMatchItemOptionEdit('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>','<%= otmpOrder.FItemList(i).ForderItemID %>','<%= otmpOrder.FItemList(i).FmatchItemOption %>');">
  	    <br><input type="button" value="삭제" onClick="delInputOrder('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>','<%= otmpOrder.FItemList(i).FOutMallOrderSerial %>','<%= otmpOrder.FItemList(i).ForderItemID %>','<%= NULL2Blank(otmpOrder.FItemList(i).ForderItemOption) %>');">
 	    <% end if %>
  	<% end if %>
  	</td>
  	<% if isNotiList then %>
  	<td>
  	<% if otmpOrder.FItemList(i).FmatchState="0" then %>
  	<input type="button" value="입력전 취소" onClick="delNotiList('<%= otmpOrder.FItemList(i).FOutMallOrderSerial %>','<%= otmpOrder.FItemList(i).FOrgDetailKey %>');">
  	<% end if %>
  	</td>
  	<% end if %>
</tr>
<%
pOrderSerial = otmpOrder.FItemList(i).FOutMallOrderSerial
%>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if otmpOrder.HasPreScroll then %>
		<a href="javascript:NextPage('<%= otmpOrder.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + otmpOrder.StartScrollPage to otmpOrder.FScrollCount + otmpOrder.StartScrollPage - 1 %>
			<% if i>otmpOrder.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if otmpOrder.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<form name="frmTmp" method="post" action="http://scm.10x10.co.kr/admin/etc/cjMall/actCjMallReq.asp">
<input type="hidden" name="cmdparam" value="">
</form>

<form name="frmXl" method="post" action="/admin/etc/orderInput/excelSongJang.asp">
<input type="hidden" name="mallid" value="">
</form>

<form name="frmXSiteOrder" method="post" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
</form>
<% Set otmpOrder = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim mallid, oOutmall, page, i
Dim itemid, apiAction, resultCode, lastUserid, sellyn, pagesize, errMsg
mallid		= request("mallid")
itemid		= request("itemid")
apiAction	= request("apiAction")
resultCode	= request("resultCode")
page 		= request("page")
lastUserid	= request("lastUserid")
sellyn		= request("sellyn")
pagesize	= request("pagesize")
errMsg		= request("errMsg")

If page = "" Then page = 1
If pagesize = "" Then pagesize = 100

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

Set oOutmall = new COutmall
	oOutmall.FPageSize 			= pagesize
	oOutmall.FCurrPage			= page
	oOutmall.FRectMallid 		= mallid
	oOutmall.FRectItemid 		= itemid
	oOutmall.FRectApiAction 	= apiAction
	oOutmall.FRectResultCode 	= resultCode
	oOutmall.FRectLastUserid 	= lastUserid
	oOutmall.FRectGSShopSellyn 	= sellyn
	oOutmall.FRectErrMsg	 	= errMsg
	oOutmall.getQueLogbackupList
%>
<script>
//크롬 업데이트로 alert 수정..2021-07-26
function systemAlert(message){
	alert(message);
}
window.addEventListener("message", (event) => {
    var data = event.data;
    if (typeof(window[data.action]) == "function") {
        window[data.action].call(null, data.message);
    } },
false);
//크롬 업데이트로 alert 수정..2021-07-26 끝

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

// 선택된 상품 삭제
function etcmallDeleteProcess(imallid) {
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	if (imallid == '11st1010'){
		if (confirm('API로 삭제하는 기능이 아닙니다.\n\n11번가 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
			if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
				document.frmSvArr.target = "xLink";
				document.frmSvArr.cmdparam.value = "DELETE";
				document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
				document.frmSvArr.submit();
			}
		}
	} else if (imallid == 'interpark'){
		if (confirm('API로 삭제하는 기능이 아닙니다.\n\인터파크 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
			if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
				document.frmSvArr.target = "xLink";
				document.frmSvArr.cmdparam.value = "DELETE";
				document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actinterparkReq.asp"
				document.frmSvArr.submit();
			}
		}
	} else if (imallid == 'gmarket1010'){
		if (confirm('API로 삭제하는 기능이 아닙니다.\n\지마켓 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
			if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
				document.frmSvArr.target = "xLink";
				document.frmSvArr.cmdparam.value = "DELETE";
				document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
				document.frmSvArr.submit();
			}
		}
	}
}
// 선택된 상품 판매여부 변경
function etcmallSellYnProcess(chkYn, imallid) {
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="판매중";break;
		case "N": strSell="품절";break;
	}

	if (imallid == 'gsshop'){
	    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        if (chkYn=="X"){
	            if (!confirm(strSell + '로 변경하면 GSShop에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
	        }
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshop/actgsshopReq.asp"
	        document.frmSvArr.submit();
	    }
	}else if (imallid == 'lotteimall'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 롯데iMall에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'lotteCom'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※롯데닷컴에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/LotteCom/actLotteComReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'cjmall'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※CJMall에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'auction1010'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※Auction과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※Auction에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'gmarket1010'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※Gmarket과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※Gmarket에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'homeplus'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※홈플러스의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※홈플러스에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'interpark'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※인터파크의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※인터파크에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actinterparkReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'nvstorefarm'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※스토어팜에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarm/actnvstorefarmReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'nvstoremoonbangu'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※스토어팜 문방구와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※스토어팜 문방구에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoremoonbangu/actnvstoremoonbanguReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'Mylittlewhoopee'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '로 변경하면 n※스토어팜에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == '11st1010'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※11st와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'ssg'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※ssg와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'coupang'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※coupang와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'halfclub'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※halfclub와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'hmall1010'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※hmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
			document.frmSvArr.submit();
		}
	}else if (imallid == 'ezwel'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※ezwel와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'lotteon'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※lotteon과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteon/actlotteonReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'lfmall'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※lfmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'WMP'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※위메프와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/wmp/actWmpReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'skstoa'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※skstoa와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp"
	        document.frmSvArr.submit();
		}
	}else if (imallid == 'shintvshopping'){
		if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※shintvshopping와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp"
	        document.frmSvArr.submit();
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td>몰구분 : <%= mallid %></td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>

		API액션 :
	<% If mallid = "auction1010" Then %>
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<option value="REGOnSale"	 <%= Chkiif(apiAction = "REGOnSale", "selected", "")%> >등록에서판매중으로</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >상품수정(판매전환)</option>
			<option value="EditInfo"	 <%= Chkiif(apiAction = "EditInfo", "selected", "")%> >상품수정(반품)</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<option value="KEEPSELL" <%= Chkiif(apiAction = "KEEPSELL", "selected", "")%> >판매유지</option>
		</select>
	<% ElseIf mallid = "gmarket1010" Then %>
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<option value="REGOnSale"	 <%= Chkiif(apiAction = "REGOnSale", "selected", "")%> >등록에서판매중으로</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >상품수정(판매전환)</option>
			<option value="EDITPOLICY"	 <%= Chkiif(apiAction = "EDITPOLICY", "selected", "")%> >상품수정(반품)</option>
			<option value="EDITIMG"	 <%= Chkiif(apiAction = "EDITIMG", "selected", "")%> >이미지수정</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<option value="KEEPSELL" <%= Chkiif(apiAction = "KEEPSELL", "selected", "")%> >판매유지</option>
		</select>
	<% ElseIf mallid = "ssg" Then %>
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >상품수정(판매전환)</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<option value="CHKSTAT" <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >신규상품조회</option>
		</select>
	<% ElseIf mallid = "hmall1010" Then %>
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >상품수정(판매전환)</option>
			<option value="EDITBATCH"	 <%= Chkiif(apiAction = "EDITBATCH", "selected", "")%> >상품수정(배치처리)</option>
			<option value="OPTEDIT" <%= Chkiif(apiAction = "OPTEDIT", "selected", "")%> >옵션수정</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<option value="OPTSTAT" <%= Chkiif(apiAction = "OPTSTAT", "selected", "")%> >재고조회</option>
			<option value="CHKSTAT" <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >상세조회</option>
		</select>
	<% ElseIf mallid = "lotteon" Then %>
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"		 <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >상품수정(판매전환)</option>
			<option value="EDITBATCH"	 <%= Chkiif(apiAction = "EDITBATCH", "selected", "")%> >상품수정(배치처리)</option>
			<option value="QTY" <%= Chkiif(apiAction = "QTY", "selected", "")%> >재고수정</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<option value="OPTSTAT" <%= Chkiif(apiAction = "OPTSTAT", "selected", "")%> >옵션상태</option>
			<option value="CHKSTAT" <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >상세조회</option>
		</select>
	<% Else %>
		<select name="apiAction" class="select">
			<option value="">전체</option>
			<option value="REG"  <%= Chkiif(apiAction = "REG", "selected", "")%> >상품등록</option>
			<% If mallid = "shintvshopping" or mallid = "skstoa" Then %>
			<option value="CONFIRM" <%= Chkiif(apiAction = "CONFIRM", "selected", "")%> >등록승인요청</option>
			<% End If %>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >상품수정</option>
			<option value="EDIT2"	 <%= Chkiif(apiAction = "EDIT2", "selected", "")%> >상품수정(판매전환)</option>
			<option value="EDITBATCH"	 <%= Chkiif(apiAction = "EDITBATCH", "selected", "")%> >상품수정(배치처리)</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >품절처리</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >가격수정</option>
			<% If (mallid <> "cjmall" and mallid <> "11stmy") Then %>
			<option value="ITEMNAME" <%= Chkiif(apiAction = "ITEMNAME", "selected", "")%> >상품명수정</option>
			<% End If %>
			<% If mallid = "gsshop" Then %>
			<option value="IMAGE"	 <%= Chkiif(apiAction = "IMAGE", "selected", "")%> >이미지수정</option>
			<option value="CONTENT"  <%= Chkiif(apiAction = "CONTENT", "selected", "")%> >상품설명수정</option>
			<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >정부고시수정</option>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >신규상품조회</option>
			<% ElseIf mallid = "11stmy" Then %>
			<option value="VIEWOPT"  <%= Chkiif(apiAction = "VIEWOPT", "selected", "")%> >옵션조회</option>
			<% ElseIf mallid = "cjmall" or mallid="coupang" or mallid="ezwel" or mallid="interpark" or mallid="lfmall" Then %>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >신규상품조회</option>
			<% ElseIf (mallid = "lotteimall") OR (mallid = "lotteCom") Then %>
			<option value="CHKSTOCK"  <%= Chkiif(apiAction = "CHKSTOCK", "selected", "")%> >재고조회</option>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >신규상품조회</option>
				<% If mallid = "lotteimall" Then %>
					<option value="DISPVIEW"  <%= Chkiif(apiAction = "DISPVIEW", "selected", "")%> >전시상품조회</option>
				<% Else %>
					<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >정부고시수정</option>
				<% End If %>
			<% End If %>
			<% If mallid = "interpark" Then %>
			<option value="DELETE"  <%= Chkiif(apiAction = "DELETE", "selected", "")%> >상품삭제</option>
			<% End If %>
		</select>
	<% End If %>
		&nbsp;
		성공여부 :
		<select name="resultCode" class="select">
			<option value="">전체</option>
			<option value="OK"  	<%= Chkiif(resultCode = "OK", "selected", "")%> >성공</option>
			<option value="ERR"		<%= Chkiif(resultCode = "ERR", "selected", "")%> >에러</option>
			<option value="QNull"	<%= Chkiif(resultCode = "QNull", "selected", "")%> >예정</option>
		</select>
		&nbsp;
		제휴판매 :
		<select name="sellyn" class="select">
			<option value="">전체</option>
			<option value="Y"  	<%= Chkiif(sellyn = "Y", "selected", "")%> >Y</option>
			<option value="N"  	<%= Chkiif(sellyn = "N", "selected", "")%> >N</option>
		</select>
		&nbsp;
		수행ID :
		<select name="lastUserid" class="select">
			<option value="">전체</option>
			<option value="system"  <%= Chkiif(lastUserid = "system", "selected", "")%> >스케줄</option>
			<option value="etc"	<%= Chkiif(lastUserid = "etc", "selected", "")%> >관리자</option>
		</select>
		&nbsp;
		표시갯수 :
		<select name="pagesize" class="select">
			<option value="20"  <%= Chkiif(pagesize = "20", "selected", "")%> >20</option>
			<option value="100"  <%= Chkiif(pagesize = "100", "selected", "")%> >100</option>
			<option value="165"  <%= Chkiif(pagesize = "165", "selected", "")%> >165</option>
			<option value="200"  <%= Chkiif(pagesize = "200", "selected", "")%> >200</option>
			<option value="500"  <%= Chkiif(pagesize = "500", "selected", "")%> >500</option>
		</select>
		<br /><br />
		Message 검색 : <input type="text" name="errMsg" id="errMsg" value="<%= errMsg %>">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<textarea id="itemidArr"></textarea>
<button onclick="copyId();">Copy</button>
<script>
function copyId() {
	var ttt = document.getElementById("itemidArr");
	ttt.select();
	document.execCommand("copy");
}
</script>

<br />
<input type="button" value="돌아가기" class="button" onclick="location.replace('/admin/etc/que/popQueLogList.asp?mallid=<%= mallid %>');">
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="subcmd" value="">
<input type="hidden" name="chgSellYn" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		검색결과 : <b><%= FormatNumber(oOutmall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOutmall.FTotalPage,0) %></b>
	</td>
	<td align="right" valign="top">
		선택상품을 품절로
		<input class="button" type="button" id="btnSellYn" value="변경" onClick="etcmallSellYnProcess('N', '<%= mallid %>');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td>몰구분</td>
	<td>API액션</td>
	<td>상품코드</td>
	<td>우선순위</td>
	<td>등록시간</td>
	<td>큐읽은시간</td>
	<td>API완료시간</td>
	<td>제휴판매</td>
	<td>옵션수</td>
	<td>실패수</td>
	<td>성공여부</td>
	<td>수행ID</td>
	<td width="300">Message</td>
</tr>
<%
	Dim itemidArr, outmallGoodnoArr
%>
<% For i = 0 To oOutmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOutmall.FItemList(i).FItemid %>"></td>
	<td><%= oOutmall.FItemlist(i).FMallid %></td>
	<td><%= oOutmall.FItemlist(i).FApiAction %></td>
	<td><%= oOutmall.FItemlist(i).FItemid %></td>
	<td><%= oOutmall.FItemlist(i).FPriority %></td>
	<td><%= oOutmall.FItemlist(i).FRegdate %></td>
	<td><%= oOutmall.FItemlist(i).FReaddate %></td>
	<td><%= oOutmall.FItemlist(i).FFindate %></td>
	<td>
		<%
			If oOutmall.FItemlist(i).FGSShopSellyn = "Y" Then
				response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			Else
				response.write "<font color='RED'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			End If
		%>
	</td>
	<td><%= oOutmall.FItemlist(i).FOptioncnt %>:<%= oOutmall.FItemlist(i).FRegedOptCnt %></td>
	<td><%= oOutmall.FItemlist(i).FAccFailCnt %></td>
	<td>
	<%
		Select Case oOutmall.FItemlist(i).FResultCode
			Case "OK"		response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case "ERR"		response.write "<font color='RED'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case Else		response.write "<font color='GRAY'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
		End Select
	%>
	</td>
	<td><%= oOutmall.FItemlist(i).FLastUserid %></td>
	<td width="300"><font title='<%= oOutmall.FItemlist(i).FLastErrMsg %>'><%= left(oOutmall.FItemlist(i).FLastErrMsg, 120) %></font></td>
</tr>
<%
		itemidArr = itemidArr & oOutmall.FItemList(i).FItemid & ","
		If oOutmall.FItemList(i).FOutmallGoodno <> "" Then
			outmallGoodnoArr = outmallGoodnoArr & oOutmall.FItemList(i).FOutmallGoodno & VBCRLF
		End If
	Next

	If Right(itemidArr,1) = "," Then
		itemidArr = Left(itemidArr, Len(itemidArr) - 1)
	End If
%>

<% If mallid = "11st1010" or mallid = "interpark" or mallid = "gmarket1010" Then %>
<textarea id="outmallGoodnoArr"><%= outmallGoodnoArr %></textarea>
<button onclick="copyId2();">Copy</button>
<script>
function copyId2() {
	var ttt2 = document.getElementById("outmallGoodnoArr");
	ttt2.select();
	document.execCommand("copy");
}
</script>
<input class="button" type="button" id="DELETE" value="삭제" onClick="etcmallDeleteProcess('<%= mallid %>');">
<% End If %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14" align="center">
	<% If oOutmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutmall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutmall.StartScrollPage To oOutmall.FScrollCount + oOutmall.StartScrollPage - 1 %>
		<% If i>oOutmall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutmall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<% Set oOutmall = Nothing %>
<script>
	document.getElementById("itemidArr").value = "<%= itemidArr %>";
</script>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="400"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

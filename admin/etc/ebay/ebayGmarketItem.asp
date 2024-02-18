<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/ebay/ebayCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, addOptErr, isSpecialPrice
Dim bestOrdMall, gmarketGoodNo, g9GoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchBrand, optAddPrcRegTypeNone, notinmakerid, notinitemid, MatchG9, sellpriceChk, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, gmarketYes10x10No, gmarketNo10x10Yes, gmarketKeepSell, reqEdit, reqExpire, failCntExists, priceOption, isextusing, scheduleNotInItemid
Dim page, i, research, GmarketGoodNoArray
Dim oEbay, gubun
Dim startMargin, endMargin
gubun					= "G"
page    				= request("page")
research				= request("research")
itemid  				= request("itemid")
makerid					= request("makerid")
itemname				= request("itemname")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
startMargin				= request("startMargin")
endMargin				= request("endMargin")
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")
infoDiv					= request("infoDiv")
morningJY				= request("morningJY")
extsellyn				= request("extsellyn")
gmarketGoodNo			= request("gmarketGoodNo")
g9GoodNo				= request("g9GoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchBrand				= request("MatchBrand")
MatchG9					= request("MatchG9")
sellpriceChk			= request("sellpriceChk")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
addOptErr				= request("addOptErr")
gmarketYes10x10No		= request("gmarketYes10x10No")
gmarketNo10x10Yes		= request("gmarketNo10x10Yes")
gmarketKeepSell			= request("gmarketKeepSell")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchBrand = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = ""

End If

'텐바이텐 상품코드 엔터키로 검색되게
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

'Gmarket 상품코드 엔터키로 검색되게
If gmarketGoodNo <> "" then
	Dim iA2, arrTemp2, arrgmarketGoodNo
	gmarketGoodNo = replace(gmarketGoodNo,",",chr(10))
	gmarketGoodNo = replace(gmarketGoodNo,chr(13),"")
	arrTemp2 = Split(gmarketGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrgmarketGoodNo = arrgmarketGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	gmarketGoodNo = left(arrgmarketGoodNo,len(arrgmarketGoodNo)-1)
End If

'G9 상품코드 엔터키로 검색되게
If g9GoodNo <> "" then
	Dim iA3, arrTemp3, arrg9GoodNo
	g9GoodNo = replace(g9GoodNo,",",chr(10))
	g9GoodNo = replace(g9GoodNo,chr(13),"")
	arrTemp3 = Split(g9GoodNo,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			arrg9GoodNo = arrg9GoodNo& "'"& trim(arrTemp3(iA3)) & "',"
		End If
		iA3 = iA3 + 1
	Loop
	g9GoodNo = left(arrg9GoodNo,len(arrg9GoodNo)-1)
End If

Set oEbay = new CEbay
	oEbay.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oEbay.FPageSize					= 100
Else
	oEbay.FPageSize					= 50
End If
	oEbay.FRectCDL					= request("cdl")
	oEbay.FRectCDM					= request("cdm")
	oEbay.FRectCDS					= request("cds")
	oEbay.FRectItemID				= itemid
	oEbay.FRectItemName				= itemname
	oEbay.FRectSellYn				= sellyn
	oEbay.FRectLimitYn				= limityn
	oEbay.FRectSailYn				= sailyn
'	oEbay.FRectonlyValidMargin		= onlyValidMargin
	oEbay.FRectStartMargin			= startMargin
	oEbay.FRectEndMargin				= endMargin
	oEbay.FRectMakerid				= makerid
	oEbay.FRectGmarketGoodNo			= gmarketGoodNo
	oEbay.FRectG9GoodNo				= g9GoodNo
	oEbay.FRectMatchCate				= MatchCate
	oEbay.FRectMatchBrand			= MatchBrand
	oEbay.FRectMatchG9				= MatchG9
	oEbay.FRectSellpriceChk			= sellpriceChk
	oEbay.FRectIsMadeHand			= isMadeHand
	oEbay.FRectIsOption				= isOption
	oEbay.FRectIsReged				= isReged
	oEbay.FRectNotinmakerid			= notinmakerid
	oEbay.FRectNotinitemid			= notinitemid
	oEbay.FRectExcTrans				= exctrans
	oEbay.FRectPriceOption			= priceOption
	oEbay.FRectIsSpecialPrice     	= isSpecialPrice
	oEbay.FRectAddOptErr				= addOptErr
	oEbay.FRectDeliverytype			= deliverytype
	oEbay.FRectMwdiv					= mwdiv
	oEbay.FRectScheduleNotInItemid	= scheduleNotInItemid
	oEbay.FRectIsextusing			= isextusing

	oEbay.FRectExtNotReg				= ExtNotReg
	oEbay.FRectExpensive10x10		= expensive10x10
	oEbay.FRectdiffPrc				= diffPrc
	oEbay.FRectGmarketYes10x10No		= gmarketYes10x10No
	oEbay.FRectGmarketNo10x10Yes		= gmarketNo10x10Yes
	oEbay.FRectGmarketKeepSell		= gmarketKeepSell
	oEbay.FRectExtSellYn				= extsellyn
	oEbay.FRectInfoDiv				= infoDiv
	oEbay.FRectFailCntOverExcept		= ""
	oEbay.FRectFailCntExists			= failCntExists
	oEbay.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oEbay.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oEbay.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oEbay.getGmarketreqExpireItemList
Else
	oEbay.getGmarketRegedItemList		'그 외 리스트
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=gmarket1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=gmarket1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=gmarket1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function onlyJY(comp){
     if ((comp.name=="morningJY")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=true;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.value="D"
			comp.form.ExtNotReg.disabled = true;
			comp.form.sellyn.value = "A";
			comp.form.extsellyn.value = "";
			comp.form.onlyValidMargin.value="";
    	}
    }

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="gmarketKeepSell")&&(frm.gmarketKeepSell.checked)){ frm.gmarketKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="gmarketYes10x10No")&&(frm.gmarketYes10x10No.checked)){ frm.gmarketYes10x10No.checked=false }
	if ((comp.name!="gmarketNo10x10Yes")&&(frm.gmarketNo10x10Yes.checked)){ frm.gmarketNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
function checkisReged(comp){
    if (comp.name=="isReged"){
    	if (document.getElementById("AR").checked == true){
    		comp.form.ExtNotReg.value = "D"
   			comp.form.ExtNotReg.disabled = true;
   		}else if(document.getElementById("QR").checked == true){
    		comp.form.ExtNotReg.value = "D"
   			comp.form.ExtNotReg.disabled = true;
			comp.form.extsellyn.value = "N";
			comp.form.sellyn.value = "Y";
   		}else{
			if (document.getElementById("NR").checked == false){
				comp.form.extsellyn.value = "Y";
			}else{
				comp.form.extsellyn.value = "";
				comp.form.sellyn.value = "Y";
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="gmarketYes10x10No")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.isReged.checked = true;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "N";
			comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="gmarketNo10x10Yes")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "Y";
			comp.form.extsellyn.value = "N";
			comp.form.notinmakerid.value = "";
			comp.form.notinitemid.value = "";
			comp.form.exctrans.value = "N";
			comp.form.failCntExists.value = "N";
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.gmarketYes10x10No.checked){
            comp.form.gmarketYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
	        comp.form.sellyn.value = "Y";
	        comp.form.onlyValidMargin.value="Y";
	        comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="gmarketKeepSell")&&(comp.checked)){
        if (comp.form.gmarketYes10x10No.checked){
            comp.form.gmarketYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
	        comp.form.sellyn.value = "Y";
	        comp.form.onlyValidMargin.value="";
	        comp.form.extsellyn.value = "Y";
    	}
    }

	if ((comp.name=="diffPrc")){
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
        }
	}

	if (comp.name=="reqEdit"){
		if (comp.checked){
			document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
		}
	}

	if (comp.name=="addOptErr"){
		if (comp.checked){
			document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.priceOption.value = "Y";
			comp.form.ExtNotReg.value="W"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="";
			comp.form.extsellyn.value = "N";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="gmarketKeepSell")&&(frm.gmarketKeepSell.checked)){ frm.gmarketKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="gmarketYes10x10No")&&(frm.gmarketYes10x10No.checked)){ frm.gmarketYes10x10No.checked=false }
	if ((comp.name!="gmarketNo10x10Yes")&&(frm.gmarketNo10x10Yes.checked)){ frm.gmarketNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//등록여부 조건 Reset
function ckeckReset(){
	document.frm.ExtNotReg.disabled = false;
	document.frm.wReset.checked=false;
	document.getElementById("AR").checked=false;
	document.getElementById("NR").checked=false;
	document.getElementById("RR").checked=false;
	document.getElementById("QR").checked=false;
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/ebay/popEbayCateList.asp","popCateManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=gmarket1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=gmarket1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function goMallgubun(v) {
	if(v == "auction"){
		location.replace('/admin/etc/ebay/ebayAuctionItem.asp?menupos=<%=menupos%>');
	}else{
		location.replace('/admin/etc/ebay/ebayGmarketItem.asp?menupos=<%=menupos%>');
	}
}
</script>
지마켓
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		몰 선택 :
		<select name="gubun" class="select" onchange="goMallgubun(this.value);";>
			<option value="auction">옥션</option>
			<option value="gmarket" selected>G마켓</option>
		</select>
		<br />

		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="http://www.esmplus.com/Home/Home" target="_blank">G마켓 Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then
				response.write "<font color='GREEN'>[ 10x10store | cube101010 ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		G마켓 상품코드 : <textarea rows="2" cols="20" name="gmarketGoodNo" id="itemid"><%= replace(replace(gmarketGoodNo,",",chr(10)), "'", "")%></textarea>
		&nbsp;
		G9 상품코드 : <textarea rows="2" cols="20" name="g9GoodNo" id="itemid"><%= replace(replace(g9GoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >G마켓 등록성공_OnSale전
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >G마켓 전송시도 중 오류
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >G마켓 등록예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >G마켓 등록완료(전시)
		</select>&nbsp;
		<input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">미등록</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">품절처리요망</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">등록상품 판매가능</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">등록여부조건Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
		카테고리
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>&nbsp;
		스케줄제외상품
		<select name="scheduleNotInItemid" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
		</select>&nbsp;
		G9등록여부
		<select name="MatchG9" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchG9="Y","selected","") %> >등록
			<option value="N" <%= CHkIIF(MatchG9="N","selected","") %> >미등록
		</select>&nbsp;
		금액
		<select name="sellpriceChk" class="select">
			<option value="">전체
			<option value="samman" <%= CHkIIF(sellpriceChk="samman","selected","") %> >3만원이상
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>G마켓 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="addOptErr" <%= ChkIIF(addOptErr="on","checked","") %> >추가금액등록오류</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="gmarketYes10x10No" <%= ChkIIF(gmarketYes10x10No="on","checked","") %> ><font color=red>G마켓판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="gmarketNo10x10Yes" <%= ChkIIF(gmarketNo10x10Yes="on","checked","") %> ><font color=red>G마켓품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="gmarketKeepSell" <%= ChkIIF(gmarketKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>

<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 상품설명에 IFRAME TAG 사용한 상품, 옵션가가 판매가 50% 이상인 상품, 옵션수 100개 초과 상품, 옵션가 0원 판매중 상품이 없음(옵션 한정수량 5개 이하 포함)<br />

<p />

<form name="frmReg" method="post" action="gmarketitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="등록 제외 브랜드" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('gmarket1010');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="카테고리" onclick="pop_CateManager();"> &nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
		<!-- #include virtual="/admin/etc/ebay/apiActions.asp"-->
    </td>
</tr>
</table>
</form>
<br>












<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oEbay.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEbay.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">Gmarket등록일<br>Gmarket최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">Gmarket<br>가격및판매</td>
	<td width="70">Gmarket<br>상품번호</td>
	<td width="70">G9<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
	<td width="100">기|고|옵<br>판매로 변경일</td>
</tr>
<% For i=0 to oEbay.FResultCount - 1 %>
<tr align="center" <% If oEbay.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oEbay.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oEbay.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oEbay.FItemList(i).FItemID %>','gmarket1010','<%=oEbay.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oEbay.FItemList(i).FItemID%>" target="_blank"><%= oEbay.FItemList(i).FItemID %></a>
		<% If oEbay.FItemList(i).FGmarketStatcd <> 7 Then %>
		<br><%= oEbay.FItemList(i).getGmarketStatName %>
		<% End If %>
		<%= oEbay.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oEbay.FItemList(i).FMakerid %> <%= oEbay.FItemList(i).getDeliverytypeName %><br><%= oEbay.FItemList(i).FItemName %></td>
	<td align="center"><%= oEbay.FItemList(i).FRegdate %><br><%= oEbay.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oEbay.FItemList(i).FGmarketRegdate %><br><%= oEbay.FItemList(i).FGmarketLastUpdate %></td>
	<td align="right">
		<% If oEbay.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oEbay.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oEbay.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oEbay.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oEbay.FItemList(i).Fsellcash = 0 Then
		elseif (oEbay.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oEbay.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oEbay.FItemList(i).FOrgSuplycash/oEbay.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oEbay.FItemList(i).Fbuycash/oEbay.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oEbay.FItemList(i).Fbuycash/oEbay.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oEbay.FItemList(i).IsSoldOut Then
			If oEbay.FItemList(i).FSellyn = "N" Then
	%>
		<font color="red">품절</font>
	<%
			Else
	%>
		<font color="red">일시<br>품절</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If oEbay.FItemList(i).FItemdiv = "06" OR oEbay.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oEbay.FItemList(i).FGmarketStatCd > 0) Then
			If Not IsNULL(oEbay.FItemList(i).FGmarketPrice) Then
				If (oEbay.FItemList(i).Mustprice <> oEbay.FItemList(i).FGmarketPrice) Then
	%>
					<strong><%= formatNumber(oEbay.FItemList(i).FGmarketPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oEbay.FItemList(i).FGmarketPrice,0)&"<br>"
				End If

				If Not IsNULL(oEbay.FItemList(i).FSpecialPrice) Then
					If (now() >= oEbay.FItemList(i).FStartDate) And (now() <= oEbay.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oEbay.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oEbay.FItemList(i).FSellyn="Y" and oEbay.FItemList(i).FGmarketSellYn<>"Y") or (oEbay.FItemList(i).FSellyn<>"Y" and oEbay.FItemList(i).FGmarketSellYn="Y") Then
	%>
					<strong><%= oEbay.FItemList(i).FGmarketSellYn %></strong>
	<%
				Else
					response.write oEbay.FItemList(i).FGmarketSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oEbay.FItemList(i).FGmarketGoodNo)) Then
			Response.Write "<a target='_blank' href='https://item.gmarket.co.kr/Item?goodscode="&oEbay.FItemList(i).FGmarketGoodNo&"'>"&oEbay.FItemList(i).FGmarketGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oEbay.FItemList(i).FG9GoodNo)) Then
			Response.Write "<a target='_blank' href='http://www.g9.co.kr/Display/VIP/Index/"&oEbay.FItemList(i).FG9GoodNo&"'>"&oEbay.FItemList(i).FG9GoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oEbay.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oEbay.FItemList(i).FItemID%>','0');"><%= oEbay.FItemList(i).FoptionCnt %>:<%= oEbay.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oEbay.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oEbay.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨(카)"
		Else
			response.write "<font color='darkred'>매칭안됨(카)</font>"
		End If

		' If oEbay.FItemList(i).FBrandCode > 0 Then
		' 	response.write "<br />매칭됨(브)"
		' Else
		' 	response.write "<br /><font color='darkred'>매칭안됨(브)</font>"
		' End If
	%>
	</td>
	<td align="center">
		<%= oEbay.FItemList(i).FinfoDiv %>
		<%
		If (oEbay.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oEbay.FItemList(i).FlastErrStr) &"'>ERR:"& oEbay.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= oEbay.FItemList(i).FDisplayDate %>
	</td>
</tr>
<% GmarketGoodNoArray = GmarketGoodNoArray & oEbay.FItemList(i).FGmarketGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= GmarketGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oEbay.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEbay.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEbay.StartScrollPage to oEbay.FScrollCount + oEbay.StartScrollPage - 1 %>
    		<% if i>oEbay.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEbay.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>









<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oEbay = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

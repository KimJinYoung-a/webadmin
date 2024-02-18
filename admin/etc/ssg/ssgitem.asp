<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ssg
' Hieditor : 김진영 생성
'            2022.09.27 한용민 수정(오류수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgitemcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, ssgGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, setMargin, exctrans
Dim expensive10x10, diffPrc, ssgYes10x10No, ssgNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, deliverytype, mwdiv, isSpecialPrice
Dim page, i, research, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim oSsg, xl, kjypageSize
dim startsell, stopsell
Dim startMargin, endMargin
Dim purchasetype
page    				= request("page")
kjypageSize				= request("kjypageSize")
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
ssgGoodNo				= request("ssgGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
ssgYes10x10No			= request("ssgYes10x10No")
ssgNo10x10Yes			= request("ssgNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
setMargin				= request("setMargin")
startsell				= requestCheckVar(request("startsell"), 1)
stopsell				= requestCheckVar(request("stopsell"), 1)
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
purchasetype			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"

	if (stopsell = "Y") then
		'// 판매중지 대상 상품목록
		ExtNotReg = "D"
		sellyn = "N"
		extsellyn = "Y"
		ssgYes10x10No = "on"
		onlyValidMargin = ""
	elseif (startsell = "Y") then
		'// 판매전환 대상 상품목록
		ExtNotReg = "D"
		sellyn = "Y"
		extsellyn = "N"
		onlyValidMargin = "Y"
		notinmakerid = "N"
		notinitemid = "N"
		ssgNo10x10Yes = "on"
	end if
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = "1844673,1844658,1840365,1840349,1840348,1836595,1836260,1835212,1834869,1833193,1833156,1832693,1832692,1832454,1830488,1830124,1830113,1829131,1829130,1829100,1829057,1829056,1828945,1828944,1828856,1828354,1826836,1826687,1826686,1826467,1826466,1825608,1824761,1824444,1823959,1823958,1823795,1823794,1823637,1823537,1823290,1823288,1823287,1823286,1823282,1823280,1823278,1823257,1823256,1823255,1823254,1823253,1823251,1823250,1823249,1823248,1823247,1823246,1823245,1823244,1823243,1823242,1823241,1823236,1823232,1823230,1823229,1823228,1823227,1823226,1823225,1823224,1823223,1823222,1823221,1823220,1823219,1823218,1823217,1822369,1820674,1820596,1819187,1819186,1819185,1819184,1819139,1818606,1818605,1818604,1818603,1818602,1817236,1817167,1816479,1816409,1816408,1815096,1815062,1814656,1814579,1814578,1813143,1813131,1812395,1812394,1812393,1812392,1811522,1811482,1811480,1811456,1811455,1811454,1811453,1811452,1811451,1811450,1811449,1811448,1811447,1811446,1811445,1811442,1811441,1811440,1811439,1811423,1811422,1811420,1811139,1810710,1810701,1808667,1808666,1808665,1808305,1808304,1805638,1805637,1805636,1805635,1805634,1805633,1805632,1805631,1804494,1804493,1804492,1804490,1804489,1804478,1804477,1803160,1803159,1803158,1803157,1802751,1800434,1800421,1800420,1798878,1798877,1798876,1798875,1796468,1796466,1795719,1795713,1795504,1795458,1795352,1795221,1795111,1792876,1791947,1779732,1778064,1777398,1777396,1772554,1771028,1767757,1764925,1764875,1764874,1680141,1622014,1533355,1413087,1404481,1396558,1363606,1361058,1196372,1143455,1143452,1143451,1143450,1143447,1143445,1135310,1097515,958669"
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
	if trim(arrItemid)<>"" and not(isnull(trim(arrItemid))) then
	itemid = left(trim(arrItemid),len(trim(arrItemid))-1)
	end if
End If

'ssg 상품코드 엔터키로 검색되게
If ssgGoodNo <> "" then
	Dim iA2, arrTemp2, arrssgGoodNo
	ssgGoodNo = replace(ssgGoodNo,",",chr(10))
	ssgGoodNo = replace(ssgGoodNo,chr(13),"")
	arrTemp2 = Split(ssgGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrssgGoodNo = arrssgGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	ssgGoodNo = left(arrssgGoodNo,len(arrssgGoodNo)-1)
End If

Set oSsg = new Cssg
	oSsg.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oSsg.FPageSize					= kjypageSize
Else
	oSsg.FPageSize					= 100
End If
	oSsg.FRectCDL					= request("cdl")
	oSsg.FRectCDM					= request("cdm")
	oSsg.FRectCDS					= request("cds")
	oSsg.FRectItemID				= itemid
	oSsg.FRectItemName				= itemname
	oSsg.FRectSellYn				= sellyn
	oSsg.FRectLimitYn				= limityn
	oSsg.FRectSailYn				= sailyn
'	oSsg.FRectonlyValidMargin		= onlyValidMargin
	oSsg.FRectStartMargin			= startMargin
	oSsg.FRectEndMargin				= endMargin
	oSsg.FRectMakerid				= makerid
	oSsg.FRectssgGoodNo				= ssgGoodNo
	oSsg.FRectMatchCate				= MatchCate
	oSsg.FRectIsMadeHand			= isMadeHand
	oSsg.FRectIsOption				= isOption
	oSsg.FRectIsReged				= isReged
	oSsg.FRectNotinmakerid			= notinmakerid
	oSsg.FRectNotinitemid			= notinitemid
	oSsg.FRectExcTrans				= exctrans
	oSsg.FRectPriceOption			= priceOption
	oSsg.FRectIsSpecialPrice        = isSpecialPrice
	oSsg.FRectDeliverytype			= deliverytype
	oSsg.FRectMwdiv					= mwdiv
	oSsg.FRectScheduleNotInItemid	= scheduleNotInItemid
	oSsg.FRectSetMargin				= setMargin
	oSsg.FRectIsextusing			= isextusing
	oSsg.FRectCisextusing			= cisextusing
	oSsg.FRectRctsellcnt			= rctsellcnt

	oSsg.FRectExtNotReg				= ExtNotReg
	oSsg.FRectExpensive10x10		= expensive10x10
	oSsg.FRectdiffPrc				= diffPrc
	oSsg.FRectssgYes10x10No			= ssgYes10x10No
	oSsg.FRectssgNo10x10Yes			= ssgNo10x10Yes
	oSsg.FRectExtSellYn				= extsellyn
	oSsg.FRectInfoDiv				= infoDiv
	oSsg.FRectFailCntOverExcept		= ""
	oSsg.FRectFailCntExists			= failCntExists
	oSsg.FRectReqEdit				= reqEdit
	oSsg.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oSsg.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oSsg.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oSsg.getssgreqExpireItemList
Else
	oSsg.getssgRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=ssgList"& replace(DATE(), "-", "") &"_xl.xls"
Else
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
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

// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=ssg","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=ssg','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=ssg','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//마진 변경 카테고리
function popMarginCateList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginCateList.asp?mallid=ssg','popMarginCateList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//마진 변경 상품
function popMarginItemList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginItemList.asp?mallid=ssg','popMarginItemList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//키워드 관리
function popKeywordItemList(){
	var popwin=window.open('/admin/etc/common/popKeywordList.asp?mallgubun=ssg','popKeywordItemList','width=1300,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//키워드 관리
function popSourceAreaList(){
	var popwin=window.open('/admin/etc/common/popSourceareaList.asp?mallgubun=ssg','popSourceAreaList','width=1300,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=ssg','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="ssgYes10x10No")&&(frm.ssgYes10x10No.checked)){ frm.ssgYes10x10No.checked=false }
	if ((comp.name!="ssgNo10x10Yes")&&(frm.ssgNo10x10Yes.checked)){ frm.ssgNo10x10Yes.checked=false }
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

    if ((comp.name=="ssgYes10x10No")&&(comp.checked)){
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
			comp.form.failCntExists.value = "N";
    	}
    }

    if ((comp.name=="ssgNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.ssgYes10x10No.checked){
            comp.form.ssgYes10x10No.checked = false;
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

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="ssgYes10x10No")&&(frm.ssgYes10x10No.checked)){ frm.ssgYes10x10No.checked=false }
	if ((comp.name!="ssgNo10x10Yes")&&(frm.ssgNo10x10Yes.checked)){ frm.ssgNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/ssg/popssgcateList.asp","popCateSSGmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//표준 카테고리 관리
function pop_stdCateManager() {
	var stdCM2 = window.open("/admin/etc/ssg/popssgStdcateList.asp","popstdCateSSGManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	stdCM2.focus();
}
//전시 카테고리 관리
function pop_dispCateManager() {
	var stdCM3 = window.open("/admin/etc/ssg/popssgdispcateList.asp","popstdCateSSGManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	stdCM3.focus();
}

// 선택된 상품 등록
function ssgREGProcess() {
	var chkSel=0;
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

    if (confirm('SSG에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}
//승인조회
function ssgConfirmProcess(){
	var chkSel=0;
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

    if (confirm('SSG에 선택하신 ' + chkSel + '개 상품 승인여부를 검색하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function ssgSellYnProcess(chkYn) {
	var chkSel=0;
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
		case "Y": strSell="판매";break;
		case "N": strSell="일시중지";break;
		case "X": strSell="영구중단";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}

//상품 수정
function ssgEditProcess(){
	var chkSel=0;
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

    if (confirm('SSG에 선택하신 ' + chkSel + '개 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}
//상품 조회
function ssgViewItemProcess(){
	var chkSel=0;
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

    if (confirm('SSG에 선택하신 ' + chkSel + '개 조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "VIEW";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}
function ssgGosiViewProcess(){
	if (confirm('SSG에 고시 정보를 조회 하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "GOSI";
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
		document.frmSvArr.submit();
	}
}
function ssgAreaViewProcess(){
	if (confirm('SSG에 원산지 정보를 조회 하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "AREA";
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
		document.frmSvArr.submit();
	}
}
function ssgDisplayCateProcess(){
	if (confirm('SSG 전시카테고리를 가져 오시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "DISPCATE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
		document.frmSvArr.submit();
	}
}
//제조, 유효일 등록 팝업
function popssgDate(iitemid){
    var pdate = window.open("/admin/etc/ssg/popssgDate.asp?itemid="+iitemid+'&mallid=ssg',"popssgDate","width=500,height=200,scrollbars=yes,resizable=yes");
	pdate.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=ssg&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

<% if request("auto") = "Y" then %>
function ssgEditProcessAuto() {
	var cnt = <%= oSsg.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oSsg.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		ssgEditProcessAuto();
		// 5분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 5*60*1000);
	}
}

$(document).ready(function() {
    $('table').hide();
});
<% end if %>

function popXL()
{
    frmXL.submit();
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="http://po.ssgadm.com/" target="_blank">SSG Admin바로가기</a>
		<%
			If ((session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") ) Then
				response.write "<font color='GREEN'>[ 0000003198 | Cube1010**! ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		SSG 상품코드 : <textarea rows="2" cols="20" name="ssgGoodNo" id="itemid"><%= replace(replace(ssgGoodNo,",",chr(10)), "'", "")%></textarea>
		&nbsp;
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >SSG 등록성공_승인대기
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >SSG 전송시도 중 오류
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >SSG 반려
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >SSG 등록예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >SSG 등록완료(전시)
		</select>&nbsp;
		<input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">미등록<font color="<%= CHKIIF(makerid="" and itemid="", "#000000", "#AAAAAA") %>">(최근 3개월 등록상품만)</font></label>&nbsp;
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
		<% If (session("ssBctID")="kjy8517") Then %>
			<input class="text" size="5" type="text" name="kjypageSize" value="<%= kjypageSize %>">
		<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>SSG 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="ssgYes10x10No" <%= ChkIIF(ssgYes10x10No="on","checked","") %> ><font color=red>SSG판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="ssgNo10x10Yes" <%= ChkIIF(ssgNo10x10Yes="on","checked","") %> ><font color=red>SSG품절&텐바이텐판매가능</font>(전송제외상품 제외) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 소비자가 대비 80% 초과할인인 경우 80% 할인가<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : <br />

<p />
<% end if %>
<form name="frmReg" method="post" action="ssgitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="매입마진변경(카테고리)" onclick="popMarginCateList();">&nbsp;
				<input class="button" type="button" value="매입마진변경(상품)" onclick="popMarginItemList();">&nbsp;
				<input class="button" type="button" value="키워드" onclick="popKeywordItemList();">&nbsp;
				<input class="button" type="button" value="원산지" onclick="popSourceAreaList();">&nbsp;
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">

			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('ssg');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="표준카테고리" onclick="pop_stdCateManager();" style=color:blue;font-weight:bold> &nbsp;
				<input class="button" type="button" value="전시카테고리" onclick="pop_dispCateManager();" style=color:blue;font-weight:bold> &nbsp;
				<!-- <input class="button" type="button" value="카테고리" onclick="pop_CateManager();"> &nbsp; -->
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
				실제상품 등록 :
				<input class="button" type="button" id="btnREG" value="등록" onClick="ssgREGProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnOEdit" value="수정" onClick="ssgEditProcess();" style=color:blue;font-weight:bold>
				<br><br>
				기타코드 조회 :
				<input class="button" type="button" id="btnStat" value="상품조회" onClick="ssgViewItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStat" value="승인" onClick="ssgConfirmProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" id="btnGosi" value="고시" onClick="ssgGosiViewProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnGosi" value="원산지" onClick="ssgAreaViewProcess();">
				<br><br>
				카테고리 조회 :
				<input class="button" type="button" id="btnGosi" value="전시" onClick="ssgDisplayCateProcess();">
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">판매중지</option>
					<option value="Y">판매</option>
					<option value="X">영구중단</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="ssgSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		검색결과 : <b><%= FormatNumber(oSsg.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSsg.FTotalPage,0) %></b>
	</td>
	<td align="right">
		<input type="button" class="button" value="엑셀받기" onClick="popXL()">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
<% If (xl <> "Y") Then %>
	<td width="50">이미지</td>
<% End If %>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">SSG등록일<br>SSG최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">SSG<br>가격및판매</td>
	<td width="100">SSG<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="50">적용마진</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oSsg.FResultCount - 1 %>
<tr align="center" <% If oSsg.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oSsg.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oSsg.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oSsg.FItemList(i).FItemID %>','ssg','<%=oSsg.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oSsg.FItemList(i).FItemID%>" target="_blank"><%= oSsg.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oSsg.FItemList(i).FssgStatcd <> 7 Then
	%>
		<br><%= oSsg.FItemList(i).getssgStatName %>
	<%
			End If
			response.write oSsg.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oSsg.FItemList(i).FMakerid %> <%= oSsg.FItemList(i).getDeliverytypeName %><br><%= oSsg.FItemList(i).FItemName %></td>
	<td align="center"><%= oSsg.FItemList(i).FRegdate %><br><%= oSsg.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oSsg.FItemList(i).FssgRegdate %><br><%= oSsg.FItemList(i).FssgLastUpdate %></td>
	<td align="right">
		<% If oSsg.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oSsg.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oSsg.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oSsg.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oSsg.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).FssgPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oSsg.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oSsg.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oSsg.FItemList(i).FOrgSuplycash/oSsg.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oSsg.FItemList(i).IsSoldOut Then
			If oSsg.FItemList(i).FSellyn = "N" Then
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
		If oSsg.FItemList(i).FItemdiv = "06" OR oSsg.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oSsg.FItemList(i).FssgStatCd > 0) Then
			If Not IsNULL(oSsg.FItemList(i).FssgPrice) Then
				If (oSsg.FItemList(i).Mustprice <> oSsg.FItemList(i).FssgPrice) Then
	%>
					<strong><%= formatNumber(oSsg.FItemList(i).FssgPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oSsg.FItemList(i).FssgPrice,0)&"<br>"
				End If

				If Not IsNULL(oSsg.FItemList(i).FSpecialPrice) Then
					If (now() >= oSsg.FItemList(i).FStartDate) And (now() <= oSsg.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oSsg.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oSsg.FItemList(i).FSellyn="Y" and oSsg.FItemList(i).FssgSellYn<>"Y") or (oSsg.FItemList(i).FSellyn<>"Y" and oSsg.FItemList(i).FssgSellYn="Y") Then
	%>
					<strong><%= oSsg.FItemList(i).FssgSellYn %></strong>
	<%
				Else
					response.write oSsg.FItemList(i).FssgSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oSsg.FItemList(i).FssgGoodNo)) Then
			Response.Write "<a target='_blank' href='http://www.ssg.com/item/itemView.ssg?itemId="&oSsg.FItemList(i).FssgGoodNo&"'>"&oSsg.FItemList(i).FssgGoodNo&"</a>"
			'Response.Write "<span style='cursor:pointer;' onclick=window.open('http://www.ssg.com/item/itemView.ssg?itemid="&oSsg.FItemList(i).FssgGoodNo&"')>"&oSsg.FItemList(i).FssgGoodNo&"</span><br>"
		End If
	%>
	</td>
	<td align="center"><%= oSsg.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oSsg.FItemList(i).FItemID%>','0');"><%= oSsg.FItemList(i).FoptionCnt %>:<%= oSsg.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oSsg.FItemList(i).FrctSellCNT %></td>
	<td align="center"><%= oSsg.FItemList(i).FSetMargin %></td>
	<td align="center">
	<%
		If oSsg.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨(카)"
		Else
			response.write "<font color='darkred'>매칭안됨(카)</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oSsg.FItemList(i).FinfoDiv %>
		<%
		If (oSsg.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oSsg.FItemList(i).FlastErrStr) &"'>ERR:"& oSsg.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oSsg.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSsg.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSsg.StartScrollPage to oSsg.FScrollCount + oSsg.StartScrollPage - 1 %>
    		<% if i>oSsg.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSsg.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="page" value= <%= page %>>
	<input type="hidden" name="research" value= <%= research %>>
	<input type="hidden" name="itemid" value= <%= itemid %>>
	<input type="hidden" name="makerid" value= <%= makerid %>>
	<input type="hidden" name="itemname" value= <%= itemname %>>
	<input type="hidden" name="bestOrd" value= <%= bestOrd %>>
	<input type="hidden" name="bestOrdMall" value= <%= bestOrdMall %>>
	<input type="hidden" name="sellyn" value= <%= sellyn %>>
	<input type="hidden" name="limityn" value= <%= limityn %>>
	<input type="hidden" name="sailyn" value= <%= sailyn %>>
	<input type="hidden" name="onlyValidMargin" value= <%= onlyValidMargin %>>
	<input type="hidden" name="startMargin" value= <%= startMargin %>>
	<input type="hidden" name="endMargin" value= <%= endMargin %>>
	<input type="hidden" name="isMadeHand" value= <%= isMadeHand %>>
	<input type="hidden" name="isOption" value= <%= isOption %>>
	<input type="hidden" name="infoDiv" value= <%= infoDiv %>>
	<input type="hidden" name="morningJY" value= <%= morningJY %>>
	<input type="hidden" name="extsellyn" value= <%= extsellyn %>>
	<input type="hidden" name="ssgGoodNo" value= <%= ssgGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="ssgYes10x10No" value= <%= ssgYes10x10No %>>
	<input type="hidden" name="ssgNo10x10Yes" value= <%= ssgNo10x10Yes %>>
	<input type="hidden" name="reqEdit" value= <%= reqEdit %>>
	<input type="hidden" name="reqExpire" value= <%= reqExpire %>>
	<input type="hidden" name="failCntExists" value= <%= failCntExists %>>
	<input type="hidden" name="optAddPrcRegTypeNone" value= <%= optAddPrcRegTypeNone %>>
	<input type="hidden" name="notinmakerid" value= <%= notinmakerid %>>
	<input type="hidden" name="priceOption" value= <%= priceOption %>>
	<input type="hidden" name="isSpecialPrice" value= <%= isSpecialPrice %>>
	<input type="hidden" name="deliverytype" value= <%= deliverytype %>>
	<input type="hidden" name="mwdiv" value= <%= mwdiv %>>
	<input type="hidden" name="setMargin" value= <%= setMargin %>>
	<input type="hidden" name="startsell" value= <%= startsell %>>
	<input type="hidden" name="stopsell" value= <%= stopsell %>>
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET oSsg = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

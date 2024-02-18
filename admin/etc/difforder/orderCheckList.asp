<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/difforder/diffOrderCls.asp"-->
<%
Dim oOrder, page, i, sellsite, snapDate, mwdiv
Dim oOrdSummary
page		= requestCheckvar(request("page"),10)
sellsite	= requestCheckvar(request("sellsite"),32)
snapDate	= requestCheckvar(request("snapDate"),10)
mwdiv		= requestCheckvar(request("mwdiv"),10)

Dim research : research = requestCheckvar(request("research"),10)
Dim ErrType1 : ErrType1 = requestCheckvar(request("ErrType1"),10)
Dim ErrType2 : ErrType2 = requestCheckvar(request("ErrType2"),10)
Dim ErrType3 : ErrType3 = requestCheckvar(request("ErrType3"),10)
Dim showimage : showimage = requestCheckvar(request("showimage"),10)
Dim showsummary : showsummary = requestCheckvar(request("showsummary"),10)
Dim bygrp : bygrp = requestCheckvar(request("bygrp"),10)
Dim isiteMatch

if (research="") and (showsummary="") then showsummary="on"
if (research="") and (snapDate="") then snapDate=LEFT(NOW(),10)

If page = "" Then page = 1
if (showsummary="on") then
	set oOrdSummary = new COrder
	oOrdSummary.FRectSnapDate	= snapDate
	oOrdSummary.getOrderCheckSummaryList
end if
SET oOrder = new COrder
	oOrder.FCurrPage		= page
	oOrder.FPageSize		= 50
	oOrder.FRectSellsite	= sellsite
	oOrder.FRectSnapDate	= snapDate
	oOrder.FRectErrType1	= ErrType1
	oOrder.FRectErrType2	= ErrType2
	oOrder.FRectErrType3	= ErrType3
	oOrder.FRectGroupByItem  = bygrp
	oOrder.FRectMwdiv		= mwdiv
	oOrder.getOrderCheckList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function goPopOutmall(isellsite, iitemid){
	var pCM;
	switch(isellsite){
		case "auction1010"	: pCM = window.open("/admin/etc/auction/auctionItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ezwel"		: pCM = window.open("/admin/etc/ezwel/ezwelItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "kakaostore"	: pCM = window.open("/admin/etc/kakaostore/kakaostoreItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "boribori1010"	: pCM = window.open("/admin/etc/boribori/boriboriItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "gmarket1010"	: pCM = window.open("/admin/etc/gmarket/gmarketItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "gseshop"		: pCM = window.open("/admin/etc/gsshop/gsshopItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");break;pCM.focus();
		case "interpark"	: pCM = window.open("/admin/etc/interpark/interparkItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "nvstorefarm"	: pCM = window.open("/admin/etc/nvstorefarm/nvstorefarmItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "nvstoregift"	: pCM = window.open("/admin/etc/nvstoregift/nvstoregiftItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "Mylittlewhoopee"	: pCM = window.open("/admin/etc/Mylittlewhoopee/MylittlewhoopeeItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteCom"		: pCM = window.open("/admin/etc/lotte/lotteItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteimall"	: pCM = window.open("/admin/etc/ltimall/lotteiMallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "cjmall"		: pCM = window.open("/admin/etc/cjmall/cjmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "benepia1010"	: pCM = window.open("/admin/etc/benepia/benepiaItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "11stmy"		: pCM = window.open("/admin/etc/my11st/my11stItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "kakaogift"	: pCM = window.open("/admin/etc/kakaogift/kakaogiftitem.asp?research=on&itemid="+iitemid+"&ExtNotReg=D&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "11st1010"		: pCM = window.open("/admin/etc/11st/11stItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ssg"			: pCM = window.open("/admin/etc/ssg/ssgItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "shintvshopping"	: pCM = window.open("/admin/etc/shintvshopping/shintvshoppingItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wetoo1300k"	: pCM = window.open("/admin/etc/wetoo1300k/wetoo1300kItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "skstoa"		: pCM = window.open("/admin/etc/skstoa/skstoaItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "halfclub"		: pCM = window.open("/admin/etc/halfclub/halfclubItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "coupang"		: pCM = window.open("/admin/etc/coupang/coupangItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "hmall1010"	: pCM = window.open("/admin/etc/hmall/hmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "WMP"			: pCM = window.open("/admin/etc/wmp/wmpItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lfmall"		: pCM = window.open("/admin/etc/lfmall/lfmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteon"		: pCM = window.open("/admin/etc/lotteon/lotteonItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "yes24"		: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wconcept1010"	: pCM = window.open("/admin/etc/wconcept/wconceptItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "withnature1010"	: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "alphamall"	: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ohou1010"		: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "casamia_good_com"		: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wadsmartstore": pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "cnhigo"		: break;
		default				: pCM = window.open("/admin/etc/orderinput/xSiteItemLink.asp?sellsite="+isellsite+"&itemidarr="+iitemid,"goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
	}
}

function rePage(sellsite,ErrType1,ErrType2,ErrType3){
	var frm = document.frm;
	frm.sellsite.value=sellsite;
	$("#ErrType1_"+ErrType1).val(ErrType1).prop("checked", true);
	$("#ErrType2_"+ErrType2).val(ErrType2).prop("checked", true);
	$("#ErrType3_"+ErrType3).val(ErrType3).prop("checked", true);

	frm.submit();
}

function popExtItemCheckUpload(){
	var popwin = window.open("/admin/etc/difforder/popExtItemCheckUpload.asp","popExtItemCheckUpload","width=800,height=600,scrollbars=yes,resizable=yes")
	popwin.focus();
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
		제휴몰 :
		<select class="select" name="sellsite">
			<option value="">-전체-</option>
			<option value="auction1010" <%= chkiif(sellsite = "auction1010", "selected", "") %> >옥션</option>
			<option value="ezwel" <%= chkiif(sellsite = "ezwel", "selected", "") %> >이지웰페어</option>
			<option value="benepia1010" <%= chkiif(sellsite = "benepia1010", "selected", "") %> >베네피아</option>
			<option value="gmarket1010" <%= chkiif(sellsite = "gmarket1010", "selected", "") %> >G마켓</option>
			<option value="gseshop" <%= chkiif(sellsite = "gseshop", "selected", "") %> >GSShop</option>
			<option value="interpark" <%= chkiif(sellsite = "interpark", "selected", "") %> >인터파크</option>
			<option value="nvstorefarm" <%= chkiif(sellsite = "nvstorefarm", "selected", "") %> >스토어팜</option>
			<option value="Mylittlewhoopee" <%= chkiif(sellsite = "Mylittlewhoopee", "selected", "") %> >스토어팜 캣앤독</option>
			<option value="nvstoregift" <%= chkiif(sellsite = "nvstoregift", "selected", "") %> >스토어팜선물하기</option>
			<option value="kakaogift" <%= chkiif(sellsite = "kakaogift", "selected", "") %> >카카오기프트</option>
			<option value="kakaostore" <%= chkiif(sellsite = "kakaostore", "selected", "") %> >카카오톡스토어</option>
			<option value="boribori1010" <%= chkiif(sellsite = "boribori1010", "selected", "") %> >보리보리</option>
			<option value="lotteimall" <%= chkiif(sellsite = "lotteimall", "selected", "") %> >롯데아이몰</option>
			<option value="lotteon" <%= chkIIF(sellsite="lotteon","selected","") %> >롯데On</option>
			<option value="cjmall" <%= chkiif(sellsite = "cjmall", "selected", "") %> >CJMall</option>
			<option value="11st1010" <%= chkiif(sellsite = "11st1010", "selected", "") %> >11번가</option>
			<option value="ssg" <%= chkiif(sellsite = "ssg", "selected", "") %> >SSG</option>
			<option value="shintvshopping" <%= chkiif(sellsite = "shintvshopping", "selected", "") %> >신세계TV쇼핑</option>
			<option value="skstoa" <%= chkiif(sellsite = "skstoa", "selected", "") %> >SKSTOA</option>
			<option value="wetoo1300k" <%= chkiif(sellsite = "wetoo1300k", "selected", "") %> >1300k</option>
			<option value="coupang" <%= chkiif(sellsite = "coupang", "selected", "") %> >쿠팡</option>
			<option value="hmall1010" <%= chkiif(sellsite = "hmall1010", "selected", "") %> >hMall</option>
			<option value="WMP" <%= chkIIF(sellsite="WMP","selected","") %> >위메프</option>
			<option value="lfmall" <%= chkIIF(sellsite="lfmall","selected","") %> >LFmall</option>
			<option value="wconcept1010" <%= chkIIF(sellsite="wconcept1010","selected","") %> >W컨셉</option>
			<option value="withnature1010" <%= chkIIF(sellsite="withnature1010","selected","") %> >자연이랑</option>
			<option value="yes24" <%= chkIIF(sellsite="yes24","selected","") %> >Yes24</option>
			<option value="alphamall" <%= chkIIF(sellsite="alphamall","selected","") %> >알파몰</option>
			<option value="ohou1010" <%= chkIIF(sellsite="ohou1010","selected","") %> >오늘의집</option>
			<option value="wadsmartstore" <%= chkIIF(sellsite="wadsmartstore","selected","") %> >와드스마트스토어</option>
			<option value="casamia_good_com" <%= chkIIF(sellsite="casamia_good_com","selected","") %> >까사미아</option>
		</select>&nbsp;
		거래구분
		<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;
		주문입력일 :
		<input id="snapDate" name="snapDate" value="<%=snapDate%>" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		&nbsp;&nbsp;|&nbsp;&nbsp;
		오차타입(품절)
		<input type="radio" name="ErrType1" id="ErrType1_" value="" <%=CHKIIF(ErrType1="","checked","")%> >전체
		<input type="radio" name="ErrType1" id="ErrType1_3" value="3" <%=CHKIIF(ErrType1="3","checked","")%> >상품/옵션품절
		<input type="radio" name="ErrType1" id="ErrType1_1" value="1" <%=CHKIIF(ErrType1="1","checked","")%> >상품품절
		<input type="radio" name="ErrType1" id="ErrType1_2" value="2" <%=CHKIIF(ErrType1="2","checked","")%> >옵션품절
		&nbsp;&nbsp;|&nbsp;&nbsp;
		오차타입(가격)
		<input type="radio" name="ErrType2" id="ErrType2_" value="" <%=CHKIIF(ErrType2="","checked","")%> >전체
		<input type="radio" name="ErrType2" id="ErrType2_7" value="7" <%=CHKIIF(ErrType2="7","checked","")%> >가격오류
		<input type="radio" name="ErrType2" id="ErrType2_-15" value="-15" <%=CHKIIF(ErrType2="-15","checked","")%> >가격(기타)
		&nbsp;&nbsp;|&nbsp;&nbsp;
		오차타입(1+1)
		<input type="radio" name="ErrType3" id="ErrType3_" value="" <%=CHKIIF(ErrType3="","checked","")%> >전체
		<input type="radio" name="ErrType3" id="ErrType3_1" value="1" <%=CHKIIF(ErrType3="1","checked","")%> >1+1
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
	<input type="checkbox" name="showsummary" <%=CHKIIF(showsummary="on","checked","")%> >서머리표시
	<input type="checkbox" name="showimage" <%=CHKIIF(showimage="on","checked","")%> >이미지표시

	<% if (FALSE) then %>
	&nbsp;&nbsp;|&nbsp;&nbsp;
	<input type="checkbox" name="bygrp" <%=CHKIIF(bygrp="on","checked","")%> >상품별그루핑
	<% end if %>
	</td>
</tr>
</form>
</table>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="5" class="a" bgcolor="#FFFFFF">
<tr>
	<td align="right">
	<input type="button" value="상품검증XL등록" onClick="popExtItemCheckUpload()">
	</td>
</tr>
</table>
-->
<br />
<% if (showsummary="on") then %>
<%
	Dim ErrType11Cnt, ErrType12Cnt, ErrType27Cnt, errTTL, PriceEtcErrCNT
	Dim ErrType31Cnt, proTTL, sellRowTTL
	for i=0 to oOrdSummary.FResultCount - 1
		ErrType11Cnt = ErrType11Cnt + oOrdSummary.FItemList(i).FItemSoldOutCNT
		ErrType12Cnt = ErrType12Cnt + oOrdSummary.FItemList(i).FOptionSoldOutCNT
		ErrType27Cnt = ErrType27Cnt + oOrdSummary.FItemList(i).FPriceErrCNT

		errTTL			 = errTTL + oOrdSummary.FItemList(i).FerrTTL
		PriceEtcErrCNT	 = PriceEtcErrCNT + oOrdSummary.FItemList(i).FPriceEtcErrCNT

		ErrType31Cnt = ErrType31Cnt + oOrdSummary.FItemList(i).FnmErrCnt

		sellRowTTL = sellRowTTL + oOrdSummary.FItemList(i).FsellRowCnt
	next
	if (sellRowTTL<>0) then
		proTTL = CLNG((ErrType11Cnt+ErrType12Cnt)/sellRowTTL*100)/10
	else
		proTTL = 0
	end if
%>
<table width="100%" align="center" cellpadding="3" cellspacing="5" class="a" bgcolor="#FFFFFF">
<tr>
	<td width="50%">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" style="cursor:pointer">
			<td onClick="rePage('','','')">사이트</td>
			<td width="60">최종시각</td>
			<td width="60">주문Rows<br>(<%=sellRowTTL%>)</td>
			<td width="11%" onClick="rePage('','1','','')">상품품절오류<br>(<%=ErrType11Cnt%>)</td>
			<td width="11%" onClick="rePage('','2','','')">옵션품절오류<br>(<%=ErrType12Cnt%>)</td>
			<td width="11%" onClick="rePage('','3','','')"><strong>품절오류(TTL)<br>(<%=ErrType11Cnt+ErrType12Cnt%>)</strong></td>
			<td width="11%" onClick="rePage('','','7','')"><strong>가격오류<br>(<%=ErrType27Cnt%>)</strong></td>
			<td width="11%" onClick="rePage('','3','7','')">오류TOTAL<br>(<%=errTTL%>)</td>
			<td width="5%">품절율<br>(<%=proTTL%>)</td>
			<td width="10%" onClick="rePage('','','-15','')">가격오류기타<br>(<%=PriceEtcErrCNT%>)</td>
			<td width="6%" onClick="rePage('','','','1')">1+1<br>(<%=ErrType31Cnt%>)</td>
		</tr>
		<% for i=0 to CLNG(oOrdSummary.FResultCount/2+0.01) -1 %>
		<% if oOrdSummary.FResultCount>i then %>
		<% isiteMatch = oOrdSummary.FItemList(i).Fsellsite=sellsite %>
		<tr align="right" bgcolor="#FFFFFF" style="cursor:pointer">
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','','')"><%=oOrdSummary.FItemList(i).Fsellsite %></td>
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','','')"><%=oOrdSummary.FItemList(i).getLastInputTime%></td>
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','','')"><%=FormatNumber(oOrdSummary.FItemList(i).FsellRowCnt,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="1","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','1','','')"><%=FormatNumber(oOrdSummary.FItemList(i).FItemSoldOutCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="2","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','2','','')"><%=FormatNumber(oOrdSummary.FItemList(i).FOptionSoldOutCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="3","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','3','','')"><%=FormatNumber(oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType2="7","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','7','')"><%=FormatNumber(oOrdSummary.FItemList(i).FPriceErrCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="3" and ErrType2="7","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','3','7','')"><%=FormatNumber(oOrdSummary.FItemList(i).FerrTTL,0) %></td>
			<td style="cursor:default">
			<% if (oOrdSummary.FItemList(i).FsellRowCnt<>0) then %>
				<% if CLNG((oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT)/oOrdSummary.FItemList(i).FsellRowCnt*100*10)/10>=1.0 then %>
				<strong><%= CLNG((oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT)/oOrdSummary.FItemList(i).FsellRowCnt*100*10)/10 %></strong>
				<% else %>
				<%= CLNG((oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT)/oOrdSummary.FItemList(i).FsellRowCnt*100*10)/10 %>
				<% end if %>
			<% end if %>
			</td>
			<td <%=CHKIIF(isiteMatch and ErrType2="-15","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','-15','')"><%=FormatNumber(oOrdSummary.FItemList(i).FPriceEtcErrCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType3="1","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','','1')" ><%=FormatNumber(oOrdSummary.FItemList(i).FnmErrCnt,0)%></td>
		</tr>
		<% end if %>
		<% next %>

		</table>
	</td>
	<td width="50%">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >사이트</td>
			<td width="60">최종시각</td>
			<td width="60">주문Rows</td>
			<td width="11%">상품품절오류<br></td>
			<td width="11%">옵션품절오류<br></td>
			<td width="11%"><strong>품절오류(TTL)<br></strong></td>
			<td width="11%"><strong>가격오류<br></strong></td>
			<td width="11%">오류TOTAL<br></td>
			<td width="5%">품절율</td>
			<td width="10%">가격오류기타<br></td>
			<td width="6%">1+1</td>
		</tr>
		<% for i=CLNG(oOrdSummary.FResultCount/2+0.01) to CLNG(oOrdSummary.FResultCount/2+0.01)*2- 1 %>
		<% if oOrdSummary.FResultCount>i then %>
		<% isiteMatch = oOrdSummary.FItemList(i).Fsellsite=sellsite %>
		<tr align="right" bgcolor="#FFFFFF" style="cursor:pointer">
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','')"><%=oOrdSummary.FItemList(i).Fsellsite %></td>
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','')"><%=oOrdSummary.FItemList(i).getLastInputTime%></td>
			<td align="center" <%=CHKIIF(isiteMatch,"bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','')"><%=FormatNumber(oOrdSummary.FItemList(i).FsellRowCnt,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="1","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','1','')"><%=FormatNumber(oOrdSummary.FItemList(i).FItemSoldOutCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="2","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','2','')"><%=FormatNumber(oOrdSummary.FItemList(i).FOptionSoldOutCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="3","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','3','')"><%=FormatNumber(oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType2="7","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','7')"><%=FormatNumber(oOrdSummary.FItemList(i).FPriceErrCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType1="3" and ErrType2="7","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','3','7')"><%=FormatNumber(oOrdSummary.FItemList(i).FerrTTL,0) %></td>
			<td style="cursor:default">
			<% if (oOrdSummary.FItemList(i).FsellRowCnt<>0) then %>
				<% if CLNG((oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT)/oOrdSummary.FItemList(i).FsellRowCnt*100*10)/10>=1.0 then %>
				<strong><%= CLNG((oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT)/oOrdSummary.FItemList(i).FsellRowCnt*100*10)/10 %></strong>
				<% else %>
				<%= CLNG((oOrdSummary.FItemList(i).FItemSoldOutCNT+oOrdSummary.FItemList(i).FOptionSoldOutCNT)/oOrdSummary.FItemList(i).FsellRowCnt*100*10)/10 %>
				<% end if %>
			<% end if %>
			</td>
			<td <%=CHKIIF(isiteMatch and ErrType2="-15","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','-15')"><%=FormatNumber(oOrdSummary.FItemList(i).FPriceEtcErrCNT,0) %></td>
			<td <%=CHKIIF(isiteMatch and ErrType3="1","bgcolor='#E6B9B8'","")%> onClick="rePage('<%=oOrdSummary.FItemList(i).Fsellsite%>','','','1')" ><%=FormatNumber(oOrdSummary.FItemList(i).FnmErrCnt,0)%></td>
		</tr>
		<% else %>
		<tr align="right" bgcolor="#FFFFFF" >
			<td>&nbsp;</td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
			<td></td>
		</tr>
		<% end if %>
		<% next %>

		</table>
	</td>
</tr>
</table>
<br />
<% set oOrdSummary = Nothing %>
<% end if %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= FormatNumber(oOrder.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOrder.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">제휴몰</td>
	<td width="50">텐바이텐<br>상품번호</td>
	<% if (showimage="on") then %>
	<td width="50">이미지</td>
	<% end if %>
	<td width="80">제휴주문번호</td>
	<td width="80">당시<br>판매여부</td>
	<td width="80">당시<br>한정여부</td>
	<td width="50">옵션코드</td>
	<td width="80">당시<br>옵션판매여부</td>
	<% if ErrType3<>"" then %>
	<td >TEN 상품명 /옵션명</td>
	<td >제휴 상품명 /옵션명</td>
	<% else %>
	<td width="80">당시<br>옵션한정갯수</td>
	<td width="80">당시<br>옵션추가금액</td>
	<td width="80">당시<br>판매가</td>
	<td width="100">주문입력가격<br />+옵션추가금액</td>
	<td width="100">오차금액</td>
	<td width="100">오차품절</td>
	<td width="100">오차가격</td>
	<% end if %>
	<td width="90">제휴상품코드</td>
	<% if (bygrp="on") then %>
	<td width="80">건수</td>
	<% else %>
	<td width="90">주문입력일</td>
	<% end if %>
	<td>연동상태체크Time</td>
	<td>연동판매상태<br>연동STAT</td>
	<td>연동판가</td>
	<td width="100">연동최종수정일<br>연동LastCheckDT</td>
</tr>
<%
	Dim DiffStat
%>
<% For i=0 to oOrder.FResultCount - 1 %>
<%
	DiffStat = ""
%>
<tr align="center" bgcolor="#FFFFFF">
	<td style="cursor:pointer;" onclick="goPopOutmall('<%= oOrder.FItemList(i).FSellsite %>', '<%= oOrder.FItemList(i).FItemID %>');"><%= oOrder.FItemList(i).FSellsite %></td>
	<td><a href="<%=vwwwURL%>/<%=oOrder.FItemList(i).FItemID%>" target="_blank"><%= oOrder.FItemList(i).FItemID %></a></td>
	<% if (showimage="on") then %>
	<td><img src="<%= oOrder.FItemList(i).FImageSmall%>" width="50"></td>
	<% end if %>
	<td><a target="_xSiteOrderInputList" href="/admin/etc/orderinput/xSiteOrderInputList.asp?menupos=1326&research=on&sellsite=<%= oOrder.FItemList(i).FSellsite %>&matchState=&csViewYn=Y&orderserial=&outmallorderserial=<%= oOrder.FItemList(i).Foutmallorderserial%>&regyyyymmdd="><%= oOrder.FItemList(i).Foutmallorderserial%></a></td>
	<td>
		<% if oOrder.FItemList(i).FSellyn<>"Y" then %>
		<strong><%= oOrder.FItemList(i).FSellyn %></strong>
		<% else %>
		<%= oOrder.FItemList(i).FSellyn %>
		<% end if %>
	</td>
	<td>
		<%= oOrder.FItemList(i).getItemLimitStatHtml() %>
	</td>
	<td>
	<%
		If oOrder.FItemList(i).FItemoption <> "0000" Then
			response.write oOrder.FItemList(i).FItemoption
		End If
	%>
	</td>
	<td>
		<% if oOrder.FItemList(i).FItemoption <> "0000" then %>
		<%= CHKIIF(isNULL(oOrder.FItemList(i).FOptsellyn),"",oOrder.FItemList(i).FOptsellyn) %>
		<% end if %>
	</td>
	<% if ErrType3<>"" then %>
	<td align="left">
		<%= oOrder.FItemList(i).Fitemname %>
		<br><br>
		<%= oOrder.FItemList(i).Foptionname %>
	</td>
	<td align="left">
		<%= oOrder.FItemList(i).getConvXsiteOrderItemName%>
		<br><br>
		<%= oOrder.FItemList(i).getConvXsiteOrderItemOptionName%>
	</td>
	<% else %>
	<td>
		<%= oOrder.FItemList(i).getOptionItemLimitStatHtml() %>
	</td>
	<td>
		<% if oOrder.FItemList(i).FItemoption <> "0000" then %>
		<%= Formatnumber(oOrder.FItemList(i).FOptaddprice, 0) %>
		<% end if %>
	</td>
	<td ><%= Formatnumber(oOrder.FItemList(i).FSellcash, 0) %></td>
	<td><%= Formatnumber(oOrder.FItemList(i).FOrderPrice, 0) %></td>
	<td>
	<%
		response.write Formatnumber(oOrder.FItemList(i).FOrderPrice-(oOrder.FItemList(i).FOptaddprice + oOrder.FItemList(i).FSellcash), 0)
	%>
	</td>
	<td>
		<%= oOrder.FItemList(i).getMapErrType1Str() %>
		<% if (FALSE) then %>
		<br>---------<br>
		<%
			If oOrder.FItemList(i).FSellyn <> "Y" Then
				DiffStat = DiffStat & "상품품절,"
			End If
			If (oOrder.FItemList(i).FLimityn = "Y") and (oOrder.FItemList(i).FLimitno - oOrder.FItemList(i).FLimitsold < 1) Then
				DiffStat = DiffStat & "상품재고품절,"
			End If
			If oOrder.FItemList(i).FOptsellyn <> "Y" Then
				DiffStat = DiffStat & "옵션품절,"
			End If
			If (oOrder.FItemList(i).FOptlimityn = "Y") and (oOrder.FItemList(i).FOptLimitno - oOrder.FItemList(i).FOptLimitsold < 1) Then
				DiffStat = DiffStat & "옵션재고품절,"
			End If
			If Right(DiffStat,1) = "," Then
				DiffStat = Left(DiffStat, Len(DiffStat) - 1)
			End If
			response.write Replace(DiffStat, ",", "<br />")
		%>
		<% end if %>
	</td>
	<td>
		<%= oOrder.FItemList(i).getMapErrType2Str() %>
	</td>
	<% end if %>
	<td>
	<%
		If Not(IsNULL(oOrder.FItemList(i).FOutMallGoodsNo)) Then
			Select Case oOrder.FItemList(i).FSellsite
				Case "auction1010"		Response.Write "<a target='_blank' href='http://itempage3.auction.co.kr/detailview.aspx?itemNo="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "ezwel"			Response.Write "<span style='cursor:pointer;' onclick=window.open('http://shop.ezwel.com/shopNew/goods/preview/goodsDetailView.ez?preview=yes&goodsBean.goodsCd="&oOrder.FItemList(i).FOutMallGoodsNo&"');>"&oOrder.FItemList(i).FOutMallGoodsNo&"</span>"
				Case "gmarket1010"		Response.Write "<a target='_blank' href='https://item.gmarket.co.kr/Item?goodscode="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "gseshop"			Response.Write "<span style='cursor:pointer;' onclick=window.open('http://www.gsshop.com/prd/prd.gs?prdid="&oOrder.FItemList(i).FOutMallGoodsNo&"');>"&oOrder.FItemList(i).FOutMallGoodsNo&"</span>"
				Case "interpark"		Response.Write "<a target='_blank' href='https://shopping.interpark.com/product/productInfo.do?prdNo="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "nvstorefarm"		Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbyten/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "nvstoregift"		Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbytengift/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "Mylittlewhoopee"	Response.Write "<a target='_blank' href='http://storefarm.naver.com/mylittlewhoopee/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "lotteimall"		Response.Write "<a target='_blank' href='http://www.lotteimall.com/product/Product.jsp?i_code="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "cjmall"			Response.Write "<a target='_blank' href='http://www.oCJMall.com/prd/detail_cate.jsp?item_cd="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "benepia1010"		Response.Write "<a target='_blank' href='https://newmall.benepia.co.kr/disp/storeMain.bene?prdId="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "11st1010"			Response.Write "<a target='_blank' href='http://www.11st.co.kr/product/SellerProductDetail.tmall?method=getSellerProductDetail&prdNo="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "ssg"				Response.Write "<a target='_blank' href='http://www.ssg.com/item/itemView.ssg?itemId="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "shintvshopping"	Response.Write "<a target='_blank' href='https://www.shinsegaetvshopping.com/display/detail/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "wetoo1300k"		Response.Write "<a target='_blank' href='http://www.1300k.com/shop/goodsDetail.html?f_goodsno="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "skstoa"			Response.Write "<a target='_blank' href='http://www.skstoa.com/display/goods/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "WMP"				Response.Write "<a target='_blank' href='https://front.wemakeprice.com/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "lfmall"			Response.Write "<a target='_blank' href='https://www.lfmall.co.kr/product.do?cmd=getProductDetail&PROD_CD="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "lotteon"			Response.Write "<a target='_blank' href='https://www.lotteon.com/p/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "kakaostore"		Response.Write "<a target='_blank' href='https://store.kakao.com/10x10/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "boribori1010"		Response.Write "<a target='_blank' href='https://www.boribori.co.kr/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case Else 				Response.Write oOrder.FItemList(i).FOutMallGoodsNo
			End Select
		End If
	%>
	</td>
	<% if (bygrp="on") then %>
	<td><%= oOrder.FItemList(i).FgrpCNT %></td>
	<% else %>
	<td><%= LEFT(oOrder.FItemList(i).FRegdate, 13) %></td>
	<% end if %>
	<td><%=oOrder.FItemList(i).FmallSnapDt %></td>
	<td><%=oOrder.FItemList(i).FmallSnapSellyn%><br><%=oOrder.FItemList(i).FmallSnapStatcd%></td>
	<td>
		<% if NOT isNULL(oOrder.FItemList(i).FmallSnapSellprice) then %>
		<%=FormatNumber(oOrder.FItemList(i).FmallSnapSellprice,0)%>
		<% end if %>
	</td>
	<td>
		[<%=oOrder.FItemList(i).FmallSnapLastUpDT%>]<br>
		[<%=oOrder.FItemList(i).FmallSnapLastCheckDT%>]
	</td>
</tr>
<% Next %>
<tr height="21">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oOrder.HasPreScroll then %>
		<a href="javascript:goPage('<%= oOrder.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oOrder.StartScrollPage to oOrder.FScrollCount + oOrder.StartScrollPage - 1 %>
    		<% if i>oOrder.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oOrder.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<script language="javascript">
$(function() {
	var CAL_Start = new Calendar({
		inputField : "snapDate", trigger    : "sDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			//CAL_End.args.min = date;
			//CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
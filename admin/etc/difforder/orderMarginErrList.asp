<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/difforder/diffOrderCls.asp"-->
<%
Dim oOrder, page, i, sellsite, orderdate, isok, optaddpriceYN, jungsanMaeip
Dim tmpMargin, ckDiffMargin, sSdate, sEdate, isOKStr
page		= request("page")
sellsite	= request("sellsite")
isok		= request("isok")
optaddpriceYN = request("optaddpriceYN")
jungsanMaeip = request("jungsanMaeip")
sSdate			= requestCheckVar(request("iSD"),10)
sEdate			= requestCheckVar(request("iED"),10)

If page = "" Then page = 1
If sSdate = "" Then
	sSdate = DateSerial(Year(dateadd("d",-4,Now())), Month(dateadd("d",-4,Now())), 1)
end if
If sEdate = "" Then sEdate = Date()

SET oOrder = new COrder
	oOrder.FCurrPage			= page
	oOrder.FPageSize			= 50
	oOrder.FRectSdate			= sSdate
	oOrder.FRectEdate			= sEdate
	oOrder.FRectSellsite		= sellsite
	oOrder.FRectIsok			= isok
	oOrder.FRectoptaddpriceYN	= optaddpriceYN
	oOrder.FRectjungsanMaeip	= jungsanMaeip
	oOrder.getOrderMarginErrList
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
function checkIsOk(v, chk){
	document.frmSvArr.target = "xLink";
	document.frmSvArr.mode.value = "CHK2";
	document.frmSvArr.chk.value = chk;
	document.frmSvArr.idx.value = v;
	document.frmSvArr.action = "/admin/etc/difforder/isOkProc.asp"
	document.frmSvArr.submit();
}
function goPopOutmall(isellsite, iitemid){
	var pCM;
	switch(isellsite){
		case "auction1010"	: pCM = window.open("/admin/etc/auction/auctionItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ezwel"		: pCM = window.open("/admin/etc/ezwel/ezwelItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "kakaostore"	: pCM = window.open("/admin/etc/kakaostore/kakaostoreItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "gmarket1010"	: pCM = window.open("/admin/etc/gmarket/gmarketItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "gseshop"		: pCM = window.open("/admin/etc/gsshop/gsshopItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");break;pCM.focus();
		case "interpark"	: pCM = window.open("/admin/etc/interpark/interparkItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "nvstorefarm"	: pCM = window.open("/admin/etc/nvstorefarm/nvstorefarmItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteCom"		: pCM = window.open("/admin/etc/lotte/lotteItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteimall"	: pCM = window.open("/admin/etc/ltimall/lotteiMallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "cjmall"		: pCM = window.open("/admin/etc/cjmall/cjmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "11stmy"		: pCM = window.open("/admin/etc/my11st/my11stItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "kakaogift"	: pCM = window.open("/admin/etc/kakaogift/kakaogiftitem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "11st1010"		: pCM = window.open("/admin/etc/11st/11stItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ssg"			: pCM = window.open("/admin/etc/ssg/ssgItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "shintvshopping"	: pCM = window.open("/admin/etc/shintvshopping/shintvshoppingItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "skstoa"		: pCM = window.open("/admin/etc/skstoa/skstoaItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wetoo1300k"	: pCM = window.open("/admin/etc/wetoo1300k/wetoo1300kItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "coupang"		: pCM = window.open("/admin/etc/coupang/coupangItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "hmall1010"	: pCM = window.open("/admin/etc/hmall/hmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "WMP"			: pCM = window.open("/admin/etc/wmp/wmpItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lfmall"		: pCM = window.open("/admin/etc/lfmall/lfmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteon"		: pCM = window.open("/admin/etc/lotteon/lotteonItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wconcept1010"		: pCM = window.open("/admin/etc/wconcept/wconceptItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "yes24"		: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "alphamall"	: pCM = window.open("/admin/etc/sabangnet/sabangnetItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "cnglob10x10"	: break;
		case "cnhigo"		: break;
		default				: pCM = window.open("/admin/etc/orderinput/xSiteItemLink.asp?sellsite="+isellsite+"&itemidarr="+iitemid,"goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
	}
}

function pop_extsitejungsan(vOptAddYn, vItemid, vItemCost, vItemoption){
	var pCM5;
	pCM5 = window.open("/admin/etc/extsitejungsan_check.asp?itemid="+vItemid+"&mallsellcash="+vItemCost+"&itemoption="+vItemoption,"pop_jungsan");
	pCM5.location.href="/admin/etc/extsitejungsan_check.asp?itemid="+vItemid+"&mallsellcash="+vItemCost+"&itemoption="+vItemoption;
	pCM5.focus();

}
function popkakaocheck(){
	var popwin = window.open("","_popkakaocheck")
	popwin.location.href="/admin/etc/difforder/kakaochecklist.asp";
	popwin.focus();
}
function HighlightRow(obj){
	var table = document.getElementById("tableId");
	var tr = table.getElementsByTagName("tr");
	for(var i=0; i < tr.length; i++){
		tr[i].style.background = "#FFFFFF";
	}
	document.getElementById("topTr").style.background = "#e6e6e6";
	obj.parentElement.style.background = "#FCE6E0";
}
function confirmProcess() {
	var chkSel=0;
	try {
		if(frmSvArr2.cksel.length>1) {
			for(var i=0;i<frmSvArr2.cksel.length;i++) {
				if(frmSvArr2.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr2.cksel.checked) chkSel++;
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

    if (confirm('선택하신 ' + chkSel + '개 상품 가격을 일괄 확인 하시겠습니까?')){
		document.frmSvArr2.target = "xLink";
		document.frmSvArr2.mode.value = "ALL";
		document.frmSvArr2.action = "/admin/etc/difforder/isOkProc.asp"
		document.frmSvArr2.submit();
    }
}
</script>
<!-- #include virtual="/admin/etc/difforder/gubunTab.asp"-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간 :
		<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "iED", trigger    : "iED_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		제휴몰 :
		<select class="select" name="sellsite">
			<option value="">-전체-</option>
			<option value="auction1010" <%= chkiif(sellsite = "auction1010", "selected", "") %> >옥션</option>
			<option value="ezwel" <%= chkiif(sellsite = "ezwel", "selected", "") %> >이지웰페어</option>
			<option value="gmarket1010" <%= chkiif(sellsite = "gmarket1010", "selected", "") %> >G마켓</option>
			<option value="gseshop" <%= chkiif(sellsite = "gseshop", "selected", "") %> >GSShop</option>
			<option value="interpark" <%= chkiif(sellsite = "interpark", "selected", "") %> >인터파크</option>
			<option value="nvstorefarm" <%= chkiif(sellsite = "nvstorefarm", "selected", "") %> >스토어팜</option>
			<option value="lotteCom" <%= chkiif(sellsite = "lotteCom", "selected", "") %> >롯데닷컴</option>
			<option value="lotteimall" <%= chkiif(sellsite = "lotteimall", "selected", "") %> >롯데아이몰</option>
			<option value="cjmall" <%= chkiif(sellsite = "cjmall", "selected", "") %> >CJMall</option>
			<option value="11stmy" <%= chkiif(sellsite = "11stmy", "selected", "") %> >11번가(말레이시아)</option>
			<option value="11st1010" <%= chkiif(sellsite = "11st1010", "selected", "") %> >11번가</option>
			<option value="WMP" <%= chkiif(sellsite = "WMP", "selected", "") %> >위메프</option>
			<option value="ssg" <%= chkiif(sellsite = "ssg", "selected", "") %> >SSG</option>
			<option value="shintvshopping" <%= chkiif(sellsite = "shintvshopping", "selected", "") %> >신세계TV쇼핑</option>
			<option value="skstoa" <%= chkiif(sellsite = "skstoa", "selected", "") %> >SKSTOA</option>
			<option value="wetoo1300k" <%= chkiif(sellsite = "wetoo1300k", "selected", "") %> >1300k</option>			
			<option value="coupang" <%= chkiif(sellsite = "coupang", "selected", "") %> >쿠팡</option>
			<option value="hmall" <%= chkiif(sellsite = "hmall", "selected", "") %> >HMall</option>
			<option value="celectory" <%= chkiif(sellsite = "celectory", "selected", "") %> >셀렉토리</option>
			<option value="kakaogift" <%= chkiif(sellsite = "kakaogift", "selected", "") %> >kakaogift</option>
			<option value="kakaostore" <%= chkiif(sellsite = "kakaostore", "selected", "") %> >kakaostore</option>
			<option value="boribori1010" <%= chkiif(sellsite = "boribori1010", "selected", "") %> >보리보리</option>
			<option value="wconcept1010" <%= chkiif(sellsite = "wconcept1010", "selected", "") %> >W컨셉</option>
		</select>&nbsp;&nbsp;
		관리여부 :
		<select class="select" name="isok">
			<option value="">-전체-</option>
			<option value="Y" <%= chkiif(isok = "Y", "selected", "") %> >Y</option>
			<option value="N" <%= chkiif(isok = "N", "selected", "") %> >N</option>
			<option value="A" <%= chkiif(isok = "A", "selected", "") %> >A(스케줄러처리)</option>
			<option value="B" <%= chkiif(isok = "B", "selected", "") %> >B(이미매입가가맞음)</option>
		</select>&nbsp;&nbsp;
		옵션추가액여부 :
		<select class="select" name="optaddpriceYN">
			<option value="">-전체-</option>
			<option value="1" <%= chkiif(optaddpriceYN = "1", "selected", "") %> >Y</option>
			<option value="0" <%= chkiif(optaddpriceYN = "0", "selected", "") %> >N</option>
		</select>&nbsp;&nbsp;
		차이 :
		<select class="select" name="jungsanMaeip">
			<option value="">-전체-</option>
			<option value="1" <%= chkiif(jungsanMaeip = "1", "selected", "") %> >정산금액 > 매입가</option>
			<option value="2" <%= chkiif(jungsanMaeip = "2", "selected", "") %> >정산금액 <= 매입가</option>
		</select>&nbsp;&nbsp;
		&nbsp;&nbsp;
		<input type="button" value="kakaogift배송비 검토" onClick="popkakaocheck();">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<input class="button" type="button" id="btnEditPrice" value="선택내역 확인" onClick="confirmProcess();">&nbsp;&nbsp;
<br />
<form name="frmSvArr2" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode">
<table id="tableId" width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="33">
		검색결과 : <b><%= FormatNumber(oOrder.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOrder.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="topTr">
	<td width="3.1%"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr2.cksel);"></td>
	<td width="3.1%">제휴몰</td>
	<td width="3.1%">주문번호</td>
	<td width="3.1%">브랜드</td>
	<td width="3.1%">텐바이텐<br>상품번호</td>
	<td width="3.1%">옵션번호</td>
	<td width="3.1%">판매가(판매시)</td>
	<td width="3.1%">정산금액</td>
	<td width="3.1%">매입가(판매시)</td>
	<td width="3.1%">매입가(특가)</td>
	<td width="3.1%">매입구분</td>
	<td width="3.1%">할인판매</td>
	<td width="3.1%">판매시마진</td>
	<td width="3.1%">브랜드2달평균마진</td>
	<td width="3.1%">옵션추가금액</td>
	<td width="3.1%">옵션매입액</td>
	<td width="3.1%">쿠폰미적용판매가</td>
	<td width="3.1%">쿠폰미적용매입가</td>
	<td width="3.1%">현재판매가</td>
	<td width="3.1%">현재매입가</td>
	<td width="3.1%">현기준판매가차이</td>
	<td width="3.1%">현기준매입가차이</td>
	<td width="3.1%">로그매입가</td>
	<td width="3.1%">로그매입가기준일</td>
	<td width="3.1%">로그매입가차이</td>
	<td width="3.1%">판매가-옵션추가액</td>
	<td width="3.1%">매입가-옵션매입액</td>
	<td width="3.1%">제외조건1</td>
	<td width="3.1%">제외조건2</td>
	<td width="3.1%">옵션추가액여부</td>
	<td width="3.1%">체크일</td>
	<td width="3.1%">관리</td>
</tr>
<% For i=0 to oOrder.FResultCount - 1 %>
<%
	If oOrder.FItemList(i).FMargin > oOrder.FItemList(i).FBrand2MonthMargin Then
		tmpMargin = oOrder.FItemList(i).FMargin - oOrder.FItemList(i).FBrand2MonthMargin
	Else
		tmpMargin = oOrder.FItemList(i).FBrand2MonthMargin - oOrder.FItemList(i).FMargin
	End If

	If Sgn(tmpMargin) = "-1" Then
		ckDiffMargin = tmpMargin * -1
	Else
		ckDiffMargin = tmpMargin
	End If
%>
<tr align="center" bgcolor="#FFFFFF" >
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oOrder.FItemList(i).FIdx %>"></td>
	<td style="cursor:pointer;" onclick="goPopOutmall('<%= oOrder.FItemList(i).FSellsite %>', '<%= oOrder.FItemList(i).FItemID %>');"><%= oOrder.FItemList(i).FSellsite %></td>
	<td><%= oOrder.FItemList(i).FOrderserial %></td>
	<td><%= oOrder.FItemList(i).FMakerid %></td>
	<td><a href="<%=vwwwURL%>/<%=oOrder.FItemList(i).FItemID%>" target="_blank"><%= oOrder.FItemList(i).FItemID %></a></td>
	<td><%= oOrder.FItemList(i).FItemoption %></td>
	<td style="cursor:pointer;" onclick="HighlightRow(this);pop_extsitejungsan('<%=oOrder.FItemList(i).FOptaddpriceYN%>','<%=oOrder.FItemList(i).FItemID%>', '<%= oOrder.FItemList(i).FItemcost %>', '<%= oOrder.FItemList(i).FItemoption %>' );"><%= oOrder.FItemList(i).FItemcost %></td>
	<td><%= oOrder.FItemList(i).FExtTenJungsanPrice %></td>
	<td><%= oOrder.FItemList(i).FBuycash %></td>
	<td><%= oOrder.FItemList(i).FMustBuyPrice %></td>
	<td><%= oOrder.FItemList(i).FMwdiv %></td>
	<td><%= oOrder.FItemList(i).FIssailitem %></td>
	<td>
	<%
		If ckDiffMargin >= 15 Then
			response.write "<strong>"&oOrder.FItemList(i).FMargin&"</strong>"
		Else
			response.write oOrder.FItemList(i).FMargin
		End If
	%>
	</td>
	<td><%= oOrder.FItemList(i).FBrand2MonthMargin %></td>
	<td><%= oOrder.FItemList(i).FOptaddprice %></td>
	<td><%= oOrder.FItemList(i).FOptaddbuyprice %></td>
	<td><%= oOrder.FItemList(i).FItemcostCouponNotApplied %></td>
	<td><%= oOrder.FItemList(i).FBuycashCouponNotApplied %></td>
	<td><%= oOrder.FItemList(i).FNowselladdoptCost %></td>
	<td><%= oOrder.FItemList(i).FNowselladdoptbuycash %></td>
	<td><%= oOrder.FItemList(i).FNowDiffCost %></td>
	<td><%= oOrder.FItemList(i).FNowDiffbuycash %></td>
	<td><%= oOrder.FItemList(i).FLogbuycash %></td>
	<td><%= oOrder.FItemList(i).FLogbuycashDate %></td>
	<td><%= oOrder.FItemList(i).FLogDiffbuycash %></td>
	<td><%= oOrder.FItemList(i).FMinusPrice %></td>
	<td><%= oOrder.FItemList(i).FMinusbuycash %></td>
	<td><%= oOrder.FItemList(i).FEtc1 %></td>
	<td><%= oOrder.FItemList(i).FEtc2 %></td>
	<td>
	<%
		Select Case oOrder.FItemList(i).FOptaddpriceYN
			Case "1"	response.write "Y"
			Case "0"	response.write "N"
		End Select
	%>
	</td>
	<td><%= oOrder.FItemList(i).FChkDate %></td>
	<td>
	<%
		If oOrder.FItemList(i).FIsOk = "Y" OR oOrder.FItemList(i).FIsOk = "A" OR oOrder.FItemList(i).FIsOk = "B" Then
			Select Case oOrder.FItemList(i).FIsOk
				Case "Y"	isOKStr = "완료"
				Case "A"	isOKStr = "스케줄완료"
				Case "B"	isOKStr = "처리완료"
			End Select
	%>
		<input type="button" class="button"  value="<%= isOKStr %>" onclick="checkIsOk('<%=oOrder.FItemList(i).FIdx%>', 'N');" style="color:blue;font-weight:bold">
	<% Else %>
		<input type="button" class="button"  value="확인" onclick="checkIsOk('<%=oOrder.FItemList(i).FIdx%>', 'Y');">
	<% End If %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="33" align="center" bgcolor="#FFFFFF">
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
<%
SET oOrder = nothing
%>
</table>
</form>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="getOrderDate">
<input type="hidden" name="chk">
<input type="hidden" name="idx">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="500"></iframe>
<script language="javascript">
	var CAL_Start = new Calendar({
		inputField : "orderdate", trigger    : "sDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});

	var CAL_Start = new Calendar({
		inputField : "getOrderdate", trigger    : "gDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
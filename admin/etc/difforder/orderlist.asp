<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/difforder/diffOrderCls.asp"-->
<%
Dim oOrder, page, i, sellsite, orderdate, orderstate, onlyapi, isok
page		= request("page")
sellsite	= request("sellsite")
orderdate	= request("orderdate")
orderstate	= request("orderstate")
onlyapi		= request("onlyapi")
isok		= request("isok")

If orderdate = "" Then orderdate = Date() - 1

If page = "" Then page = 1
SET oOrder = new COrder
	oOrder.FCurrPage		= page
	oOrder.FPageSize		= 100
	oOrder.FRectSellsite	= sellsite
	oOrder.FRectOrderdate	= orderdate
	oOrder.FRectOrderstate	= orderstate
	oOrder.FRectOnlyapi		= onlyapi
	oOrder.FRectIsok		= isok
	oOrder.getDiffOrderList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
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
function checkIsOk(v, chk){
	document.frmSvArr.target = "xLink";
	document.frmSvArr.mode.value = "CHK";
	document.frmSvArr.chk.value = chk;
	document.frmSvArr.idx.value = v;
	document.frmSvArr.action = "/admin/etc/difforder/isOkProc.asp"
	document.frmSvArr.submit();
}
function getOrder(){
	document.frmSvArr.target = "xLink";
	document.frmSvArr.mode.value = "getOrder";
	document.frmSvArr.getOrderDate.value = $("#getOrderdate").val();
	document.frmSvArr.action = "/admin/etc/difforder/isOkProc.asp"
	document.frmSvArr.submit();
}
function goPopOutmall(isellsite, iitemid){
	var pCM;
	switch(isellsite){
		case "auction1010"	: pCM = window.open("/admin/etc/auction/auctionItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ezwel"		: pCM = window.open("/admin/etc/ezwel/ezwelItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "gmarket1010"	: pCM = window.open("/admin/etc/gmarket/gmarketItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "gseshop"		: pCM = window.open("/admin/etc/gsshop/gsshopItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");break;pCM.focus();
		case "interpark"	: pCM = window.open("/admin/etc/interpark/interparkItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "nvstorefarm"	: pCM = window.open("/admin/etc/nvstorefarm/nvstorefarmItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "nvstorefarmclass"	: pCM = window.open("/admin/etc/nvstorefarmclass/nvClassItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "nvstoremoonbangu"	: pCM = window.open("/admin/etc/nvstoremoonbangu/nvstoremoonbanguItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "WMP"			: pCM = window.open("/admin/etc/wmp/wmpItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wmpfashion"	: pCM = window.open("/admin/etc/wmpfashion/wmpfashionItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteimall"	: pCM = window.open("/admin/etc/ltimall/lotteiMallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "cjmall"		: pCM = window.open("/admin/etc/cjmall/cjmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "kakaogift"	: pCM = window.open("/admin/etc/gift/index.asp?gubun=giftting&itemid="+iitemid,"goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "kakaostore"	: pCM = window.open("/admin/etc/kakaostore/kakaostoreItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "11st1010"		: pCM = window.open("/admin/etc/11st/11stItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "ssg"			: pCM = window.open("/admin/etc/ssg/ssgItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "coupang"		: pCM = window.open("/admin/etc/coupang/coupangItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "hmall1010"	: pCM = window.open("/admin/etc/hmall/hmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lfmall"		: pCM = window.open("/admin/etc/lfmall/lfmallItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "shintvshopping"	: pCM = window.open("/admin/etc/shintvshopping/shintvshoppingItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wetoo1300k"	: pCM = window.open("/admin/etc/wetoo1300k/wetoo1300kItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "skstoa"		: pCM = window.open("/admin/etc/skstoa/skstoaItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "lotteon"		: pCM = window.open("/admin/etc/lotteon/lotteonItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "wconcept1010"	: pCM = window.open("/admin/etc/wconcept/wconceptItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		case "boribori1010"	: pCM = window.open("/admin/etc/boribori/boriboriItem.asp?research=on&itemid="+iitemid+"&isReged=A&sellyn=A","goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;		
		default				: pCM = window.open("/admin/etc/orderinput/xSiteItemLink.asp?sellsite="+isellsite+"&itemidarr="+iitemid,"goPopOutmall","width=1400,height=600,scrollbars=yes,resizable=yes");pCM.focus();break;
		
	}
}
function orderEditProcess() {
	var chkSel=0;
	try {
		if(frmlist.cksel.length>1) {
			for(var i=0;i<frmlist.cksel.length;i++) {
				if(frmlist.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmlist.cksel.checked) chkSel++;
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

    if (confirm('수정 하시겠습니까?')){
    	var mallId = "<%=sellsite%>";
        document.frmlist.target = "xLink";
        document.frmlist.cmdparam.value = "EDIT";
		switch(mallId){
			case "11st1010"		: document.frmlist.action = "<%=apiURL%>/outmall/11st/act11stReq.asp";  document.frmlist.submit(); break;
			case "auction1010"	: document.frmlist.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp";  document.frmlist.submit(); break;
			case "cjmall"		: document.frmlist.action = "<%=apiURL%>/outmall/cjmall/actCjmallReq.asp";  document.frmlist.submit(); break;
			case "ezwel"		: document.frmlist.action = "/admin/etc/ezwel/actezwelNewReq.asp";  document.frmlist.submit(); break;
			case "kakaostore"	: document.frmlist.action = "/admin/etc/kakaostore/actkakaostoreReq.asp";  document.frmlist.submit(); break;
			case "gseshop"		: document.frmlist.action = "<%=apiURL%>/outmall/gsshop/actgsshopReq.asp";  document.frmlist.submit(); break;
			case "interpark"	: document.frmlist.action = "<%=apiURL%>/outmall/interpark/actinterparkReq.asp";  document.frmlist.submit(); break;
			case "lotteimall"	: document.frmlist.action = "<%=apiURL%>/outmall/ltimall/actlotteiMallReq.asp";  document.frmlist.submit(); break;
			case "ssg"			: document.frmlist.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp";  document.frmlist.submit(); break;
			case "gmarket1010"	: document.frmlist.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp";  document.frmlist.submit(); break;
			case "nvstorefarm"	: document.frmlist.action = "<%=apiURL%>/outmall/nvstorefarm/actnvstorefarmReq.asp";  document.frmlist.submit(); break;
			case "nvstorefarmclass"	: document.frmlist.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp";  document.frmlist.submit(); break;
			case "nvstoremoonbangu"	: document.frmlist.action = "<%=apiURL%>/outmall/nvstoremoonbangu/actnvstoremoonbanguReq.asp";  document.frmlist.submit(); break;
			case "WMP"			: document.frmlist.action = "<%=apiURL%>/outmall/wmp/actWmpReq.asp";  document.frmlist.submit(); break;
			case "wmpfashion"	: document.frmlist.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp";  document.frmlist.submit(); break;
			case "hmall1010"	: document.frmlist.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp";  document.frmlist.submit(); break;
			case "lfmall"		: document.frmlist.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp";  document.frmlist.submit(); break;
			case "lotteon"		: document.frmlist.action = "<%=apiURL%>/outmall/lotteon/actlotteonReq.asp";  document.frmlist.submit(); break;
			case "wconcept1010"	: document.frmlist.action = "/admin/etc/wconcept/actwconceptReq.asp";  document.frmlist.submit(); break;
			case "shintvshopping"	: document.frmlist.action = "<%=apiURL%>/outmall/shintvshopping/actshintvshoppingReq.asp";  document.frmlist.submit(); break;
			case "wetoo1300k"	: document.frmlist.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp";  document.frmlist.submit(); break;
			case "skstoa"		: document.frmlist.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";  document.frmlist.submit(); break;
			case "WMP"			: document.frmlist.action = "<%=apiURL%>/outmall/wmp/actWmpReq.asp";  document.frmlist.submit(); break;
			case "boribori1010"		: document.frmlist.action = "/admin/etc/boribori/actBoriboriReq.asp";  document.frmlist.submit(); break;
		}
    }
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
			<option value="gmarket1010" <%= chkiif(sellsite = "gmarket1010", "selected", "") %> >G마켓</option>
			<option value="gseshop" <%= chkiif(sellsite = "gseshop", "selected", "") %> >GSShop</option>
			<option value="interpark" <%= chkiif(sellsite = "interpark", "selected", "") %> >인터파크</option>
			<option value="nvstorefarm" <%= chkiif(sellsite = "nvstorefarm", "selected", "") %> >스토어팜</option>
			<option value="nvstorefarmclass" <%= chkiif(sellsite = "nvstorefarmclass", "selected", "") %> >스토어팜클래스</option>
			<option value="nvstoremoonbangu" <%= chkiif(sellsite = "nvstoremoonbangu", "selected", "") %> >스토어팜문방구</option>
			<option value="WMP" <%= chkiif(sellsite = "WMP", "selected", "") %> >위메프</option>
			<option value="wmpfashion" <%= chkiif(sellsite = "wmpfashion", "selected", "") %> >위메프W패션</option>
			<option value="lotteimall" <%= chkiif(sellsite = "lotteimall", "selected", "") %> >롯데아이몰</option>
			<option value="lotteon" <%= chkiif(sellsite = "lotteon", "selected", "") %> >롯데On</option>
			<option value="shintvshopping" <%= chkiif(sellsite = "shintvshopping", "selected", "") %> >신세계TV쇼핑</option>
			<option value="skstoa" <%= chkiif(sellsite = "skstoa", "selected", "") %> >SKSTOA</option>
			<option value="wetoo1300k" <%= chkiif(sellsite = "wetoo1300k", "selected", "") %> >1300k</option>
			<option value="cjmall" <%= chkiif(sellsite = "cjmall", "selected", "") %> >CJMall</option>
			<option value="11st1010" <%= chkiif(sellsite = "11st1010", "selected", "") %> >11번가</option>
			<option value="ssg" <%= chkiif(sellsite = "ssg", "selected", "") %> >SSG</option>
			<option value="coupang" <%= chkiif(sellsite = "coupang", "selected", "") %> >쿠팡</option>
			<option value="hmall1010" <%= chkiif(sellsite = "hmall1010", "selected", "") %> >HMall</option>
			<option value="lfmall" <%= chkiif(sellsite = "lfmall", "selected", "") %> >LFmall</option>
			<option value="kakaostore" <%= chkiif(sellsite = "kakaostore", "selected", "") %> >카카오톡스토어</option>
			<option value="wconcept1010" <%= chkiif(sellsite = "wconcept1010", "selected", "") %> >W컨셉</option>
			<option value="boribori1010" <%= chkiif(sellsite = "boribori1010", "selected", "") %> >보리보리</option>
		</select>&nbsp;&nbsp;
		오차 :
		<select class="select" name="orderstate">
			<option value="">-전체-</option>
			<option value="S" <%= chkiif(orderstate = "S", "selected", "") %> >판매상태오차</option>
			<option value="P" <%= chkiif(orderstate = "P", "selected", "") %> >판매가격오차</option>
		</select>&nbsp;&nbsp;
		업로드날짜 :
		<input id="orderdate" name="orderdate" value="<%=orderdate%>" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		&nbsp;&nbsp;
		API몰만 :
		<select class="select" name="onlyapi">
			<option value="">-전체-</option>
			<option value="Y" <%= chkiif(onlyapi = "Y", "selected", "") %> >Y</option>
			<option value="N" <%= chkiif(onlyapi = "N", "selected", "") %> >N</option>
		</select>&nbsp;&nbsp;
		관리여부 :
		<select class="select" name="isok">
			<option value="">-전체-</option>
			<option value="Y" <%= chkiif(isok = "Y", "selected", "") %> >Y</option>
			<option value="N" <%= chkiif(isok = "N", "selected", "") %> >N</option>
		</select>&nbsp;&nbsp;
		<br /><br />
		<input id="getOrderdate" name="getOrderdate" value="" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="gDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<input type="button" value="가져오기" class="button" onclick="getOrder();">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br /><br />
<% If sellsite <> "" Then %>
<input class="button" type="button" id="btnCommcd" value="수정" onClick="orderEditProcess();" >
<% End If %>
<form name="frmlist" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oOrder.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOrder.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="2%"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmlist.cksel);"></td>
	<td width="5%">제휴몰</td>
	<td width="5%">텐바이텐<br>상품번호</td>
	<td width="5%">브랜드</td>
	<td width="5%">판매여부</td>
	<td width="5%">한정여부</td>
	<td width="5%">한정갯수</td>
	<td width="5%">옵션코드</td>
	<td width="5%">옵션판매여부</td>
	<td width="5%">옵션한정갯수</td>
	<td width="5%">옵션추가금액</td>
	<td width="5%">판매가</td>
	<td width="5%">제휴판매가</td>
	<td width="5%">오차금액</td>
	<td width="5%">제휴상품코드</td>
	<td width="5%">업로드날짜</td>
	<td width="5%">오차</td>
	<td width="4%">관리</td>
</tr>
<% For i=0 to oOrder.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOrder.FItemList(i).FItemID %>"></td>
	<td><%= oOrder.FItemList(i).FSellsite %></td>
	<td><a href="<%=vwwwURL%>/<%=oOrder.FItemList(i).FItemID%>" target="_blank"><%= oOrder.FItemList(i).FItemID %></a></td>
	<td><%= oOrder.FItemList(i).FMakerid %></td>
	<td><%= oOrder.FItemList(i).FSellyn %></td>
<% If oOrder.FItemList(i).FLimityn = "Y" Then %>
	<td><%= oOrder.FItemList(i).FLimityn %></td>
	<td>
	<%
		If oOrder.FItemList(i).FLimityn = "Y" Then
			response.write "<font color='Blue'>"&oOrder.FItemList(i).FLimitNo - oOrder.FItemList(i).FLimitSold&"</font>"
		Else
			response.write "비한정"
		End If
	%>
	</td>
<% Else %>
	<td colspan="2">한정아님</td>
<% End If %>
	<td>
	<%
		If oOrder.FItemList(i).FMatchitemoption <> "0000" Then
			response.write oOrder.FItemList(i).FMatchitemoption
		Else
			response.write "단품"
		End If
	%>
	</td>
<% If isnull(oOrder.FItemList(i).FOptsellyn) Then %>
	<td colspan="2"> </td>
<% Else %>
	<td><%= oOrder.FItemList(i).FOptsellyn %></td>
	<td>
	<%
		If oOrder.FItemList(i).FLimityn = "Y" Then
			response.write "<font color='Blue'>"&oOrder.FItemList(i).FOptlimitno - oOrder.FItemList(i).FOptlimitsold&"</font>"
		End If
	%>
	</td>
<% End If %>
	<td><%= Formatnumber(oOrder.FItemList(i).FOptaddprice, 0) %></td>
	<td style="cursor:pointer;" onclick="goPopOutmall('<%= oOrder.FItemList(i).FSellsite %>', '<%= oOrder.FItemList(i).FItemID %>');"><%= Formatnumber(oOrder.FItemList(i).FSellcash, 0) %></td>
	<td><%= Formatnumber(oOrder.FItemList(i).FOutmallsellprice, 0) %></td>
	<td>
		<%
			If oOrder.FItemList(i).FDiffprice <> "0" Then
				If oOrder.FItemList(i).FDiffprice > 0 Then
					response.write "<font color='red'>"&Formatnumber(oOrder.FItemList(i).FDiffprice, 0)&"</font>"
				Else
					response.write "<font color='blue'>"&Formatnumber(oOrder.FItemList(i).FDiffprice, 0)&"</font>"
				End If
			Else
				response.write Formatnumber(oOrder.FItemList(i).FDiffprice, 0)
			End If

		%>
	</td>
	<td>
	<%
		If Not(IsNULL(oOrder.FItemList(i).FOutMallGoodsNo)) Then
			Select Case oOrder.FItemList(i).FSellsite
				Case "auction1010"	Response.Write "<a target='_blank' href='http://itempage3.auction.co.kr/detailview.aspx?itemNo="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "ezwel"		Response.Write "<span style='cursor:pointer;' onclick=window.open('http://shop.ezwel.com/shopNew/goods/preview/goodsDetailView.ez?preview=yes&goodsBean.goodsCd="&oOrder.FItemList(i).FOutMallGoodsNo&"');>"&oOrder.FItemList(i).FOutMallGoodsNo&"</span>"
				Case "gmarket1010"	Response.Write "<a target='_blank' href='https://item.gmarket.co.kr/Item?goodscode="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "gseshop"		Response.Write "<span style='cursor:pointer;' onclick=window.open('http://www.gsshop.com/prd/prd.gs?prdid="&oOrder.FItemList(i).FOutMallGoodsNo&"');>"&oOrder.FItemList(i).FOutMallGoodsNo&"</span>"
				Case "interpark"	Response.Write "<a target='_blank' href='https://shopping.interpark.com/product/productInfo.do?prdNo="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "nvstorefarm"	Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbyten/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "nvstorefarmclass"	Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbytenclass/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "nvstoremoonbangu"	Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbytenclass/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "WMP"			Response.Write "<a target='_blank' href='https://front.wemakeprice.com/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "WMPfashion"	Response.Write "<a target='_blank' href='https://front.wemakeprice.com/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "lotteimall"	Response.Write "<a target='_blank' href='http://www.lotteimall.com/product/Product.jsp?i_code="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "cjmall"		Response.Write "<a target='_blank' href='http://www.oCJMall.com/prd/detail_cate.jsp?item_cd="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "11st1010"		Response.Write "<a target='_blank' href='http://www.11st.co.kr/product/SellerProductDetail.tmall?method=getSellerProductDetail&prdNo="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "hmall1010"	Response.Write "<a target='_blank' href='https://www.hyundaihmall.com/front/pda/itemPtc.do?slitmCd="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "lfmall"		Response.Write "<a target='_blank' href='https://www.lfmall.co.kr/product.do?cmd=getProductDetail&PROD_CD="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "kakaostore"	Response.Write "<a target='_blank' href='https://store.kakao.com/10x10/products/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "wconcept1010"	Response.Write "<a target='_blank' href='https://www.wconcept.co.kr/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "boribori1010"	Response.Write "<a target='_blank' href='https://www.boribori.co.kr/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "lotteon"		Response.Write "<a target='_blank' href='https://www.lotteon.com/p/product/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "shintvshopping"		Response.Write "<a target='_blank' href='https://www.shinsegaetvshopping.com/display/detail/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "skstoa"		Response.Write "<a target='_blank' href='http://www.skstoa.com/display/goods/"&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case "wetoo1300k"		Response.Write "<a target='_blank' href='http://www.1300k.com/shop/goodsDetail.html?f_goodsno="&oOrder.FItemList(i).FOutMallGoodsNo&"'>"&oOrder.FItemList(i).FOutMallGoodsNo&"</a>"
				Case Else 			Response.Write oOrder.FItemList(i).FOutMallGoodsNo
			End Select
		End If
	%>
	</td>
	<td><%= LEFT(oOrder.FItemList(i).FOrderdate, 10) %></td>
	<td>
		<%
			Select Case oOrder.FItemList(i).FOrderstate
				Case "S"	response.write "<font color='red'>품절</font>"
				Case "P"	response.write "<font color='blue'>가격</font>"
				Case Else	response.write "<font color='green'>미정의</font>"
			End Select
		%>
	</td>
	<td>
	<% If oOrder.FItemList(i).FIsOk = "Y" Then %>
		<input type="button" class="button"  value="완료" onclick="checkIsOk('<%=oOrder.FItemList(i).FIdx%>', 'N');" style="color:blue;font-weight:bold">
	<% Else %>
		<input type="button" class="button"  value="확인" onclick="checkIsOk('<%=oOrder.FItemList(i).FIdx%>', 'Y');">
	<% End If %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
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
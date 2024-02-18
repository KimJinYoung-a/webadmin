<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 오프라인 바코드출력(a4,formtec)
' Hieditor : 2009.04.07 서동석 생성
'			 2012.04.23 한용민 수정(소스 표준코딩으로 수정)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim itemgubun, isusingyn, research ,makerid, iitemid, barcode ,obarcode ,i ,makeriddispyn ,printpriceyn, page
	makerid = requestCheckVar(request("makerid"),32)
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),1)
	printpriceyn 	= requestCheckVar(request("printpriceyn"),1)
	iitemid = request("iitemid")
	barcode = request("barcode")
	research = request("research")
	itemgubun = request("itemgubun")
	isusingyn = request("isusingyn")
	page 			= requestCheckVar(request("page"),32)

if page="" then page="1"
if makeriddispyn = "" then makeriddispyn = "Y"
if (research="") and (isusingyn="") then isusingyn="Y"
if printpriceyn="" then printpriceyn="R"

'/매장
if (C_IS_SHOP) then
	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		'shopid = C_STREETSHOPID
	'end if
else
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

set obarcode = new COffShopItem
	obarcode.FCurrpage = page
	obarcode.FPageSize = 100
	obarcode.FRectItemgubun = itemgubun
	obarcode.FRectDesigner = makerid
	obarcode.FRectBarCode = barcode
	obarcode.FRectItemId = iitemid
	obarcode.FRectOnlyUsing = ChkIIF(isusingyn="Y","on","")
	obarcode.GetBarCodeList_paging

%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">

function SelectCk(opt){
	$(document.frmList.cksel).prop('checked',opt.checked);
}

function printBarcode(){
	var frm = document.frmList;

	if(!$(frm.cksel).is(':checked')) {
		alert('선택된 상품이 없습니다.');
		return;
	}

	var browser = navigator.userAgent.toLowerCase();
	if ( -1 != browser.indexOf('chrome') ){
		frm.target = "FrameCKP" ;
		frm.action = "/common/barcode/CssBarcodeprint_A4.asp" ;
		frm.submit() ;
	}else if ( -1 != browser.indexOf('trident') ){
		try{
			AddArr();
		}catch (e) {
			alert("- 도구 > 인터넷 옵션 > 보안 탭 > 신뢰할 수 있는 사이트 선택\n   1. 사이트 버튼 클릭 > 사이트 추가\n   2. 사용자 지정 수준 클릭 > 스크립팅하기 안전하지 않은 것으로 표시된 ActiveX 컨트롤 (사용)으로 체크\n\n※ 위 설정은 프린트 기능을 사용하기 위함임");
		}
	}else{
		alert("사용하시는 브라우저를 확인해주세요.\n- 크로미움 엔진을 사용한 브라우저(Chrome, Edge, Whale 등)에서 출력하실 수 있습니다.");
	}
}

function CheckThis(tn){
	var cksel = $("#frmList #cksel"+tn);
	cksel.prop("checked", true);
}

function AddData(itemid, itemoption, prdname, prdoptionname, socname, itemprice, itemtype, itemno){
	iaxobject.AddData(itemid, itemoption, prdname, prdoptionname, socname, itemprice, itemtype, itemno);
}

//AddData(v,'0000','아이템명','옵션명','브랜드',3000,'T','5')
function AddArr(){
	var makeriddisp;
	var printprice; var showpriceyn; var saleyn;
	var frm = document.frmList;

	iaxobject.ClearItem();
	//iaxobject.setTitleVisible(true);

	$("input[name='cksel']:checked").each(function(){
		var vid = $(this).val()-1; // 체크id

		//브랜드표시
		if (frm.makeriddispyn.value != 'N'){
			makeriddisp = makeriddisp = $(frm.socname).eq(vid).val();
		}else{
			makeriddisp = '';
		}

		//가격표시
		switch (frm.printpriceyn) {
			case 'C':	//할인가표시
				if(frm.saleyn.value=="Y") {
					//할인가
					printprice = $(frm.saleprice).eq(vid).val().trim();
				} else {
					//소비자가
					printprice = $(frm.customerprice).eq(vid).val().trim();
				}
				break;
			case 'R':	//판매가표시
				printprice = $(frm.sellprice).eq(vid).val().trim();
				break;
			default:
				//소비자가 표시
				printprice = $(frm.customerprice).eq(vid).val().trim();
				break;
		}

		// 데이터 추가
		if ($(frm.itemid).eq(vid).val()*1>=1000000){
			AddData($(frm.itemid).eq(vid).val(),
				$(frm.itemoption).eq(vid).val(),
				$(frm.prdname).eq(vid).val(),
				$(frm.prdoptionname).eq(vid).val(),
				makeriddisp, printprice,
				$(frm.itemgubun).eq(vid).val()*10,
				$(frm.itemno).eq(vid).val());
		}else{
			AddData($(frm.itemid).eq(vid).val(),
				$(frm.itemoption).eq(vid).val(),
				$(frm.prdname).eq(vid).val(),
				$(frm.prdoptionname).eq(vid).val(),
				makeriddisp, printprice,
				$(frm.itemgubun).eq(vid).val(),
				$(frm.itemno).eq(vid).val());
		}

	});

	iaxobject.ShowFrm();
}

function onlyNumberInput(){
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}

function reg(page) {
	if (frm.barcode.value != "") {
		if (frm.barcode.value.length < 12) {
			alert("잘못된 바코드입니다.");
			return;
		}
	}

	frm.page.value=page;
	frm.submit();
}

</script>

<OBJECT
	  id=iaxobject
	  classid="clsid:5D776FEA-8C6B-4C53-8EC3-3585FC040BDB"
	  codebase="http://webadmin.10x10.co.kr/common/cab/tenbarPrint.cab#version=1,0,0,29"
	  width=0
	  height=0
	  align=center
	  hspace=0
	  vspace=0
>
</OBJECT>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on" %>
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% if (C_IS_Maker_Upche) then %>
			* 브랜드 : <%= makerid %>
			<input type="hidden" name="makerid" value="<%= makerid %>">
			&nbsp;&nbsp;
		<% else %>
			* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;&nbsp;
		<% end if %>

		* 상품코드 : <input type="text" class="text" name="iitemid" value="<%= iitemid %>" maxlength="7" size="7" onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
		&nbsp;&nbsp;
		* 바코드 : <input type="text" class="text" name="barcode" value="<%= barcode %>" maxlength="14" size="14">
	<!--	&nbsp;
		주문코드 : <input type="text" class="text" name="" value="" maxlength="8" size="9">(코딩해야함)
    -->
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reg('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 상품구분:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		&nbsp;&nbsp;
		* 사용여부:
		<select class="select" name="isusingyn">
			<option value="">전체</option>
			<option value="Y" <%= CHKIIF(isusingyn="Y","selected","") %> >사용함</option>
		</select>
	 </td>
</tr>
</table>
</form>
<br>

<!-- 액션 시작 -->
<form name="frmList" id="frmList" method="POST" tyle="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ 프린터 설정 :
		<br>
		<input type="hidden" name="printpriceyn" value="R">
		<!--* 금액표시방식 :
		<select name="printpriceyn" id="printpriceyn">
			<option value="Y" <% 'if (printpriceyn = "Y") then %>selected<% 'end if %>>소비자가표시</option>
			<option value="C" <% 'if (printpriceyn = "C") then %>selected<% 'end if %>>할인가표시</option>
			<option value="R" <% 'if (printpriceyn = "R") then %>selected<% 'end if %>>판매가표시</option>
			<option value="S" <% 'if (printpriceyn = "S") then %>selected<% 'end if %>>심플금액표시</option>
			<option value="N" <% 'if (printpriceyn = "N") then %>selected<% 'end if %>>금액표시안함</option>
		</select>-->
		<select name="makeriddispyn" id="makeriddispyn">
			<option value="Y" <% if (makeriddispyn = "Y") then %>selected<% end if %>>브랜드표시</option>
			<option value="N" <% if (makeriddispyn = "N") then %>selected<% end if %>>브랜드표시안함</option>
		</select>
	</td>
	<td align="right">
		폼텍 용지 65칸 용 : LA-3100,LB-3100 등

		<% if obarcode.FResultCount>0 then %>
			<input type="button" class="button" value="선택상품 바코드출력" onclick="printBarcode()">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="15">
        검색결과 : <b><%= obarcode.FTotalCount %></b>
        &nbsp;
        <b><%= page %> / <%= obarcode.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="SelectCk(this)" id="ckall"></td>
	<td>이미지</td>
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td>판매가</td>
	<td>출력수량</td>
</tr>
<% if obarcode.FResultCount > 0 then %>
<% for i=0 to obarcode.FResultCount-1 %>
<input type="hidden" name="itemid" value="<%= obarcode.FItemList(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= obarcode.FItemList(i).Fitemoption %>">
<input type="hidden" name="prdname" value="<%= Replace(obarcode.FItemList(i).Fshopitemname,Chr(34),"") %>">
<input type="hidden" name="prdoptionname" value="<%= obarcode.FItemList(i).Fshopitemoptionname %>">
<input type="hidden" name="prdname_foreign" value="">
<input type="hidden" name="prdoptionname_foreign" value="">
<input type="hidden" name="socname" value="<%= Replace(obarcode.FItemList(i).FSocName,Chr(34),"") %>">
<input type="hidden" name="socname_kor" value="<%= Replace(obarcode.FItemList(i).FSocName_Kor,Chr(34),"") %>">
<input type="hidden" name="customerprice" value="<%= obarcode.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="sellprice" value="<%= obarcode.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="saleprice" value="<%= obarcode.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="saleyn" value="N">
<input type="hidden" name="itemgubun" value="<%= obarcode.FItemList(i).Fitemgubun %>">
<input type="hidden" name="prdcode" value="<%= BF_MakeTenBarcode(obarcode.FItemList(i).Fitemgubun, obarcode.FItemList(i).Fshopitemid, obarcode.FItemList(i).Fitemoption) %>">
<input type="hidden" name="generalbarcode" value="<%= obarcode.FItemList(i).Fgeneralbarcode %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width=30><input type="checkbox" id="cksel<%= i %>" name="cksel" value="<%=i+1%>" onClick="AnCheckClick(this);" ></td>
	<td width=50>
		<% if obarcode.FItemList(i).Fitemgubun="10" then %>
			<img src="<%= obarcode.FItemList(i).FimageSmall %>" width=50 height=50>
		<% else %>
			<img src="<%= obarcode.FItemList(i).FOffimgSmall %>" width=50 height=50>
		<% end if %>
	</td>
	<td width=100>
		<%= obarcode.FItemList(i).GetBarCode %>
	</td>
	<td><%= obarcode.FItemList(i).fmakerid %></td>
	<td align="left"><%= obarcode.FItemList(i).Fshopitemname %></td>
	<td width=80><%= obarcode.FItemList(i).Fshopitemoptionname %></td>
	<td align="right" width=80><%= FormatNumber(obarcode.FItemList(i).Fshopitemprice,0) %></td>
	<td width=80><input type="text" class="text" id="printno_<%= i %>" name="fixedno" value="1" maxlength="4" size="3" onKeyPress="CheckThis(<%= i %>)" onFocus="this.select()"></td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if obarcode.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=obarcode.StartScrollPage-1%>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + obarcode.StartScrollPage to obarcode.StartScrollPage + obarcode.FScrollCount - 1 %>
			<% if (i > obarcode.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(obarcode.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if obarcode.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="10" align="center">검색 결과가 없습니다.</td>
</tr>

<% end if %>
</table>
</form>

<%
set obarcode = Nothing

session.codePage = 949
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

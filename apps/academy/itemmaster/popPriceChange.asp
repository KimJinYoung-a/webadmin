<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Dim pageTitle
pageTitle="더핑거스 - 작품 가격수정"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/lectureadmin/lib/classes/itemcls_upche_2014.asp"-->  
<%
dim clsItem, itemid, i
dim arrList, intLoop
intLoop = 0
itemid = RequestCheckVar(request("itemid"),10)

if itemid="" then
	Response.Write "<script>alert('잘못된 파라메터');fnAPPclosePopup();</script>"
end if

if request.cookies("partner")("userid")="" then
	Response.Write "<script>alert('로그인이 필요합니다.');fnAPPclosePopup();</script>"
end if

'업체배송 상품정보 가져오기(텐배제외)	
set clsItem = new CItem
	clsItem.FRectMakerId = request.cookies("partner")("userid")
	clsItem.FRectItemId = itemid
	clsItem.FRectSort = "ID"
	clsItem.FRectCheckEX = "1"
	clsItem.FCurrPage		= 1
	clsItem.FPageSize		= 10
	if (clsItem.FRectMakerId<>"") then
		arrList = clsItem.fnGetItemUpcheBaesongList
	end if
set clsItem = nothing

if Not isArray(arrList) then
	Response.Write "<script>alert('상품이 없거나 변경할 수 없는 상품입니다.');fnAPPclosePopup();</script>"
	Response.end
end if
%>
<script type="text/javascript" src="/apps/academy/lib/confirm.js"></script>
<script>
	//공급가 자동설정
	function jsSetSupplyCash(){
		//공백체크,100원 이하체크
		document.frm.sellcash.value = document.frm.sellcash.value.replace(/\,/g,"");

		if(!IsDigit(document.frm.sellcash.value)){
			alert("판매가는 숫자만 입력 가능합니다.");
			document.frm.sellcash.focus();
			return;
		}

		document.frm.buycash.value =  document.frm.sellcash.value  - parseInt(document.frm.sellcash.value*document.frm.iMargin.value/100);

		// App의 버튼 활성화
		if(document.frm.etcStr.value!="") {
			fnAPPShowRightConfirmBtns();
		}
	}

	function fnAppCallWinConfirm(){
		if(!document.frm.sellcash.value) {
			alert("판매가를 입력해주요.");
			document.frm.sellcash.focus();
			return;
		}

		if(!document.frm.etcStr.value) {
			alert("상품수정변경사유를 입력해 주세요.");
			document.frm.etcStr.focus();
			return;
		}

		if(confirm("가격 변경을 요청하시겠습니까?\n변경 요청된 가격은 관리자의 승인 후 반영됩니다.")) {
			document.frm.submit();
		}
	}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">가격 변경 요청</h1>
			<div class="priceChange">
			<form name="frm" method="post" action="popPriceChange_Proc.asp">
			<input type="hidden" name="hidM" value="P">
			<input type="hidden" name="itemid" value="<%=itemid%>">
				<ul class="list">
					<li class="">
						<dfn><b>공급 마진</b></dfn>
						<input type="hidden" name="iMargin" value="<%=FormatNumber(1-(clng(arrList(16,intLoop))/clng(arrList(15,intLoop))))*100 %>">
						<div class="aftPrice"><%=fnPercent(arrList(16,intLoop),arrList(15,intLoop),1)%></div>
					</li>
					<li class="critical">
						<dfn><b>판매가</b></dfn>
						<input type="hidden" name="oldsellcash" value="<%=arrList(15,intLoop)%>">
						<div class="prePrice"><%=formatnumber(arrList(15,intLoop),0)%> 원</div>
						<div class="chgArw"></div>
						<div class="aftPrice"><input type="number" name="sellcash" value="" pattern="[0-9]*" placeholder="입력" onKeyUp="jsSetSupplyCash();"></div>
						<div style="width:1.6rem">원</div>
					</li>
					<li class="">
						<dfn><b>공급가 <span class="fs1-1r">(부가세 포함)</span></b></dfn>
						<input type="hidden" name="oldbuycash" value="<%=arrList(16,intLoop)%>">
						<div class="prePrice"><%=formatnumber(arrList(16,intLoop),0)%> 원</div>
						<div class="chgArw"></div>
						<div class="aftPrice"><input type="number" name="buycash" readonly value="" placeholder="<%=arrList(16,intLoop)%>" /></div>
						<div style="width:1.6rem">원</div>
					</li>
				</ul>
				<div class="linkInsert tMar2r">
					<textarea rows="8" name="etcStr" placeholder="변경사유를 입력해주세요." onKeyUp="jsSetSupplyCash();"></textarea>
				</div>
			</form>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
	fnAPPShowRightConfirmBtns();
});
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<%
'#############################################
' PageName : /diyshop/shop_prd.asp	
' Description : DIY Shop 상품상세
' History : 2016.07.11 이종화 생성
'#############################################
%>
<!-- #include virtual="/apps/academy/preview/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/apps/academy/preview/diyItemCategoryCls.asp" -->
<!-- #include virtual="/apps/academy/preview/diyItemInfoCls.asp" -->
<!-- #include virtual="/apps/academy/preview/diyItemPrdAppCls.asp" -->
<!-- #include virtual="/apps/academy/preview/diyItemOptionCls.asp" -->
<!-- #include virtual="/apps/academy/preview/diyItemEvaluateCls.asp" -->
<!-- #include virtual="/apps/academy/preview/searchcls.asp" -->
<!-- #include virtual="/apps/academy/preview/lecture_cls.asp" -->
<!-- #include virtual="/apps/academy/preview/commlib.asp" -->
<!-- #include virtual="/apps/academy/preview/function.asp" -->
<!-- #include virtual="/apps/academy/preview/inc_const.asp" -->
<!-- #include virtual="/apps/academy/preview/tenEncUtil.asp" -->
<%
	strPageTitle = "상세 정보"
	''//db
	dim itemid, i , lp, returnqna
	itemid = requestCheckVar(request("itemid"),6)
	returnqna=requestCheckVar(request("returnqna"),2)

	if itemid="" or itemid="0" then
		Call Alert_Return("상품번호가 없습니다.")
		response.End
	elseif Not(isNumeric(itemid)) then
		Call Alert_Return("잘못된 상품번호입니다.")
		response.End
	else
		'정수형태로 변환
		itemid=CLng(itemid)
	end If
	
	dim LoginUserid
	
	LoginUserid = GetEncLoginUserID()

	dim chkPlusItem : if chkPlusItem="" then chkPlusItem=false		'추가상품 존재 여부(inc_PlusDIYItem.asp안에서 값 지정)
	
	dim oItem
	set oItem = new DIYItemPrdCls
	oItem.GetItemData itemid , "N"

	if oItem.FResultCount=0 then
		Call Alert_Return("존재하지 않는 상품입니다.")
		response.End
	end if

	'// 파라메터 접수
	Dim vDisp
	vDisp = requestCheckVar(getNumeric(Request("disp")),18)
	if vDisp="" or (len(vDisp) mod 3)<>0 then vDisp = oItem.Prd.FcateCode

	'// 추가 이미지
	dim oADDimage
	set oADDimage = new DIYItemPrdCls
	oADDimage.FRectItemId = itemid
	oADDimage.GetOneItemAddImageList 

	'//상품 상세 이미지
	dim oADD
	set oADD = new DIYItemPrdCls
	oADD.getAddImage itemid ,"N"

	'//옵션 HTML생성
	dim ioptionBoxHtml
	IF (oitem.Prd.FOptionCnt>0) then
		ioptionBoxHtml = GetOptionBoxHTML2016(itemid, oitem.Prd.IsSoldOut)
	End IF

	function ImageExists(byval iimg)
		if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
			ImageExists = false
		else
			ImageExists = true
		end if
	end function

	'//상품설명 추가 (고시정보
	dim addEx
	set addEx = new DIYItemPrdCls
		addEx.getItemAddExplain itemid ,"N"

	Dim tempsource , tempsize

	tempsource = oItem.Prd.FItemSource
	tempsize = oItem.Prd.FItemSize

	'// 상품상세설명 동영상 추가
	Dim itemVideos
	Set itemVideos = New DIYItemPrdCls
		itemVideos.fnGetItemVideos itemid, "video1" , "N"

	Dim isEvaluateCnt , isQnACnt
	'//구매후기 카운트
		isEvaluateCnt = getIsEvaluateCnt(itemid)
	'//QnA카운트
		isQnACnt	= getItemIsQnACnt(itemid)

	'// 페이지 타이틀 및 페이지 설명 작성 (strHeadTitle)
	strHeadTitle = "더핑거스 : " & Replace(oItem.Prd.FItemName,"""","")
	strPageKeyword = "TheFingers"
	if Not(oItem.Prd.FKeyWords="" or isNull(oItem.Prd.FKeyWords)) then
		strPageKeyword = strPageKeyword & "," & oItem.Prd.FKeyWords
	else
		strPageKeyword = strPageKeyword & ",작품,수공예,핸드메이드"
	end if
	strPageDesc = "더핑거스 작품 - "& Replace(oItem.Prd.FItemName,"""","")		'페이지 설명
	strPageImage = oitem.Prd.FImageBasic		'페이지 요약 이미지(SNS 퍼가기용)
	strPageUrl = "http://www.thefingers.co.kr/diyshop/shop_prd.asp?itemid="& itemid	'페이지 URL(SNS 퍼가기용)
%>
<!-- #include virtual="/apps/academy/preview/head.asp" -->
<script>
$(function(){
	// restinfo show/hide
	$(".fingerDetail .restInfo dt").click(function(){
		if ($(this).parents("dl").hasClass("current")) {
			$(".fingerDetail .restInfo dl").removeClass('current');
		} else {
			$(".fingerDetail .restInfo dl").removeClass('current');
			$(this).parents("dl").addClass('current');
		}
	});

	// detail image swipe
	var swiper01 = new Swiper(".detailImg .swiper-container", {
		speed: 300,
		pagination: '.swiper-pagination',
		paginationClickable: true
	});

	// another item swipe
	var swiper02 = new Swiper(".anotherItem .thumbList .swiper-container", {
		speed: 300,
		freeMode: true,
		slidesPerView:'auto',
		pagination:false
	});

	// fingers choice swipe
	var swiper03 = new Swiper(".fingersChoice .thumbList .swiper-container", {
		speed: 300,
		freeMode: true,
		slidesPerView:'auto',
		pagination:false
	});

	// add option scroll
	var layerScroll01 = new Swiper(".addOption .scrollArea .swiper-container", {
		scrollbar:'.addOption .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});

	// select option
	$(".btnBuy").click(function() {
		$(".floatingBar").addClass("showOption");
		$(".layerMask").show();
	});
	$(".layerMask").click(function(){
		$(".floatingBar").removeClass("showOption");
	});
	var selectScroll01 = new Swiper(".select1 .scrollArea .swiper-container", {
		scrollbar:'.select1 .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});
	var selectScroll02 = new Swiper(".select2 .scrollArea .swiper-container", {
		scrollbar:'.select2 .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});
	var selectScroll03 = new Swiper(".select3 .scrollArea .swiper-container", {
		scrollbar:'.select3 .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});
	$(".selectbox p").click(function(){
		if ($(this).parent().parent(".selectWrap").hasClass("current")) {
			$(".selectWrap").removeClass("current");
			$(".odrOption").removeClass("openOpt");
		} else {
			$(".selectWrap").removeClass("current");
			$(this).parent().parent(".selectWrap").addClass('current');
			if ($(this).parent().parent().hasClass("focus"))
			{
				$(".odrOption").addClass("openOpt");
			}
		}
	});

	//옵션 선택
	//1번 2번 옵션 disabled
	$('.select1').addClass("focus");
	$('.select2 .scrollArea,.select3 .scrollArea').css('display','none');

	$(".selectbox li").click(function(){
		var selectedVal = $(this).text();
		var selValue = $(this).attr('value'); //<%' li값 찾기 %>
		var selid = parseInt($(this).closest(".scrollArea").find("input[name='item_option']").attr("id"))+1;
		$(this).parent().parent().find("input[name='item_option']").val(selValue); //<% 'li값 hidden에 넣기 li ul 상위에 hidden 위치 parent().parent() %>
		$(this).closest(".scrollArea").prev("p").text(selectedVal);
		$(".selectWrap").removeClass("current");
		setTimeout(function(){//<%' 간이장바구니 레이어 리사이즈 적용 %>
			if (selValue != ''){
				$('.select'+parseInt(selid)).removeClass("focus");
				$('.select'+parseInt(selid+1)+' .scrollArea').css('display',''); //<%' value 값이 있을경우 다음 옵션 보여줌 %>
				$('.select'+parseInt(selid+1)).addClass("current");
				if (selid == 1){ //클릭되어 넘어 갈때 옵션별 슬라이드 리로드
					if ($('.select'+parseInt(selid+1)+' .scrollArea').length > 0  ){
						selectScroll02.onResize();
					}
				}else if (selid == 2){
					if ($('.select'+parseInt(selid+1)+' .scrollArea').length > 0 ){
						selectScroll03.onResize();
					}
				}
			}else{
				$('.select'+parseInt(selid+1)+' .scrollArea').css('display','none'); //<%' value 값이 없거나 없는걸 눌렀을 경우 다음 옵션 안보여줌 %>
				$('.select'+parseInt(selid)).addClass("focus");
				$(".odrOption").removeClass("openOpt"); //간격 조절
			}
			//<%'제작 문구 없이 옵션 만으로 간이장바구니 들어 갈때 레이어 리사이즈%>
			if($('.addOption .option').length > 0){
				layerScroll01.onResize();
			}
		}, 0);
	});
	
	//확인버튼-간이장바구니 레이아웃
	$(".btnYgn").click(function(){
		if($('.addOption .option').length > 0){
			layerScroll01.onResize();
		}
	});

	// share popup
	var layerScroll02 = new Swiper(".sharePop .scrollArea .swiper-container", {
		scrollbar:'.sharePop .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});
	$(".btnShare").click(function() {
//		$(".sharePop").show();
//		$(".layerMask").show();
//		layerScroll02.onResize();
//		var lyrH = $(".layerPopup").outerHeight();
//		$(".layerPopup").css('margin-top',-lyrH/2);
	});

	// delivery charge popup
	var layerScroll03 = new Swiper(".deliveryPop .scrollArea .swiper-container", {
		scrollbar:'.deliveryPop .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});
	$(".btnDelivery").click(function() {
		$(".deliveryPop").show();
		$(".layerMask").show();
		layerScroll03.onResize();
		var lyrH = $(".layerPopup").outerHeight();
		$(".layerPopup").css('margin-top',-lyrH/2);
	});

	// custom period popup
	var layerScroll04 = new Swiper(".customPop .scrollArea .swiper-container", {
		scrollbar:'.customPop .swiper-scrollbar',
		direction:'vertical',
		slidesPerView:'auto',
		mousewheelControl: true,
		freeMode: true
	});
	$(".btnCustom").click(function() {
		$(".customPop").show();
		$(".layerMask").show();
		layerScroll04.onResize();
		var lyrH = $(".layerPopup").outerHeight();
		$(".layerPopup").css('margin-top',-lyrH/2);
	});

	$(window).resize(function() {
		var lyrH = $(".layerPopup").outerHeight();
		$(".layerPopup").css('margin-top',-lyrH/2);
	});
});
window.onload = function(){
	// floating tab
	var tabTop = $("#detailView").offset().top;
	$(window).scroll(function(){
		if( $(window).scrollTop()>=tabTop ) {
			$("#detailView").addClass("stickyTab");
		} else {
			$("#detailView").removeClass("stickyTab");
		}
	});
	$(window).resize(function() {
		tabTop = $("#detailView").offset().top;
	});
}
//-------------------------------------------------------------------------------------------
//상품 관련 가격관련 JS
//-------------------------------------------------------------------------------------------
function jsItemea(plusminus)
{
//	var vmin = parseInt(<%=chkIIF(oItem.Prd.IsLimitItemReal and oItem.Prd.FRemainCount<=0,"0",oItem.Prd.ForderMinNum)%>);
//	var vmax = parseInt(<%=chkIIF(oItem.Prd.IsLimitItemReal,CHKIIF(oItem.Prd.FRemainCount<=oItem.Prd.ForderMaxNum,oItem.Prd.FRemainCount,oItem.Prd.ForderMaxNum),oItem.Prd.ForderMaxNum)%>);

	var vmin = 1;
	var vmax = 100;

	var v = parseInt(sbagfrm.itemea.value);
	if(plusminus == "+") {
		v++;
		if(v > vmax) v--;
	}
	else if(plusminus == "-") {
		if(v > 1) {
			v--;
		} else {
			v = 1;
		}
		if(v < vmin) v++;
	}
	sbagfrm.itemea.value = v;
	sbagfrm.optItemEa.value = v;

	var p = parseInt(sbagfrm.itemPrice.value);

	$("#spTotalPrc").text(plusComma(parseInt(v * p))+"원");
	$("#subtot").text(plusComma(parseInt(v * p))+"원");
}
</script>
<script type="application/x-javascript" src="http://m.thefingers.co.kr/lib/js/diyitem_shoppingbag.js"></script>
<script type="application/x-javascript" src="http://m.thefingers.co.kr/lib/js/jquery.numspinner_m.js"></script>
<script type="text/javascript">
<!-- #include virtual="/apps/academy/preview/shop_prd_javascript.asp" -->
</script>
<script type="text/javaScript" src="http://m.thefingers.co.kr/lib/js/todayviewdiy.js"></script>
</head>
<body>
<div class="wrap">
	<div class="container headB bgGry1 diyDetail fingerDetail">
		<%' content %>
		<div class="content">
			<div class="detailImg <%=chkiif(Not(itemVideos.Prd.FvideoFullUrl="")," isVideo","")%>">
				<div class="swiper-container">
					<div class="swiper-wrapper">
						<% if ImageExists(oitem.Prd.FImageBasic) then %>
						<div class="swiper-slide">
							<img src="<%= oitem.Prd.FImageBasic %>" alt="<%= oItem.Prd.FItemName %>" />
							<%' badge %>
							<div class="fingerBadge">
								<% if (oItem.Prd.FItemDiv = "06") then %>
								<span class="badge custom"><em>주문제작</em></span>
								<% End If %>
								<% IF oItem.Prd.isLimitItem and not (oItem.Prd.isSoldout or oItem.Prd.isTempSoldOut) And oItem.Prd.FRemainCount < 100 Then %>
								<span class="badge limited"><em>한정 <% = oItem.Prd.FRemainCount %>개</em></span>
								<% End If %>
							</div>
							<%' badge %>
						</div>
						<% end if %>

						<% IF oAddimage.FResultCount > 0 THEN %>
						<% FOR i= 0 to oAddimage.FResultCount-1  %>
						<%
						IF oAddimage.FItemList(i).FIMGTYPE=0 and ImageExists(oAddimage.FItemList(i).FADDIMAGE_400) THEN
							If i = 3 Then Exit for
						%>
							<div class="swiper-slide"><img src="<%= oAddimage.FItemList(i).FADDIMAGE_400 %>" alt="" /></div>
						<% End IF %>
						<% NEXT %>
						<% END IF %>

						<% If Not(itemVideos.Prd.FvideoFullUrl="") Then %>
						<div class="swiper-slide">
							<div class="videoWrap">
								<div class="video">
									<iframe src="<%=itemVideos.Prd.FvideoUrl%>" frameborder="0" allowfullscreen></iframe>
								</div>
							</div>
						</div>
						<% End If %>
					</div>
					<div class="swiper-pagination"></div>
				</div>
			</div>
			<div class="fingerCont">
				<p class="title"><a href="/corner/lectureDetail.asp?lecturer_id=<%= oItem.Prd.FMakerid %>">[<%= UCase(oItem.Prd.FBrandName) %>]</a> <%= oItem.Prd.FItemName %></p>
				<div class="price">
					<% IF oItem.Prd.IsSaleItem Or oitem.Prd.isCouponItem THEN %>
					<p class="cRed2"><s><%= FormatNumber(oItem.Prd.getOrgPrice,0) %>원</s> <span class="cGry5 fs1-1r">(<%= FormatNumber(oItem.Prd.FMileage,0) %>P 적립)</span></p>
					<% Else %>
					<p class="cRed2"><%= FormatNumber(oItem.Prd.getOrgPrice,0) %>원 <span class="cGry5 fs1-1r">(<%= FormatNumber(oItem.Prd.FMileage,0) %>P 적립)</span></p>
					<% End If %>
					<% IF oItem.Prd.IsSaleItem THEN %>
					<p class="cRed2">할인가 <%= FormatNumber(oItem.Prd.getRealPrice,0) %>원 [<% = oItem.Prd.getSalePro %>]</p>
					<% End If %>
					<% if oitem.Prd.isCouponItem Then %>
					<p class="cYgn1">쿠폰가 <%= FormatNumber(oItem.Prd.GetCouponAssignPrice,0) %>원 [<%= oItem.Prd.GetCouponDiscountStr %>]</p>
					<% End If %>
					<% if oItem.Prd.IsSoldOut then %>
					<p class="cRed2">품절 되었습니다.</p>
					<% end if %>
					<% if oitem.Prd.isCouponItem Then %>
					<button type="button" class="btn btnM4 btnYgn" onclick="DlDiyItemCoupon('<%= oitem.Prd.FCurrItemCouponIdx %>','<%= Server.URLEncode(CurrURLQ()) %>');">쿠폰받기</button>
					<% End If %>
					<form name="frmC" method="post" action="/myfingers/DownloaditemCoupon_Process.asp" style="margin:0px;">
					<input type="hidden" name="itemcouponidx" value="" />
					</form>
				</div>
				<div class="btnGroup">
					<%' 관심작품 등록시 클래스 favorOn %>
					<button type="button" id="favbtn" class="btnFavor">
						<span>관심작품</span><em id="favbtnCnt"><%=oItem.Prd.FfavCount%></em>
					</button>
					<button type="button" class="btnShare"><span>공유하기</span></button>
				</div>
			</div>
			<%''// 작품정보 / 상품후기 / QnA %>
			<!-- #include virtual="/apps/academy/preview/shop_prd_tabs.asp" -->
		</div>

	

		<%' <!-- 배송비 안내 레이어팝업 --> %>
		<div class="layerPopup deliveryPop" style="display:none">
			<div class="layerCont">
				<h2>배송비 안내</h2>
				<button type="button" class="layerClose" onclick="closeLayer();"><span>닫기</span></button>
				<div class="scrollArea">
					<div class="swiper-container">
						<div class="swiper-wrapper">
							<div class="swiper-slide">
								<div class="itemAddInfo">
									<ul class="info1">
										<li><span>기본 배송비</span><%=FormatNumber(oItem.Prd.FDefaultDeliverPay,0)%>원</li>
										<li><span>무료배송 조건</span><%=FormatNumber(oItem.Prd.FDefaultFreeBeasongLimit,0)%>원 이상</li>
									</ul>
									<div class="info2">
										<%=html2db(nl2br(oItem.Prd.FOrderComment))%>
									</div>
								</div>
							</div>
						</div>
						<div class="swiper-scrollbar"></div>
					</div>
				</div>
			</div>
		</div>
		<%' <!--// 배송비 안내 레이어팝업 --> '%>

		<%' <!-- 제작기간 레이어팝업 --> '%>
		<div class="layerPopup customPop" style="display:none">
			<div class="layerCont">
				<h2>제작기간</h2>
				<button type="button" class="layerClose" onclick="closeLayer();"><span>닫기</span></button>
				<div class="scrollArea">
					<div class="swiper-container">
						<div class="swiper-wrapper">
							<div class="swiper-slide">
								<div class="itemAddInfo">
									<ul class="info1">
										<li><span>제작 및 발송기간</span><%=oItem.Prd.Frequiremakeday%>일 이내</li>
									</ul>
									<div class="info2">
										<%=nl2br(oItem.Prd.Frequirecontents)%>
									</div>
								</div>
							</div>
						</div>
						<div class="swiper-scrollbar"></div>
					</div>
				</div>
			</div>
		</div>
		<%' <!--// 제작기간 레이어팝업 --> %>
		<iframe src="" name="iiBagWin" frameborder="0" width="0" height="0"></iframe>
		<div id="tmpopt" style="display:none;"></div>
		<div id="tmpopLimit" style="display:none;"></div>
		<div id="tmpitemCnt" style="display:none;"></div>
		<!-- # include virtual="/apps/academy/preview/incFooter.asp" -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
<script>
	setTimeout(function(){
		$('html, body').animate({scrollTop:$(".content").offset().top}, 'fast');
	}, 300);
</script>
</html>
<% 
	Set oADD = nothing
	Set oADDimage = nothing
	set oItem = Nothing
	Set addEx = Nothing
	Set itemVideos = Nothing
%>
<!-- #include virtual="/apps/academy/preview/dbclose.asp" -->	
<%@ language=vbscript %>
<% option Explicit %>
<%
'###########################################################
' Description : 핑거스 다이 상품 등록 대기 상품 
' Hieditor : 2016.08.08 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/waitDIYitemCls_preView.asp"-->

<%
'response.write "수정중입니다. "
'dbget.close()	:	response.End

Dim oupcheitemedit,ix,page
page = request("page")

if page="" then page=1


dim i
dim cdl_mnu,cdm_mnu,cdn_mnu
dim cdl,cdm,cdn
cdl = request("cdl")
cdm = request("cdm")
cdn = request("cdn")

dim cmi,cmj,k

dim itemid, oitem, designer
itemid = requestCheckvar(request("itemid"),10)

'// 상품설명 이미지
dim oADD
set oADD = new CWaitItemAddImage
'oADD.FRectItemId = itemid
oADD.getAddImage itemid
'oADD.GetOneItemAddImageList 

'// 추가 이미지
dim oADDimage
set oADDimage = new CWaitItemAddImage
oADDimage.FRectItemId = itemid
oADDimage.GetOneItemAddImageList 

set oitem = New CWaitItem
oitem.FRectItemid = itemid

if (C_IS_Maker_Upche) then
    oitem.FRectMakerid = session("ssBctID")
end if

oitem.GetOneItem

'designer = oitem.FOneItem.Fmakerid

if (oitem.FResultCount<1) then
    response.write "검색 결과가 없습니다."
    response.End
end if

function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

'// 추가 이미지-메인 이미지
Function getFirstAddimage()
	if ImageExists(oitem.Prd.FImageBasic) then
		getFirstAddimage= oitem.Prd.FImageBasic

	elseif (oAdd.FResultCount>0) then
		if ImageExists(oAdd.FADD(0).FAddimage) then
			getFirstAddimage= oAdd.FADD(0).FAddimage
		end if
	else
		getFirstAddimage= oitem.Prd.FImageMain
	end if
end Function

dim iOptionBoxHTML

iOptionBoxHTML = getOptionBoxHTML_FrontType(oitem.FOneItem.FWaititemid)

''동영상
Dim itemVideos
Set itemVideos = New CWaitItem
	itemVideos.fnGetItemVideos oitem.FOneItem.FWaititemid, "video1"
	
	
'//상품설명 추가 (고시정보
dim addEx
set addEx = new CWaitItem
	addEx.getItemAddExplain itemid
%>

<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/jquery-publishing.js"></script>
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/fingerscommon.js"></script>
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/swiper-2.1.min.js"></script>
<script type="text/javascript" src="/academy/itemmaster/viewDIYitem/js/jquery.slides.min.js"></script>
<link rel="stylesheet" type="text/css" href="/academy/itemmaster/viewDIYitem/css/common.css" />
<link rel="stylesheet" type="text/css" href="/academy/itemmaster/viewDIYitem/css/content.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/academy/itemmaster/viewDIYitem/css/ie.css" />
<![endif]-->
<!--[if IE 6] //-->
<title>더 핑거스</title>
<script>
$(function(){
	// slide
	if($('.imgSlide div').length>1) {
		$('.imgSlide').slidesjs({
			width:390,
			height:260,
			navigation:false,
			pagination:{active:true, effect:"fade"},
			effect:{
				fade:{speed:200, crossfade:true}
			}
		});
		$('.imgSlide .slidesjs-pagination > li a').append('<span></span>');
		$('.imgSlide .slidesjs-pagination > li a').mouseenter(function(){
			$('a[data-slidesjs-item="' + $(this).attr("data-slidesjs-item") + '"]').trigger('click');
		});
	} else {
		$('.imgSlide').append('<ul class="slidesjs-pagination"><li><a href="" class="active" onclick="return false;"><span></span></a></li></ul>');
	}

	// for dev msg : 슬라이드 이미지 갯수만큼 아래 background-image 에 넣어주세요(텐텐 상품상세와 동일) / 이미지사이즈 72px * 48px
	<% if ImageExists(oitem.FOneItem.FBasicImage) then %>
		$('.imgSlide .slidesjs-pagination > li').eq(0).children("a").css('background-image', 'url(<%= oitem.FOneItem.FBasicImage %>)');
	<% end if %>
	<% IF oAddimage.FResultCount > 0 THEN %>
	<% FOR i= 0 to oAddimage.FResultCount-1 %>
	<%
	IF oAddimage.FItemList(i).FIMGTYPE=0 and ImageExists(oAddimage.FItemList(i).FADDIMAGE_400) THEN
		If i = 3 Then Exit for
	%>
		$('.imgSlide .slidesjs-pagination > li').eq(<%=i+1 %>).children("a").css('background-image', 'url(<%= oAddimage.FItemList(i).FADDIMAGE_400 %>)');
	<% End IF %>
	<% NEXT %>
	<% END IF %>
    $('.imgSlide.isVideo .slidesjs-pagination > li:last-child').children("a").css('background-image', 'url(http://image.thefingers.co.kr/2016/common/bg_video.png)');
});
</script>
</head>
<body style="background:none; padding-bottom:50px;">

<div class="container fingerDetailV16">
	<!-- include virtual="/html/lib/inc/incHeader.asp" -->
	<div id="contentWrap">
		<div class="innerContent fullRange productDetailV16">
			<div class="contents">
				<div class="detailInfoV16" style="border-top:0;">
					<div class="detailImgV16">
						<div class="imgSlide <%=chkiif(Not(itemVideos.FOneItem.FvideoFullUrl="")," isVideo","")%>">
						<% if ImageExists(oitem.FOneItem.FBasicImage) then %>
							<div><img src="<%= oitem.FOneItem.FBasicImage %>" alt="" /></div>
						<% end if %>
						<% IF oAddimage.FResultCount > 0 THEN %>
						<% FOR i= 0 to oAddimage.FResultCount-1  %>
						<%
						IF oAddimage.FItemList(i).FIMGTYPE=0 and ImageExists(oAddimage.FItemList(i).FADDIMAGE_400) THEN
							If i = 3 Then Exit for
						%>
							<div><img src="<%= oAddimage.FItemList(i).FADDIMAGE_400 %>" alt="" /></div>
						<% End IF %>
						<% NEXT %>
						<% END IF %>
						<% If Not(itemVideos.FOneItem.FvideoFullUrl="") Then %>
						<div><iframe width="390" height="260" src="<%=itemVideos.FOneItem.FvideoUrl%>" frameborder="0" allowfullscreen></iframe></div>
						<% End If %>
						</div>
					</div>
					<!--// 작품 이미지 -->

					<!-- 가격 및 할인, 옵션선택 -->
					<div class="saleInfoV16">
						<div class="fingerBadge">
							<% if oitem.FOneItem.IsSoldOut then %>
								<span class="badge soldout"><em>일시품절</em></span>
							<% else %>
								<% if oitem.FOneItem.IsSaleItem then %>
									<span class="badge custom"><em>주문제작</em></span>
								<% end if %>

								<% if oitem.FOneItem.IsLimitItem and not (oItem.FOneItem.isSoldout or oItem.FOneItem.isTempSoldOut) then %>
									<span class="badge limited"><em>한정 <% = oitem.FOneItem.FRemainCount %>개</em></span>
								<% end if %>
								
							<% end if %>
						</div>
						<h2><%= UCase(oitem.FOneItem.Fitemname) %></h2>
						<div class="saleGroupV16 tMar10">
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_sale_price.gif" alt="판매가" /></dt>
								<dd><strong class="cBlk1"><%= FormatNumber(oitem.FOneItem.getOrgPrice,0) %>원</strong></dd>
							</dl>

    						<% IF oitem.FOneItem.IsSaleItem THEN %>
								<dl>
									<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_discount.gif" alt="할인판매가" /></dt>
									<dd><strong class="cRed1"><%= FormatNumber(oitem.FOneItem.getRealPrice,0) %>원 [<% = oitem.FOneItem.getSalePro %>]</strong></dd>
								</dl>
    						<% End If %>

							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_mileage.gif" alt="마일리지" /></dt>
								<dd><strong class="cGry1"><% = oitem.FOneItem.FMileage %> Point</strong></dd>
							</dl>
						</div>
						<div class="saleGroupV16">
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_product_code.gif" alt="작품코드" /></dt>
								<dd><%= itemid %></dd>
							</dl>
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_delivery_group.gif" alt="배송구분" /></dt>
								<dd class="cRed1"><% = oitem.FOneItem.GetDeliveryName %></dd>
							</dl>
						</div>

						<% if oitem.FOneItem.IsLimitItem and not (oItem.FOneItem.isSoldout or oItem.FOneItem.isTempSoldOut) then %>
							<div class="saleGroupV16">
								<dl>
									<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_limit_product.gif" alt="한정판매상품" /></dt>
									<dd><strong><% = oitem.FOneItem.FRemainCount %>개</strong> 남았습니다.</dd>
								</dl>
							</div>
						<% end if %>

						<div class="saleGroupV16">
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_order_amount.gif" alt="주문수량" /></dt>
								<dd>
									<div class="spinnerV16">
										<input type="text" class="txtBasic" value="1" />
										<div class="buttons">
											<div class="up">갯수 늘리기</div>
											<div class="down">갯수 줄이기</div>
										</div>
									</div>
								</dd>
							</dl>
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_option.gif" alt="옵션선택" /></dt>
								<dd>
									<div>
										<%= ioptionBoxHtml %>
									</div>
								</dd>
							</dl>
						</div>

						<% if oitem.FOneItem.FItemDiv = "06" then %>
							<div class="saleGroupV16">
								<dl>
									<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_custom_message.gif" alt="제작메세지" /></dt>
									<dd><textarea cols="20" rows="5" placeholder="제작 메세지를 입력해주세요" style="width:228px;" class="cBlk1"></textarea></dd>
								</dl>
							</div>
						<% End If %>

						<div class="btnGroupV16">
							<div><button type="button" class="btn btnM1 btnRed">바로구매</button></div>
							<div><button type="button" class="btn btnM1 btnWht">장바구니</button></div>
						</div>
					</div>

				</div>

				<div class="detailContWrapV16" style="background:none;">
					<div class="detailContV16">
						<!-- 상품정보 -->
						<div id="itemInfo" class="itemInfo">
							<div class="tab1">
								<ul>
									<li class="current"><a href="#itemInfo">작품정보</a></li>
								</ul>
							</div>
							<div class="tabCont">
								<!-- 상품 설명 입력영역 -->
								<div class="detailAreaV16">
									<% IF oAdd.FResultCount > 0 THEN %>
									<% FOR i= 0 to oAdd.FResultCount-1  %>
										<% IF oAdd.FItemList(i).FAddImageType=2 THEN %>
										<div class="image"><img src="<%= oAdd.FItemList(i).FAddimage %>" alt="<%= oitem.FOneItem.Fitemname %>" /></div>
										<% If oAdd.FItemList(i).FAddimgText <> "" Then %>
										<div class="txt"><%=nl2br(oAdd.FItemList(i).FAddimgText)%></div>
										<% Else %>
										<br>
										<% End IF %>
										<% End IF %>
									<% NEXT %>
									<% END IF %>

								</div>
								<!--// 상품 설명 입력영역 -->

								<div class="restInfo">
									<dl class="type1">
										<dt>배송비 안내</dt>
										<dd>
											<ul>
												<li><strong>기본 배송비</strong><%=FormatNumber(oItem.FOneItem.FDefaultDeliverPay,0)%>원</li>
												<li><strong>무료 배송 조건</strong><%=FormatNumber(oItem.FOneItem.FDefaultFreeBeasongLimit,0)%>원 이상</li>
											</ul>
											<p><%=nl2br(oItem.FOneItem.FOrderComment)%></p>
											<!-- p>배송기간은 주문일로부터 1일~5일정도 소요됩니다. 업체배송 상품은 무료배송 되며, 업체조건배송 상품은 특정 브랜드 배송 기준으로 배송비가 부여되며 업체착불 배송은 특정 브랜드 배송기준으로 고객님의 배송지에 따라 배송비가 착불로 부과됩니다.</p -->
										</dd>
									</dl>

									<% if (oItem.FOneItem.FItemDiv = "06") then %>
										<dl class="type1 lMar30">
											<dt>제작기간 안내</dt>
											<dd>
												<ul>
													<li><strong>제작 및 발송기간</strong><%=oItem.FOneItem.Frequiremakeday%>일 이내</li>
												</ul>
												<p><%=oItem.FOneItem.Frequirecontents%></p>
											</dd>
										</dl>
									<% End If %>

									<dl class="type2">
										<dt>상품 필수 정보<span>전자상거래 등에서의 상품정보 제공 고시에 따라 작성 되었습니다.</span></dt>
										<dd>
											<div class="list">
												<%
												IF addEx.FResultCount > 0 THEN
													FOR i= 0 to addEx.FResultCount-1
												%>
													<span>- <strong><%=addEx.FItemList(i).FInfoname%> :</strong><%=addEx.FItemList(i).FInfoContent%></span><br>
												<%
														Next
													End If
												%>
											</div>
										</dd>
									</dl>
									<dl class="type2">
										<dt>교환/환불 정책</dt>
										<dd>
										    <p><%=(nl2br(oItem.FOneItem.Frefundpolicy))%></p>
											<!-- <p>상품 수령일로부터 7일 이내 반품/환불 가능합니다. 변심 반품의 경우 왕복배송비를 차감한 금액이 환불되며, 제품 및 포장 상태가 재판매 가능하여야 합니다. 상품 불량인 경우는 배송비를 포함한 전액이 환불됩니다. 출고 이후 환불요청 시 상품 회수 후 처리됩니다. 완제품으로 수입된 상품의 경우 A/S가 불가합니다. 특정브랜드의 교환/환불/AS에 대한 개별기준이 상품페이지에 있는 경우 브랜드의 개별기준이 우선 적용 됩니다.</p> -->
										</dd>
									</dl>
								</div>
							</div>
						</div>
					</div>
				</div>
				<!--// 작품 상세 -->
			</div>
		</div>
	</div>
	<!-- include virtual="/html/lib/inc/incFooter.asp" -->
</div>
</body>
</html>

<%
Set itemVideos = Nothing
set oitem = Nothing
set oADD = Nothing
set addEx = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
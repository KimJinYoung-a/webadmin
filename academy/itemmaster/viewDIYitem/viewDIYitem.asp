<%@ language=vbscript %>
<% option Explicit %>
<%
'###########################################################
' Description : �ΰŽ� ���� ��ǰ ��� ��� ��ǰ 
' Hieditor : 2016.08.08 �ѿ�� ����
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
'response.write "�������Դϴ�. "
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

'// ��ǰ���� �̹���
dim oADD
set oADD = new CWaitItemAddImage
'oADD.FRectItemId = itemid
oADD.getAddImage itemid
'oADD.GetOneItemAddImageList 

'// �߰� �̹���
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
    response.write "�˻� ����� �����ϴ�."
    response.End
end if

function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

'// �߰� �̹���-���� �̹���
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

''������
Dim itemVideos
Set itemVideos = New CWaitItem
	itemVideos.fnGetItemVideos oitem.FOneItem.FWaititemid, "video1"
	
	
'//��ǰ���� �߰� (�������
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
<title>�� �ΰŽ�</title>
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

	// for dev msg : �����̵� �̹��� ������ŭ �Ʒ� background-image �� �־��ּ���(���� ��ǰ�󼼿� ����) / �̹��������� 72px * 48px
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
					<!--// ��ǰ �̹��� -->

					<!-- ���� �� ����, �ɼǼ��� -->
					<div class="saleInfoV16">
						<div class="fingerBadge">
							<% if oitem.FOneItem.IsSoldOut then %>
								<span class="badge soldout"><em>�Ͻ�ǰ��</em></span>
							<% else %>
								<% if oitem.FOneItem.IsSaleItem then %>
									<span class="badge custom"><em>�ֹ�����</em></span>
								<% end if %>

								<% if oitem.FOneItem.IsLimitItem and not (oItem.FOneItem.isSoldout or oItem.FOneItem.isTempSoldOut) then %>
									<span class="badge limited"><em>���� <% = oitem.FOneItem.FRemainCount %>��</em></span>
								<% end if %>
								
							<% end if %>
						</div>
						<h2><%= UCase(oitem.FOneItem.Fitemname) %></h2>
						<div class="saleGroupV16 tMar10">
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_sale_price.gif" alt="�ǸŰ�" /></dt>
								<dd><strong class="cBlk1"><%= FormatNumber(oitem.FOneItem.getOrgPrice,0) %>��</strong></dd>
							</dl>

    						<% IF oitem.FOneItem.IsSaleItem THEN %>
								<dl>
									<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_discount.gif" alt="�����ǸŰ�" /></dt>
									<dd><strong class="cRed1"><%= FormatNumber(oitem.FOneItem.getRealPrice,0) %>�� [<% = oitem.FOneItem.getSalePro %>]</strong></dd>
								</dl>
    						<% End If %>

							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_mileage.gif" alt="���ϸ���" /></dt>
								<dd><strong class="cGry1"><% = oitem.FOneItem.FMileage %> Point</strong></dd>
							</dl>
						</div>
						<div class="saleGroupV16">
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_product_code.gif" alt="��ǰ�ڵ�" /></dt>
								<dd><%= itemid %></dd>
							</dl>
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_delivery_group.gif" alt="��۱���" /></dt>
								<dd class="cRed1"><% = oitem.FOneItem.GetDeliveryName %></dd>
							</dl>
						</div>

						<% if oitem.FOneItem.IsLimitItem and not (oItem.FOneItem.isSoldout or oItem.FOneItem.isTempSoldOut) then %>
							<div class="saleGroupV16">
								<dl>
									<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_limit_product.gif" alt="�����ǸŻ�ǰ" /></dt>
									<dd><strong><% = oitem.FOneItem.FRemainCount %>��</strong> ���ҽ��ϴ�.</dd>
								</dl>
							</div>
						<% end if %>

						<div class="saleGroupV16">
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_order_amount.gif" alt="�ֹ�����" /></dt>
								<dd>
									<div class="spinnerV16">
										<input type="text" class="txtBasic" value="1" />
										<div class="buttons">
											<div class="up">���� �ø���</div>
											<div class="down">���� ���̱�</div>
										</div>
									</div>
								</dd>
							</dl>
							<dl>
								<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_option.gif" alt="�ɼǼ���" /></dt>
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
									<dt><img src="http://image.thefingers.co.kr/academy2012/common/title/conttit_custom_message.gif" alt="���۸޼���" /></dt>
									<dd><textarea cols="20" rows="5" placeholder="���� �޼����� �Է����ּ���" style="width:228px;" class="cBlk1"></textarea></dd>
								</dl>
							</div>
						<% End If %>

						<div class="btnGroupV16">
							<div><button type="button" class="btn btnM1 btnRed">�ٷα���</button></div>
							<div><button type="button" class="btn btnM1 btnWht">��ٱ���</button></div>
						</div>
					</div>

				</div>

				<div class="detailContWrapV16" style="background:none;">
					<div class="detailContV16">
						<!-- ��ǰ���� -->
						<div id="itemInfo" class="itemInfo">
							<div class="tab1">
								<ul>
									<li class="current"><a href="#itemInfo">��ǰ����</a></li>
								</ul>
							</div>
							<div class="tabCont">
								<!-- ��ǰ ���� �Է¿��� -->
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
								<!--// ��ǰ ���� �Է¿��� -->

								<div class="restInfo">
									<dl class="type1">
										<dt>��ۺ� �ȳ�</dt>
										<dd>
											<ul>
												<li><strong>�⺻ ��ۺ�</strong><%=FormatNumber(oItem.FOneItem.FDefaultDeliverPay,0)%>��</li>
												<li><strong>���� ��� ����</strong><%=FormatNumber(oItem.FOneItem.FDefaultFreeBeasongLimit,0)%>�� �̻�</li>
											</ul>
											<p><%=nl2br(oItem.FOneItem.FOrderComment)%></p>
											<!-- p>��۱Ⱓ�� �ֹ��Ϸκ��� 1��~5������ �ҿ�˴ϴ�. ��ü��� ��ǰ�� ������ �Ǹ�, ��ü���ǹ�� ��ǰ�� Ư�� �귣�� ��� �������� ��ۺ� �ο��Ǹ� ��ü���� ����� Ư�� �귣�� ��۱������� ������ ������� ���� ��ۺ� ���ҷ� �ΰ��˴ϴ�.</p -->
										</dd>
									</dl>

									<% if (oItem.FOneItem.FItemDiv = "06") then %>
										<dl class="type1 lMar30">
											<dt>���۱Ⱓ �ȳ�</dt>
											<dd>
												<ul>
													<li><strong>���� �� �߼۱Ⱓ</strong><%=oItem.FOneItem.Frequiremakeday%>�� �̳�</li>
												</ul>
												<p><%=oItem.FOneItem.Frequirecontents%></p>
											</dd>
										</dl>
									<% End If %>

									<dl class="type2">
										<dt>��ǰ �ʼ� ����<span>���ڻ�ŷ� ����� ��ǰ���� ���� ��ÿ� ���� �ۼ� �Ǿ����ϴ�.</span></dt>
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
										<dt>��ȯ/ȯ�� ��å</dt>
										<dd>
										    <p><%=(nl2br(oItem.FOneItem.Frefundpolicy))%></p>
											<!-- <p>��ǰ �����Ϸκ��� 7�� �̳� ��ǰ/ȯ�� �����մϴ�. ���� ��ǰ�� ��� �պ���ۺ� ������ �ݾ��� ȯ�ҵǸ�, ��ǰ �� ���� ���°� ���Ǹ� �����Ͽ��� �մϴ�. ��ǰ �ҷ��� ���� ��ۺ� ������ ������ ȯ�ҵ˴ϴ�. ��� ���� ȯ�ҿ�û �� ��ǰ ȸ�� �� ó���˴ϴ�. ����ǰ���� ���Ե� ��ǰ�� ��� A/S�� �Ұ��մϴ�. Ư���귣���� ��ȯ/ȯ��/AS�� ���� ���������� ��ǰ�������� �ִ� ��� �귣���� ���������� �켱 ���� �˴ϴ�.</p> -->
										</dd>
									</dl>
								</div>
							</div>
						</div>
					</div>
				</div>
				<!--// ��ǰ �� -->
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
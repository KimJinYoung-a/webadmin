<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader_xhtml.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim i, j
dim getPdfDownLinkUrlAdm
dim orderserial, ekey,webImgUrl, webImgSSLUrl
orderserial = requestCheckvar(request("vos"),16)
ekey =  requestCheckvar(request("ekey"),32)

if (ekey="") then
    response.write "암호화 키가 올바르지 않습니다.1"
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5(orderserial))) then
    response.write "암호화 키가 올바르지 않습니다.2"
    response.end
end if

IF application("Svr_Info")="Dev" THEN
	webImgUrl		= "http://testwebimage.10x10.co.kr"
	webImgSSLUrl	= "http://testwebimage.10x10.co.kr"
ELSE
	webImgUrl		= "http://webimage.10x10.co.kr"
	webImgSSLUrl	= "http://webimage.10x10.co.kr"
END IF

dim oordermaster, oorderdetail
dim addparam

set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial
oordermaster.QuickSearchOrderMaster

if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if

set oorderdetail = new COrderMaster
oorderdetail.FRectOldOrder = oordermaster.FRectOldOrder
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

dim totalBeasongPay : totalBeasongPay = 0

%>
<style>
		#orderPrint .baseTable th, #orderPrint .baseTable td {padding:4px 0;}
		.companyInfo {margin:40px 0 0 0;}
		.popWrap .popHeader {background:transparent; text-align:center; font-size:20px;}
		h1 {font-size:24px; font-weight:bold;}
		h1 img {display:none;}
		h1 span {display:block;}
		.fs14 {font-size:12px;}
		.estimateTo {font-size:11px;}
		.estimateTo dl {overflow:hidden; _zoom:1; margin-top:3px;}
		.estimateTo dl.tMar20 {margin-top:10px;}
		.estimateTo dl dt, .estimateTo dl dd {float:left;}
		.estimateTo dl dt {width:70px;}
		#orderPrint h2 {font-size:12px; font-weight:bold; margin:20px 0 5px 0;}
	</style>
<script type="text/javascript">
	function jsGoPDF(iUri){
		  var popwin = window.open(iUri,'dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes');
	}
</script>
	<div class="heightgird" id="orderPrint"><!-- 2013.09.24 : id="mediaPrint" 추가 -->
		<div class="popWrap">
			<div class="popHeader">
				<h1><span>견 적 서</span></h1>
			</div>
			<div class="popContent">
				<!-- content -->
				<div class="mySection">
					<div class="orderDetail">
						<div class="overHidden">
							<div class="ftLt estimateTo" style="width:45%;">
								<dl>
									<dt>견 적 일 :</dt>
									<dd><%= left(oordermaster.FOneItem.FRegDate,10) %></dd>
								</dl>
								<dl>
									<dt>회 사 명 :</dt>
									<dd><%= oordermaster.FOneItem.FBuyName %></dd>
								</dl>
								<dl>
									<dt>담 당 자 :</dt>
									<dd><%= oordermaster.FOneItem.FBuyName %></dd>
								</dl>
								<dl>
									<dt>전화번호 :</dt>
									<dd><%= oordermaster.FOneItem.FBuyPhone %>/ <%= oordermaster.FOneItem.FBuyHp %></dd>
								</dl>
								<dl>
									<dt>팩스번호 :</dt>
									<dd></dd>
								</dl>
								<dl>
									<dt>이 메 일 :</dt>
									<dd><%=oordermaster.FOneItem.Fbuyemail%></dd>
								</dl>
								<dl class="tMar20">
									<dt><strong>주문번호 :</strong></dt>
									<dd><strong><%= oordermaster.FOneItem.FOrderSerial %>
										 <% If oordermaster.FOneItem.IsForeignDeliver Then %>
									 		 (<strong>해외배송</strong>)
								 		 <% End If %></strong></dd>
								</dl>
								<dl>
									<dt><strong>합계금액 :</strong></dt>
									<dd><strong>￦ <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></strong></dd>
								</dl>
							</div>
							<div class="ftRt" style="width:55%;">
								<div class="title">
									<h2 class="ftLt" style="margin-top:0;">공급자 정보</h2>
								</div>
								<table class="baseTable btmLine">
									<caption>공급자 정보</caption>
									<colgroup>
										<col width="18%" /> <col width="30%" /> <col width="18%" /> <col width="30%" />
									</colgroup>
									<tbody>
									<tr>
										<th scope="row">등록번호</th>
										<td colspan="3">211-87-00620</td>
									</tr>
									<tr>
										<th scope="row">상호</th>
										<td>(주)텐바이텐</td>
										<th scope="row">대표자</th>
										<td>최은희</td>
									</tr>
									<tr>
										<th scope="row">소재지</th>
										<td colspan="3">(03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐</td>
									</tr>
									<tr>
										<th scope="row">업태</th>
										<td>서비스, 도소매 등</td>
										<th scope="row">업종</th>
										<td>전자상거래 등</td>
									</tr>
									<tr>
										<th scope="row">담당자</th>
										<td></td>
										<th scope="row">연락처</th>
										<td> </td>
									</tr>
									</tbody>
								</table>
							</div>
						</div>

						<div class="title">
							<h2>주문상품정보</h2>
						</div>
						<table class="baseTable btmLine">
						<caption>주문상품정보 목록</caption>
						<colgroup>
							<col width="90" /> <col width="*" /> <col width="90" /> <col width="60" /> <col width="100" />
						</colgroup>
						<thead>
						<tr>
							<th scope="col">상품코드</th>
							<th scope="col">상품명[옵션명]</th>
							<th scope="col">단가</th>
							<th scope="col">수량</th>
							<th scope="col">금액</th>
						</tr>
						</thead>
						<tbody>
					 <% for i=0 to oorderdetail.FResultCount-1 %>
						 <% if oorderdetail.FItemList(i).Fitemid <>0 then %>
						<tr>
							<td><%= oorderdetail.FItemList(i).FItemid %></td>
							<td class="lt">
								<div><%= oorderdetail.FItemList(i).FItemName %></div>
								<div><font color="blue"><%= oorderdetail.FItemList(i).FItemoptionName %></font></div>
							</td>
							<td><% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
							<% if (oorderdetail.FItemList(i).IsSaleItem) then %>
                                    <strike><%= FormatNumber(oorderdetail.FItemList(i).Forgitemcost,0) %></strike><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %><br>
                                    <strong class="crRed"><%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %></strong><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %>
                                <% else %>
                                    <% if (oorderdetail.FItemList(i).IsItemCouponAssignedItem) then %>
                                    <strike><%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %></strike><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %>
                                    <% else %>
                                    <%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %>
                                    <% end if %>
                                <% end if %>

                                <% if (oorderdetail.FItemList(i).IsItemCouponAssignedItem) then %>
                                    <br><strong class="crGrn"><%= FormatNumber(oorderdetail.FItemList(i).FItemCost,0) %>원</strong>
                                <% else %>

                                <% end if %>

                                <% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
                                <p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %></p>
                                <% end if %>
                        <% else %>
                        	<font color="red">취소</font>
                        <% end if %></td>
							<td><%= oorderdetail.FItemList(i).FItemNo %></td>
							<td>
								<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
									<%= FormatNumber(oorderdetail.FItemList(i).FItemCost*oorderdetail.FItemList(i).FItemNo,0) %> <%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %>
									<% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
									<p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","원") %></p>
									<% end if %>
								 <% else %>
		                        	<font color="red">취소</font>
		                        <% end if %>
                       	 </td>
						</tr>
						<%
							else
								if (oorderdetail.FItemList(i).Fcancelyn <> "Y") then
									totalBeasongPay = totalBeasongPay + oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo
								end if
							end if
						next
						%>
						</tbody>
						<tfoot>
						<tr>
							<td colspan="5">
								<div class="orderTotal">
									총 합계금액 : <strong><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></strong>원
									<% if (totalBeasongPay <> 0) then %>
									(배송비 : <%= FormatNumber(totalBeasongPay,0) %>원)
									<% end if %>
								</div>
							</td>
						</tr>
						</tfoot>
						</table>
					</div>

					<dl class="tMar30">
						<dt class="bBdr2 bPad05 fs14"><strong>참고사항</strong></dt>
						<dd class="tPad05 tBdr3">1. 합계금액은 부가세 포함 금액입니다.</dd>
						<dd>2. 본 견적은 배송비를 포함합니다.</dd>
						<dd>3. 견적유효기간은 견적일로부터 15일입니다.</dd>
					</dl>

					<div class="companyInfo">
						<p><img src="http://fiximage.10x10.co.kr/web2020/my10x10/img_company_info.png" alt="company info" /></p>
					</div>
				</div>
				<!-- //content -->
			</div>
		</div>
	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

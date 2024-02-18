<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/cscenter/lib/popheader_xhtml.asp"-->

<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"--> 
<!-- #include virtual="/lib/util/md5.asp"--> 
<%
dim i, j
dim getPdfDownLinkUrlAdm
dim orderserial
orderserial = requestCheckvar(request("orderserial"),16)

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
addparam = "vos="&orderserial&"&ekey="&md5(orderserial)
if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnOrderReceiptPdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnOrderReceiptPdf.asp?"&addparam
end if
%>	
<script type="text/javascript">
	function jsGoPDF(iUri){
		  var popwin = window.open(iUri,'dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes'); 
	}
</script>
	<div class="heightgird" id="orderPrint"><!-- 2013.09.24 : id="mediaPrint" 추가 -->
		<div class="popWrap">
			<div class="popHeader">
				<h1><img src="http://fiximage.10x10.co.kr/web2013/my10x10/tit_order_receipt.gif" alt="주문확인서" /></h1>
			</div>
			<div class="popContent">
				<!-- content -->
				<div class="mySection">
					<div class="orderDetail">
						<div class="title">
							<h2 class="ftLt" style="margin-top:0;">주문정보</h2>
							<!--p class="ftRt"><img src="http://fiximage.10x10.co.kr/web2013/@temp/img_barcode.gif" alt="바코드이미지" /></p-->
						</div>
						<table class="baseTable rowTable">
						<caption>주문정보 내역</caption>
						<colgroup>
							<col width="15%" /> <col width="35%" /> <col width="15%" /> <col width="35%" />
						</colgroup>
						<tbody>
						<tr>
							<th scope="row">주문번호</th>
							<td><%= oordermaster.FOneItem.FOrderSerial %>
							 <% If oordermaster.FOneItem.IsForeignDeliver Then %>
									  (<strong>해외배송</strong>)
								  <% End If %>
							</td>
							<th scope="row">주문일자</th>
							<td><%= left(oordermaster.FOneItem.FRegDate,10) %></td>
						</tr>
						<tr>
							<th scope="row">결제방법</th>
							<td><%= oordermaster.FOneItem.JumunMethodName %></td>
							<th scope="row">결제일자</th>
							<td><% if IsNULL(oordermaster.FOneItem.FIpkumDate) then %>
                                <strong class="crRed">입금전</strong>
                                <% else %>
                                <%= left(oordermaster.FOneItem.FIpkumDate,10) %>
                                <% end if %>
							</td>
						</tr>
						<tr>
							<% if oordermaster.FOneItem.FAccountdiv = 7 then %>
							<th scope="row"><%= CHKIIF(IsNULL(oordermaster.FOneItem.FIpkumDate),"결제하실금액","결제금액") %></th>
							<td><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.TotalMajorPaymentPrice,0) %></strong>원</em></td>
							<th scope="row">입금하실 계좌</th>
							<td><%= Replace(oordermaster.FOneItem.Faccountno,"X","") %></td>
						 <% else %>
							<th scope="row"><%= CHKIIF(IsNULL(oordermaster.FOneItem.FIpkumDate),"결제하실금액","결제금액") %></th>
							<td colspan="3"><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.TotalMajorPaymentPrice,0) %></strong>원</em>
							 <% if (oordermaster.FOneItem.FAccountDiv="100") or (oordermaster.FOneItem.FAccountDiv="110") then %>
		                        <% if (oordermaster.FOneItem.FokcashbagSpend<>0) then %>
			                        : <span class="red_11px">신용카드 <%= FormatNumber(oordermaster.FOneItem.TotalMajorPaymentPrice-oordermaster.FOneItem.FokcashbagSpend,0) %> 원
			                        , OK캐쉬백 사용 : <%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %> 원
			                	   <% end if %>
		                         </span>
		                    <% end if %>
                    		</td>
                    	<% end if %>
						</tr> 
						 <% if oordermaster.FOneItem.FspendTenCash<>0 then %>
		                <tr>
		                  <th scope="row">예치금사용</th>
		                  <td colspan="3"><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.FspendTenCash,0) %></strong> 원</em></td>
		                </tr>
		                 <% end if %>
		                 <% if oordermaster.FOneItem.Fspendgiftmoney<>0 then %>
		                <tr>
		                  <th scope="row">Gift카드사용</th>
		                     <td colspan="3"><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.Fspendgiftmoney,0) %></strong> 원</em></td>
		                </tr>
		                 <% end if %>
						<tr>
							<th scope="row">주문자 정보</th>
							<td colspan="3"><%= oordermaster.FOneItem.FBuyName %> (휴대전화번호 : <%= oordermaster.FOneItem.FBuyHp %> /  전화번호 : <%= oordermaster.FOneItem.FBuyPhone %>)</td>
						</tr>
						<tr>
							<th scope="row">수령자 정보</th>
							<td colspan="3">
								<div><%= oordermaster.FOneItem.FReqName %> (휴대전화번호 : <%= oordermaster.FOneItem.FReqHp %> /  전화번호 : <%=oordermaster.FOneItem.FReqPhone  %>)</div>
								<div><%= oordermaster.FOneItem.Freqzipaddr %></div>
								<div><%= oordermaster.FOneItem.Freqaddress %></div>
							</td>
						</tr>
						<% if (oordermaster.FOneItem.IsReceiveSiteOrder) then %>
						<tr>
							<th scope="row">수령 날짜</th>
							<td colspan="3"><%= oordermaster.FOneItem.Freqdate %></td>
						</tr>
						<tr>
							<th scope="row">수령 장소</th>
							<td colspan="3">
							    <!--
				                  서울시 송파구 방이동 88-2 올림픽 체조경기장 2-1번 게이트 앞 텐바이텐 예약판매 현장수령 부스
				                  <br>* 지정한 날짜에 지정된 장소에서만 상품수령가능
				                  -->
							</td>
						</tr>
						<% End If %>
						</tbody>
						</table>

						<div class="title">
							<h2>주문상품정보</h2>
						</div>
						<table class="baseTable btmLine">
						<caption>주문상품정보 목록</caption>
						<colgroup>
							<col width="98" /> <col width="70" /> <col width="*" /> <col width="90" /> <col width="50" /> <col width="90" /> <col width="80" />
						</colgroup>
						<thead>
						<tr>
							<th scope="col">상품코드/배송</th>
							<th scope="col" colspan="2">상품정보</th>
							<th scope="col">판매가</th>
							<th scope="col">수량</th>
							<th scope="col">소계금액</th>
							<th scope="col">주문상태</th>
						</tr>
						</thead>
						<tfoot>
						<tr>
							<td colspan="7">
								<div class="orderSummary">
									<span>주문상품수 <strong><%=oorderdetail.FTotItemKind%>종 (<%=oorderdetail.FTotItemNo%>개)</strong></span>
									<span>적립 마일리지 <strong><%=FormatNumber(oordermaster.FOneItem.Ftotalmileage,0)%>P</strong></span>
									<span>상품구매 총액 <strong><%= FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) %>원</strong></span>
								</div>
								<div class="orderTotal">
									총 결제금액 : 상품구매총액 <strong><%= FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) %></strong>원 
									+ 배송비 <%= FormatNumber(oorderdetail.BeasongPay,0) %>원  

			    					<% IF (oordermaster.FOneItem.Fmiletotalprice<>0) then %>
			    					- 마일리지 <%= FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) %>P
			    					<% end if %>
									<% IF (oordermaster.FOneItem.Ftencardspend<>0) then %>
									- 보너스쿠폰할인 <%= FormatNumber(oordermaster.FOneItem.Ftencardspend,0) %>원
									<% end if %> 
			    					<% if (oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership<>0) then %>
			    					- 기타할인 <%= FormatNumber((oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership),0) %>원
			    					<% end if %> 
									= <strong class="crRed"><%= FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0) %></strong>원
								</div> 
							</td>
						</tr>
						</tfoot>
						<tbody>
						 <% for i=0 to oorderdetail.FResultCount-1 %> 
						 <% if oorderdetail.FItemList(i).Fitemid <>0 then %>
						<tr>
							<td>
								<div> <%= oorderdetail.FItemList(i).FItemid %></div>
								<div>
									<% if oorderdetail.FItemList(i).Fisupchebeasong="N" then %>
									텐바이텐
									<% elseif oorderdetail.FItemList(i).Fisupchebeasong="Y" then %>
									<font color="red">업체개별</font>
									<% end if %>
								</div>
							</td>
							<td><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%= oorderdetail.FItemList(i).FItemid %>"  target="_blank"><img src="<%= oorderdetail.FItemList(i).FSmallImage %>" width="50" height="50" alt="<%= oorderdetail.FItemList(i).FItemName %>" /></a></td>
							<td class="lt">
								<div><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%= oorderdetail.FItemList(i).FItemid %>" target="_blank"><%= oorderdetail.FItemList(i).FItemName %></a></div>
								<div><font color="blue"><%= oorderdetail.FItemList(i).FItemoptionName %></font></div>
							</td>
							<td>
						<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
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
                        <% end if %>
							</td>
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
							<td><%= oorderdetail.FItemList(i).GetItemDeliverStateName(oordermaster.FOneItem.FIpkumDiv, oordermaster.FOneItem.FCancelyn) %></td>
						</tr> 
						<%end if%>
						 <% next %>
						</tbody>
						</table>
					</div>

					<div class="companyInfo">
						<p><img src="http://fiximage.10x10.co.kr/web2020/my10x10/img_company_info.png" alt="텐바이텐 10X10 / 판매처 안내 : (주)텐바이텐 사업자등록번호 : 211-87-00620 / 대표이사 : 최은희 / 소재지 : 우)110-510 서울시 종로구 동숭도 1-45 자유빌딩 5층 / 텐바이텐 고객센터안내 TEL : 1644-6030 / AM 09 :00~PM 06:00 점심시간 PM 12:00~01:00 주말,공휴일 휴무 / E-mail : customer@10x10.co.kr " /></p>
					</div>

					<div class="btnArea tMar30 ct">
						<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">인쇄하기</button>
						<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF 전환</button>
					</div>
				</div>
				<!-- //content -->
			</div>
		</div>
		<div class="popFooter">
			<div class="btnArea">
				<button type="button" class="btn btnS1 btnGry2" onclick="window.close();">닫기</button>
			</div>
		</div>
	</div>
<%
set oorderdetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외출고_인보이스
' History : 2014.4.18 정윤정 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
Dim i, currencyChar, currencyunit, cUserInfo, vUserID, IsConfirmedOrder, IsForeignOrder, IsForeign_confirmed, isfixed, jumunwait
dim subdamount, totaldamount, idx, mode, loginsite, addparam, getPdfDownLinkUrlAdm, tplcompanyid, ekey
	loginsite= requestCheckVar(request("ls"),32)
	idx=requestCheckvar(request("idx"),10)
	mode = requestCheckvar(request("mode"),32)
	tplcompanyid    = requestCheckVar(request("tpl"),32)
	currencyunit    = requestCheckVar(request("cunit"),32)
	ekey =  requestCheckvar(request("ekey"),32)

IsConfirmedOrder = False
IsForeignOrder = false		'/업체접수주문
IsForeign_confirmed = false		'/업체접수주문 컨펌완료여부

'// 마스터
Dim oOneOrder, oDetail
SET oOneOrder = new CStorageMaster
	oOneOrder.frectsitename = "WSLWEB"
	oOneOrder.FRectOrderIDX=idx
	oOneOrder.FRectAuthMode = "none"

	if (mode <> "getpdffooter") then
		oOneOrder.getShopOneOrderMaster

		isfixed = oOneOrder.FOneItem.IsFixed

		if oOneOrder.FOneItem.FStatecd=" " then
			jumunwait = true
		end if

		if oOneOrder.FOneItem.FStatecd=" " then
			jumunwait = true
		end if

		if oOneOrder.FOneItem.fforeign_statecd<>"" then
			IsForeignOrder=true

			if oOneOrder.FOneItem.fforeign_statecd="7" then
				IsForeign_confirmed = true
			end if
		else
			IsForeign_confirmed = true
		end if

		currencyChar = oOneOrder.FOneItem.FcurrencyChar
		currencyunit = oOneOrder.FOneItem.FcurrencyUnit
	end if
'// ============================================================================
'// 디테일
SET oDetail = new CStorageDetail
	oDetail.frectsitename = "WSLWEB"
	oDetail.FRectOrderIDX=idx

'	if (Not IsUserLoginOK) then
'		'// 아이피 인증
'		oDetail.FRectAuthMode = "none"
'	end if

	if (mode <> "getpdffooter") then
		if (oOneOrder.FresultCount>0) then
			oDetail.getShopOneOrderDetailList
		end if
	end if

if (mode <> "getpdffooter") then
	if (oOneOrder.FresultCount<1) then
		response.write "Invalid Order"
		dbget.close() : response.end
	end if
end if

if (mode <> "getpdffooter") then
	currencyChar = oOneOrder.FOneItem.FcurrencyChar
end if
'// ============================================================================
'// 주문자정보
vUserID = oOneOrder.FOneItem.Fbaljuid

if (vUserID = "") then
    response.write "Invalid Order"
    dbget.close() : response.end
end if

'// ============================================================================
'// 주문상태
if (mode <> "getpdffooter") then
	if (oOneOrder.FOneItem.Fstatecd >= "7") and (oOneOrder.FOneItem.IsInvoiceExistsState) then
		'// 출고완료 이후는 인보이스 정보
		IsConfirmedOrder = True
	end if
end if

%>

<% if (false) then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<% end if %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
 
	<style type="text/css">
		html, body, blockquote, caption, dd, div, dl, dt, h1, h2, h3, h4, h5, h6, hr, ol, p, pre, q, select, table, textarea, tr, td, ul {margin:0; padding:0;}
		ol, ul {list-style:none;}
		img {border:0;}
		body, h1, h2, h3 ,h4 {font-size:10px; letter-spacing:0; font-family:tahoma, verdana, sans-serif; line-height:14px; color:#333;}
		table {display:block; border-collapse:collapse; border:0; empty-cells:show; width:100%; border-top:1px solid #ccc;}
		th {border-bottom:2px solid #000; padding:2px 5px; font-size:10px;}
		td {text-align:center; padding:3px 5px; border-bottom:1px dotted #cecece; font-size:10px;}
		.subtotal td {border-top:1px solid #cecece; border-bottom:none;}
		.total td {border-top:2px solid #000; border-bottom:none; padding:5px; font-size:10px;}
		.wrapper {width:100%; margin:0 auto;}
		.container {width:100%; position:relative; overflow:hidden; _zoom:1;}
		.ci {width:170px;}
		.ci img {width:170px; height:25px;}
		.header {overflow:hidden; _zoom:1; min-height:102px; padding:5px 0;}
		.title {width:100%; text-align:center; font-size:15px; border-bottom:2px solid #000; padding-bottom:10px;}
		.w30 {width:33%;}
		.w40 {width:40%;}
		.w50 {width:49%;}
		.w60 {width:59%;}
		.w75 {width:75%;}
		.w100 {width:100%;}
		.ftLt {float:left;}
		.ftRt {float:right;}
		.hor {overflow:hidden; _zoom:1;}
		.hor dt {float:left; width:35%;}
		.hor dd {float:left; width:65%}
		.bPad03 {padding-bottom:3px;}
		.tPad05 {padding-top:5px;}
		.tPad10 {padding-top:10px;}
		.vPad1 {padding:1px 0;}
		.vPad10 {padding:10px 0;}
		.vPad15 {padding:15px 0;}
		.bxInner {padding:10px;}
		.tMar10 {margin-top:10px;}
		.tMar15 {margin-top:15px;}
		.tMar20 {margin-top:20px;}
		.lt {text-align:left !important;}
		.rt {text-align:right !important;}
		.fs10 {font-size:10px;}
		.bgGry {background-color:#ebebeb;}
		.bdrBtm {border-bottom:1px dotted #ccc;}
		.bdrBtm2 {border-bottom:1px solid #ccc;}
		.cGry {color:#666;}
	</style>
</head>
<body>
<div class="wrapper">
		<!-- 01 -->
		<div class="container">
			<div class="header">
				<div class="ftLt w60">
					<p class="ci"><img src="/images/10x10_ci.jpg" alt="TENBYTEN" /></p>
					<dl class="ver tMar10">
						<dt><strong>TENBYTEN Inc.</strong></dt>
						<dd class="hor">
							<div class="ftLt w50">
								<p>14F(GyoYukDong)</p>
								<p>57, Daehak-ro, Jongno-gu</p>
								<p>Seoul, Korea [03082]</p>
								<p>VAT Reg.No. : 211-87-00620</p>
							</div>
							<ul class="ftLt w50">
								<li>Tel : +82 2 554 2033</li>
								<li>Fax : +82 2 2179 9244</li>
								<li>Mail : wholesale@10x10.co.kr</li>
								<li>Website : wholesale.10x10.co.kr</li>
							</ul>
						</dd>
					</dl>
				</div>
				<div class="ftRt w40">
					<h1 class="title">
						<% if IsConfirmedOrder then %>
							PROFORMA INVOICE
						<% else %>
							QUOTATION SHEET
						<% end if %>
					</h1>
					<div class="bgGry bxInner bdrBtm2">
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Order No.</strong></dt>
							<dd class="rt">
								<%= oOneOrder.FOneItem.Fbaljucode %>
							</dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Date</strong></dt>
							<dd class="rt">
								<% if IsConfirmedOrder then %>
									<% if Not IsNull(oOneOrder.FOneItem.Fbeasongdate) then %>
										<%= Left(oOneOrder.FOneItem.Fbeasongdate, 10) %>
									<% end if %>
								<% else %>
									<%= Left(oOneOrder.FOneItem.Fregdate, 10) %>
								<% end if %>
							</dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Wholesale ID</strong></dt>
							<dd class="rt"><%= vUserID %></dd>
						</dl>
						<!--<dl class="hor bdrBtm vPad1">
							<dt><strong>Freight Term</strong></dt>
							<dd class="rt"><% '= oOneOrder.FOneItem.FfreightTerm %></dd>
						</dl>-->
						<dl class="tMar10 hor bdrBtm vPad1">
							<dt><strong>
								<% ' 밑에 배송비 주석 처리 문제로 우선 Sub Amount -> Total Amount %>
								<!-- Sub Amount -->
								Total Amount
							</strong></dt>
							<dd class="rt"><strong>
								<% if IsConfirmedOrder then %>
									<%= getdisp_price_currencyChar(oOneOrder.FOneItem.Ftotalforeign_suplycash, oOneOrder.FOneItem.FcurrencyChar) %>
								<% else %>
									<%= getdisp_price_currencyChar(oOneOrder.FOneItem.Fjumunforeign_suplycash,  oOneOrder.FOneItem.FcurrencyChar) %>
								<% end if %>
							</strong></dd>
						</dl>
						<!--<dl class="hor bdrBtm vPad1">
							<dt><strong>Freight charge</strong></dt>
							<dd class="rt"><strong>
								<% 'if IsConfirmedOrder then %>
									<%'= getdisp_price_currencyChar(oOneOrder.FOneItem.FtotalDeliverPriceForeign, oOneOrder.FOneItem.FcurrencyChar) %>
								<% 'else %>
									--
								<% 'end if %>
							</strong></dd>
						</dl>-->
						<!--<dl class="hor vPad1">
							<dt><strong>Total Amount</strong></dt>
							<dd class="rt"><strong>
								<% 'if IsConfirmedOrder then %>
									<%'= 'getdisp_price_currencyChar( oOneOrder.FOneItem.Ftotalforeign_suplycash + oOneOrder.FOneItem.FtotalDeliverPriceForeign , oOneOrder.FOneItem.FcurrencyChar) %>
								<% 'else %>
									<%'= getdisp_price_currencyChar(oOneOrder.FOneItem.Fjumunforeign_suplycash, oOneOrder.FOneItem.FcurrencyChar) %>
								<% 'end if %>
							</strong></dd>
						</dl>-->
					</div>
				</div>
			</div>
			<div class="vPad10">
				<!--<dl class="ftLt ver w50">
					<dt><strong>Invoice address</strong></dt>
					<dd><p>
						<% 'if (oOneOrder.FOneItem.FinvoiceAddress = "Same as Above") then %>
							<p><%'= nl2br(oOneOrder.FOneItem.FshippingAddress) %></p>
						<% 'else %>
							<p><%'= nl2br(oOneOrder.FOneItem.FinvoiceAddress) %></p>
						<% 'end if %>
					</p></dd>
				</dl>-->
				<!--<dl class="ftRt ver w50">
					<dt><strong>Shipping Address</strong></dt>
					<dd>
						<%'= nl2br(oOneOrder.FOneItem.FshippingAddress) %>
					</dd>
				</dl>-->
			</div>
			<div class="vPad10">
				<table>
					<colgroup>
						<col width="70px" /><col width="" /><col width="60px" /><col width="70px" />

						<%
						'/주문서작성중이 아닌거
						if not(jumunwait) then
						%>
							<col width="50px" /><col width="50px" />
						<% else %>
							<col width="50px" />
						<% end if %>

						<col width="" />
					</colgroup>
					<thead>
						<tr>
							<th>Item Code</th>
							<th>Description</th>
							<th>Option</th>
							<th class="rt">Price</th>

							<%
							'/주문서작성중이 아닌거
							if not(jumunwait) then
							%>
								<th>ORDER<br>Quantity</th>
								<th>CONFIRM<br>Quantity</th>
							<% else %>
								<th>ORDER<br>Quantity</td>
							<% end if %>

							<th class="rt">Amount</th>
						</tr>
					</thead>
					<tbody>
						<%
						subdamount = 0
						totaldamount = 0

						if oDetail.FResultCount > 0 then
						for i=0 to oDetail.FResultCount-1
						%>
						<tr>
							<td><%= BF_MakeTenBarcode(oDetail.FItemList(i).Fitemgubun, oDetail.FItemList(i).Fitemid, oDetail.FItemList(i).Fitemoption) %></td>
							<td class="lt"><%=oDetail.FItemList(i).FmLitemname%></td>
							<td class="lt"><%= Replace(oDetail.FItemList(i).getOptionDpFormat, "Option : ", "") %></td>
							<td class="rt">
							    <%= getdisp_price_currencyChar(oDetail.FItemList(i).Fforeign_suplycash, currencyChar) %>
							</td>

							<%
							'/주문서작성중이 아닌거
							if not(jumunwait) then
							%>
								<td>
									<%= oDetail.FItemList(i).Fbaljuitemno %>
								</td>
								<td>
									<%=FormatNumber( getstateitemno(oOneOrder.FOneItem.Fstatecd, oOneOrder.FOneItem.Fforeign_statecd, oDetail.FItemList(i).Fbaljuitemno, oDetail.FItemList(i).Frealitemno) ,0)%>
								</td>
							<% else %>
								<td>
									<%=FormatNumber( getstateitemno(oOneOrder.FOneItem.Fstatecd, oOneOrder.FOneItem.Fforeign_statecd, oDetail.FItemList(i).Fbaljuitemno, oDetail.FItemList(i).Frealitemno) ,0)%>
								</td>
							<% end if %>

							<td class="rt">
								<% if IsConfirmedOrder then %>
									<%
									subdamount = subdamount + (oDetail.FItemList(i).Fforeign_suplycash*oDetail.FItemList(i).Frealitemno)
									%>
									<%= getdisp_price_currencyChar( oDetail.FItemList(i).Fforeign_suplycash * oDetail.FItemList(i).Frealitemno , currencyChar) %>
								<% else %>
									<%
									subdamount = subdamount + (oDetail.FItemList(i).Fforeign_suplycash*oDetail.FItemList(i).Fbaljuitemno)
									%>
									<%= getdisp_price_currencyChar( oDetail.FItemList(i).Fforeign_suplycash * oDetail.FItemList(i).Fbaljuitemno , currencyChar) %>
								<% end if %>
							</td>
						</tr>
						<%
						Next
						end if

						'totaldamount  = getdisp_price_currencyChar(subdamount+totalDeliverPriceForeign,currencyunit)
						totaldamount  = getdisp_price_currencyChar(subdamount,currencyunit)
						subdamount    = getdisp_price_currencyChar(subdamount,currencyunit)
						%>
					</tbody>
					<tfoot>
						<tr class="subtotal">
							<td class="rt" <% if not(jumunwait) then %>colspan="5"<% else %>colspan="4"<% end if %>><strong>Sub Amount</strong></td>
							<td class="rt" <% if not(jumunwait) then %>colspan="3"<% else %>colspan="2"<% end if %>><strong><%= subdamount %></strong></td>
						</tr>
						<!--<tr>
							<td class="rt" colspan="4">Freight charge</td>
							<td class="rt" colspan="2">
								<strong class="cGry">
								<% 'if IsConfirmedOrder then %>
									<%'= getdisp_price_currencyChar(oOneOrder.FOneItem.FtotalDeliverPriceForeign, oOneOrder.FOneItem.FcurrencyChar) %>
								<% 'else %>
									--
								<% 'end if %>
								</strong>
							</td>
						</tr>-->
						<tr class="total">
							<td class="rt bgGry" <% if not(jumunwait) then %>colspan="5"<% else %>colspan="4"<% end if %>><strong>Total Amount</strong></td>
							<td class="rt bgGry" <% if not(jumunwait) then %>colspan="3"<% else %>colspan="2"<% end if %>><strong><%=totaldamount%></strong></td>
						</tr>
					</tfoot>
				</table>
			</div>
			<!--<div class="vPad10">
				<dl class="ftLt ver w100">
					<dt class="bdrBtm2 bPad03"><strong>Note</strong></dt>
					<dd class="tPad05">
						(note area)
					</dd>
				</dl>
			</div>-->
		</div>
		<!-- //01 -->

	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

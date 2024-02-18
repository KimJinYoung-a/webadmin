<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외 주문서
' History : 2017.06.15 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx ,loginsite, i, j, tplcompanyid
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount ,currencychar, currencyunit, isfixed, ekey
dim countryLangCD,arrQS   ,intQS
	idx = requestCheckVar(getNumeric(request("idx")),10) '--cartoonbox idx
	loginsite= requestCheckVar(request("ls"),32)
	currencyunit = requestCheckVar(request("cunit"),32) '--shop 기준 화폐
	tplcompanyid = requestCheckVar(request("tpl"),32)
	ekey =  requestCheckvar(request("ekey"),32)

isfixed = false
if idx="" then idx=0

if (ekey="") then
    response.write "암호화 키가 올바르지 않습니다.1"
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5(idx&loginsite))) then
    response.write "암호화 키가 올바르지 않습니다.2"
    response.end
end if

dim ojumunmaster
set ojumunmaster = new COrderSheet
	ojumunmaster.FRectIdx = idx
	ojumunmaster.GetOneOrderSheetMaster

if ojumunmaster.ftotalcount < 1 then
	response.write "<script type='text/javascript>"
	response.write "	alert('해당되는 주문건이 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if

isfixed = ojumunmaster.FOneItem.FStatecd >= 7

dim ojumundetail
set ojumundetail= new COrderSheet
	ojumundetail.FRectIdx = idx
	ojumundetail.GetOrderSheetDetail_foreign

if ojumundetail.ftotalcount > 0 then
	for i = 0 to ojumundetail.FResultCount - 1
	totalPriceForeign = totalPriceForeign + (ojumundetail.FItemList(i).flcprice * ojumundetail.FItemList(i).frealitemno)
	next
end if

totalamount  = getdisp_price_currencyChar(totalPriceForeign ,currencyunit)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
	<style type="text/css">
		html, body, blockquote, caption, dd, div, dl, dt, h1, h2, h3, h4, h5, h6, hr, ol, p, pre, q, select, table, textarea, tr, td, ul {margin:0; padding:0;}
		ol, ul {list-style:none;}
		img {border:0;}
		body, h1, h2, h3 ,h4 {font-size:10px; letter-spacing:0; font-family:tahoma, verdana, sans-serif; line-height:14px; color:#333;}
		table {border-collapse:collapse; border:0; empty-cells:show; width:100%; border-top:1px solid #ccc;}
		th {border-bottom:2px solid #000; padding:2px 5px; font-size:10px;}
		td {text-align:center; padding:3px 5px; border-bottom:1px dotted #cecece; font-size:10px;}
		.subtotal td {border-top:1px solid #cecece; border-bottom:none;}
		.total td {border-top:2px solid #000; border-bottom:none; padding:5px; font-size:10px;}
		.wrapper {width:638px; margin:0 auto;}
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

<div class="wrapper">
		<!-- 01 -->
		<div class="container">
			<div class="header">
				<div class="ftLt w60">
					<% IF tplcompanyid <> "" THEN %>
						<p><img src="/images/logo_ithinkso.jpg" alt="ithinkso" width="182" height="36"/></p>
						<dl class="ver tMar10">
							<dt><strong>S&T works Inc. </strong></dt>
							<dd class="hor">
								<div class="ftLt w50">
									<p>4F, 52, Daehak-ro 8ga-gil,</p>
									<p>Jongno-gu, Seoul,</p>
									<p>Korea [03086]</p>
									<p>VAT Reg.No. : 101-86-84103</p>
								</div>
								<ul class="ftLt w50">
									<li>Tel : +82 70 4821 1903</li>
									<li>Fax : +82 2 2179 8631</li>
									<% '<li>Mail : salesmanger@ithinksoweb.com</li> %>
									<% '<li>Website : www.ithinksoweb.com</li> %>
								</ul>
							</dd>
						</dl>
					<% ELSE %>
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
									<% '<li>Mail : wholesale@10x10.co.kr</li> %>
									<% '<li>Website : wholesale.10x10.co.kr</li> %>
								</ul>
							</dd>
						</dl>
					<% END IF %>
				</div>
				<div class="ftRt w40">
					<h1 class="title">ORDER SHEET</h1>
					<div class="bgGry bxInner bdrBtm2">
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Order No.</strong></dt>
							<dd class="rt">
								<%= ojumunmaster.FOneItem.Fbaljucode %>
							</dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Date</strong></dt>
							<dd class="rt">
								<%= left(ojumunmaster.FOneItem.Fregdate,10) %>
							</dd>
						</dl>
						<!--<dl class="hor vPad1">
							<dt><strong>Total Amount</strong></dt>
							<dd class="rt"><strong><%'=totalamount%></strong></dd>
						</dl>-->
					</div>
				</div>
			</div>
			<div class="vPad10" style="overflow:hidden;">
				<dl class="ftLt ver w50">
					<dt><strong>Invoice address</strong></dt>
					<dd>
						<p>
							<% '우리쪽에서 보내는거라 그냥 박아넣으면됨. 고정임 %>
							TENBYTEN Inc.
							<br>5F, Jayu Bldg , 31, Daehak-ro 12-gil
							<br>Jongno-gu, Seoul, Korea (ZIP : 03086)
							<br>Attn.  Yeeun Kim 
							<br>82-70-4000-4330 / nmmkye@10x10.co.kr
						</p>
					</dd>
					</dd>
				</dl>
				<dl class="ftRt ver w50">
					<dt><strong>Shipping Address</strong></dt>
					<dd>
						<p>
							<% '우리쪽에서 보내는거라 그냥 박아넣으면됨. 고정임 %>
							TENBYTEN Inc.
							<br>4F, Yoein dotcom, 31,  Dobong-ro 180-gil
							<br>Dobong-gu, Seoul, Korea (ZIP :  01319)
							<br>Attn.  Yeeun Kim 
							<br>82-70-4000-4330 / nmmkye@10x10.co.kr
						</p>
					</dd>
				</dl>
			</div>
			<div class="vPad10">
				<table>
					<colgroup>
						<col width="80px" /><col width="80px" /><col width="80px" /><col width="" /><col width="" /><col width="70px" />
					</colgroup>
					<thead>
						<tr>
							<th>Item Code(Buyer)</th>
							<th>BarCode</th>
							<th>Item Code(Seller)</th>
							<th>Description</th>
							<th>Option</th>
							<th>Quantity</th>
						</tr>
					</thead>
					<tbody>
						<%
						if ojumundetail.ftotalcount > 0 then
							 for i = 0 to ojumundetail.FResultCount - 1
							%>
							<tr>
								<td><%=BF_MakeTenBarcode(ojumundetail.FItemList(i).Fitemgubun,ojumundetail.FItemList(i).Fitemid,ojumundetail.FItemList(i).Fitemoption)%></td>
								<td><%= ojumundetail.FItemList(i).FPublicBarcode %></td>
								<td><%= ojumundetail.FItemList(i).FUpcheManageCode %></td>
								<td class="lt"><%= ojumundetail.FItemList(i).flcitemname %></td>
								<td>
									<%= ojumundetail.FItemList(i).flcitemoptionname %>
								</td>
								<!--<td class="rt">
								    <%'= getdisp_price_currencyChar(ojumundetail.FItemList(i).flcprice,currencyunit) %>
								</td>-->
								<td>
									<% IF isfixed THEN %>
										<%= ojumundetail.FItemList(i).Frealitemno %>
									<% else %>
										<%= ojumundetail.FItemList(i).Fbaljuitemno %>
									<% end if %>
								</td>
								<!--<td class="rt">
									<% 'IF isfixed THEN %>
										<%'= getdisp_price_currencyChar(ojumundetail.FItemList(i).flcprice*ojumundetail.FItemList(i).Frealitemno,currencyunit) %>
									<% 'else %>
										<%'= getdisp_price_currencyChar(ojumundetail.FItemList(i).flcprice*ojumundetail.FItemList(i).Fbaljuitemno,currencyunit) %>
									<% 'end if %>
								</td>-->
							</tr>
						<%
							next
						end if
						%>
					</tbody>
					<!--<tfoot>
						<tr class="total">
							<td class="rt bgGry" colspan="5"><strong>Total Amount</strong></td>
							<td class="rt bgGry" colspan="2"><strong><%'=totalamount%></strong></td>
						</tr>
					</tfoot>-->
				</table>
			</div>
			<div class="vPad10">
				<dl class="ftLt ver w100">
					<dt class="bdrBtm2 bPad03"><strong>Note</strong></dt>
					<dd class="tPad05">
						(note area)
					</dd>
				</dl>
			</div>
		</div>
		<!-- //01 -->
	</div>
</body>
</html>

<%
set ojumundetail = nothing
set ojumunmaster = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

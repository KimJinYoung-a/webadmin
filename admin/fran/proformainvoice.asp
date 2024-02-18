<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외출고_인보이스
' History : 2014.4.18 정윤정 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/admin/lib/popheader_xhtml.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx , loginsite,invoiceNo,invoicedate
dim i, j
dim ClsOI , arrList , intLoop
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount, freightcharge , currencychar, currencyunit,subdamount, totaldamount
dim countryLangCD,getPdfDownLinkUrlAdm,addparam
dim tplcompanyid, jungsanidx, workidx

tplcompanyid = requestCheckVar(request("tpl"),32)
idx = requestCheckVar(request("idx"),10) '--invoice idx
loginsite= requestCheckVar(request("ls"),32)
currencyunit = requestCheckVar(request("cunit"),32)
jungsanidx = requestCheckVar(request("jungsanidx"),10)
workidx = requestCheckVar(request("workidx"),10)
if idx="" then idx=0
'================================================================================
 set ClsOI  = new COffInvoice
 ClsOI.FRectMasterIdx = idx
 ClsOI.FRectLoginsite = loginsite
 ClsOI.FRectJungsanidx = jungsanidx
 ClsOI.FRectWorkidx = workidx
 ClsOI.fnGetFranInvoice
 invoiceNo					= ClsOI.Finvoiceno
 invoicedate				= ClsOI.Finvoicedate
 baljucode                  = ClsOI.Fbaljucode
 baljuid                	= ClsOI.Fbaljuid
 beasongdate                = ClsOI.Fbeasongdate
 regdate					= ClsOI.Fregdate
 priceunit				    = ClsOI.Fpriceunit
 totalGoodsPriceWon		    = ClsOI.FtotalGoodsPriceWon
 totalDeliverPriceWon 	    = ClsOI.FtotalDeliverPriceWon
 totalPriceWon  			= ClsOI.FtotalPriceWon
 totalGoodsPriceForeign	    = ClsOI.FtotalGoodsPriceForeign
 totalDeliverPriceForeign   = ClsOI.FtotalDeliverPriceForeign
 totalPriceForeign		    = ClsOI.FtotalPriceForeign
 freightTerm                = ClsOI.FfreightTerm
 openState                  = ClsOI.FopenState
 shippingAddress            = ClsOI.FshippingAddress
 invoiceAddress             = ClsOI.FinvoiceAddress
 currencychar				= ClsOI.Fcurrencychar
'' currencyunit				= ClsOI.Fcurrencyunit

 if baljucode <> "" then
 ClsOI.FRectbaljucode	= baljucode
 arrList = ClsOI.fnGetFranItemList  '주문확정 수량 기준 상품리스트
 end if
  set ClsOI = nothing

' if currencyunit = "WON" THEN   '--샵 기준 화폐가 원일때
' 	currencychar = "원"
' 	subamount 	 = FormatNumber(totalGoodsPriceWon,0)&currencychar 'master 상의 합계금액(cf. subdamount:detail 리스트 합계)
' 	freightcharge= FormatNumber(totalDeliverPriceWon,0)&currencychar
' 	totalamount  = FormatNumber(totalPriceWon,0)&currencychar
'else
'	subamount 	 = currencychar&FormatNumber(totalGoodsPriceForeign,2)
' 	freightcharge= currencychar&FormatNumber(totalDeliverPriceForeign,2)
' 	totalamount  = currencychar&FormatNumber(totalPriceForeign ,2)
'end if

'' 무조건 foreign PRICE로 2016/10/18  아래 arrList(8,intLoop) => arrList(13,intLoop) 으로 수정.
subamount 	 = getdisp_price_currencyChar(totalGoodsPriceForeign,currencyunit)
freightcharge= getdisp_price_currencyChar(totalDeliverPriceForeign,currencyunit)
totalamount  = getdisp_price_currencyChar(totalPriceForeign ,currencyunit)

'--- pdf 전환처리----------------------------------------
addparam = "idx="&idx&"&ls="&loginsite&"&cunit="&currencyunit&"&tpl="&tplcompanyid&"&ekey="&md5(idx&loginsite)&"&jungsanidx=" & jungsanidx & "&workidx=" & workidx
if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnproformainvoicePdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnproformainvoicePdf.asp?"&addparam
end if
'---------------------------------------------------------
%>
	<style type="text/css">
		html, body, blockquote, caption, dd, div, dl, dt, h1, h2, h3, h4, h5, h6, hr, ol, p, pre, q, select, table, textarea, tr, td, ul {margin:0; padding:0;}
		ol, ul {list-style:none;}
		img {border:0;}
		body, h1, h2, h3 ,h4 {font-size:10px; letter-spacing:0; font-family:tahoma, verdana, sans-serif; line-height:14px; color:#333;}
		div {overflow:hidden; _zoom:1;}
		table {border-collapse:collapse; border:0; empty-cells:show; width:100%; border-top:1px solid #ccc;}
		th {border-bottom:2px solid #000; padding:2px 5px;}
		td {text-align:center; padding:3px 5px; border-bottom:1px dotted #cecece;}
		.subtotal td {border-top:1px solid #cecece; border-bottom:none;}
		.total td {border-top:2px solid #000; border-bottom:none; padding:5px;}
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
<script type="text/javascript">
	function jsGoPDF(iUri){
		  var popwin = window.open(iUri,'dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes');
	}
</script>
<div class="wrapper">
		<!-- 01 -->
		<div class="container">
			<div class="header">
				<div class="ftLt w60">
					<%IF tplcompanyid <> "" THEN%>
					<% if (idx >= 1263) then '' 주소가 달라진다. %>
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
								<li>Mail : salesmanger@ithinksoweb.com</li>
								<li>Website : www.ithinksoweb.com</li>
							</ul>
						</dd>
					</dl>
					<% else %>
					<p><img src="/images/logo_ithinkso.jpg" alt="ithinkso" width="182" height="36"/></p>
					<dl class="ver tMar10">
						<dt><strong>S&T works Inc. </strong></dt>
						<dd class="hor">
							<div class="ftLt w50">
								<p>5F, ERH bldg, 1-74, </p>
								<p>Dongsung-dong, Jongno-gu,</p>
								<p>Seoul, Korea [110-809]</p>
								<p>VAT Reg.No. : 101-86-84103</p>
							</div>
							<ul class="ftLt w50">
								<li>Tel : +82 70 4821 1903</li>
								<li>Fax : +82 2 2179 8631</li>
								<li>Mail : salesmanger@ithinksoweb.com</li>
								<li>Website : www.ithinksoweb.com</li>
							</ul>
						</dd>
					</dl>
					<% end if %>
					<%ELSE%>
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
					<%END IF%>
				</div>
				<div class="ftRt w40">
					<h1 class="title">PROFORMA INVOICE</h1>
					<div class="bgGry bxInner bdrBtm2">
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Invoice No.</strong></dt>
							<dd class="rt"><%=invoiceNo%></dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Invoice Date</strong></dt>
							<dd class="rt"><%=invoicedate%></dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Wholesale ID</strong></dt>
							<dd class="rt"><%=baljuid%></dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Freight Term</strong></dt>
							<dd class="rt"><%=freightTerm%></dd>
						</dl>
						<dl class="tMar10 hor bdrBtm vPad1">
							<dt><strong>Sub Amount</strong></dt>
							<dd class="rt"><strong><%=subamount%></strong></dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Shipping</strong></dt>
							<dd class="rt"><strong><%=freightcharge%></strong></dd>
						</dl>
						<dl class="hor vPad1">
							<dt><strong>Total Amount</strong></dt>
							<dd class="rt"><strong><%=totalamount%></strong></dd>
						</dl>
					</div>
				</div>
			</div>
			<div class="vPad10">
				<dl class="ftLt ver w50">
					<dt><strong>Invoice address</strong></dt>
					<dd> <p><%= nl2br(invoiceAddress) %></p>
					</dd>
				</dl>
				<dl class="ftRt ver w50">
					<dt><strong>Shipping Address</strong></dt>
					<dd>
						<% if (trim(replace(shippingAddress,chr(13)&chr(10),"")) = "Same as Above") then %>
							<p><%= nl2br(invoiceAddress) %></p>
						<% else %>
							<p><%= nl2br(shippingAddress) %></p>
						<% end if %>
					</dd>
				</dl>
			</div>
			<div class="vPad10">
				<table>
					<colgroup>
						<col width="80px" /><col width="" /><col width="" /><col width="70px" /><col width="50px" /><col width="" />
					</colgroup>
					<thead>
						<tr>
							<th>Item Code</th>
							<th>Description</th>
							<th>Option</th>
							<th class="rt">Price</th>
							<th>Quantity</th>
							<th class="rt">Amount</th>
						</tr>
					</thead>
					<tbody>
						<% subdamount = 0
						   totaldamount = 0
						 if isArray(arrList) then
							 for intLoop = 0 To UBound(arrList,2)
							%>
						<tr>
							<td><%=BF_MakeTenBarcode(arrList(1,intLoop),arrList(3,intLoop),arrList(4,intLoop))%></td>
							<td class="lt"><%=arrList(5,intLoop)%></td>
							<td><%=arrList(6,intLoop)%></td>
							<td class="rt"><% '--shop기준 화폐단위가 won 이 아닐땐 해외값으로 계산%>
							    <%= getdisp_price_currencyChar(arrList(13,intLoop),currencyunit) %>
				                <% if (FALSE) then %>
								<%IF  currencyunit <> "WON" THEN%>
								 <%=currencyChar%><%=FormatNumber(arrList(13,intLoop),2)%>
								<%else%>
								<%=FormatNumber(arrList(8,intLoop),0)%><%=currencyChar%>
								<%end if%>
								<%end if%>
							</td>
							<td><%=arrList(11,intLoop)%></td>
							<td class="rt">
							    <%
                			    subdamount = subdamount + (arrList(13,intLoop)*arrList(11,intLoop))
                			    %>
                			    <%= getdisp_price_currencyChar(arrList(13,intLoop)*arrList(11,intLoop),currencyunit) %>
                			    <% if (FALSE) then %>
    								<%IF  currencyunit <> "WON" THEN '--shop기준 화폐단위가 won 이 아닐땐 해외값으로 계산
    									subdamount = subdamount + (arrList(13,intLoop)*arrList(11,intLoop)) '상품 detail 리스트 가격 합계처리
    								%>
    									<%=currencyChar%><%=FormatNumber(arrList(13,intLoop)*arrList(11,intLoop),2)%>
    								<%ELSE
    									subdamount = subdamount + (arrList(8,intLoop)*arrList(11,intLoop)) '상품 detail 리스트 가격 합계처리
    								%>
    									<%=FormatNumber(arrList(8,intLoop)*arrList(11,intLoop),0)%> <%=currencyChar%>
    								<%END IF%>
								<%END IF%>
							</td>
						</tr>
						<%	Next
						 end if

						 totaldamount  = getdisp_price_currencyChar(subdamount+totalDeliverPriceForeign,currencyunit)
                         subdamount    = getdisp_price_currencyChar(subdamount,currencyunit)

                         if (FALSE) then
    						 if currencyunit <> "WON" THEN
    						 	totaldamount  = currencychar&FormatNumber(subdamount+totalDeliverPriceForeign,2)
    						 	subdamount 	 = currencychar&FormatNumber(subdamount,2)
    						 else
    							totaldamount  = FormatNumber(subdamount+totalDeliverPriceWon,0)&currencychar
    						 	subdamount 	 = FormatNumber(subdamount,0)&currencychar
    						 end if
						 end if
						 %>
					</tbody>
					<tfoot>
						<tr class="subtotal">
							<td class="rt" colspan="4"><strong>Sub Amount</strong></td>
							<td class="rt" colspan="2"><strong><%=subdamount%></strong></td>
						</tr>
						<tr>
							<td class="rt" colspan="4">Freight charge</td>
							<td class="rt" colspan="2"><strong class="cGry"><%=freightcharge%></strong></td>
						</tr>
						<tr class="total">
							<td class="rt bgGry" colspan="4"><strong>Total Amount</strong></td>
							<td class="rt bgGry" colspan="2"><strong><%=totaldamount%></strong></td>
						</tr>
					</tfoot>
				</table>
			</div>
			<div class="vPad10">
				<dl class="ftLt ver w50">
					<dt class="bdrBtm2 bPad03"><strong>Note</strong></dt>
					<dd class="tPad05">
						(note area)
					</dd>
				</dl>
				<dl class="ftRt ver w50">
					<dt class="bdrBtm2 bPad03"><strong>Payment</strong></dt>
					<dd class="tPad05">
						<%IF tplcompanyid <> "" THEN%>
						<dl class="hor vPad1">
							<dt>Bebeficiary</dt>
							<dd>S&T works Inc.</dd>
						</dl>
						<dl class="hor vPad1">
							<dt>Bank</dt>
							<dd>Industrial Bank of Korea</dd>
						</dl>
						<dl class="hor vPad1">
							<dt>Acc. No.</dt>
							<dd>277-050381-56-00011</dd>
						</dl>
						<dl class="hor vPad1">
							<dt>SWIFT/BIC</dt>
							<dd>IBKOKRSE</dd>
						</dl>
						<%else %>
						<dl class="hor vPad1">
							<dt>Bebeficiary</dt>
							<dd>TenbyTen Inc.</dd>
						</dl>
						<dl class="hor vPad1">
							<dt>Bank</dt>
							<dd>Industrial Bank of Korea</dd>
						</dl>
						<dl class="hor vPad1">
							<dt>Acc. No.</dt>
							<dd>27703918804031</dd>
						</dl>
						<dl class="hor vPad1">
							<dt>SWIFT/BIC</dt>
							<dd>IBKOKRSE</dd>
						</dl>
						<%end if%>
					</dd>
					<dd class="fs10 tPad10">* Please pay by bank transfer without deduction of bank charges.</dd>
				</dl>
			</div>
		</div>
		<!-- //01 -->
	</div>
	<div class="btnArea tMar30 ct">
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">인쇄하기</button>
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF 전환</button>
	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

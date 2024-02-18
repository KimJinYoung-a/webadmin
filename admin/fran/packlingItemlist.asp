<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외출고 패킹리스트
' History : 2014.04.18 정윤정 생성
'			2016.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<!-- #include virtual="/admin/lib/popheader_xhtml.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim idx , loginsite,invoiceNo,invoicedate, boxidx, i, j, ClsOI , arrList , intLoop
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount, freightcharge , currencychar, currencyunit, countryLangCD,getPdfDownLinkUrlAdm,addparam, tplcompanyid
	tplcompanyid = requestCheckVar(request("tpl"),32)
	idx = requestCheckVar(request("idx"),10) '--invoice idx
	loginsite= requestCheckVar(request("ls"),32)
	boxidx= requestCheckVar(request("boxidx"),10)
	currencyunit = requestCheckVar(request("cunit"),32)

if idx="" then idx=0
if 	boxidx = "" then boxidx = 0

set ClsOI  = new COffInvoice
	ClsOI.FRectMasterIdx = idx
	ClsOI.FRectLoginsite = loginsite
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
	' currencyunit				= ClsOI.Fcurrencyunit
set ClsOI = nothing

' if currencyunit = "WON" THEN   '--샵 기준 화폐가 원일때
' 	currencychar = " KRW"
' 	subamount 	 = FormatNumber(totalGoodsPriceWon,0)&currencychar
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
addparam = "idx="&idx&"&ls="&loginsite&"&boxidx="&boxidx&"&cunit="&currencyunit&"&tpl="&tplcompanyid&"&ekey="&md5(idx&loginsite&boxidx)

if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnpacklingitemlistPdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnpacklingitemlistPdf.asp?"&addparam
end if

dim ocartoonboxdetail
set ocartoonboxdetail = new CCartoonBox
	ocartoonboxdetail.FPageSize = 1000
	ocartoonboxdetail.FRectMasterIdx = boxidx
	ocartoonboxdetail.FRectShopid = baljuid
	ocartoonboxdetail.GetDetailItemList
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
	.title {width:100%; text-align:center; font-size:18px; border-bottom:2px solid #000; padding-bottom:10px;}
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

function jsGoEXCEL(){
	  var popwin = window.open('/admin/fran/viewpacklingitemlist_excel.asp?<%= addparam %>','dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes');
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
					<h1 class="title">PACKING LIST</h1>
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
			</div>
			<div class="vPad10">
				<table>
					<colgroup>
						<% ' 밑에 하위 테이블이 짜여져서 크기 조율을 미세하게 이쁘게 잘해야함 %>
						<col width="20px" />
						<col width="70px" />
						<col width="" />
						<col width="100px" />
						<col width="30px" />
						<col width="30px" />
						<col width="40px" />
						<col width="40px" />
					</colgroup>
					<thead>
						<tr>
							<th>BOX<Br>NO.</th>
							<th>Item Code</th>
							<th>Description</th>
							<th>Option</th>
							<th>Qty</th>
							<th>weight</th>
							<th>N<Br>weight</th>
							<th>G<Br>weight</th>
						</tr>
					</thead>
					<tbody>
					<%
					dim currcartoonboxno, suminnerboxweight, sumcartoonboxNweight, sumcartoonboxweight, isnewcartoonbox, sumcartoonboxcbm
					dim itemweight
					currcartoonboxno = ""
					suminnerboxweight = 0
					sumcartoonboxNweight = 0
					sumcartoonboxweight = 0
					sumcartoonboxcbm = 0
					itemweight = 0

					for i=0 to ocartoonboxdetail.FResultCount-1

						if (ocartoonboxdetail.FItemList(i).Fcartoonboxno <> currcartoonboxno) then
							isnewcartoonbox = true
							currcartoonboxno = ocartoonboxdetail.FItemList(i).Fcartoonboxno
						else
							isnewcartoonbox = false
						end if

						if IsNull(ocartoonboxdetail.FItemList(i).FcartoonboxNweight) then
							ocartoonboxdetail.FItemList(i).FcartoonboxNweight = 0
						end if

						if (isnewcartoonbox = true) then
							sumcartoonboxNweight = sumcartoonboxNweight + ocartoonboxdetail.FItemList(i).FcartoonboxNweight
							sumcartoonboxweight = sumcartoonboxweight + ocartoonboxdetail.FItemList(i).Fcartoonboxweight

						end if

						suminnerboxweight = suminnerboxweight + ocartoonboxdetail.FItemList(i).Finnerboxweight
						if isnewcartoonbox = true then

							if ocartoonboxdetail.FItemList(i).FcartoonboxType <> "" then
								sumcartoonboxcbm = sumcartoonboxcbm + getcartoonboxtype(ocartoonboxdetail.FItemList(i).FcartoonboxType, 1)
							end if
						%>
							<% if i > 0 then %>
									</table>
								</td>
								<td><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).FcartoonboxNweight, 2) %>Kgs</td>
								<td><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).Fcartoonboxweight, 2) %>Kgs</td>
								<!--td class="rt">
									<% if ocartoonboxdetail.FItemList(i-1).FcartoonboxType <> "" then %>
										<%= getcartoonboxtype(ocartoonboxdetail.FItemList(i-1).FcartoonboxType, 1) %>
									<% end if %>
								</td-->
							</tr>
							<%end if %>

							<tr valign="top">
								<td valign="top"><%= currcartoonboxno%></td>
								<!--td class="lt"  valign="top">
									<%= ocartoonboxdetail.FItemList(i).FcartoonboxType %>

									<% if ocartoonboxdetail.FItemList(i).FcartoonboxType <> "" then %>
										(<%= getcartoonboxtype(ocartoonboxdetail.FItemList(i).FcartoonboxType, 0) %>)
									<% end if %>
								</td-->
								<td colspan="5">
									<table class="fs10" style="border-top:0px;border-bottom:0px;" cellpadding="0" cellspacing="0" border="0">
										<colgroup>
											<col width="60px" /><col width="" /><col width="80px" /><col width="30px" /><col width="30px" />
										</colgroup>
					<%end if %>
										<tr>
											<td style="border-bottom:0px;">
												<%'=i	'이게 왜 들어가 있는거?? %>
												<%=ocartoonboxdetail.FItemList(i).Fitemgubun%><%=ocartoonboxdetail.FItemList(i).Fitemid%><%=ocartoonboxdetail.FItemList(i).Fitemoption%>
											</td>
											<td style="border-bottom:0px;" class="lt"><%=ocartoonboxdetail.FItemList(i).Fitemname%></td>
											<td style="border-bottom:0px;" class="lt"><%=ocartoonboxdetail.FItemList(i).Fitemoptionname%></td>
											<td style="border-bottom:0px;"><%=ocartoonboxdetail.FItemList(i).Frealitemno%></td>
											<td style="border-bottom:0px;"><%=ocartoonboxdetail.FItemList(i).Fitemweight/1000%></td>
										</tr>
						<%
							if ocartoonboxdetail.FItemList(i).Fitemweight <> "" then
							 itemweight = itemweight + ((ocartoonboxdetail.FItemList(i).Fitemweight*ocartoonboxdetail.FItemList(i).Frealitemno)/1000)
							end if
						 next
						%>
						</table>
						</td>
							<td><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).FcartoonboxNweight, 2) %>Kgs</td>
							<td><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).Fcartoonboxweight, 2) %>Kgs</td>
							<!--td class="rt">
								<% if ocartoonboxdetail.FItemList(i-1).FcartoonboxType <> "" then %>
									<%= getcartoonboxtype(ocartoonboxdetail.FItemList(i-1).FcartoonboxType, 1) %>
								<% end if %>
							</td-->

						</tr>
					</tbody>
					<tfoot>
						<tr class="total">
							<td class="bgGry"><strong>Total</strong></td>
							<td class="bgGry" colspan="5"></td>
							<td class="bgGry"><strong><%= FormatNumber(sumcartoonboxNweight, 2) %>Kgs</strong></td>
							<td class="rt bgGry"><strong><%= FormatNumber(sumcartoonboxweight, 2) %>Kgs</strong></td>
						</tr>
					</tfoot>
				</table>
			</div>
		</div>
		<!-- //01 -->
	</div>
	<div class="btnArea tMar30 ct">
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">인쇄하기</button>
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF 전환</button>
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoEXCEL();">엑셀 전환</button>
	</div>
</body>
</html>

<%
set ocartoonboxdetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

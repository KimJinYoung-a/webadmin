<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 견적서
' History : 2014.4.18 정윤정 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx ,loginsite
dim i, j
dim ClsQS , arrList , intLoop
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount, freightcharge ,currencychar, currencyunit,subdamount, totaldamount, dfreightcharge
dim countryLangCD,getPdfDownLinkUrlAdm,addparam,arrQS   ,intQS
dim tplcompanyid, jungsanidx, workidx

idx = requestCheckVar(request("idx"),10) '--cartoonbox idx
loginsite= requestCheckVar(request("ls"),32)
currencyunit = requestCheckVar(request("cunit"),32) '--shop 기준 화폐
tplcompanyid = requestCheckVar(request("tpl"),32)
jungsanidx = requestCheckVar(request("jungsanidx"),10)
workidx = requestCheckVar(request("workidx"),10)

if idx="" then idx=0
'================================================================================
 set ClsQS  = new COffInvoice
 ClsQS.FRectMasterIdx = idx
 ClsQS.FRectLoginsite = loginsite
 ClsQS.FRectJungsanidx = jungsanidx
 ClsQS.FRectWorkidx = workidx
 arrQS = ClsQS.fnGetQuotationSheet
IF isArray(arrQS) THEN
 baljuid                	= arrQS(1,0)
 beasongdate                = arrQS(2,0)
 regdate					= arrQS(3,0)
 priceunit				    = arrQS(4,0)
 totalGoodsPriceWon		    = arrQS(5,0)
 totalDeliverPriceWon 	    = arrQS(6,0)
 totalPriceWon  			= arrQS(7,0)
 totalGoodsPriceForeign	    = arrQS(8,0)
 totalDeliverPriceForeign   = arrQS(9,0)
 totalPriceForeign		    = arrQS(10,0)
 freightTerm                = arrQS(11,0)
 openState                  = arrQS(12,0)
 shippingAddress            = arrQS(13,0)
 invoiceAddress             = arrQS(14,0)
' currencyunit        		= arrQS(16,0)
 currencychar				= arrQS(17,0)
 countryLangCD				= arrQS(15,0)
END IF
 arrList = ClsQS.fnGetQuotationSheetItemList '--견적서는 주문상품 수량 기준으로


 set clsQS = nothing

' if currencyunit = "WON" THEN  '--샵 기준 화폐가 원일때
' 	currencychar = "원"
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
addparam = "idx="&idx&"&ls="&loginsite&"&cunit="&currencyunit&"&tpl="&tplcompanyid&"&ekey="&md5(idx&loginsite)&"&jungsanidx=" & jungsanidx & "&workidx=" & workidx

if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnquotationsheetPdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnquotationsheetPdf.asp?"&addparam
end if
'---------------------------------------------------------
%>
<!-- #include virtual="/admin/lib/popheader_xhtml.asp"-->
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
		.ci img {width:170px; height:27px;}
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
		.vPad1 {padding:1px 0;}
		.vPad10 {padding:10px 0;}
		.vPad15 {padding:15px 0;}
		.bxInner {padding:10px;}
		.tMar10 {margin-top:10px;}
		.tMar15 {margin-top:15px;}
		.tMar20 {margin-top:20px;}
		.lt {text-align:left !important;}
		.rt {text-align:right !important;}
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
					<h1 class="title">QUOTATION SHEET</h1>
					<div class="bgGry bxInner bdrBtm2">
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Order No.</strong></dt>
							<dd class="rt">
								<%For intQS = 0 To UBound(arrQS,2) %>
								 <%=arrQS(0,intQS)%><%IF intQS<UBound(arrQS,2) THEN%>,<%END IF%>
								 <% if (intQS mod 3) = 2 THEN
								 	%>
								 <br>
								 <%END IF%>
								<%Next%>
							</dd>
						</dl>
						<dl class="hor bdrBtm vPad1">
							<dt><strong>Date</strong></dt>
							<dd class="rt">
								<%IF openState = "Open" THEN%>
								<% if Not IsNull(beasongdate) then %>
										<%= Left(beasongdate, 10) %>
									<% end if %>
								<%ELSE%>
									<%= Left(regdate, 10) %>
								<%END IF%>
							</dd>
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
							<dt><strong>Freight charge</strong></dt>
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
						<col width="80px" /><col width="" /><col width="" /><col width="70px" /><col width="70px" /><col width="" />
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
							<td><%IF  arrList(14,intLoop) <> "" THEN%><%=arrList(14,intLoop)%><%ELSE%><%=arrList(6,intLoop)%><%END IF%></td>
							<td class="rt">
							    <%= getdisp_price_currencyChar(arrList(13,intLoop),currencyunit) %>
				                <% if (FALSE) then %>
    								<%IF  currencyunit <> "WON" THEN%>
    								 <%=currencyChar%><%=FormatNumber(arrList(13,intLoop),2)%>
    								<%else%>
    								<%=FormatNumber(arrList(8,intLoop),0)%><%=currencyChar%>
    								<%end if%>
								<%end if%>
							</td>
							<td><%IF openState = "Open" THEN %><%=arrList(11,intLoop)%><%else%><%=arrList(10,intLoop)%><%end if%></td>
							<td class="rt">
							    <%
							    IF openState = "Open" THEN
                			        subdamount = subdamount + (arrList(13,intLoop)*arrList(11,intLoop))
                			    %>
                			    <%= getdisp_price_currencyChar(arrList(13,intLoop)*arrList(11,intLoop),currencyunit) %>
                			    <%else
									subdamount = subdamount + (arrList(13,intLoop)*arrList(10,intLoop))
								%>
								<%= getdisp_price_currencyChar(arrList(13,intLoop)*arrList(10,intLoop),currencyunit) %>
                			    <%end if%>
                			    <% if (FALSE) then %>
								<%IF  currencyunit <> "WON" THEN%>
									<%IF openState = "Open" THEN
										subdamount = subdamount + (arrList(13,intLoop)*arrList(11,intLoop))
										%>
										 <%=currencyChar%><%=FormatNumber(arrList(13,intLoop)*arrList(11,intLoop),2)%>
									<%else
										subdamount = subdamount + (arrList(13,intLoop)*arrList(10,intLoop))
									%>
										<%=currencyChar%><%=FormatNumber(arrList(13,intLoop)*arrList(10,intLoop),2)%>
									<%end if%>
								<%ELSE%>
									<%IF openState = "Open" THEN
										subdamount = subdamount + (arrList(8,intLoop)*arrList(11,intLoop))
										%>
										<%=FormatNumber(arrList(8,intLoop)*arrList(11,intLoop),0)%> <%=currencyChar%>
									<%else
										subdamount = subdamount + (arrList(8,intLoop)*arrList(10,intLoop))
									%>
										<%=FormatNumber(arrList(8,intLoop)*arrList(10,intLoop),0)%> <%=currencyChar%>
									<%end if%>
								<%END IF%>
								<%END IF%>
							</td>
						</tr>
						<%
							Next
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
	<div class="btnArea tMar30 ct">
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">인쇄하기</button>
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF 전환</button>
	</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

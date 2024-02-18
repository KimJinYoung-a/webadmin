<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<% Session.codepage="65001" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  해외 주문서
' History : 2017.06.15 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx ,loginsite, i, j, tplcompanyid
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount ,currencychar, currencyunit, isfixed, reguser, userNameEN, userTel, userMail, oRegUser
dim countryLangCD,getPdfDownLinkUrlAdm,addparam,arrQS   ,intQS
	idx = requestCheckVar(getNumeric(request("idx")),10) '--cartoonbox idx
	loginsite= requestCheckVar(request("ls"),32)
	currencyunit = requestCheckVar(request("cunit"),32) '--shop 기준 화폐
	tplcompanyid = requestCheckVar(request("tpl"),32)

isfixed = false
if idx="" then idx=0

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

reguser = ojumunmaster.FOneItem.Freguser
isfixed = ojumunmaster.FOneItem.FStatecd >= 7

set oRegUser = new CTenByTenMember
	oRegUser.frectUserId = reguser
	oRegUser.fnGetMemberData

userNameEN = trim(oRegUser.fuserNameEN)
userTel = trim(oRegUser.Fdirect070)
userMail = trim(oRegUser.Fusermail)

dim ojumundetail
set ojumundetail= new COrderSheet
	ojumundetail.FRectIdx = idx
	ojumundetail.GetOrderSheetDetail_ENG()

if ojumundetail.ftotalcount > 0 then
	for i = 0 to ojumundetail.FResultCount - 1
		totalPriceForeign = totalPriceForeign + (ojumundetail.FItemList(i).flcprice * ojumundetail.FItemList(i).frealitemno)
	next
	if (ojumundetail.FItemList(0).FcurrencyUnit <> currencyunit) then
		'// 상품 화폐단위가 USD가 아닌 경우
		currencyunit = ojumundetail.FItemList(0).FcurrencyUnit
	end if
end if

totalamount  = getdisp_price_currencyChar(totalPriceForeign ,currencyunit)

'--- pdf 전환처리----------------------------------------
addparam = "idx="&idx&"&ls="&loginsite&"&cunit="&currencyunit&"&tpl="&tplcompanyid&"&ekey="&md5(idx&loginsite)&"&isUtf8=Y"

if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnordersheetPdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnordersheetPdf.asp?"&addparam
end if
'---------------------------------------------------------

dim buymakerid

if ojumundetail.ftotalcount > 0 then
	buymakerid = ojumundetail.FItemList(0).Fmakerid
end if

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
<script src="https://code.jquery.com/jquery-1.11.2.min.js"></script>
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
									<p>14F(GyoYukDong), 57</p>
									<p>Daehak-ro, Jongno-gu</p>
									<p>Seoul, Republic of Korea [03082]</p>
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
						<dl class="hor vPad1">
							<dt><strong>Brand</strong></dt>
							<dd class="rt">
								<%= buymakerid %>
							</dd>
						</dl>
						<!--<dl class="hor vPad1">
							<dt><strong>Total Amount</strong></dt>
							<dd class="rt"><strong><%'=totalamount%></strong></dd>
						</dl>-->
					</div>
				</div>
			</div>
			<div class="vPad10">
				<dl class="ftLt ver w50">
					<dt><strong>Invoice address</strong></dt>
					<dd>
						<p>
							<% '우리쪽에서 보내는거라 그냥 박아넣으면됨. 고정임 %>
							TENBYTEN Inc.
							<br>14F(GyoYukDong), 57, Daehak-ro
							<br>Jongno-gu, Seoul, Republic of Korea (ZIP : 03082)						
							<% if userNameEN<>"" and not(isnull(userNameEN)) then %><br>Attn. <%= userNameEN %><% end if %>
							<br>
							<% if userNameEN<>"" and not(isnull(userNameEN)) then %>Tel：82-<%= userTel %><% end if %>
							<% if userMail<>"" and not(isnull(userMail)) then %>
								<% if userNameEN<>"" and not(isnull(userNameEN)) then %>
									/ 
								<% end if %>
								<%= userMail %>
							<% end if %>
						</p>
					</dd>
				</dl>
				<dl class="ftRt ver w50">
					<dt><strong>Shipping Address</strong></dt>
					<dd>
						<p>
							<% '우리쪽에서 보내는거라 그냥 박아넣으면됨. 고정임 %>
							Tenbyten Logistics Center.
							<br>83, Yongjeonggyeongje-ro 2-gil, Gunnae-myeon,
							<br>Pocheon-si, Gyeonggi-do, Republic of Korea (ZIP :  11154)
							<% if userNameEN<>"" and not(isnull(userNameEN)) then %><br>Attn. <%= userNameEN %><% end if %>
							<br>
							<% if userNameEN<>"" and not(isnull(userNameEN)) then %>Tel：82-<%= userTel %><% end if %>
							<% if userMail<>"" and not(isnull(userMail)) then %>
								<% if userNameEN<>"" and not(isnull(userNameEN)) then %>
									/ 
								<% end if %>
								<%= userMail %>
							<% end if %>
						</p>
					</dd>
				</dl>
			</div>
			<div class="vPad10" id="tableWrap">
				<table id='tblExport'>
					<colgroup>
						<col width="80px" />
						<!--<col width="80px" />-->
						<col width="80px" />
						<col />
						<col />
						<col width="60px" />
						<col width="50px" />
						<col />
					</colgroup>
					<thead>
						<tr>
							<th>Item Code(Buyer)</th>
							<!--<th>BarCode</th>-->
							<th>Item Code(Seller)</th>
							<th>Description</th>
							<th>Option</th>
							<th class="rt">Price<br />(<%= currencyunit %>)</th>
							<th>Qty</th>
							<th class="rt">Amount<br />(<%= currencyunit %>)</th>
						</tr>
					</thead>
					<tbody>
						<%
						if ojumundetail.ftotalcount > 0 then
							 for i = 0 to ojumundetail.FResultCount - 1
							%>
							<tr>
								<td style='mso-number-format:"\@";'><%=BF_MakeTenBarcode(ojumundetail.FItemList(i).Fitemgubun,ojumundetail.FItemList(i).Fitemid,ojumundetail.FItemList(i).Fitemoption)%></td>
								<!--<td><%= ojumundetail.FItemList(i).FPublicBarcode %></td>-->
								<td><%= ojumundetail.FItemList(i).FUpcheManageCode %></td>
								<td class="lt"><%= ojumundetail.FItemList(i).flcitemname %></td>
								<td>
									<%= ojumundetail.FItemList(i).flcitemoptionname %>
								</td>
								<td class="rt" style='mso-number-format:"\@";'>
								    <%= getdisp_price(ojumundetail.FItemList(i).flcprice,currencyunit) %>
								</td>
								<td>
									<% IF isfixed THEN %>
										<%= ojumundetail.FItemList(i).Frealitemno %>
									<% else %>
										<%= ojumundetail.FItemList(i).Fbaljuitemno %>
									<% end if %>
								</td>
								<td class="rt" style='mso-number-format:"\@";'>
									<% IF isfixed THEN %>
										<%= getdisp_price(ojumundetail.FItemList(i).flcprice*ojumundetail.FItemList(i).Frealitemno,currencyunit) %>
									<% else %>
										<%= getdisp_price(ojumundetail.FItemList(i).flcprice*ojumundetail.FItemList(i).Fbaljuitemno,currencyunit) %>
									<% end if %>
								</td>
							</tr>
						<%
							next
						end if
						%>
					</tbody>
					<tfoot>
						<tr class="total">
							<td class="rt bgGry" colspan="5"><strong>Total Amount</strong></td>
							<td class="rt bgGry" colspan="2"><strong><%=totalamount%></strong></td>
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
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" id="btnExport">엑셀받기</button>
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF 전환</button>
	</div>
	<script>
    $(document).ready(function(){
		function itoStr($num) {
            $num < 10 ? $num = '0'+$num : $num;
            return $num.toString();
        }

		function exportTableToExcel(tableID, filename) {
			var downloadLink;
			var dataType = 'application/vnd.ms-excel';
			var tableSelect = document.getElementById(tableID);
			var tableHTML = tableSelect.outerHTML.replace(/<br>/g, '');

			// Specify file name
			filename = filename?filename+'.xls':'excel_data.xls';

			// Create download link element
			downloadLink = document.createElement("a");

			document.body.appendChild(downloadLink);

			if(navigator.msSaveOrOpenBlob){
				var blob = new Blob([tableHTML], {
					type: dataType
				});
				navigator.msSaveOrOpenBlob( blob, filename);
			}else{
				// Create a link to the file
				downloadLink.href = 'data:' + dataType + ', ' + tableHTML.replace(/ /g, '%20');

				// Setting the file name
				downloadLink.download = filename;

				//triggering the function
				downloadLink.click();
			}
		}

        var btn = $('#btnExport');
        var tbl = 'tblExport';

        btn.click(function(e){
            var dt = new Date();
            var year =  itoStr( dt.getFullYear() );
            var month = itoStr( dt.getMonth() + 1 );
            var day =   itoStr( dt.getDate() );
            var hour =  itoStr( dt.getHours() );
            var mins =  itoStr( dt.getMinutes() );

            var postfix = year + month + day + "_" + hour + "_" + mins;
            var fileName = "<%= buymakerid %>_"+ postfix;

			exportTableToExcel(tbl, fileName);

            e.preventDefault();
        });
    });
    </script>
</body>
</html>

<%
set ojumundetail = nothing
set ojumunmaster = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% Session.codepage="949" %>

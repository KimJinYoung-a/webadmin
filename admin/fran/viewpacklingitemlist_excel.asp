<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외출고_인보이스
' History : 2017.03.28 한용민 생성
' pdf 파일 엑셀로 변경시 css 모두 안먹음 테이블 구조로 바꿈.	통 테이블로 안할경우 셀 깨짐
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx , loginsite,invoiceNo,invoicedate, boxidx,ekey, i, j, ClsOI , arrList , intLoop
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount, freightcharge , currencychar, currencyunit, tplcompanyid
dim currcartoonboxno, suminnerboxweight, sumcartoonboxNweight, sumcartoonboxweight, isnewcartoonbox, sumcartoonboxcbm

	tplcompanyid = requestCheckVar(request("tpl"),32)
	idx = requestCheckVar(request("idx"),10)
	loginsite= requestCheckVar(request("ls"),32)
	boxidx= requestCheckVar(request("boxidx"),10)
	currencyunit = requestCheckVar(request("cunit"),32)
	ekey =  requestCheckvar(request("ekey"),32)

if idx="" then idx=0
if 	boxidx = "" then boxidx = 0

if (ekey="") then
    response.write "암호화 키가 올바르지 않습니다.1"
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5(idx&loginsite&boxidx))) then
    response.write "암호화 키가 올바르지 않습니다.2"
    response.end
end if

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
	'currencyunit				= ClsOI.Fcurrencyunit
set ClsOI = nothing

' if currencyunit = "WON" THEN
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

dim ocartoonboxdetail
set ocartoonboxdetail = new CCartoonBox
	ocartoonboxdetail.FRectMasterIdx = boxidx
	ocartoonboxdetail.FRectShopid = baljuid
	ocartoonboxdetail.GetDetailItemList

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<title></title>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
	.a {font-size:11px;}
</style>
</head>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<body>

<table width="800" border="0" cellpadding="3" cellspacing="1">

<% IF tplcompanyid <> "" THEN %>
	<% if (idx >= 1263) then '' 주소가 달라진다. %>
		<tr>
			<td width="50%" colspan=3><img src="http://webadmin.10x10.co.kr/images/logo_ithinkso.jpg" alt="ithinkso" width="182" height="36"/></td>
			<td width="50%" colspan=2 align="center"><strong>PACKING LIST</strong></td>
		</tr>
		<tr>
			<td colspan=2></td>
			<td></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice No.</strong></font></td>
			<td bgcolor="#e1e1e1" align="right" class='txt'><font size=1><%=invoiceNo%></font></td>
		</tr>
		<tr>
			<td colspan=2></td>
			<td></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice Date</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=invoicedate%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1><b>S&T works Inc.</b></font></td>
			<td></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Wholesale ID</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=baljuid%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>4F, 52, Daehak-ro 8ga-gil,</font></td>
			<td><font size=1>Tel : +82 70 4821 1903</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Freight Term</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=freightTerm%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>Jongno-gu, Seoul,</font></td>
			<td><font size=1>Fax : +82 2 2179 8631</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Sub Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=subamount%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>Korea [03086]</font></td>
			<td><font size=1>Mail : salesmanger@ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Shipping</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=freightcharge%></strong></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>VAT Reg.No. : 101-86-84103</font></td>
			<td><font size=1>Website : www.ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=totalamount%></strong></font></td>
		</tr>
	<% else %>
		<tr>
			<td width="50%" colspan=3><img src="http://webadmin.10x10.co.kr/images/logo_ithinkso.jpg" alt="ithinkso" width="182" height="36"/></td>
			<td width="50%" colspan=2 align="center"><strong>PACKING LIST</strong></td>
		</tr>
		<tr>
			<td colspan=2></td>
			<td></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice No.</strong></font></td>
			<td bgcolor="#e1e1e1" align="right" class='txt'><font size=1><%=invoiceNo%></font></td>
		</tr>
		<tr>
			<td colspan=2></td>
			<td></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice Date</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=invoicedate%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1><b>S&T works Inc.</b></font></td>
			<td></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Wholesale ID</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=baljuid%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>5F, ERH bldg, 1-74,</font></td>
			<td><font size=1>Tel : +82 70 4821 1903</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Freight Term</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=freightTerm%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>Dongsung-dong, Jongno-gu,</font></td>
			<td><font size=1>Fax : +82 2 2179 8631</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Sub Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=subamount%></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>Seoul, Korea [110-809]</font></td>
			<td><font size=1>Mail : salesmanger@ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Shipping</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=freightcharge%></strong></font></td>
		</tr>
		<tr>
			<td colspan=2><font size=1>VAT Reg.No. : 101-86-84103</font></td>
			<td><font size=1>Website : www.ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=totalamount%></strong></font></td>
		</tr>
	<% end if %>
<% else %>
	<tr>
		<td width="50%" colspan=3><img src="http://webadmin.10x10.co.kr/images/10x10_ci.jpg" alt="TENBYTEN" /></td>
		<td width="50%" colspan=2 align="center"><strong>PACKING LIST</strong></td>
	</tr>
	<tr>
		<td colspan=3></td>
		<td height=2 colspan=2 bgcolor="black"></td>
	</tr>
	<tr>
		<td colspan=2></td>
		<td></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Invoice No.</strong></font></td>
		<td bgcolor="#e1e1e1" align="right" class='txt'><font size=1><%=invoiceNo%></font></td>
	</tr>
	<tr>
		<td colspan=2></td>
		<td></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Invoice Date</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><%=invoicedate%></font></td>
	</tr>
	<tr>
		<td colspan=2><font size=1><b>TENBYTEN Inc.</b></font></td>
		<td></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Wholesale ID</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><%=baljuid%></font></td>
	</tr>
	<tr>
		<td colspan=2><font size=1>14F(GyoYukDong)</font></td>
		<td><font size=1>Tel : +82 2 554 2033</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Freight Term</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><%=freightTerm%></font></td>
	</tr>
	<tr>
		<td colspan=2><font size=1>57, Daehak-ro, Jongno-gu</font></td>
		<td><font size=1>Fax : +82 2 2179 9244</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Sub Amount</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=subamount%></strong></font></td>
	</tr>
	<tr>
		<td colspan=2><font size=1>Seoul, Korea [03082]</font></td>
		<td><font size=1>Mail : wholesale@10x10.co.kr</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Shipping</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=freightcharge%></strong></font></td>
	</tr>
	<tr>
		<td colspan=2><font size=1>VAT Reg.No. : 211-87-00620</font></td>
		<td><font size=1>Website : wholesale.10x10.co.kr</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=totalamount%></strong></font></td>
	</tr>
<% end if %>

<tr><td height=30 colspan=5></td></tr>
<tr>
	<td width="50%" colspan=3><font size=1><strong>Invoice address</strong></font></td>
	<td width="50%" colspan=2><font size=1><strong>Shipping Address</strong></font></td>
</tr>
<tr>
	<td width="50%" colspan=3 valign="top">
		<font size=1>
			<%= nl2br(invoiceAddress) %>
		</font>
	</td>
	<td width="50%" colspan=2 valign="top">
		<font size=1>
			<% if (trim(replace(shippingAddress,chr(13)&chr(10),"")) = "Same as Above") then %>
				<p><%= nl2br(invoiceAddress) %></p>
			<% else %>
				<p><%= nl2br(shippingAddress) %></p>
			<% end if %>
		</font>
	</td>
</tr>
<tr><td height=30 colspan=5></td></tr>
<tr><td height=1 bgcolor="gray" colspan=5></td></tr>
<tr> 
<td colspan=5>
<table class="a">
					<colgroup>
						<col width="60px" />
						<col width="120px" />
						<col width="" />
						<col width="100px" />
						<col width="30px" />
						<col width="80px" />
						<col width="" />
						<col width="" />
					</colgroup>
					<thead>
						<tr>
							<th>BOX NO.</th> 
							<th>Item Code</th>
							<th>Description</th>
							<th>Option</th>
							<th>Qty</th>
							<th>weight</th>
							<th>N weight</th>
							<th>G weight</th> 
						</tr>
					</thead>
					<tbody>
					<%

				
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
							<%if i > 0 then%>
										 
										</table>
							</td>
							
								<td><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).FcartoonboxNweight, 2) %>Kgs</td>
								<td class="rt"><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).Fcartoonboxweight, 2) %>Kgs</td>	
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
									<table class="a" style="border-top:0px;border-bottom:0px;" cellpadding="0" cellspacing="0" border="0">
										<colgroup>
											<col width="100px" /><col width="" /><col width="100px" /><col width="30px" /><col width="80px" />
										</colgroup>
					<%end if %>					
										<tr>
											<td style="border-bottom:0px;"><%=i%><%=ocartoonboxdetail.FItemList(i).Fitemgubun%><%=ocartoonboxdetail.FItemList(i).Fitemid%><%=ocartoonboxdetail.FItemList(i).Fitemoption%></td>
											<td style="border-bottom:0px;" class="lt"><%=ocartoonboxdetail.FItemList(i).Fitemname%></td>
											<td style="border-bottom:0px;"  class="lt"><%=ocartoonboxdetail.FItemList(i).Fitemoptionname%></td>
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
								<td class="rt"><%= FormatNumber(ocartoonboxdetail.FItemList(i-1).Fcartoonboxweight, 2) %>Kgs</td>
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
	</td>
	</tr>
</table>
</body>
</html>

<%
set ocartoonboxdetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

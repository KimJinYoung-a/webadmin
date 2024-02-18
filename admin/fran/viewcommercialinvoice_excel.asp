<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외출고_인보이스
' History : 2016.08.03 한용민 생성
' pdf 파일 엑셀로 변경시 css 모두 안먹음 테이블 구조로 바꿈.	통 테이블로 안할경우 셀 깨짐
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx , loginsite,invoiceNo,invoicedate,ekey
dim i, j
dim ClsOI , arrList , intLoop
dim baljucode,baljuid,beasongdate,regdate,priceunit,totalGoodsPriceWon,totalDeliverPriceWon,totalPriceWon
Dim totalGoodsPriceForeign,totalDeliverPriceForeign, totalPriceForeign,freightTerm,openState,shippingAddress,invoiceAddress
dim subamount, totalamount, freightcharge , currencychar, currencyunit,subdamount, totaldamount,comment
dim tplcompanyid, jungsanidx, workidx

tplcompanyid = requestCheckVar(request("tpl"),32)
idx = requestCheckVar(request("idx"),10)
loginsite= requestCheckVar(request("ls"),32)
currencyunit = requestCheckVar(request("cunit"),32)
ekey =  requestCheckvar(request("ekey"),32)
jungsanidx = requestCheckVar(request("jungsanidx"),10)
workidx = requestCheckVar(request("workidx"),10)
if idx="" then idx=0

if (ekey="") then
    response.write "암호화 키가 올바르지 않습니다.1"
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5(idx&loginsite))) then
    response.write "암호화 키가 올바르지 않습니다.2"
    response.end
end if
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
' currencyunit				= ClsOI.Fcurrencyunit
 comment						= ClsOI.Fcomment
 if baljucode <> "" then
 ClsOI.FRectbaljucode	= baljucode
 arrList = ClsOI.fnGetFranItemList
 end if
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
</style>
</head>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<body>

<table width="800" border="0" cellpadding="3" cellspacing="1">

<% IF tplcompanyid <> "" THEN %>
	<% if (idx >= 1263) then '' 주소가 달라진다. %>
		<tr>
			<td width="50%" colspan=5><img src="http://webadmin.10x10.co.kr/images/logo_ithinkso.jpg" alt="ithinkso" width="182" height="36"/></td>
			<td width="50%" colspan=2 align="center"><strong>COMMERCIAL INVOICE</strong></td>
		</tr>
		<tr>
			<td colspan=3></td>
			<td colspan=2></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice No.</strong></font></td>
			<td bgcolor="#e1e1e1" align="right" class='txt'><font size=1><%=invoiceNo%></font></td>
		</tr>
		<tr>
			<td colspan=3></td>
			<td colspan=2></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice Date</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=invoicedate%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1><b>S&T works Inc.</b></font></td>
			<td colspan=2></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Wholesale ID</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=baljuid%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>4F, 52, Daehak-ro 8ga-gil,</font></td>
			<td colspan=2><font size=1>Tel : +82 70 4821 1903</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Freight Term</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=freightTerm%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>Jongno-gu, Seoul,</font></td>
			<td colspan=2><font size=1>Fax : +82 2 2179 8631</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Sub Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=subamount%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>Korea [03086]</font></td>
			<td colspan=2><font size=1>Mail : salesmanger@ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Shipping</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=freightcharge%></strong></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>VAT Reg.No. : 101-86-84103</font></td>
			<td colspan=2><font size=1>Website : www.ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=totalamount%></strong></font></td>
		</tr>
	<% else %>
		<tr>
			<td width="50%" colspan=5><img src="http://webadmin.10x10.co.kr/images/logo_ithinkso.jpg" alt="ithinkso" width="182" height="36"/></td>
			<td width="50%" colspan=2 align="center"><strong>COMMERCIAL INVOICE</strong></td>
		</tr>
		<tr>
			<td colspan=3></td>
			<td colspan=2></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice No.</strong></font></td>
			<td bgcolor="#e1e1e1" align="right" class='txt'><font size=1><%=invoiceNo%></font></td>
		</tr>
		<tr>
			<td colspan=3></td>
			<td colspan=2></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Invoice Date</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=invoicedate%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1><b>S&T works Inc.</b></font></td>
			<td colspan=2></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Wholesale ID</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=baljuid%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>5F, ERH bldg, 1-74,</font></td>
			<td colspan=2><font size=1>Tel : +82 70 4821 1903</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Freight Term</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=freightTerm%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>Dongsung-dong, Jongno-gu,</font></td>
			<td colspan=2><font size=1>Fax : +82 2 2179 8631</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Sub Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><%=subamount%></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>Seoul, Korea [110-809]</font></td>
			<td colspan=2><font size=1>Mail : salesmanger@ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Shipping</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=freightcharge%></strong></font></td>
		</tr>
		<tr>
			<td colspan=3><font size=1>VAT Reg.No. : 101-86-84103</font></td>
			<td colspan=2><font size=1>Website : www.ithinksoweb.com</font></td>
			<td bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
			<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=totalamount%></strong></font></td>
		</tr>
	<% end if %>
<% else %>
	<tr>
		<td width="50%" colspan=5><img src="http://webadmin.10x10.co.kr/images/10x10_ci.jpg" alt="TENBYTEN" /></td>
		<td width="50%" colspan=2 align="center"><strong>COMMERCIAL INVOICE</strong></td>
	</tr>
	<tr>
		<td colspan=5></td>
		<td height=2 colspan=2 bgcolor="black"></td>
	</tr>
	<tr>
		<td colspan=3></td>
		<td colspan=2></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Invoice No.</strong></font></td>
		<td bgcolor="#e1e1e1" align="right" class='txt'><font size=1><%=invoiceNo%></font></td>
	</tr>
	<tr>
		<td colspan=3></td>
		<td colspan=2></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Invoice Date</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><%=invoicedate%></font></td>
	</tr>
	<tr>
		<td colspan=3><font size=1><b>TENBYTEN Inc.</b></font></td>
		<td colspan=2></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Wholesale ID</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><%=baljuid%></font></td>
	</tr>
	<tr>
		<td colspan=3><font size=1>14F(GyoYukDong)</font></td>
		<td colspan=2><font size=1>Tel : +82 2 554 2033</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Freight Term</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><%=freightTerm%></font></td>
	</tr>
	<tr>
		<td colspan=3><font size=1>57, Daehak-ro, Jongno-gu</font></td>
		<td colspan=2><font size=1>Fax : +82 2 2179 9244</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Sub Amount</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=subamount%></strong></font></td>
	</tr>
	<tr>
		<td colspan=3><font size=1>Seoul, Korea [03082]</font></td>
		<td colspan=2><font size=1>Mail : wholesale@10x10.co.kr</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Shipping</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=freightcharge%></strong></font></td>
	</tr>
	<tr>
		<td colspan=3><font size=1>VAT Reg.No. : 211-87-00620</font></td>
		<td colspan=2><font size=1>Website : wholesale.10x10.co.kr</font></td>
		<td bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
		<td bgcolor="#e1e1e1" align="right"><font size=1><strong><%=totalamount%></strong></font></td>
	</tr>
<% end if %>



<tr><td height=30></td></tr>
<tr>
	<td width="50%" colspan=4><font size=1><strong>Invoice address</strong></font></td>
	<td width="50%" colspan=3><font size=1><strong>Shipping Address</strong></font></td>
</tr>
<tr>
	<td width="50%" colspan=4 valign="top">
		<font size=1>
			<% if (invoiceAddress = "Same as Above") then %>
				<%= nl2br(shippingAddress) %>
			<% else %>
				<%= nl2br(invoiceAddress) %>
			<% end if %>
		</font>
	</td>
	<td width="50%" colspan=3 valign="top"><font size=1><%= nl2br(shippingAddress) %></font></td>
</tr>



<tr colspan=7><td height=30></td></tr>
<tr><td height=1 bgcolor="gray" colspan=7></td></tr>
<tr align="center">
	<td bgcolor="#e1e1e1" width="8%"><font size=1><strong>Item Code</strong></font></td>
	<td bgcolor="#e1e1e1" width="4%"></td>
	<td bgcolor="#e1e1e1"><font size=1><strong>Description</strong></font></td>
	<td bgcolor="#e1e1e1" width="8%"><font size=1><strong>Option</strong></font></td>
	<td bgcolor="#e1e1e1" width="9%"><font size=1><strong>Price</strong></font></td>
	<td bgcolor="#e1e1e1" width="8%"><font size=1><strong>Quantity</strong></font></td>
	<td bgcolor="#e1e1e1" width="10%"><font size=1><strong>Amount</strong></font></td>
</tr>
<%
subdamount = 0
totaldamount = 0

if isArray(arrList) then
	for intLoop = 0 To UBound(arrList,2)
%>
	<tr>
		<td class='txt' align="left"><font size=1><%=BF_MakeTenBarcode(arrList(1,intLoop),arrList(3,intLoop),arrList(4,intLoop))%></font></td>
		<td></td>
		<td align="left"><font size=1><%=arrList(5,intLoop)%></font></td>
		<td align="left"><font size=1><%=arrList(6,intLoop)%></font></td>
		<td align="right">
			<font size=1>
			    <%= getdisp_price_currencyChar(arrList(13,intLoop),currencyunit) %>
				<% if (FALSE) then %>
    				<% IF currencyunit <> "WON" THEN %>
    					<%=currencyChar%><%=FormatNumber(arrList(13,intLoop),2)%>
    				<% else %>
    					<%=FormatNumber(arrList(8,intLoop),0)%><%=currencyChar%>
    				<% end if %>
				<% end if %>
			</font>
		</td>
		<td align="right"><font size=1><%=arrList(11,intLoop)%></font></td>
		<td align="right">
			<font size=1>
			    <%
			    subdamount = subdamount + (arrList(13,intLoop)*arrList(11,intLoop))
			    %>
			    <%= getdisp_price_currencyChar(arrList(13,intLoop)*arrList(11,intLoop),currencyunit) %>
			    <% if (FALSE) then %>
    				<%
    				IF  currencyunit <> "WON" THEN
    					subdamount = subdamount + (arrList(13,intLoop)*arrList(11,intLoop))
    				%>
    					<%=currencyChar%><%=FormatNumber(arrList(13,intLoop)*arrList(11,intLoop),2)%>
    				<%
    				ELSE
    					 subdamount = subdamount + (arrList(8,intLoop)*arrList(11,intLoop))
    				 %>
    					<%=FormatNumber(arrList(8,intLoop)*arrList(11,intLoop),0)%> <%=currencyChar%>
    				<% END IF %>
				<% END IF %>
			</font>
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
<tr>
	<td colspan=6 align="right"><font size=1><strong>Sub Amount</strong></font></td>
	<td align="right"><font size=1><strong><%=subdamount%></strong></font></td>
</tr>
<tr>
	<td colspan=6 align="right"><font size=1>Freight charge</font></td>
	<td align="right"><font size=1><strong class="cGry"><%=freightcharge%></strong></font></td>
</tr>
<tr><td height=2 bgcolor="black" colspan=7></td></tr>
<tr>
	<td colspan=6 align="right" bgcolor="#e1e1e1"><font size=1><strong>Total Amount</strong></font></td>
	<td align="right" bgcolor="#e1e1e1"><font size=1><strong><%=totaldamount%></strong></font></td>
</tr>
<tr><td height=20 colspan=7></td></tr>
<tr><td colspan=7><font size=1><strong>Remarks</strong></font></td></tr>
<tr><td height=1 bgcolor="gray" colspan=7></td></tr>
<tr><td colspan=7><font size=1><%=nl2br(comment)%></font></td></tr>
<tr><td height=60 colspan=7></td></tr>
<tr>
	<td colspan=5 align="right"><font size=1><strong>SIGNED BY</strong></font></td>
	<td colspan=2></td>
</tr>
<tr>
	<td height=1 colspan=4></td>
	<td height=1 bgcolor="gray" colspan=3></td>
</tr>
</table>

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

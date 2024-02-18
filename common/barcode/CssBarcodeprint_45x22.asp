<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'####################################################
' Description : 바코드 출력 45X22
' History : 2023.03.09 한용민 생성
'####################################################
%>
<!-- include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim iloop, chkCount, maxCount, i, printpriceyn, makeriddispyn, itemid, itemoption, barcodetype, generalbarcode
dim cksel, prdname, prdoptionname, customerprice, sellprice, saleprice, saleyn, itemgubun, fixedno, customerprice_foreign
dim prdcode, socname, socname_kor, isforeignprint, currencychar, prdname_foreign, prdoptionname_foreign
	printpriceyn = requestcheckvar(request.Form("printpriceyn"),1)		' 금액표시방식
	makeriddispyn = requestcheckvar(request.Form("makeriddispyn"),1)	' 브랜드표시
	isforeignprint = requestcheckvar(request.Form("isforeignprint"),1)	' 표시상품명
	barcodetype = requestcheckvar(request.Form("barcodetype"),1)	' 바코드구분
	currencychar = requestcheckvar(request.Form("currencychar"),1)	' 화폐구분
	set cksel = request.Form("cksel")
	set socname = request.Form("socname")	' 브랜드명영문
	set socname_kor = request.Form("socname_kor")	' 브랜드명한글
	set prdname = request.Form("prdname")	' 상품명
	set prdoptionname = request.Form("prdoptionname")		' 옵션명
	set prdname_foreign = request.Form("prdname_foreign")		' 해외상품명
	set prdoptionname_foreign = request.Form("prdoptionname_foreign")		' 해외옵션명
	set customerprice = request.Form("customerprice")	' 소비자가
	set sellprice = request.Form("sellprice")	' 판매가
	set saleprice = request.Form("saleprice")	' 할인가
	set customerprice_foreign = request.Form("customerprice_foreign")	' 해외 or 매장별 소비자가
	set saleyn = request.Form("saleyn")		' 할인여부
	set itemgubun = request.Form("itemgubun")
	set itemid = request.Form("itemid")
	set itemoption = request.Form("itemoption")
	set fixedno = request.Form("fixedno")	' 수량
	set prdcode = request.Form("prdcode")	' 물류코드
	set generalbarcode = request.Form("generalbarcode")		' 범용바코드

if printpriceyn="" then printpriceyn="Y"	'(Y:소비자가, C:할인가, R:판매가 표시, S:심플금액표시, N:금액표시안함)
if makeriddispyn="" then makeriddispyn="Y"	'브랜드표시여부
if isforeignprint="" then isforeignprint="N"'	 표시상품명
if currencychar="" then currencychar="￦"	' 화폐구분
if printpriceyn="S" then currencychar=""
if barcodetype="" then barcodetype="T"	' 바코드구분
chkCount = 0
maxCount = 0

if cksel.count<1 then
	Call Alert_close("출력할 상품이 없습니다.")
	Response.End
end if

for iloop=1 to cksel.count
	maxCount = maxCount + fixedno(cksel(iloop))
next	
%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>10x10 Barcode Print</title>
<style>
body {
	margin: 0;
	padding: 0;
	font: 10pt "Tahoma";
}

* {
	box-sizing: border-box;
	-moz-box-sizing: border-box;
}

.page {
	width: 4.5cm;
	min-height: 2.2cm;
	padding: 0 0 0 0;
	margin: 0 auto;
	border-radius: 1px;
	background: white;
	box-shadow: 0 0 5px rgba(0, 0, 0, 0.1);
}

.subpage {
	height: 0mm;
}

.labelItem {
	width: 4.5cm;
	height: 2.2cm;
	padding: 0 0 0 0;
	border: 0px solid #FFF;
	border-radius: 1px;
	display: inline-block;
	margin-right: 0mm;
	text-align:center;
}

.labelItem:nth-child(5n) {
  margin-right: 0;
}

.barcode {
	margin:auto;
	width:100% !important;
	margin-top: -4mm;
}
.barcode img{
	//max-width: 30mm;
	//min-height: 3mm;
}

.barcodeText {
	margin-top: -1mm;
	font-size: 7pt;
	font-weight: bold;
}

.barcodeDesc {
	margin: -5mm 0 0 -20mm;
	width:200%;
	-webkit-transform:scale(0.5);
}
.barcodeDesc .brandName {
	width: 100%;
	font: 11pt "Malgun Gothic";
	text-align: left;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
}
.barcodeDesc .prdname {
	width: 100%;
	font: 10pt "Malgun Gothic";
	text-align: left;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
}
.barcodeDesc .itemPrice {
	width: 100%;
	font: 14pt "Malgun Gothic";
	text-align: right;
	display: inline-block;
	font-weight: bold;
}
.barcodeDesc .line {
	width: 100%;
	font: 7pt "Malgun Gothic";
	text-align: left;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	margin-top: -3mm;
}
@page {
	size: 4.5cm 2.2cm;
	margin: 0;
	//size: landscape;
    //size: portrait;
}

@media print {
	.page {
		margin: 0;
		border: initial;
		border-radius: initial;
		width: initial;
		min-height: initial;
		box-shadow: initial;
		background: initial;
		page-break-after: always;
	    //size: landscape;
        //size: portrait;
	}
}
</style>
<script type="text/javascript" src="/js/jquery-2.2.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-barcode.min.js"></script>
<script type="text/javascript">
window.onload = function(){
	window.print();
}
</script>
</head>
<body>
<div class="book">
	<div class="page">
		<div class="subpage">
			<%
			for iloop=1 to cksel.count
				for i=1 to fixedno(cksel(iloop))
					if prdname(cksel(iloop))<>"" then
			%>
						<div class="labelItem">
							<div class="barcodeDesc">
								<% if makeriddispyn="Y" then %>
									<div class="brandName"><%= socname_kor(cksel(iloop)) & " " & socname(cksel(iloop)) %></div>
								<% end if %>
								<div class="prdname">
									<% if isforeignprint="Y" then %>
										<%= prdname_foreign(cksel(iloop)) %>
										<%= chkIIF(prdoptionname_foreign(cksel(iloop))<>"","<br>["&prdoptionname_foreign(cksel(iloop))&"]","<br>&nbsp;") %>
									<% else %>
										<%= prdname(cksel(iloop)) %>
										<%= chkIIF(prdoptionname(cksel(iloop))<>"","<br>["&prdoptionname(cksel(iloop))&"]","<br>&nbsp;") %>
									<% end if %>
								</div>
								<div class="itemPrice">
									<% if isforeignprint="Y" then %>
										<%= currencychar %>&nbsp;<%= FormatNumber(customerprice_foreign(cksel(iloop)),0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<% else %>
										<% if printpriceyn="Y" then %>
											<%= currencychar %>&nbsp;<%= FormatNumber(customerprice(cksel(iloop)),0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% elseif printpriceyn="C" then %>
											<%= currencychar %>&nbsp;<%= FormatNumber(saleprice(cksel(iloop)),0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% else %>
											<%= currencychar %>&nbsp;<%= FormatNumber(sellprice(cksel(iloop)),0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% end if %>
									<% end if %>
								</div>
								<div class="line">____________________________________________________________________________</div>
							</div>
							<div class="barcode" id="label<%=chkCount%>"></div>
							<div class="barcodeText">
								<%
								' 물류코드
								if barcodetype="T" then
								%>
									<%=itemgubun(cksel(iloop)) &"-"& BF_GetFormattedItemId(itemid(cksel(iloop))) &"-"& itemoption(cksel(iloop))%>
								<% else %>
									<%= generalbarcode(cksel(iloop)) %>
								<% end if %>
							</div>
							<script type="text/javascript">
								<%
								' 물류코드
								if barcodetype="T" then
								%>
									$("#label<%=chkCount%>").barcode("<%= prdcode(cksel(iloop)) %>","code128",{barWidth:1,barHeight:15,showHRI:false,output:"svg"});
								<% else %>
									$("#label<%=chkCount%>").barcode("<%= generalbarcode(cksel(iloop)) %>","code128",{barWidth:1,barHeight:15,showHRI:false,output:"svg"});
								<% end if %>
							</script>
						</div>
			<%
						chkCount = chkCount +1
					end if
				next
			next
			%>
		</div>
	</div>
</div>
</body>
</html>
<%
set cksel = nothing
set socname = nothing
set socname_kor = nothing
set prdname = nothing
set prdoptionname = nothing
set prdname_foreign = nothing
set prdoptionname_foreign = nothing
set customerprice = nothing
set sellprice = nothing
set saleprice = nothing
set customerprice_foreign = nothing
set saleyn = nothing
set itemgubun = nothing
set itemid = nothing
set itemoption = nothing
set fixedno = nothing
set prdcode = nothing
set generalbarcode = nothing

session.codePage = 949
%>
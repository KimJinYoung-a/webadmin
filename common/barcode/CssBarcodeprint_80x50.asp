<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'####################################################
' Description : 바코드 출력 80X50
' History : 2023.09.12 한용민 생성
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
dim brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode
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
    set brandrackcode = request.Form("prtidx")  ' 브랜드랙코드
    set itemrackcode = request.Form("itemrackcode")  ' 상품랙코드
    set itemoptionrackcode = request.Form("itemoptionrackcode")  ' 상품옵션랙코드
    set subitemrackcode = request.Form("subitemrackcode")  ' 상품보조랙코드

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
p {
	float: left;
}

body {
	margin: 0;
	padding: 0;
	font: 15pt "Tahoma";
}

* {
	box-sizing: border-box;
	-moz-box-sizing: border-box;
}

.page {
	width: 8.0cm;
	min-height: 5.0cm;
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
	width: 8.0cm;
	height: 5.0cm;
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
	margin: -17mm 0 0 0;
	font: 11pt "Malgun Gothic";
	text-align: center;
}
.barcode img{
	//max-width: 30mm;
	//min-height: 3mm;
}

.barcodeDesc {
	margin: -17mm 0 0 -35mm;
	width:200%;
	-webkit-transform:scale(0.5);
}
.barcodeDesc .rackCode {
	width: 100%;
	font: 18pt "Malgun Gothic";
	text-align: left;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
}
.barcodeDesc .brandName {
	width: 100%;
	font: 15pt "Malgun Gothic";
	text-align: left;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
}
.barcodeDesc .prdname {
	width: 100%;
	font: 15pt "Malgun Gothic";
	text-align: left;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;ssssssss
	font-weight: bold;
}
.barcodeDesc .itemId {
	width: 100%;
	font: 80pt "Malgun Gothic";
	text-align: center;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
	margin-top: -6mm;	
}
.barcodeDesc .prdCode {
	width: 100%;
	font: 20pt "Malgun Gothic";
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
	margin-top: -6mm;
}
.barcodeDesc .prdCode .itemGubunItemId{
	float: left;
}
.barcodeDesc .prdCode .ItemOption{
	font: 30pt "Malgun Gothic";
	text-align: left;
	font-weight: bold;
}
.barcodeDesc .generalBarcode {
	margin: -2mm 0 0 0;
	width: 100%;
	font: 20pt "Malgun Gothic";
	text-align: right;
	display: block;
	white-space: nowrap;
	overflow: hidden;
	text-overflow: ellipsis;
	font-weight: bold;
}

@page {
	size: 8.0cm 5.0cm;
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
				                <% ' 브랜드랙과 보조랙에 값이 있는데 브랜드랙과 보조랙 값이 틀린거 %>
                                <% if brandrackcode<>"" and subitemrackcode<>"" and brandrackcode<>subitemrackcode then %>
                                    <div class="rackCode">보조랙 : <%= subitemrackcode(cksel(iloop)) %></div>
                                <% elseif brandrackcode<>"" then %>
                                    <div class="rackCode">브랜드랙 : <%= brandrackcode(cksel(iloop)) %></div>
                                <% end if %>

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
								<div class="itemId">
                                    <%= BF_GetFormattedItemId(itemid(cksel(iloop))) %>
								</div>
								<div class="prdCode">
									<div class="itemGubunItemId">
                                    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    	<%= itemgubun(cksel(iloop)) &"-"& BF_GetFormattedItemId(itemid(cksel(iloop))) %>
									</div>
									<div class="ItemOption">
										<%= "-"& itemoption(cksel(iloop)) %>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									</div>
								</div>
								<div class="generalBarcode">
                                    <%= generalbarcode(cksel(iloop)) %>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								</div>
							</div>
							<%
							' 테스트용 z들어간 이중옵션 상품코드 17728886
							%>
							<div class="barcode" id="label<%=chkCount%>"></div>
							<script type="text/javascript">
								$("#label<%=chkCount%>").barcode("<%= prdcode(cksel(iloop)) %>","code128",{barWidth:1,barHeight:20,showHRI:false,output:"svg"});
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
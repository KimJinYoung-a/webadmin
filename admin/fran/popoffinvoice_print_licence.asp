<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<%

dim idx, mode
dim i, j

idx = request("idx")
mode = request("mode")



if idx="" then idx=0

'================================================================================
dim ocoffinvoice

set ocoffinvoice = new COffInvoice

ocoffinvoice.FRectMasterIdx = idx

ocoffinvoice.GetMasterOne

'================================================================================
if C_ADMIN_USER then
elseif (C_IS_SHOP = true) then

	if (ocoffinvoice.FOneItem.Fshopid <> C_STREETSHOPID) then
		response.write "<script>alert('잘못된 접근입니다.');</script>"
		response.end
	end if

end if
%>
<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + CStr(ocoffinvoice.FOneItem.Fshopname) + "_" + CStr(ocoffinvoice.FOneItem.Finvoiceno) + "_summary.xls"
end if
%>
<%

'================================================================================
dim ocoffinvoicedetail

set ocoffinvoicedetail = new COffInvoice

ocoffinvoicedetail.FRectMasterIdx = idx
ocoffinvoicedetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid

ocoffinvoicedetail.GetDetailList



'================================================================================
dim ocoffinvoiceproductdetail

set ocoffinvoiceproductdetail = new COffInvoice

ocoffinvoiceproductdetail.FRectMasterIdx = idx
ocoffinvoiceproductdetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid

ocoffinvoiceproductdetail.GetProductDetailList



'================================================================================
dim avggoodsprice

if (ocoffinvoicedetail.FResultCount = "") then
	ocoffinvoicedetail.FResultCount = 0
end if

if (ocoffinvoice.FOneItem.Ftotalprice = "") then
	ocoffinvoice.FOneItem.Ftotalprice = 0
end if

if (ocoffinvoicedetail.FResultCount = 0) then
	avggoodsprice = 0
else
	avggoodsprice = ocoffinvoice.FOneItem.Ftotalprice / ocoffinvoicedetail.FResultCount
end if



'================================================================================
'// 원화
dim totalgoodsprice
dim totalboxprice
dim totalprice

'// 외화
dim totalgoodspricecalc
dim totalboxpricecalc
dim totalpricecalc

dim exchangerate
dim pointprice

totalgoodsprice = ocoffinvoice.FOneItem.Ftotalgoodsprice
totalboxprice = ocoffinvoice.FOneItem.Ftotalboxprice
totalprice = totalgoodsprice + totalboxprice

pointprice = 2
if (ocoffinvoice.FOneItem.Fpriceunit = "JPY") then
	'// 엔화는 100을 나눠준다.
	exchangerate = ocoffinvoice.FOneItem.Fexchangerate / 100
	pointprice = 0
else
	exchangerate = ocoffinvoice.FOneItem.Fexchangerate
end if

'totalgoodspricecalc = FormatNumber((totalgoodsprice / exchangerate), pointprice)
'totalboxpricecalc = FormatNumber((totalboxprice / exchangerate), pointprice)
'totalpricecalc = FormatNumber(((totalgoodsprice + totalboxprice) / exchangerate), pointprice)
totalgoodspricecalc =FormatNumber(ocoffinvoice.FOneItem.FtotalGoodsPriceForeign,2)
totalboxpricecalc = FormatNumber(ocoffinvoice.FOneItem.FtotalDeliverPriceForeign,2)
totalpricecalc = FormatNumber(ocoffinvoice.FOneItem.FtotalPriceForeign,2)

totalgoodsprice = FormatNumber(totalgoodsprice, 0)
totalboxprice = FormatNumber(totalboxprice, 0)
totalprice = FormatNumber(totalprice, 0)



'================================================================================
dim totalnweight, totalgweight, totalemsprice

totalnweight = 0
totalgweight = 0
totalemsprice = 0
for i=0 to ocoffinvoicedetail.FResultCount-1
	totalnweight = totalnweight + ocoffinvoicedetail.FItemList(i).Fnweight
	totalgweight = totalgweight + ocoffinvoicedetail.FItemList(i).Fgweight
	totalemsprice = totalemsprice + ocoffinvoicedetail.FItemList(i).Femsprice
next



'================================================================================
dim invoicedate, tmpdate, tmpyear, tmpmonth, tmpday
dim carrierdate

dim arrMonth(12)
arrMonth(1) = "January"
arrMonth(2) = "February"
arrMonth(3) = "March"
arrMonth(4) = "April"
arrMonth(5) = "May"
arrMonth(6) = "June"
arrMonth(7) = "July"
arrMonth(8) = "August"
arrMonth(9) = "September"
arrMonth(10) = "October"
arrMonth(11) = "November"
arrMonth(12) = "December"

if (ocoffinvoice.FOneItem.Finvoicedate <> "") then
	tmpdate = CDate(ocoffinvoice.FOneItem.Finvoicedate)

	tmpyear = Right(Year(tmpdate), 2)
	tmpmonth = Left(arrMonth(Month(tmpdate)), 3)
	tmpday = Right(("0" & Day(tmpdate)), 2)

	invoicedate = tmpday & "-" & tmpmonth & "-" & tmpyear
end if

if (ocoffinvoice.FOneItem.Fcarrierdate <> "") then
	tmpdate = CDate(ocoffinvoice.FOneItem.Fcarrierdate)

	tmpyear = Right(Year(tmpdate), 2)
	tmpmonth = Left(arrMonth(Month(tmpdate)), 3)
	tmpday = Right(("0" & Day(tmpdate)), 2)

	carrierdate = tmpday & "-" & tmpmonth & "-" & tmpyear
end if

%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<link rel=File-List href="수출신고작성정보.files/filelist.xml">
<style id="수출신고작성정보_24178_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font524178
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
	mso-font-charset:129;}
.font624178
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;}
.xl23924178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0\.00_ ";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24024178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24124178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24224178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24324178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#D8D8D8;
	mso-pattern:black none;
	white-space:nowrap;}
.xl24424178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24524178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0_ ";
	text-align:general;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24624178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"0_ ";
	text-align:right;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24724178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"0\.00_ ";
	text-align:right;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24824178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\0022US$\0022\#\,\#\#0\.00";
	text-align:right;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl24924178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\0022\20A9\0022\#\,\#\#0\;\[Red\]\\-\0022\20A9\0022\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl25024178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl25124178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl25224178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl25324178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl25424178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"0\.00_ ";
	text-align:general;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl25524178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#D8D8D8;
	mso-pattern:black none;
	white-space:nowrap;}
.xl25624178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"0\.0_ ";
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#D8D8D8;
	mso-pattern:black none;
	white-space:nowrap;}
.xl25724178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0_ ";
	text-align:general;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#D8D8D8;
	mso-pattern:black none;
	white-space:nowrap;}
.xl25824178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"0\.00_ ";
	text-align:general;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#D8D8D8;
	mso-pattern:black none;
	white-space:nowrap;}
.xl25924178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\0022\20A9\0022\#\,\#\#0\;\[Red\]\\-\0022\20A9\0022\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	border:.5pt solid gray;
	background:yellow;
	mso-pattern:black none;
	white-space:nowrap;}
.xl26024178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:Percent;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl26124178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"0\.0_ ";
	text-align:right;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl26224178
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0_ ";
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
	mso-font-charset:129;
	mso-char-type:none;}
	
	/* button */
.btnArea a,
.btnArea input,
.btnArea img {margin:0 8px;}
.btn {display:inline-block; text-align:center; font-weight:bold; vertical-align:middle; cursor:pointer; font-family:/*verdana, tahoma,*/ dotum, dotumche, '돋움', '돋움체', sans-serif;}
.btn:link, .btn:active, .btn:visited {color:#fff;}
.btn:hover {text-decoration:none;}
.btnB1 {font-size:12px; line-height:13px; padding:18px 45px;} 
.btnWhite {color:#d50c0c; background:#fff; border:1px solid #d50c0c;}
.btnW185 {width:183px; padding-left:0; padding-right:0;}

.tMar30 {margin-top:30px !important;}
.lMar10 {margin-left:10px;}
.ct {text-align:center !important;}	
-->
</style>
</head>

<body>
<!--[if !excel]>　　<![endif]-->
<!--다음 내용은 Microsoft Office Excel의 웹 페이지 마법사를 사용하여 작성되었습니다.-->
<!--같은 내용의 항목이 다시 게시되면 DIV 태그 사이에 있는 내용이 변경됩니다.-->
<!----------------------------->
<!--Excel의 웹 페이지 마법사로 게시해서 나온 결과의 시작 -->
<!----------------------------->

<div id="수출신고작성정보_24178" align=center x:publishsource="Excel">

<table border=0 cellpadding=0 cellspacing=0 width=357 class=xl25124178
 style='border-collapse:collapse;table-layout:fixed;width:268pt'>
 <col class=xl25124178 width=80 style='width:60pt'>
 <col class=xl25124178 width=90 style='mso-width-source:userset;mso-width-alt:
 2560;width:68pt'>
 <col class=xl25124178 width=83 style='mso-width-source:userset;mso-width-alt:
 2360;width:62pt'>
 <col class=xl25124178 width=104 style='mso-width-source:userset;mso-width-alt:
 2958;width:78pt'>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24224178 width=80 style='height:15.0pt;width:60pt'>기본정보</td>
  <td class=xl25124178 width=90 style='width:68pt'></td>
  <td class=xl25124178 width=83 style='width:62pt'></td>
  <td class=xl25124178 width=104 style='width:78pt'></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt'>운송방법</td>
  <td class=xl25024178 style='border-left:none'><%= ocoffinvoice.FOneItem.GetDeliverMethodName %></td>
  <td class=xl24324178 style='border-left:none'>운임부담</td>
  <td class=xl25024178 style='border-left:none'><%= ocoffinvoice.FOneItem.GetExportMethodName %></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt;border-top:none'>기준통화</td>
  <td class=xl25024178 style='border-top:none;border-left:none'><%= ocoffinvoice.FOneItem.Fpriceunit %></td>
  <td class=xl24324178 style='border-top:none;border-left:none'>정산방법</td>
  <td class=xl25024178 style='border-top:none;border-left:none'><%= ocoffinvoice.FOneItem.GetJungsanTypeName %></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt;border-top:none'>기준환율</td>
  <td class=xl26224178 style='border-top:none;border-left:none'><%= FormatNumber(ocoffinvoice.FOneItem.Fexchangerate, 0) %></td>
  <td class=xl25324178 style='border-top:none;border-left:none'>　</td>
  <td class=xl25024178 style='border-top:none;border-left:none'>　</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl25124178 style='height:15.0pt'></td>
  <td class=xl25124178></td>
  <td class=xl25124178></td>
  <td class=xl25124178></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt'><%= ocoffinvoice.FOneItem.GetExportMethodName %></td>
  <td class=xl24824178 style='border-left:none'><%= ocoffinvoice.FOneItem.Fpriceunitstring %><%= totalgoodspricecalc %></td>
  <td class=xl25924178 style='border-left:none'>&#8361;<%= totalgoodsprice %></td>
  <td class=xl25024178 style='border-left:none'>　</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt;border-top:none'>F.charge</td>
  <td class=xl24824178 style='border-top:none;border-left:none'><%= ocoffinvoice.FOneItem.Fpriceunitstring %><%= totalboxpricecalc %></td>
  <td class=xl24924178 style='border-top:none;border-left:none'>&#8361;<%= totalboxprice %></td>
  <td class=xl26024178 style='border-top:none;border-left:none'>　</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt;border-top:none'>결제금액</td>
  <td class=xl24824178 style='border-top:none;border-left:none'><%= ocoffinvoice.FOneItem.Fpriceunitstring %><%= totalpricecalc %></td>
  <td class=xl24924178 style='border-top:none;border-left:none'>&#8361;<%= totalprice %></td>
  <td class=xl25024178 style='border-top:none;border-left:none'>　</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl25124178 style='height:15.0pt'></td>
  <td class=xl23924178></td>
  <td class=xl24124178></td>
  <td class=xl24024178></td>
 </tr>
 <tr class=xl25224178 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt'>Q'ty / BOX</td>
  <td class=xl24624178 style='border-left:none'><%= ocoffinvoice.FOneItem.Ftotalboxno %></td>
  <td class=xl24324178 style='border-left:none'>G.weight(KG)</td>
  <td class=xl24724178 style='border-left:none'><%= FormatNumber(totalgweight, 2) %></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl25124178 style='height:15.0pt'></td>
  <td class=xl23924178></td>
  <td class=xl24124178></td>
  <td class=xl24024178></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24224178 colspan=2 style='height:15.0pt'>박스별 중량 및 EMS요금</td>
  <td class=xl25124178></td>
  <td class=xl25124178></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24324178 style='height:15.0pt'>BOX NO</td>
  <td class=xl24324178 style='border-left:none'>N.weight(KG)</td>
  <td class=xl24324178 style='border-left:none'>G.weight(KG)</td>
  <td class=xl24324178 style='border-left:none'>EMS요금</td>
 </tr>
 <% for i=0 to ocoffinvoicedetail.FResultCount-1 %>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
	<td height=20 class=xl24424178 style='height:15.0pt;border-top:none'><%= (i + 1) %></td>
	<td class=xl25424178 align=right style='border-top:none;border-left:none'><%= FormatNumber(ocoffinvoicedetail.FItemList(i).Fnweight, 2) %></td>
	<td class=xl26124178 style='border-top:none;border-left:none'><%= FormatNumber(ocoffinvoicedetail.FItemList(i).Fgweight, 2) %></td>
	<td class=xl24524178 align=right style='border-top:none;border-left:none'><%= FormatNumber(ocoffinvoicedetail.FItemList(i).FemsPrice, 0) %></td>
 </tr>
 <% next %>
<% if (ocoffinvoicedetail.FResultCount < 19) then %>
	 <% for i=ocoffinvoicedetail.FResultCount to 19 %>
	 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
	  <td height=20 class=xl24424178 style='height:15.0pt;border-top:none'><%= (i + 1) %></td>
	  <td class=xl25424178 style='border-top:none;border-left:none'>　</td>
	  <td class=xl26124178 style='border-top:none;border-left:none'>　</td>
	  <td class=xl24524178 style='border-top:none;border-left:none'>　</td>
	 </tr>
	  <% next %>
<% end if %>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl25524178 style='height:15.0pt;border-top:none'>합계</td>
  <td class=xl25824178 align=right style='border-top:none;border-left:none'><%= FormatNumber(totalnweight, 2) %></td>
  <td class=xl25624178 style='border-top:none;border-left:none'><%= FormatNumber(totalgweight, 2) %></td>
  <td class=xl25724178 align=right style='border-top:none;border-left:none'><%= FormatNumber(totalemsprice, 0) %></td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=80 style='width:60pt'></td>
  <td width=90 style='width:68pt'></td>
  <td width=83 style='width:62pt'></td>
  <td width=104 style='width:78pt'></td>
 </tr>
 <![endif]>
</table>

</div>


<!----------------------------->
<!--Excel의 웹 페이지 마법사로 게시해서 나온 결과의 끝-->
<!----------------------------->
<%if request("xl")="" then%>
	<div class="btnArea tMar30 ct">
		<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">인쇄하기</button>
		<!--<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('');">PDF 전환</button>-->
	</div>
<%end if%>
</body>

</html>

<%
set ocoffinvoice = Nothing
set ocoffinvoicedetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

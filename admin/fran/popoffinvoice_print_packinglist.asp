<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
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
Response.AddHeader "Content-Disposition", "attachment; filename=" + CStr(ocoffinvoice.FOneItem.Fshopname) + "_" + CStr(ocoffinvoice.FOneItem.Finvoiceno) + "_packinglist.xls"
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
dim avggoodsprice

if (ocoffinvoicedetail.FResultCount = "") then
	ocoffinvoicedetail.FResultCount = 0
end if

if (ocoffinvoice.FOneItem.Ftotalgoodsprice = "") then
	ocoffinvoice.FOneItem.Ftotalgoodsprice = 0
end if

if (ocoffinvoicedetail.FResultCount = 0) then
	avggoodsprice = 0
else
	avggoodsprice = ocoffinvoice.FOneItem.Ftotalgoodsprice / ocoffinvoicedetail.FResultCount
end if



'================================================================================
dim ocartoonboxdetail
set ocartoonboxdetail = new CCartoonBox
	ocartoonboxdetail.FRectMasterIdx = ocoffinvoice.FOneItem.Fworkidx
	ocartoonboxdetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid
	ocartoonboxdetail.GetDetailList



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
<link rel=File-List href="222.files/filelist.xml">
<style id="Stationery_Island_875_2212_20111013_packinglist 1 _1641_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.xl631641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl641641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl651641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl661641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl671641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl681641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl691641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl701641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl711641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl721641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl731641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl741641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl751641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0\.00_ ";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl761641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl771641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl781641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0\.00_ ";
	text-align:right;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl791641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:Fixed;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl801641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:20.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl811641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl821641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl831641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid black;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl841641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl851641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl861641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl871641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid black;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl881641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\@";
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl891641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\@";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl901641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl911641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid black;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl921641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\@";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl931641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl941641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl951641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl961641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl971641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl981641
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
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

<div id="Stationery_Island_875_2212_20111013_packinglist 1 _1641" align=center
x:publishsource="Excel"><!--다음 내용은 Microsoft Office Excel의 웹 페이지 마법사를 사용하여 작성되었습니다.--><!--같은 내용의 항목이 다시 게시되면 DIV 태그 사이에 있는 내용이 변경됩니다.--><!-----------------------------><!--Excel의 웹 페이지 마법사로 게시해서 나온 결과의 시작 --><!----------------------------->

<table border=0 cellpadding=0 cellspacing=0 width=690 style='border-collapse:
 collapse;table-layout:fixed;width:518pt'>
 <col class=xl631641 width=23 style='mso-width-source:userset;mso-width-alt:
 736;width:17pt'>
 <col class=xl631641 width=183 style='mso-width-source:userset;mso-width-alt:
 5856;width:137pt'>
 <col class=xl631641 width=23 style='mso-width-source:userset;mso-width-alt:
 736;width:17pt'>
 <col class=xl631641 width=153 style='mso-width-source:userset;mso-width-alt:
 4896;width:115pt'>
 <col class=xl631641 width=23 style='mso-width-source:userset;mso-width-alt:
 736;width:17pt'>
 <col class=xl631641 width=52 style='mso-width-source:userset;mso-width-alt:
 1664;width:39pt'>
 <col class=xl631641 width=50 style='mso-width-source:userset;mso-width-alt:
 1600;width:38pt'>
 <col class=xl631641 width=61 style='mso-width-source:userset;mso-width-alt:
 1952;width:46pt'>
 <col class=xl631641 width=31 style='mso-width-source:userset;mso-width-alt:
 992;width:23pt'>
 <col class=xl631641 width=62 style='mso-width-source:userset;mso-width-alt:
 1984;width:47pt'>
 <col class=xl631641 width=29 style='mso-width-source:userset;mso-width-alt:
 928;width:22pt'>
 <tr class=xl631641 height=46 style='mso-height-source:userset;height:35.1pt'>
  <td colspan=11 height=46 class=xl801641 width=690 style='height:35.1pt;
  width:518pt'>PACKING LIST</td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=4 height=24 class=xl811641 style='border-right:.5pt solid black;
  height:18.0pt'>1.Shipper/Exporter</td>
  <td colspan=7 class=xl831641 style='border-left:none'>8.No.&amp; date of
  invoice</td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td colspan=3 class=xl841641 rowspan=5 width=359 style='border-right:.5pt solid black; width:269pt'>
  	<%= nl2br(ocoffinvoice.FOneItem.Fexporteraddr) %>
  </td>
  <td class=xl651641></td>
  <td colspan=6 class=xl881641 width=285 style='width:215pt' rowspan=2>
  	<%= ocoffinvoice.FOneItem.Finvoiceno %><br>
  	<%= invoicedate %>
  </td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td colspan=7 class=xl911641 style='border-left:none'>9.Remarks :</td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl661641></td>
  <td colspan=6 rowspan=18 class=xl881641 width=285 style='border-bottom:1.0pt solid black;
  width:215pt'>
  	  <%= nl2br(ocoffinvoice.FOneItem.Fcomment) %><br />
      Freight Term : <%= ocoffinvoice.FOneItem.GetExportMethodName %><br />
      Terms of Payment : <%= ocoffinvoice.FOneItem.GetJungsanTypeName %>
  </td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=4 height=24 class=xl901641 style='border-right:.5pt solid black;
  height:18.0pt'>2.For account &amp; rist of Messers.</td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td colspan=3 rowspan=5 class=xl841641 width=359 style='border-right:.5pt solid black;
  width:269pt'>
  	<%= nl2br(ocoffinvoice.FOneItem.Friskmesseraddr) %>
  </td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'><span
  style='mso-spacerun:yes'>&nbsp;</span></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=4 height=24 class=xl901641 style='border-right:.5pt solid black;
  height:18.0pt'>3.Notify party</td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl651641 style='height:18.0pt'></td>
  <td colspan=3 rowspan=5 class=xl651641 style='border-right:.5pt solid black;
  border-bottom:.5pt solid black'>
  	<%= nl2br(ocoffinvoice.FOneItem.Fnotifyaddr) %>
  </td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl651641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl651641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl651641 style='height:18.0pt'></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl671641 style='height:18.0pt'>　</td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl901641 style='border-right:.5pt solid black;
  height:18.0pt'>4.Port of loading<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
  <td colspan=2 class=xl911641 style='border-right:.5pt solid black;border-left:
  none'>5.Final destination</td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl631641 style='height:18.0pt'></td>
  <td class=xl651641><%= nl2br(ocoffinvoice.FOneItem.Fportname) %></td>
  <td class=xl681641>　</td>
  <td class=xl631641><%= nl2br(ocoffinvoice.FOneItem.Fdestinationname) %></td>
  <td class=xl691641>　</td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl901641 style='border-right:.5pt solid black;
  height:18.0pt'>6.Carrier</td>
  <td colspan=2 class=xl911641 style='border-right:.5pt solid black;border-left:
  none'>7.Sailing on or about</td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl701641 style='height:18.0pt'>　</td>
  <td class=xl701641><%= nl2br(ocoffinvoice.FOneItem.Fcarriername) %></td>
  <td class=xl711641>　</td>
  <td class=xl721641><%= nl2br(carrierdate) %></td>
  <td class=xl731641>　</td>
 </tr>
 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl811641 style='border-right:.5pt solid black;
  height:18.0pt'>10.Marks &amp; number of pkgs</td>
  <td colspan=2 class=xl951641 style='border-right:.5pt solid black;border-left:
  none'>11.Description of Goods</td>
  <td colspan=3 class=xl971641 style='border-right:.5pt solid black;border-left:
  none'>12.Quantity/unit</td>
  <td colspan=2 class=xl971641 style='border-right:.5pt solid black;border-left:
  none'>13.N weight</td>
  <td colspan=2 class=xl971641 style='border-left:none'>14.G weight</td>
 </tr>

 <%
 dim totalnweight, totalgweight
 dim currcartoonboxno
 totalnweight = 0
 totalgweight = 0

 %>
 <% for i=0 to ocartoonboxdetail.FResultCount-1 %>
 <%
 if (ocartoonboxdetail.FItemList(i).Fcartoonboxno <> currcartoonboxno) then
     totalnweight = totalnweight + ocartoonboxdetail.FItemList(i).FcartoonboxNweight
	 totalgweight = totalgweight + ocartoonboxdetail.FItemList(i).Fcartoonboxweight
	 currcartoonboxno = ocartoonboxdetail.FItemList(i).Fcartoonboxno
 %>

 <tr class=xl631641 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl641641 style='height:18.0pt'><%= ocartoonboxdetail.FItemList(i).Fcartoonboxno %></td>
  <td class=xl631641>BOX</td>
  <td colspan=2 class=xl651641>Stationary & Gifts</td>
  <td class=xl651641></td>
  <td class=xl741641></td>
  <td class=xl651641></td>
  <td class=xl751641><%= FormatNumber(ocartoonboxdetail.FItemList(i).FcartoonboxNweight, 2) %></td>
  <td class=xl651641>Kgs</td>
  <td class=xl751641><%= FormatNumber(ocartoonboxdetail.FItemList(i).Fcartoonboxweight, 2) %></td>
  <td class=xl651641>Kgs</td>
 </tr>
<% end if %>
<% next %>

 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl641641 style='height:15.0pt'></td>
  <td class=xl631641></td>
  <td colspan=2 class=xl671641>　</td>
  <td class=xl651641></td>
  <td class=xl741641></td>
  <td class=xl651641></td>
  <td class=xl751641></td>
  <td class=xl651641></td>
  <td class=xl751641></td>
  <td class=xl651641></td>
 </tr>
 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl981641 style='height:15.0pt'>Total</td>
  <td class=xl761641 style='border-top:none'>　</td>
  <td class=xl761641 style='border-top:none'>　</td>
  <td class=xl761641>　</td>
  <td class=xl771641>　</td>
  <td class=xl761641>　</td>
  <td class=xl781641><%= FormatNumber(totalnweight, 2) %></td>
  <td class=xl761641>Kgs</td>
  <td class=xl781641><%= FormatNumber(totalgweight, 2) %></td>
  <td class=xl761641>Kgs</td>
 </tr>
 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl631641 style='height:15.0pt'></td>
  <td class=xl631641></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl791641></td>
 </tr>
 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl631641 style='height:15.0pt'></td>
  <td class=xl631641></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl791641></td>
 </tr>
 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl631641 style='height:15.0pt'></td>
  <td class=xl631641></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl791641></td>
 </tr>
 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl651641 style='height:15.0pt'></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
  <td class=xl631641></td>
 </tr>
 <tr class=xl631641 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl651641 style='height:15.0pt'></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td class=xl651641></td>
  <td colspan=2 class=xl701641>SIGNED BY</td>
  <td class=xl701641>　</td>
  <td class=xl701641>　</td>
  <td class=xl701641>　</td>
  <td class=xl701641>　</td>
  <td class=xl701641>　</td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=23 style='width:17pt'></td>
  <td width=183 style='width:137pt'></td>
  <td width=23 style='width:17pt'></td>
  <td width=153 style='width:115pt'></td>
  <td width=23 style='width:17pt'></td>
  <td width=52 style='width:39pt'></td>
  <td width=50 style='width:38pt'></td>
  <td width=61 style='width:46pt'></td>
  <td width=31 style='width:23pt'></td>
  <td width=62 style='width:47pt'></td>
  <td width=29 style='width:22pt'></td>
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

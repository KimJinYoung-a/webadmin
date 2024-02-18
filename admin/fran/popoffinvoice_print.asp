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
Response.AddHeader "Content-Disposition", "attachment; filename=" + CStr(ocoffinvoice.FOneItem.Fshopname) + "_" + CStr(ocoffinvoice.FOneItem.Finvoiceno) + "_invoice.xls"
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
dim ocartoonboxdetail
set ocartoonboxdetail = new CCartoonBox
	ocartoonboxdetail.FRectMasterIdx = ocoffinvoice.FOneItem.Fworkidx
	ocartoonboxdetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid
	ocartoonboxdetail.GetDetailList

dim totalCartonBoxCount : totalCartonBoxCount = 0
dim currCartonBox

for i = 0 to ocartoonboxdetail.FResultCount - 1
    if (currCartonBox <> ocartoonboxdetail.FItemList(i).Fcartoonboxno) then
        currCartonBox = ocartoonboxdetail.FItemList(i).Fcartoonboxno
        totalCartonBoxCount = totalCartonBoxCount + 1
    end if
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
<link rel=File-List href="2.files/filelist.xml">
<style id="&#4363;&#4457;&#4369;&#4467;&#4357;&#4449;&#4363;&#4469;&#4523; &#4370;&#4450;&#4363;&#4460;&#4366;&#4462;&#4527;&#4352;&#4457;&#4352;&#4469;&#4370;&#4460;&#4520;_3115_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font53115
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
	mso-font-charset:0;}
.font63115
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:돋움, monospace;
	mso-font-charset:0;}
.xl2393115
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
.xl2403115
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
.xl2413115
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
.xl2423115
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
.xl2433115
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
.xl2443115
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
.xl2453115
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
.xl2463115
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
.xl2473115
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
.xl2483115
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
.xl2493115
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
	mso-number-format:\@;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2503115
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
.xl2513115
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
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
.xl2523115
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
	mso-number-format:"_-\0022US$\0022* \#\,\#\#0\.00_ \;_-\0022US$\0022* \\-\#\,\#\#0\.00\\ \;_-\0022US$\0022* \0022-\0022??_ \;_-\@_ ";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2533115
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
	mso-number-format:"\0022\20A9\0022\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2543115
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2553115
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
	mso-number-format:"_-\0022\20A9\0022* \#\,\#\#0_-\;\\-\0022\20A9\0022* \#\,\#\#0_-\;_-\0022\20A9\0022* \0022-\0022_-\;_-\@_-";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2563115
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
	mso-number-format:"\[$\00A5-411\]\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2573115
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
	mso-number-format:"\[$\00A5-411\]\#\,\#\#0";
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2583115
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
.xl2593115
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
.xl2603115
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
	mso-number-format:"_-\0022US$\0022* \#\,\#\#0\.00_ \;_-\0022US$\0022* \\-\#\,\#\#0\.00\\ \;_-\0022US$\0022* \0022-\0022??_ \;_-\@_ ";
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2613115
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Malgun Gothic", monospace;
	mso-font-charset:129;
	mso-number-format:"\#\,\#\#0";
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2623115
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
	mso-number-format:"\#\,\#\#0";
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2633115
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
.xl2643115
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
.xl2653115
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
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2663115
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
.xl2673115
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
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2683115
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
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2693115
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
.xl2703115
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
.xl2713115
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
	mso-number-format:"\[ENG\]dd\\ mmm\\\,\\ yyyy";
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2723115
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
.xl2733115
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
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2743115
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
.xl2753115
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
	mso-number-format:"\[ENG\]dd\\ mmm\\\,\\ yyyy";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl2763115
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
.xl2773115
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
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
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
	mso-font-charset:0;
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

<div id="&#4363;&#4457;&#4369;&#4467;&#4357;&#4449;&#4363;&#4469;&#4523; &#4370;&#4450;&#4363;&#4460;&#4366;&#4462;&#4527;&#4352;&#4457;&#4352;&#4469;&#4370;&#4460;&#4520;_3115" align=center x:publishsource="Excel">

<table border=0 cellpadding=0 cellspacing=0 width=693 class=xl2393115 style='border-collapse:collapse;table-layout:fixed;width:522pt'>
 <col class=xl2393115 width=26 style='mso-width-source:userset;mso-width-alt: 739;width:20pt'>
 <col class=xl2393115 width=183 style='mso-width-source:userset;mso-width-alt: 5205;width:137pt'>
 <col class=xl2393115 width=26 style='mso-width-source:userset;mso-width-alt: 739;width:20pt'>
 <col class=xl2393115 width=183 style='mso-width-source:userset;mso-width-alt: 5205;width:137pt'>
 <col class=xl2393115 width=26 style='mso-width-source:userset;mso-width-alt: 739;width:20pt'>
 <col class=xl2393115 width=109 style='mso-width-source:userset;mso-width-alt: 3100;width:82pt'>
 <col class=xl2393115 width=26 style='mso-width-source:userset;mso-width-alt: 739;width:20pt'>
 <col class=xl2393115 width=114 style='mso-width-source:userset;mso-width-alt: 3242;width:86pt'>
 <tr height=46 style='mso-height-source:userset;height:35.1pt'>
  <td colspan=8 height=46 class=xl2703115 width=693 style='height:35.1pt; width:522pt'>COMMERCIAL INVOICE</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=4 height=24 class=xl2643115 style='border-right:.5pt solid black; height:18.0pt'>Shipper/Exporter</td>
  <td colspan=4 class=xl2633115 style='border-left:none'>No.&amp; date of invoice</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td colspan=3 rowspan=5 class=xl2743115 width=392 style='border-right:.5pt solid black; border-bottom:.5pt solid black;width:294pt'>
    <%= nl2br(ocoffinvoice.FOneItem.Fexporteraddr) %>
  </td>
  <td class=xl2413115 style='border-left:none'>　</td>
  <td colspan=3 rowspan=2 class=xl2763115 width=249 style='border-bottom:.5pt solid black; width:188pt'>
  	<%= ocoffinvoice.FOneItem.Finvoiceno %><br>
  	<%= invoicedate %>
  </td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td colspan=4 class=xl2653115 style='border-left:none'>No.&amp; date of L/C</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td class=xl2773115 style='border-left:none'>　</td>
  <td colspan=3 rowspan=4 class=xl2403115 style='border-bottom:.5pt solid black'>
  	<%= ocoffinvoice.FOneItem.Flccomment %>
  </td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=4 height=24 class=xl2663115 style='height:18.0pt'>For account &amp; Risk of Messers.</td>
  <td class=xl2413115>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td colspan=3 rowspan=5 class=xl2743115 width=392 style='border-right:.5pt solid black; border-bottom:.5pt solid black;width:294pt'>
  	<%= nl2br(ocoffinvoice.FOneItem.Friskmesseraddr) %>
  </td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td colspan=4 class=xl2653115 style='border-left:none'>L/C issuing bank</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'>　</td>
  <td colspan=3 rowspan=3 class=xl2403115 style='border-bottom:.5pt solid black'>
    <%= ocoffinvoice.FOneItem.Flcbank %>
  </td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'><span style='mso-spacerun:yes'>&nbsp;</span></td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=4 height=24 class=xl2663115 style='height:18.0pt'>Notify party</td>
  <td colspan=4 class=xl2653115>Remarks :</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2403115 style='height:18.0pt'></td>
  <td colspan=3 rowspan=5 class=xl2403115 style='border-right:.5pt solid black; border-bottom:.5pt solid black'>
    <%= nl2br(ocoffinvoice.FOneItem.Fnotifyaddr) %>
  </td>
  <td class=xl2413115 style='border-left:none'>　</td>
  <td colspan=3 rowspan=9 class=xl2403115 style='border-bottom:1.0pt solid black'>
      <%= nl2br(ocoffinvoice.FOneItem.Fcomment) %><br />
      Freight Term : <%= ocoffinvoice.FOneItem.GetExportMethodName %><br />
      Terms of Payment : <%= ocoffinvoice.FOneItem.GetJungsanTypeName %>
  </td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2403115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'></td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2403115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2403115 style='height:18.0pt'></td>
  <td class=xl2413115 style='border-left:none'></td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2443115 style='height:18.0pt'>　</td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl2663115 style='border-right:.5pt solid black; height:18.0pt'>Port of loading<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span></td>
  <td colspan=2 class=xl2653115 style='border-right:.5pt solid black; border-left:none'>Final destination</td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2393115 style='height:18.0pt'></td>
  <td class=xl2403115>
    <%= nl2br(ocoffinvoice.FOneItem.Fportname) %>
  </td>
  <td class=xl2453115>　</td>
  <td class=xl2393115>
    <%= nl2br(ocoffinvoice.FOneItem.Fdestinationname) %>
  </td>
  <td class=xl2413115>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl2663115 style='border-right:.5pt solid black; height:18.0pt'>Carrier</td>
  <td colspan=2 class=xl2653115 style='border-right:.5pt solid black; border-left:none'>Sailing on or about</td>
  <td class=xl2413115 style='border-left:none'>　</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2473115 style='height:18.0pt'>　</td>
  <td class=xl2473115><%= nl2br(ocoffinvoice.FOneItem.Fcarriername) %></td>
  <td class=xl2483115></td>
  <td class=xl2493115>
    <%= nl2br(carrierdate) %>
  </td>
  <td class=xl2503115></td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl2693115 style='height:18.0pt'>Description of Goods</td>
  <td colspan=2 class=xl2633115>Q'ty / Carton BOX</td>
  <td colspan=2 class=xl2633115>Price / BOX</td>
  <td colspan=2 class=xl2633115>Amount</td>
 </tr>

<% for i=0 to ocoffinvoiceproductdetail.FResultCount-1 %>
	<%
	if (ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox <> "") then
		ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox = FormatNumber(ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox, 2)
	end if
	if (ocoffinvoiceproductdetail.FItemList(i).Ftotalprice <> "") then
		ocoffinvoiceproductdetail.FItemList(i).Ftotalprice = FormatNumber(ocoffinvoiceproductdetail.FItemList(i).Ftotalprice, 2)
	end if
	%>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl2403115 style='height:18.0pt'></td>
  <td class=xl2513115>
    <%= ocoffinvoiceproductdetail.FItemList(i).Fgoodscomment %>
  </td>
  <td class=xl2393115></td>
  <td class=xl2613115>
      <!--
      <%= ocoffinvoiceproductdetail.FItemList(i).Ftotalboxno %>
      -->
      <% if i = 0 then %>
      <%= totalCartonBoxCount %>
      <% end if %>
  </td>
  <td class=xl2403115></td>
  <td class=xl2523115>
  	<% if (ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox <> "") then %>
  	<span style='mso-spacerun:yes'>&nbsp;</span><%= ocoffinvoice.FOneItem.Fpriceunitstring %><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span><%= ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox %>
	<% end if %>
  </td>
  <td class=xl2533115></td>
  <td class=xl2523115>
  	<% if (ocoffinvoiceproductdetail.FItemList(i).Ftotalprice <> "") then %>
  	<span style='mso-spacerun:yes'>&nbsp;</span><%= ocoffinvoice.FOneItem.Fpriceunitstring %><span style='mso-spacerun:yes'>&nbsp;&nbsp; </span><%= ocoffinvoiceproductdetail.FItemList(i).Ftotalprice %>
  	<% end if %>
  </td>
 </tr>
<% next %>

 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2403115 style='height:15.0pt'></td>
  <td class=xl2513115></td>
  <td class=xl2393115></td>
  <td class=xl2543115></td>
  <td class=xl2393115></td>
  <td class=xl2573115></td>
  <td class=xl2573115></td>
  <td class=xl2553115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2403115 style='height:15.0pt'></td>
  <td class=xl2513115></td>
  <td class=xl2393115></td>
  <td class=xl2543115></td>
  <td class=xl2463115></td>
  <td class=xl2563115></td>
  <td class=xl2563115></td>
  <td class=xl2553115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2403115></td>
  <td class=xl2403115></td>
  <td class=xl2463115></td>
  <td class=xl2463115></td>
  <td class=xl2463115></td>
  <td class=xl2463115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2403115></td>
  <td class=xl2403115></td>
  <td class=xl2463115></td>
  <td class=xl2463115></td>
  <td class=xl2463115></td>
  <td class=xl2463115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2403115></td>
  <td class=xl2403115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2433115 colspan=2 style='height:15.0pt'>Total</td>
  <td class=xl2423115>　</td>
  <td class=xl2623115><%= totalCartonBoxCount %></td>
  <td class=xl2423115 colspan=2>(<%= ocoffinvoice.FOneItem.GetExportMethodName %>)</td>
  <td class=xl2423115>　</td>
  <td class=xl2603115><span style='mso-spacerun:yes'>&nbsp;</span><%= ocoffinvoice.FOneItem.Fpriceunitstring %><span style='mso-spacerun:yes'>&nbsp;&nbsp; </span><%= FormatNumber(ocoffinvoice.FOneItem.Ftotalprice, 2) %></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2403115 style='height:15.0pt'></td>
  <td class=xl2403115></td>
  <td class=xl2403115></td>
  <td class=xl2403115></td>
  <td class=xl2473115 colspan=2>SIGNED BY</td>
  <td class=xl2473115>　</td>
  <td class=xl2473115>　</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl2393115 style='height:15.0pt'></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
  <td class=xl2393115></td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=26 style='width:20pt'></td>
  <td width=183 style='width:137pt'></td>
  <td width=26 style='width:20pt'></td>
  <td width=183 style='width:137pt'></td>
  <td width=26 style='width:20pt'></td>
  <td width=109 style='width:82pt'></td>
  <td width=26 style='width:20pt'></td>
  <td width=114 style='width:86pt'></td>
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

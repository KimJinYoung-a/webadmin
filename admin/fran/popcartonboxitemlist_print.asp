<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 샵별패킹리스트(박스별)
' History : 2012.02.02 이상구 생성
'			2012.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim jungsanidx, shopid, shopname ,i, j
	jungsanidx = request("jungsanidx")
	shopid = request("shopid")
	shopname = request("shopname")

if jungsanidx="" then jungsanidx=0
if shopname="" then shopname=shopid

if (C_IS_SHOP = true) then
	shopid = C_STREETSHOPID
end if

dim ocartoonboxmaster
set ocartoonboxmaster = new CCartoonBox
	ocartoonboxmaster.FRectShopid = shopid
	ocartoonboxmaster.FRectJungsanIdx = jungsanidx
	ocartoonboxmaster.GetJungsanItemList

if request("xl")<>"" then
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=" + CStr(shopname) + "_" + CStr(jungsanidx) + "_itemlist.xls"
end if
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<link rel=File-List href="&#4370;&#4457;&#4364;&#4462;Lemnis_2248.files/filelist.xml">
<style id="&#4370;&#4457;&#4364;&#4462;Lemnis_2248_2366_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font52366
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;}
.xl152366
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
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
.xl822366
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"Short Date";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl832366
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
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl842366
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
.xl852366
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
	white-space:normal;}
.xl862366
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
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl872366
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
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#FFCC00;
	mso-pattern:black none;
	white-space:nowrap;}
.xl882366
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
	background:#FFCC00;
	mso-pattern:black none;
	white-space:normal;}
.xl892366
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
	background:#FFCC00;
	mso-pattern:black none;
	white-space:normal;}
.xl902366
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
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border:.5pt solid gray;
	background:#FFCC00;
	mso-pattern:black none;
	white-space:normal;}
.xl912366
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
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
	font-family:"맑은 고딕", monospace;
	mso-font-charset:129;
	mso-char-type:none;}
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

<div id="&#4370;&#4457;&#4364;&#4462;Lemnis_2248_2366" align=center x:publishsource="Excel">

<table border=0 cellpadding=0 cellspacing=0 width=1513 style='border-collapse: collapse;table-layout:fixed;width:1136pt'>
 <tr height=48 style='height:36.0pt'>
  <td height=48 class=xl872366 width=89 style='height:36.0pt;width:67pt'><%= CTX_Order_code %></td>
  <td class=xl902366 width=89 style='border-left:none;width:67pt'><%= CTX_release_code %></td>
  <td class=xl872366 width=89 style='height:36.0pt;width:67pt'><%= CTX_Order_Date %></td>
  <td class=xl902366 width=89 style='border-left:none;width:67pt'><%= CTX_Real_Order_Date %><br>(<%= CTX_workday %>)</td>
  <td class=xl882366 width=68 style='border-left:none;width:51pt'><%= CTX_INNERBOX %></td>
  <td class=xl892366 width=72 style='border-left:none;width:54pt'><%= CTX_CARTONBOX %></td>
  <td class=xl862366 width=106 style='border-left:none;width:80pt'><%= CTX_Brand %></td>
  <td class=xl832366 width=28 style='border-left:none;width:21pt'><%= CTX_divide %></td>
  <td class=xl832366 width=60 style='border-left:none;width:45pt'><%= CTX_Item_Code %></td>
  <td class=xl832366 width=60 style='border-left:none;width:45pt'><%= CTX_Description_Option %></td>
  <td class=xl832366 width=146 style='border-left:none;width:110pt'><%= CTX_Barcode %></td>
  <td class=xl832366 width=316 style='border-left:none;width:237pt'><%= CTX_Description %></td>
  <td class=xl832366 width=184 style='border-left:none;width:138pt'><%= CTX_Description_Option_name %></td>
  <td class=xl842366 width=28 style='border-left:none;width:21pt'><%= CTX_quantity %></td>
  <td class=xl852366 width=60 style='border-left:none;width:45pt'><%= CTX_selling_price %></td>
  <td class=xl842366 width=60 style='border-left:none;width:45pt'><%= CTX_Supply_price %></td>
  <td class=xl852366 width=79 style='border-left:none;width:59pt'><%= CTX_Supply %>&nbsp;<%= CTX_margin %></td>
  <td class=xl852366 width=68 style='border-left:none;width:51pt'><%= CTX_total_Supply_price %></td>
 </tr>
<% for i=0 to ocartoonboxmaster.FResultCount-1 %>
 <tr height=21 style='mso-height-source:userset;height:15.75pt'>
  <td height=21 class=xl822366 align=right style='height:15.75pt'><%= ocartoonboxmaster.FItemList(i).Fbaljucode %></td>
  <td class=xl822366 align=right><%= ocartoonboxmaster.FItemList(i).Fchulgocode %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fjumundate %></td>
  <td class=xl152366><%= ocartoonboxmaster.FItemList(i).Fbaljudate %></td>
  <td class=xl152366><%= ocartoonboxmaster.FItemList(i).Finnerboxno %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fcartoonboxno %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fmakerid %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fitemgubun %></td>
  <td class=xl152366>&nbsp;<%= ocartoonboxmaster.FItemList(i).Fitemid %></td>
  <td class=xl152366><%= ocartoonboxmaster.FItemList(i).Fitemoption %></td>
  <td class=xl152366>&nbsp;<%= ocartoonboxmaster.FItemList(i).Fbarcode %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fitemname %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fitemoptionname %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Frealitemno %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fsellcash %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Fsuplycash %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Foffmargin %></td>
  <td class=xl152366 align=right><%= ocartoonboxmaster.FItemList(i).Ftotsuplycash %></td>
 </tr>
<% next %>

</table>
</div>

<!----------------------------->
<!--Excel의 웹 페이지 마법사로 게시해서 나온 결과의 끝-->
<!----------------------------->
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->

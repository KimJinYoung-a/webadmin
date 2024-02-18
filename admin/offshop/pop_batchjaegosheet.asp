<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 재고
' History : 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<%
dim shopid, idx
	shopid = requestCheckVar(request("shopid"),32)
	idx = requestCheckVar(request("idx"),10)

dim oshoporder
set oshoporder = new CShopOrder
oshoporder.FRectShopID = shopid
oshoporder.FRectIdx = idx
oshoporder.FPageSize = 2000
oshoporder.GetShopOrderDetail

dim i

%>
<%
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + CStr(idx) + "-" + Left(CStr(now()),10) + ".xls"
%>
<html xmlns:x="urn:schemas-microsoft-com:office:excel">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<style>
  .big_title
    {
    mso-style-parent:style0;
	white-space:normal;
    font-size:18.0pt;
    font-weight:700;
    }
  .mid_title
    {
    mso-style-parent:style0;
	white-space:normal;
    font-size:12.0pt;
    font-weight:700;
    }
  .title_center
	{mso-style-parent:style0;
	text-align:center;
	vertical-align:middle;
	white-space:normal;}
  .normal
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	white-space:normal;
	font-size:8.0pt;
	}
  .normal_b
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	white-space:normal;
	font-size:8.0pt;
	font-weight:700;
	}
  .currency
	{mso-style-parent:style0;
 	mso-number-format:"\#\,\#\#0\.00_\)\;\[Red\]\\\(\#\,\#\#0\.00\\\)";
	border:0.5pt solid black;
	white-space:normal;}
   .Format_Y1
	{mso-style-parent:style0;
	mso-number-format:"yyyy\0022\/\0022m\0022\/\0022d\;\@";
 	white-space:normal;}
   .Format_Y2
	{mso-style-parent:style0;
	mso-number-format:"yyyy\/mm\;\@";
	text-align:center;
	border:0.5pt solid black;
 	white-space:normal;}
   .Format_number
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	mso-number-format:"\#\,\#\#0";
	white-space:normal;
	font-size:8.0pt;
	}
   .Format_number_L
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	mso-number-format:"\#\,\#\#0";
	white-space:normal;
	font-size:12.0pt;
	}
  .Format_T1
	{mso-style-parent:style0;
	mso-number-format:"hh\:mm\:ss\;\@";
	text-align:center;
 	white-space:normal;}  </style>
</head>

<body leftmargin="10">
<table width=700 cellspacing=0 cellpadding=1 border=0>
 <tr height=17 align=center >
      <td width="40" height=17 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>NO</td>
      <td width="60" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>브랜드ID</td>
      <td width="40" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>구분</td>
      <td width="60" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상품ID</td>
      <td width="120" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상품명</td>
      <td width="100" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>옵션</td>
      <td width="50" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>판매가</td>
      <td width="50" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>매입가</td>
      <td width="50" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>갯수</td>
 </tr>
<% for i=0 to oshoporder.FResultCount-1 %>
 <tr height=17 align=center >
      <td width="40" height=17 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= (i + 1) %></td>
      <td width="60" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oshoporder.FItemList(i).Fmakerid %></td>
      <td width="40" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oshoporder.FItemList(i).Fitemgubun %></td>
      <td width="60" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oshoporder.FItemList(i).Fitemid %></td>
      <td width="120" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oshoporder.FItemList(i).Fitemname %></td>
      <td width="100" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oshoporder.FItemList(i).Fitemoptionname %></td>
      <td width="50" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= FormatNumber(oshoporder.FItemList(i).Frealsellprice,0) %></td>
      <td width="50" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= FormatNumber(oshoporder.FItemList(i).Fsuplyprice,0) %></td>
      <td width="50" class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oshoporder.FItemList(i).Fitemno %></td>
 </tr>
<% next %>
</table>
</body>
</html>

<%
set oshoporder = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

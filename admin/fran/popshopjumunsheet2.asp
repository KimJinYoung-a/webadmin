<%@ language=vbscript %>
<% option explicit %>
<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + request("idx") + ".xls"
end if
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim webImgUrl
IF application("Svr_Info")="Dev" THEN
  webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지
else
  webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지
end if

dim idx,itype
idx = request("idx")
itype = request("itype")


dim oordersheetmaster, oordersheet
set oordersheetmaster = new COrderSheet
oordersheetmaster.FRectIdx = idx
oordersheetmaster.GetOneOrderSheetMaster

dim isFixed
isFixed = oordersheetmaster.FOneItem.IsFixed


set oordersheet = new COrderSheet
oordersheet.FrectisFixed = isFixed
oordersheet.FRectIdx = idx
oordersheet.GetOrderSheetDetail


dim obrand
set obrand = new CBrandShopInfoItem

obrand.FRectChargeId = oordersheetmaster.FOneItem.Ftargetid
obrand.GetBrandShopInFo


dim i

dim scheduleorexedate
	if IsNULL(oordersheetmaster.FOneItem.FScheduleDate) then
		scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FRegDate),10),"-","/")
	else
		scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FScheduleDate),10),"-","/")
	end if

dim ttlsuplycash, ttlsellcash,  ttlcount
ttlsuplycash=0
ttlsellcash=0
ttlcount=0
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
 <tr height=35 style='height:26.25pt '>
  <td colspan=5 height=35 class=big_title align=left style='border-bottom:0.5pt solid black;'>주문내역서(<%= obrand.FChargeName %>)</td>
  <td width=240 colspan=3 align=right class=mid_title style='border-bottom:0.5pt solid black;' ><%= oordersheetmaster.FOneItem.Fbaljuname %> (<%= oordersheetmaster.FOneItem.FBaljuCode %>)</td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal ></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >날 짜</td>
  <td rowspan=2 colspan=2 class=Format_Y1 align=left ><b><%= scheduleorexedate %></b></td>
  <td rowspan=6></td>
  <td width=74 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>등 록 번 호</td>
  <td colspan=3 class=normal style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FSocNo %>&nbsp;</td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상 호</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FSocName %></td>
  <td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>성  명</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FCeoName %></td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >상 호</td>
  <td rowspan=2 colspan=2 class=mid_title><%= oordersheetmaster.FOneItem.Fbaljuname %></td>
  <td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>사업장소재지</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FAddress %></td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>업 태</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FUptae %></td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >총금액</td>
  <td rowspan=2 colspan=2 align=left class=Format_number_L x:num="<%= oordersheetmaster.FOneItem.FTotalSuplycash %>" ><b>\<%= ForMatNumber(oordersheetmaster.FOneItem.FTotalSuplycash,0) %></b></td>
  <td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>종 목</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FUpjong %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>담 당 자</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FManagerName %>&nbsp;</td>
  <td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>연 락 처</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FManagerHp %>&nbsp;</td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal style='height:12.0pt'></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal style='height:12.0pt'></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
 </tr>
 <tr height=17 align=center >
  <td height=17 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상품코드</td>
  <td width=76 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>브랜드</td>
  <td width=160 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상품명</td>
  <td width=76 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>옵션</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>소비자가</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>공급가</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>수량</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>합계</td>
 </tr>
 <% for i=0 to oordersheet.FResultCount -1 %>
 <%
 	ttlsellcash = ttlsellcash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FSellcash
 	ttlsuplycash = ttlsuplycash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FSuplycash
 	ttlcount = ttlcount + oordersheet.FItemList(i).Frealitemno
 %>
 <tr height=17 style='mso-height-source:userset;height:20pt'>
  <td height=17 class=normal width=97 style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
  <%= oordersheet.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %>-<%= oordersheet.FItemList(i).FItemOption %></td>
  <td class=normal width=76 style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oordersheet.FItemList(i).FMakerid %>&nbsp;</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
  	<%= oordersheet.FItemList(i).FItemName %>
  </td>
  <td class=normal width=76 style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oordersheet.FItemList(i).FItemOptionName %>&nbsp;</td>
  <td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FSellcash %>"><%= FormatNumber(oordersheet.FItemList(i).FSellcash,0) %></td>
  <td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FSuplycash %>"><%= FormatNumber(oordersheet.FItemList(i).FSuplycash,0) %></td>
  <td align=center class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FRealItemno %>" ><%= oordersheet.FItemList(i).FRealItemno %></td>
  <td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FRealItemno*oordersheet.FItemList(i).FSuplycash %>" ><%= FormatNumber(oordersheet.FItemList(i).FRealItemno*oordersheet.FItemList(i).FSuplycash,0) %></td>
 </tr>
 <% next %>
 <tr height=22 >
  <td align=center height=22 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>비고</td>
  <td colspan=5 class=normal  style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= nl2br(oordersheetmaster.FoneItem.FComment) %>&nbsp;</td>
  <td align=center class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= ttlcount %>" ><%= ttlcount %></td>
  <td class=Format_number align=right style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= ttlsuplycash %>"><b>\<%= ForMatNumber(ttlsuplycash,0) %></b></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal style='height:12.0pt'></td>
  <td class=normal></td>
  <td colspan=3 class=normal>　</td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
 </tr>
 <tr height=30 style='mso-height-source:userset;height:22.5pt'>
  <td align=center height=30 class=normal_b style='border-top:0.5pt solid black; border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>인계자</td>
  <td colspan=3 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>(인)&nbsp;</td>
  <td align=center colspan=2 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>인수자</td>
  <td colspan=2 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>(인)&nbsp;</td>
 </tr>
</table>
</body>
</html>

<%
set obrand = Nothing
set oordersheetmaster = Nothing
set oordersheet = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

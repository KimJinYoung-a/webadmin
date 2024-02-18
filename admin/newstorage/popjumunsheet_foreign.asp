<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주문내역서 영문
' History : 2017.06.12 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%

dim webImgUrl
IF application("Svr_Info")="Dev" THEN
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지
else
	webImgUrl		= "/webimage"									'웹이미지
end if

dim idx, itype, isFixed, i
	idx = requestCheckVar(getNumeric(request("idx")),10)
	itype = requestCheckVar(request("itype"),10)

dim oordersheetmaster, oordersheet
set oordersheetmaster = new COrderSheet
	oordersheetmaster.FRectIdx = idx
	oordersheetmaster.GetOneOrderSheetMaster

isFixed = oordersheetmaster.FOneItem.IsFixed

set oordersheet = new COrderSheet
	oordersheet.FrectisFixed = isFixed
	oordersheet.FRectIdx = idx
	oordersheet.GetOrderSheetDetail

dim obrand
set obrand = new CBrandShopInfoItem
	obrand.FRectChargeId = oordersheetmaster.FOneItem.Ftargetid
	obrand.GetBrandShopInFo

dim scheduleorexedate
	scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FScheduleDate),10),"-","/")

dim ttlsellcash, ttlbuycash, ttlcount
ttlsellcash = 0
ttlbuycash  = 0
ttlcount    = 0

if request("xl")<>"" then
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=" + oordersheetmaster.FOneItem.Ftargetid + Left(CStr(now()),10) + ".xls"
end if

dim tenbytenyn
	tenbytenyn = "N"
if obrand.FSocNo="211-87-00620" then
	tenbytenyn = "Y"
end if
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
<tr >
	<td colspan=4 class=big_title align=left style='border-bottom:0.5pt solid black;'>
		<% if tenbytenyn="Y" then %>
			ORDER SHEET(TENBYTEN)
		<% else %>
			ORDER SHEET(<%= obrand.FChargeName %>)
		<% end if %>
	</td>
	<td width=240 colspan=3 align=right class=mid_title style='border-bottom:0.5pt solid black;'>TENBYTEN (<%= oordersheetmaster.FOneItem.FBaljuCode %>)</td>
</tr>
<tr height=16 >
	<td height=16 class=normal ></td>
	<td class=normal></td>
	<td class=normal></td>
	<td class=normal></td>
	<td class=normal></td>
	<td class=normal></td>
	<td class=normal></td>
</tr>
<tr>
	<td rowspan=2 class=mid_title >Date</td>
	<td rowspan=2 class=Format_Y1 align=left ><b><%= scheduleorexedate %></b></td>
	<td rowspan=6></td>
	<td width=74 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>VAT Reg.No</td>
	<td colspan=3 class=normal style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FSocNo %></td>
</tr>
<tr>
	<td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Name of Company</td>
	<td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<% if tenbytenyn="Y" then %>
			TENBYTEN
		<% else %>
			<%= obrand.FSocName %>
		<% end if %>
	</td>
	<td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Name of Representative</td>
	<td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<% if tenbytenyn="Y" then %>
			ChoiEunHee
		<% else %>
			<%= obrand.FCeoName %>
		<% end if %>
	</td>
</tr>
<tr>
	<td rowspan=2 class=mid_title >Name of Company</td>
	<td rowspan=2 class=mid_title>
		<% if tenbytenyn="Y" then %>
			TENBYTEN
		<% else %>
			<%= oordersheetmaster.FOneItem.Fbaljuname %>
		<% end if %>
	</td>
	<td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Business Location</td>
	<td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<% if tenbytenyn="Y" then %>
			14F(GyoYukDong) 57, Daehak-ro, Jongno-gu Seoul, Korea [03082]
		<% else %>
			<%= obrand.FAddress %>
		<% end if %>
	</td>
</tr>
<tr>
	<td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Type of Business</td>
	<td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<% if tenbytenyn="Y" then %>
			Service,Wholesale,Sleeve
		<% else %>
			<%= obrand.FUptae %>
		<% end if %>
	</td>
</tr>
<tr>
	<td rowspan=2 class=mid_title >Total Amount</td>
	<td rowspan=2 align=left class=Format_number_L x:num="<%= oordersheetmaster.FOneItem.FTotalBuycash %>" ><b>\<%= ForMatNumber(oordersheetmaster.FOneItem.FTotalBuycash,0) %></b></td>
	<td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Items of Business</td>
	<td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<% if tenbytenyn="Y" then %>
			e-commerce
		<% else %>
			<%= obrand.FUpjong %>
		<% end if %>
	</td>
</tr>
<tr style='mso-height-source:userset;'>
	<td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Attn</td>
	<td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FManagerName %></td>
	<td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Tel</td>
	<td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FManagerHp %></td>
</tr>
<tr height=16 style='mso-height-source:userset;height:12.0pt'>
	<td height=16 class=normal style='height:12.0pt'></td>
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
</tr>
<tr align=center >
	<td class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Item Code</td>
	<td width=160 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Item Name</td>
	<td width=76 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Item Option</td>
	<td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Consumer Price</td>
	<td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Supply price </td>
	<td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Quantity</td>
	<td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Amount</td>
</tr>
<% for i=0 to oordersheet.FResultCount -1 %>
<%
ttlsellcash = ttlsellcash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FSellcash
ttlbuycash = ttlbuycash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FBuycash
ttlcount = ttlcount + oordersheet.FItemList(i).Frealitemno
%>
<tr style='mso-height-source:userset;'>
	<td class=normal width=105 style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<%= oordersheet.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %>-<%= oordersheet.FItemList(i).FItemOption %>
		<% if (oordersheet.FItemList(i).FUpcheManageCode <> "") then %>
		<br><%= oordersheet.FItemList(i).FUpcheManageCode %>
		<% end if %>
	</td>
	<td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
		<%= oordersheet.FItemList(i).FItemName %>
	</td>
	<td class=normal width=76 style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oordersheet.FItemList(i).FItemOptionName %>&nbsp;</td>
	<td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FSellcash %>"><%= FormatNumber(oordersheet.FItemList(i).FSellcash,0) %></td>
	<td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FBuycash %>"><%= FormatNumber(oordersheet.FItemList(i).FBuycash,0) %></td>
	<td align=center class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FRealItemno %>" ><%= oordersheet.FItemList(i).FRealItemno %></td>
	<td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oordersheet.FItemList(i).FRealItemno*oordersheet.FItemList(i).FBuycash %>" ><%= FormatNumber(oordersheet.FItemList(i).FRealItemno*oordersheet.FItemList(i).FBuycash,0) %></td>
</tr>
<% next %>
<tr>
	<td align=center class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Note</td>
	<td colspan=4 class=normal  style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= nl2br(oordersheetmaster.FoneItem.FComment) %>&nbsp;</td>
	<td align=center class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= ttlcount %>" ><%= ttlcount %></td>
	<td class=Format_number align=right style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= ttlbuycash %>"><b>\<%= ForMatNumber(ttlbuycash,0) %></b></td>
</tr>
<tr height=16 style='mso-height-source:userset;height:12.0pt'>
	<td height=16 class=normal style='height:12.0pt'></td>
	<td class=normal></td>
	<td colspan=2 class=normal>　</td>
	<td class=normal></td>
	<td class=normal></td>
	<td class=normal></td>
</tr>
<tr style='mso-height-source:userset;'>
	<td align=center class=normal_b style='border-top:0.5pt solid black; border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Sender</td>
	<td colspan=2 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><!--(인)-->&nbsp;</td>
	<td align=center colspan=2 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>Acceptor</td>
	<td colspan=2 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><!--(인)-->&nbsp;</td>
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

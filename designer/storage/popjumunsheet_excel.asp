<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim idx,itype
idx = requestCheckVar(request("idx"),20)
itype = requestCheckVar(request("itype"),50)


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
if not IsNULL(oordersheetmaster.FOneItem.FScheduleDate) then
scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FScheduleDate),10),"-","/")
end if

dim ttlsellcash, ttlbuycash, ttlcount
ttlsellcash = 0
ttlbuycash  = 0
ttlcount    = 0
%>

<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + oordersheetmaster.FOneItem.Ftargetid + Left(CStr(now()),10) + ".xls"
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
 <tr height=35 style='height:26.25pt '>
  <td colspan=4 height=35 class=big_title align=left style='border-bottom:0.5pt solid black;'>�ֹ�������(<%= obrand.FChargeName %>)</td>
  <td width=240 colspan=3 align=right class=mid_title style='border-bottom:0.5pt solid black;' >�ٹ����� (<%= oordersheetmaster.FOneItem.FBaljuCode %>)</td>
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
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >�� ¥</td>
  <td rowspan=2 class=Format_Y1 align=left ><b><%= scheduleorexedate %></b></td>
  <td rowspan=6></td>
  <td width=74 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�� �� �� ȣ</td>
  <td colspan=3 class=normal style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FSocNo %></td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�� ȣ</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FSocName %></td>
  <td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>��  ��</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FCeoName %></td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >�� ȣ</td>
  <td rowspan=2 class=mid_title><%= oordersheetmaster.FOneItem.Fbaljuname %></td>
  <td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>����������</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FAddress %></td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�� ��</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FUptae %></td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >�ѱݾ�</td>
  <td rowspan=2 align=left class=Format_number_L x:num="<%= oordersheetmaster.FOneItem.FTotalBuycash %>" ><b>\<%= ForMatNumber(oordersheetmaster.FOneItem.FTotalBuycash,0) %></b></td>
  <td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�� ��</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FUpjong %></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�� �� ��</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oordersheetmaster.FOneItem.Fregname %></td>
  <td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�� �� ó</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>1644-1851</td>
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
 <tr height=17 align=center >
  <td height=17 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>��ǰ�ڵ�</td>
  <td width=160 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>��ǰ��</td>
  <td width=76 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�ɼ�</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�Һ��ڰ�</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>���ް�</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>����</td>
  <td width=80 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�հ�</td>
 </tr>
 <% for i=0 to oordersheet.FResultCount -1 %>
 <%
 	ttlsellcash = ttlsellcash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FSellcash
 	ttlbuycash = ttlbuycash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FBuycash
 	ttlcount = ttlcount + oordersheet.FItemList(i).Frealitemno
 %>
 <tr height=17 style='mso-height-source:userset;height:20pt'>
  <td height=17 class=normal width=97 style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
  <%= oordersheet.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %>-<%= oordersheet.FItemList(i).FItemOption %></td>
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
 <tr height=22 >
  <td align=center height=22 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>���</td>
  <td colspan=4 class=normal  style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= nl2br(oordersheetmaster.FoneItem.FComment) %>&nbsp;</td>
  <td align=center class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= ttlcount %>" ><%= ttlcount %></td>
  <td class=Format_number align=right style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= ttlbuycash %>"><b>\<%= ForMatNumber(ttlbuycash,0) %></b></td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal style='height:12.0pt'></td>
  <td class=normal></td>
  <td colspan=2 class=normal>��</td>
  <td class=normal></td>
  <td class=normal></td>
  <td class=normal></td>
 </tr>
 <tr height=30 style='mso-height-source:userset;height:22.5pt'>
  <td align=center height=30 class=normal_b style='border-top:0.5pt solid black; border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�ΰ���</td>
  <td colspan=2 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>(��)&nbsp;</td>
  <td align=center colspan=2 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>�μ���</td>
  <td colspan=2 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>(��)&nbsp;</td>
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

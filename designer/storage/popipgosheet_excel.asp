<%
'Response.AddHeader "Cache-Control","no-cache"
'Response.AddHeader "Expires","0"
'Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim idx,itype, ibrandname
idx = requestCheckVar(request("idx"),20)
itype = requestCheckVar(request("itype"),50)
ibrandname = requestCheckVar(request("ibrandname"),100)

'==============================================================================
dim oipchul, oipchuldetail
set oipchul = new CIpChulStorage
oipchul.FRectId = idx
oipchul.GetIpChulMaster

set oipchuldetail = new CIpChulStorage
oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
oipchuldetail.GetIpChulDetail

'==============================================================================
dim obrand
set obrand = new CPartnerUser
obrand.FRectDesignerID = oipchul.FOneItem.Fsocid
obrand.GetOnePartnerNUser



dim i

dim executedate

if (oipchul.FOneItem.Fexecutedt <> "") then
	executedate = replace(Left(CstR(oipchul.FOneItem.Fexecutedt),10),"-","/")
else
	executedate = "<font color='red'>미입고</font>"
end if

dim ttlsellcash, ttlsuplycash, ttlcount
ttlsellcash = 0
ttlsuplycash  = 0
ttlcount    = 0
%>
<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + oipchul.FOneItem.Fsocid + Left(CStr(now()),10) + ".xls"
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
  <td colspan=5 height=35 class=big_title align=left style='border-bottom:0.5pt solid black;'>입고내역서(<%= obrand.FOneItem.FSocName_Kor %>)</td>
  <td width=240 colspan=3 align=right class=mid_title style='border-bottom:0.5pt solid black;' >텐바이텐 (<%= oipchul.FOneItem.Fcode %>)</td>
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
  <td rowspan=2 colspan=2 class=Format_Y1 align=left ><b><%= executedate %></b></td>
  <td rowspan=6></td>
  <td width=74 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>등 록 번 호</td>
  <td colspan=3 class=normal style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FOneItem.Fcompany_no %>&nbsp;</td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상 호</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FOneItem.Fcompany_name %>&nbsp;</td>
  <td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>성  명</td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FOneItem.FCeoname %>&nbsp;</td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >상 호</td>
  <td rowspan=2 colspan=2  class=mid_title><%= obrand.FOneItem.Fcompany_name %>&nbsp;</td>
  <td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>사업장소재지</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FOneItem.Faddress %><br><%= obrand.FOneItem.Fmanager_address %></td>
 </tr>
 <tr height=16 >
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>업 태</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FOneItem.Fcompany_uptae %>&nbsp;</td>
 </tr>
 <tr height=16 >
  <td rowspan=2 height=32 class=mid_title >총금액</td>
  <td rowspan=2 colspan=2  align=left class=Format_number_L x:num="<%= oipchul.FOneItem.Ftotalsuplycash %>" ><b>\<%= ForMatNumber(oipchul.FOneItem.Ftotalsuplycash,0) %></b></td>
  <td class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>종 목</td>
  <td colspan=3 class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= obrand.FOneItem.Fcompany_upjong %>&nbsp;</td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'></td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>&nbsp;</td>
  <td class=normal_b style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'></td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>&nbsp;</td>
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
  <td width=70 class=normal_b style='border-left:0.5pt solid black; border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>브랜드ID</td>
  <td height=17 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상품코드</td>
  <td width=160 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>상품명</td>
  <td width=76 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>옵션</td>
  <td width=70 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>소비자가</td>
  <td width=70 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>입고가</td>
  <td width=40 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>수량</td>
  <td width=70 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>합계</td>
 </tr>
 <% for i=0 to oipchuldetail.FResultCount -1 %>
 <%
 	ttlsellcash = ttlsellcash + oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsellcash
 	ttlsuplycash = ttlsuplycash + oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash
 	ttlcount = ttlcount + oipchuldetail.FItemList(i).Fitemno
 %>
 <% if ibrandname=oipchuldetail.FItemList(i).Fimakerid then %>
 <tr height=17 style='mso-height-source:userset;height:20pt' bgcolor="#CCCCCC">
 <% else %>
 <tr height=17 style='mso-height-source:userset;height:20pt'>
 <% end if %>
  <td class=normal style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
  	<%= oipchuldetail.FItemList(i).FIMakerid %>
  </td>
  <td height=17 class=normal width=97 style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
  <%= oipchuldetail.FItemList(i).Fiitemgubun %>-<%= CHKIIF(oipchuldetail.FItemList(i).FItemId>=1000000,Format00(8,oipchuldetail.FItemList(i).FItemId),Format00(6,oipchuldetail.FItemList(i).FItemId)) %>-<%= oipchuldetail.FItemList(i).FItemOption %></td>
  <td class=normal style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>
  	<%= oipchuldetail.FItemList(i).FIItemName %>
  </td>
  <td class=normal width=76 style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= oipchuldetail.FItemList(i).FIItemoptionName %>&nbsp;</td>
  <td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oipchuldetail.FItemList(i).Fsellcash %>"><%= FormatNumber(oipchuldetail.FItemList(i).Fsellcash,0) %></td>
  <td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oipchuldetail.FItemList(i).Fsuplycash %>"><%= FormatNumber(oipchuldetail.FItemList(i).Fsuplycash,0) %></td>
  <td align=center class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oipchuldetail.FItemList(i).Fitemno %>" ><%= oipchuldetail.FItemList(i).Fitemno %></td>
  <td align=right class=Format_number style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;' x:num="<%= oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash %>" ><%= FormatNumber(oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash,0) %></td>
 </tr>
 <% next %>
 <tr height=22 >
  <td align=center height=22 class=normal_b style='border-left:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>비고</td>
  <td colspan=5 class=normal  style='border-right:0.5pt solid black; border-bottom:0.5pt solid black;'><%= nl2br(oipchul.FOneItem.Fcomment) %>&nbsp;</td>
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
  <td colspan=2 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>(인)&nbsp;</td>
  <td align=center colspan=2 class=normal_b style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>인수자</td>
  <td colspan=3 class=normal align=right style='border-top:0.5pt solid black; border-right:0.5pt solid black; border-bottom:0.5pt solid black;'>(인)&nbsp;</td>
 </tr>
</table>
</body>
</html>

<%
set obrand = Nothing
set oipchul = Nothing
set oipchuldetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

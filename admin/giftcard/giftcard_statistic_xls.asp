<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim GiftStatisticNew, iCurrentpage, page, vGubun, vSDate, vEDate, arrList, intLoop, iTotCnt, i, vTotal1, vTotal2, vTotal3, vTotal4
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vGubun			= NullFillWith(requestCheckVar(request("gubun"),10),"")
	vSDate			= NullFillWith(requestCheckVar(request("sdate"),10),DateAdd("m",-1,date))
	vEDate			= NullFillWith(requestCheckVar(request("edate"),10),date)
	vTotal1 = 0
	vTotal2 = 0
	vTotal3 = 0
	vTotal4 = 0
	
	
	Set GiftStatisticNew = new ClsGift
	GiftStatisticNew.FGubun = vGubun
	GiftStatisticNew.FSDate = vSDate
	GiftStatisticNew.FEDate = vEDate
	GiftStatisticNew.FGiftStatisticNew
	
	iTotCnt = GiftStatisticNew.ftotalcount
	
	Response.ContentType = "application/x-msexcel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=Gift카드구매및사용내역.xls"
%>

<html>
<head></head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<body>
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center" rowspan="2">금액권</td>
	<td align="center" colspan="2">구매</td>
	<td align="center" colspan="2">등록</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td align="center">구매 건수</td>
	<td align="center">구매액</td>
	<td align="center">등록 건수</td>
	<td align="center">등록액</td>
</tr>
<%
	If GiftStatisticNew.FResultCount <> 0 Then
		For i = 0 To GiftStatisticNew.FResultCount -1
%>
		<tr bgcolor="FFFFFF">
			<td width="100" align="center"><%= GetCardName(GiftStatisticNew.FItemList(i).fsubtotalprice) %></td>
			<td width="100" align="center"><%=GiftStatisticNew.FItemList(i).foidx%></td>
			<td width="100" align="center"><%=FormatNumber(GiftStatisticNew.FItemList(i).fsubtotalprice,0) %></td>
			<td width="100" align="center"><%=GiftStatisticNew.FItemList(i).fridx%></td>
			<td width="100" align="center"><%=FormatNumber(GiftStatisticNew.FItemList(i).fcardprice,0) %></td>
		</tr>
<%
			vTotal1 = CLng(vTotal1) + CLng(GiftStatisticNew.FItemList(i).foidx)
			vTotal2 = CLng(vTotal2) + CLng(GiftStatisticNew.FItemList(i).fsubtotalprice)
			vTotal3 = CLng(vTotal3) + CLng(GiftStatisticNew.FItemList(i).fridx)
			vTotal4 = CLng(vTotal4) + CLng(GiftStatisticNew.FItemList(i).fcardprice)
		Next
		
		Response.Write "<tr bgcolor=""FFFFFF"" height=""30""><td align=""center"" bgcolor=""#E6E6E6"">합계</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal1,0) & "</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal2,0) & "</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal3,0) & "</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal4,0) & "</td>"
		Response.Write "</tr>"
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>

<% Set GiftStatisticNew = Nothing %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
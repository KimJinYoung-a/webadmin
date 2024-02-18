<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim iCurrentpage, GiftStatisticlist, i, iTotCnt, vSDate, vEDate, page, vGubun, vOrderSerial, vUserID, vUserName, vReqHP, vReqHP1, vReqHP2, vReqHP3, vTotalSum, vParam, vSumTemp, vNoCouponNo
	vTotalSum = "x"
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vGubun			= NullFillWith(requestCheckVar(request("gubun"),10),"")
	vOrderSerial	= NullFillWith(requestCheckVar(request("orderserial"),30),"")
	vUserID			= NullFillWith(requestCheckVar(request("userid"),50),"")
	vUserName		= NullFillWith(requestCheckVar(request("username"),100),"")
	vReqHP1			= NullFillWith(requestCheckVar(request("reqhp1"),3),"")
	vReqHP2			= NullFillWith(requestCheckVar(request("reqhp2"),4),"")
	vReqHP3			= NullFillWith(requestCheckVar(request("reqhp3"),4),"")
	If vReqHP1 <> "" AND vReqHP2 <> "" AND vReqHP3 <> "" Then
		vReqHP = vReqHP1 & "-" & vReqHP2 & "-" & vReqHP3
	End If
	vSDate			= NullFillWith(requestCheckVar(request("sdate"),10),"")
	vEDate			= NullFillWith(requestCheckVar(request("edate"),10),"")
	vNoCouponNo		= NullFillWith(requestCheckVar(request("nocouponno"),1),"")
	

	Set GiftStatisticlist = new ClsGift
	If vSDate <> "" OR vEDate <> "" Then
		vTotalSum = "o"
		GiftStatisticlist.FPageSize = "1000"
	End IF
	GiftStatisticlist.FCurrPage = page
	GiftStatisticlist.FGubun = vGubun
	GiftStatisticlist.FTCouponNo = vOrderSerial
	GiftStatisticlist.FUserID = vUserID
	GiftStatisticlist.FUSerName = vUserName
	GiftStatisticlist.FReqHP = vReqHP
	GiftStatisticlist.FSDate = vSDate
	GiftStatisticlist.FEDate = vEDate
	GiftStatisticlist.FNoCouponno = vNoCouponNo
	GiftStatisticlist.FCouponStatisticList
	
	iTotCnt = GiftStatisticlist.ftotalcount

	
	Response.ContentType = "application/x-msexcel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=기프티콘_기프팅_10x10쿠폰내역.xls"
%>

<html>
<head></head>
<body>
<table cellpadding="3" cellspacing="1" border="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center">결제방법</td>
	<td align="center">티콘/팅 쿠폰번호</td>
	<td align="center">텐바이텐쿠폰번호</td>
	<td align="center">텐바이텐쿠폰명</td>
	<td align="center">쿠폰금액</td>
	<td align="center">UserID</td>
	<td align="center">회원명</td>
	<td align="center">주문번호</td>
	<td align="center">등록일</td>
</tr>
<%
	If GiftStatisticlist.FResultCount <> 0 Then
		vSumTemp = 0
		For i = 0 To GiftStatisticlist.FResultCount -1

%>
		<tr bgcolor="FFFFFF">
			<td width="70" align="center"><%=GiftStatisticlist.FItemList(i).fgubun%></td>
			<td width="110" align="center"><%=GiftStatisticlist.FItemList(i).fcouponno%></td>
			<td width="110" align="center"><%=GiftStatisticlist.FItemList(i).fcouponidx%></td>
			<td width="150" align="center"><%=GiftStatisticlist.FItemList(i).fcouponname%></td>
			<td width="80" align="center"><%=GiftStatisticlist.FItemList(i).fcouponvalue%></td>
			<td width="100" align="center"><%=GiftStatisticlist.FItemList(i).fuserid%></td>
			<td width="80" align="center"><%=GiftStatisticlist.FItemList(i).fusername%></td>
			<td width="100" align="center"><%=GiftStatisticlist.FItemList(i).forderserial%></td>
			<td width="150"> <%=GiftStatisticlist.FItemList(i).fregdate %></td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>

<%
	set GiftStatisticlist = nothing
%>

</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
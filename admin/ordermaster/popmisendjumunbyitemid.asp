<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%
dim itemid,obalju

itemid = request("itemid")

set obalju = New COldMiSend
obalju.FRectItemid = itemid
obalju.GetMiSendOrderByitemid

dim i
%>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	상품코드 : <input type="text" name="orderserial" value="<%= obalju.FRectItemid %>" size="12">
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80" >주문번호</td>
		<td width="80">구매자 /<br>수령인</td>
	    <td width="80">주문일 /<br>결제일</td>
	   	<td width="60">사이트명</td>
		<td width="80">아이디</td>
		<td width="60">결제금액</td>
		<td width="70">주문상태/<br>송장No</td>
		<td width="50">주문수량</td>
		<td width="50">부족수량</td>
		<td width="70">지연<br>사유</td>
		<td>요청사항</td>
		<td width="70">처리<br>결과</td>
		<td width="70">처리<br>구분</td>
	</tr>
	<% for i=0 to obalju.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= obalju.FItemList(i).Forderserial %></td>
		<td><%= obalju.FItemList(i).FBuyName %> <br><%= obalju.FItemList(i).FReqName %></td>
	    <td><%= Left(obalju.FItemList(i).FRegDate,10) %> <br><%= Left(obalju.FItemList(i).FIpkumDate,10) %></td>
	   	<td><%= obalju.FItemList(i).FSiteName %></td>
		<td><%= obalju.FItemList(i).FUserId %></td>
		<td><%= FormatNumber(obalju.FItemList(i).FSubTotalPrice,0) %></td>
		<td><font color="<%= obalju.FItemList(i).IpkumDivColor %>"><%= obalju.FItemList(i).IpkumDivName %></font><br><%= obalju.FItemList(i).FDeliveryNo %></td>
		<td><%= obalju.FItemList(i).Fitemno %></td>
		<td><font color="red"><b><%= obalju.FItemList(i).FItemLackNo %></b></font></td>
		<td><%= obalju.FItemList(i).getMiSendCodeName %></td>
		<td><%= obalju.FItemList(i).FrequestString  %></td>
		<td><%= obalju.FItemList(i).FfinishString  %></td>
		<td><%= obalju.FItemList(i).GetStateString  %></td>
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
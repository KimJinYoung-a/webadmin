<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mdMenu/catemanageCls.asp" -->
<%
Dim olist
Dim dispCate, maxDepth
Dim page, i, Depth1Code, Depth1Name, j
Dim yyyy, viewList, strParam, viewGubun

maxDepth = 1
dispCate	= requestCheckVar(Request("disp"),16)
page 		= requestCheckVar(Request("page"),2)
yyyy		= requestCheckVar(Request("yyyy"),4)
viewList	= requestCheckVar(Request("viewList"),1)
viewGubun	= requestCheckVar(Request("gubun"),3)

If yyyy = "" Then yyyy = LEFT(date(), 4)
If page = "" Then page = 1

strParam = "?menupos="&menupos&"&disp="&dispCate&"&page="&page&"&yyyy="&yyyy
SET olist = new CMDCategory
	olist.FPageSize		= 500
	olist.FCurrPage		= 1
	olist.FRectCatecode	= dispCate
	olist.FRectYyyy		= yyyy
	olist.FRectGubun	= viewGubun
	olist.getMDPurposeRegedList
%>
<script type="text/javascript">
function form_check(f){
	f.submit();
}
function popRegPrice(code, mon, gubun){
    var pwin = window.open("/admin/mdMenu/purpose/popRegPrice.asp?catecode="+code+'&yyyy=<%=yyyy%>&mm='+mon+'&gubun='+gubun,"popOptionAddPrc","width=800,height=700,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function goValuePage(cd){
	var gg;
	gg = "ON"
	switch(cd){
		case "A" : location.replace('/admin/mdMenu/purpose/index.asp<%=strParam%>&gubun='+gg); break;
		case "B" : location.replace('/admin/mdMenu/purpose/target.asp<%=strParam%>&gubun='+gg); break;
		case "C" : location.replace('/admin/mdMenu/purpose/result.asp<%=strParam%>&gubun='+gg); break;
	}
}
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" height="50" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td bgcolor="<%= adminColor("gray") %>" align="left">
		검색연월 : <% DrawPurposeDateBox yyyy  %>
		&nbsp;&nbsp;
		전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		&nbsp;&nbsp;
		<select name="viewList" class="select" onchange="goValuePage(this.value);">
			<option value="A" >목표+실적보기
			<option value="B" selected>목표보기
			<option value="C" >실적보기
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색" onclick="form_check(this.form)">
	</td>
</tr>
</form>
</table>
<br><br>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="8%"><%= Chkiif(viewGubun= "ON", "온라인", "오프라인") %></td>
    <td colspan="14">월</td>
    <td rowspan="2" width="8%">합계</td>
    <td rowspan="2" width="8%">비중</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>카테고리</td>
	<td width="5%">&nbsp;</td>
	<td width="8%">구분</td>
	<% For i = 1 to 12 %>
    <td width="120"><%= i %>월</td>
	<% Next %>
</tr>
<%
Dim isreg1, isreg2, isreg3, isreg4, isreg5, isreg6, isreg7, isreg8, isreg9, isreg10, isreg11, isreg12
Dim profitper1, profitper2, profitper3, profitper4, profitper5, profitper6, profitper7, profitper8, profitper9, profitper10, profitper11, profitper12, totProfitper
Dim TotTarget1to12, TotProfit1to12

If olist.FResultCount > 0 Then
	TotTarget1to12 = olist.FOneItem.FTotTarget1 + olist.FOneItem.FTotTarget2 + olist.FOneItem.FTotTarget3 + olist.FOneItem.FTotTarget4 + olist.FOneItem.FTotTarget5 + olist.FOneItem.FTotTarget6 + olist.FOneItem.FTotTarget7 + olist.FOneItem.FTotTarget8 + olist.FOneItem.FTotTarget9 + olist.FOneItem.FTotTarget10 + olist.FOneItem.FTotTarget11 + olist.FOneItem.FTotTarget12
	TotProfit1to12 = olist.FOneItem.FTotProfit1 + olist.FOneItem.FTotProfit2 + olist.FOneItem.FTotProfit3 + olist.FOneItem.FTotProfit4 + olist.FOneItem.FTotProfit5 + olist.FOneItem.FTotProfit6 + olist.FOneItem.FTotProfit7 + olist.FOneItem.FTotProfit8 + olist.FOneItem.FTotProfit9 + olist.FOneItem.FTotProfit10 + olist.FOneItem.FTotProfit11 + olist.FOneItem.FTotProfit12
	For i = 0 to olist.FResultCount -1
		If olist.FItemList(i).FTarget1 = 0 AND olist.FItemList(i).FProfit1 = 0 Then isreg1 = False Else isreg1 = True End If
		If olist.FItemList(i).FTarget2 = 0 AND olist.FItemList(i).FProfit2 = 0 Then isreg2 = False Else isreg2 = True End If
		If olist.FItemList(i).FTarget3 = 0 AND olist.FItemList(i).FProfit3 = 0 Then isreg3 = False Else isreg3 = True End If
		If olist.FItemList(i).FTarget4 = 0 AND olist.FItemList(i).FProfit4 = 0 Then isreg4 = False Else isreg4 = True End If
		If olist.FItemList(i).FTarget5 = 0 AND olist.FItemList(i).FProfit5 = 0 Then isreg5 = False Else isreg5 = True End If
		If olist.FItemList(i).FTarget6 = 0 AND olist.FItemList(i).FProfit6 = 0 Then isreg6 = False Else isreg6 = True End If
		If olist.FItemList(i).FTarget7 = 0 AND olist.FItemList(i).FProfit7 = 0 Then isreg7 = False Else isreg7 = True End If
		If olist.FItemList(i).FTarget8 = 0 AND olist.FItemList(i).FProfit8 = 0 Then isreg8 = False Else isreg8 = True End If
		If olist.FItemList(i).FTarget9 = 0 AND olist.FItemList(i).FProfit9 = 0 Then isreg9 = False Else isreg9 = True End If
		If olist.FItemList(i).FTarget10 = 0 AND olist.FItemList(i).FProfit10 = 0 Then isreg10 = False Else isreg10 = True End If
		If olist.FItemList(i).FTarget11 = 0 AND olist.FItemList(i).FProfit11 = 0 Then isreg11 = False Else isreg11 = True End If
		If olist.FItemList(i).FTarget12 = 0 AND olist.FItemList(i).FProfit12 = 0 Then isreg12 = False Else isreg12 = True End If

		If olist.FItemList(i).FTarget1 = 0 OR olist.FItemList(i).FProfit1 = 0 Then profitper1 = "0%" Else profitper1 = formatnumber(olist.FItemList(i).FProfit1 / olist.FItemList(i).FTarget1 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget2 = 0 OR olist.FItemList(i).FProfit2 = 0 Then profitper2 = "0%" Else profitper2 = formatnumber(olist.FItemList(i).FProfit2 / olist.FItemList(i).FTarget2 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget3 = 0 OR olist.FItemList(i).FProfit3 = 0 Then profitper3 = "0%" Else profitper3 = formatnumber(olist.FItemList(i).FProfit3 / olist.FItemList(i).FTarget3 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget4 = 0 OR olist.FItemList(i).FProfit4 = 0 Then profitper4 = "0%" Else profitper4 = formatnumber(olist.FItemList(i).FProfit4 / olist.FItemList(i).FTarget4 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget5 = 0 OR olist.FItemList(i).FProfit5 = 0 Then profitper5 = "0%" Else profitper5 = formatnumber(olist.FItemList(i).FProfit5 / olist.FItemList(i).FTarget5 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget6 = 0 OR olist.FItemList(i).FProfit6 = 0 Then profitper6 = "0%" Else profitper6 = formatnumber(olist.FItemList(i).FProfit6 / olist.FItemList(i).FTarget6 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget7 = 0 OR olist.FItemList(i).FProfit7 = 0 Then profitper7 = "0%" Else profitper7 = formatnumber(olist.FItemList(i).FProfit7 / olist.FItemList(i).FTarget7 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget8 = 0 OR olist.FItemList(i).FProfit8 = 0 Then profitper8 = "0%" Else profitper8 = formatnumber(olist.FItemList(i).FProfit8 / olist.FItemList(i).FTarget8 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget9 = 0 OR olist.FItemList(i).FProfit9 = 0 Then profitper9 = "0%" Else profitper9 = formatnumber(olist.FItemList(i).FProfit9 / olist.FItemList(i).FTarget9 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget10 = 0 OR olist.FItemList(i).FProfit10 = 0 Then profitper10 = "0%" Else profitper10 = formatnumber(olist.FItemList(i).FProfit10 / olist.FItemList(i).FTarget10 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget11 = 0 OR olist.FItemList(i).FProfit11 = 0 Then profitper11 = "0%" Else profitper11 = formatnumber(olist.FItemList(i).FProfit11 / olist.FItemList(i).FTarget11 * 100, 1) & "%" End If
		If olist.FItemList(i).FTarget12 = 0 OR olist.FItemList(i).FProfit12 = 0 Then profitper12 = "0%" Else profitper12 = formatnumber(olist.FItemList(i).FProfit12 / olist.FItemList(i).FTarget12 * 100, 1) & "%" End If
		
		If olist.FItemList(i).FSumTartgetMoney = 0 OR olist.FItemList(i).FSumProfitMoney = 0 Then totProfitper = "0%" Else totProfitper = formatnumber(olist.FItemList(i).FSumProfitMoney / olist.FItemList(i).FSumTartgetMoney * 100, 1) & "%" End If
%>
<!--													매출목표													-->
<tr align="center" bgcolor="FFFFFF" height="25">
	<td rowspan="3" <%= Chkiif(olist.FItemList(i).FDepth = 1, "bgcolor='#FFFFD7'","bgcolor='FFFFFF'") %> ><%= olist.FItemList(i).FCatename %></td>
	<td rowspan="3">목표</td>
	<td>매출목표</td>
<% If isreg1 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 1, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 1, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget1) %></strong></td>
<% End If %>

<% If isreg2 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 2, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 2, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget2) %></strong></td>
<% End If %>

<% If isreg3 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 3, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 3, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget3) %></strong></td>
<% End If %>

<% If isreg4 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 4, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 4, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget4) %></strong></td>
<% End If %>

<% If isreg5 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 5, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 5, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget5) %></strong></td>
<% End If %>

<% If isreg6 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 6, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 6, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget6) %></strong></td>
<% End If %>

<% If isreg7 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 7, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 7, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget7) %></strong></td>
<% End If %>

<% If isreg8 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 8, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 8, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget8) %></strong></td>
<% End If %>

<% If isreg9 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 9, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 9, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget9) %></strong></td>
<% End If %>

<% If isreg10 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 10, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 10, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget10) %></strong></td>
<% End If %>

<% If isreg11 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 11, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 11, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget11) %></strong></td>
<% End If %>

<% If isreg12 = False Then %>
	<td rowspan="3"><input type="button" class="button_s" value="등록" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 12, '<%=viewgubun%>');"></td>
<% Else %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 12, '<%=viewgubun%>');"><strong><%= NullOrCurrFormat(olist.FItemList(i).FTarget12) %></strong></td>
<% End If %>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FItemList(i).FSumTartgetMoney) %></strong></td>
	<td align="right">
		<strong><% If olist.FItemList(i).FSumTartgetMoney = 0 or TotTarget1to12 = 0 Then response.write "0%" Else response.write formatnumber(olist.FItemList(i).FSumTartgetMoney / TotTarget1to12 * 100)&"%" End If %></strong>
	</td>
</tr>
<!--													수익목표													-->
<tr align="center"  bgcolor="FFFFFF" height="25">
	<td>수익목표</td>
<% If isreg1 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 1, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit1) %></td>
<% End If %>

<% If isreg2 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 2, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit2) %></td>
<% End If %>

<% If isreg3 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 3, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit3) %></td>
<% End If %>

<% If isreg4 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 4, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit4) %></td>
<% End If %>

<% If isreg5 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 5, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit5) %></td>
<% End If %>

<% If isreg6 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 6, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit6) %></td>
<% End If %>

<% If isreg7 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 7, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit7) %></td>
<% End If %>

<% If isreg8 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 8, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit8) %></td>
<% End If %>

<% If isreg9 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 9, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit9) %></td>
<% End If %>

<% If isreg10 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 10, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit10) %></td>
<% End If %>

<% If isreg11 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 11, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit11) %></td>
<% End If %>

<% If isreg12 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 12, '<%=viewgubun%>');"><%= NullOrCurrFormat(olist.FItemList(i).FProfit12) %></td>
<% End If %>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FSumProfitMoney) %></td>
	<td align="right"><% If olist.FItemList(i).FSumProfitMoney = 0 or TotProfit1to12 = 0 Then response.write "0%" Else response.write formatnumber(olist.FItemList(i).FSumProfitMoney / TotProfit1to12 * 100)&"%" End If %></td>
</tr>
<!--													수익율														-->
<tr align="center"  bgcolor="FFFFFF" height="25">
	<td>수익율</td>
<% If isreg1 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 1, '<%=viewgubun%>');"><%= profitper1 %></td>
<% End If %>

<% If isreg2 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 2, '<%=viewgubun%>');"><%= profitper2 %></td>
<% End If %>

<% If isreg3 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 3, '<%=viewgubun%>');"><%= profitper3 %></td>
<% End If %>

<% If isreg4 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 4, '<%=viewgubun%>');"><%= profitper4 %></td>
<% End If %>

<% If isreg5 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 5, '<%=viewgubun%>');"><%= profitper5 %></td>
<% End If %>

<% If isreg6 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 6, '<%=viewgubun%>');"><%= profitper6 %></td>
<% End If %>

<% If isreg7 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 7, '<%=viewgubun%>');"><%= profitper7 %></td>
<% End If %>

<% If isreg8 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 8, '<%=viewgubun%>');"><%= profitper8 %></td>
<% End If %>

<% If isreg9 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 9, '<%=viewgubun%>');"><%= profitper9 %></td>
<% End If %>

<% If isreg10 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 10, '<%=viewgubun%>');"><%= profitper10 %></td>
<% End If %>

<% If isreg11 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 11, '<%=viewgubun%>');"><%= profitper11 %></td>
<% End If %>

<% If isreg12 = True Then %>
	<td align="right" style="cursor:pointer;" onclick="popRegPrice(<%=olist.FItemList(i).FCateCode%>, 12, '<%=viewgubun%>');"><%= profitper12 %></td>
<% End If %>
	<td align="right"><%= totProfitper %></td>
	<td align="right">&nbsp;</td>
</tr>
<%
	Next
%>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td rowspan="3" bgcolor='#FF7E7E'>TOTAL</td>
	<td rowspan="3">목표</td>
	<td>월별매출목표</td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget1) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget2) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget3) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget4) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget5) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget6) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget7) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget8) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget9) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget10) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget11) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotTarget12) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(TotTarget1to12) %></strong></td>
	<td align="right"><strong><%= Chkiif(TotTarget1to12 = 0, "0%", "100%") %></strong></td>
</tr>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td>월별수익목표</td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit1) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit2) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit3) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit4) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit5) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit6) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit7) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit8) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit9) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit10) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit11) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotProfit12) %></td>
	<td align="right"><%= NullOrCurrFormat(TotProfit1to12) %></td>
	<td align="right"><%= Chkiif(TotProfit1to12 = 0, "0%", "100%") %></td>
</tr>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td>월별수익율</td>
	<td align="right"><% If olist.FOneItem.FTotTarget1=0 or olist.FOneItem.FTotProfit1=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit1 / olist.FOneItem.FTotTarget1 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget2=0 or olist.FOneItem.FTotProfit2=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit2 / olist.FOneItem.FTotTarget2 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget3=0 or olist.FOneItem.FTotProfit3=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit3 / olist.FOneItem.FTotTarget3 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget4=0 or olist.FOneItem.FTotProfit4=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit4 / olist.FOneItem.FTotTarget4 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget5=0 or olist.FOneItem.FTotProfit5=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit5 / olist.FOneItem.FTotTarget5 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget6=0 or olist.FOneItem.FTotProfit6=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit6 / olist.FOneItem.FTotTarget6 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget7=0 or olist.FOneItem.FTotProfit7=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit7 / olist.FOneItem.FTotTarget7 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget8=0 or olist.FOneItem.FTotProfit8=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit8 / olist.FOneItem.FTotTarget8 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget9=0 or olist.FOneItem.FTotProfit9=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit9 / olist.FOneItem.FTotTarget9 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget10=0 or olist.FOneItem.FTotProfit10=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit10 / olist.FOneItem.FTotTarget10 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget11=0 or olist.FOneItem.FTotProfit11=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit11 / olist.FOneItem.FTotTarget11 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotTarget12=0 or olist.FOneItem.FTotProfit12=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotProfit12 / olist.FOneItem.FTotTarget12 * 100)&"%" End If %></td>
	<td align="right"><% If TotTarget1to12=0 or TotProfit1to12=0 Then response.write "0%" Else response.write formatnumber(TotProfit1to12 / TotTarget1to12 * 100)&"%" End If %></td>
	<td align="right"></td>
</tr>
<%
End If
%>
</table>
<% SET olist = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
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
function popRegPrice(code, mon){
    var pwin = window.open("/admin/mdMenu/purpose/popRegPrice.asp?catecode="+code+'&yyyy=<%=yyyy%>&mm='+mon,"popOptionAddPrc","width=800,height=700,scrollbars=yes,resizable=yes");
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
			<option value="A">목표+실적보기
			<option value="B" >목표보기
			<option value="C" selected>실적보기
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색" onclick="form_check(this.form)">
	</td>
</tr>
</form>
</table>
<br><center><strong><font color="RED" size="3">※ 당월 (<%= month(now())%>월) 실적데이터는 6시간 전 데이터 입니다.</font></strong></center><br>
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
Dim itemcostper1, itemcostper2, itemcostper3, itemcostper4, itemcostper5, itemcostper6, itemcostper7, itemcostper8, itemcostper9, itemcostper10, itemcostper11, itemcostper12
Dim maechulper1, maechulper2, maechulper3, maechulper4, maechulper5, maechulper6, maechulper7, maechulper8, maechulper9, maechulper10, maechulper11, maechulper12, totMaechulper, totMaechulProfitper
Dim realPer1, realPer2, realPer3, realPer4, realPer5, realPer6, realPer7, realPer8, realPer9, realPer10, realPer11, realPer12, totrealper

Dim TotTarget1to12, TotProfit1to12, TotItemCost1to12, TotMaechul1to12

If olist.FResultCount > 0 Then
	TotTarget1to12 		= olist.FOneItem.FTotTarget1 + olist.FOneItem.FTotTarget2 + olist.FOneItem.FTotTarget3 + olist.FOneItem.FTotTarget4 + olist.FOneItem.FTotTarget5 + olist.FOneItem.FTotTarget6 + olist.FOneItem.FTotTarget7 + olist.FOneItem.FTotTarget8 + olist.FOneItem.FTotTarget9 + olist.FOneItem.FTotTarget10 + olist.FOneItem.FTotTarget11 + olist.FOneItem.FTotTarget12
	TotProfit1to12 		= olist.FOneItem.FTotProfit1 + olist.FOneItem.FTotProfit2 + olist.FOneItem.FTotProfit3 + olist.FOneItem.FTotProfit4 + olist.FOneItem.FTotProfit5 + olist.FOneItem.FTotProfit6 + olist.FOneItem.FTotProfit7 + olist.FOneItem.FTotProfit8 + olist.FOneItem.FTotProfit9 + olist.FOneItem.FTotProfit10 + olist.FOneItem.FTotProfit11 + olist.FOneItem.FTotProfit12
	TotItemCost1to12	= olist.FOneItem.FTotItemcost1 + olist.FOneItem.FTotItemcost2 + olist.FOneItem.FTotItemcost3 + olist.FOneItem.FTotItemcost4 + olist.FOneItem.FTotItemcost5 + olist.FOneItem.FTotItemcost6 + olist.FOneItem.FTotItemcost7 + olist.FOneItem.FTotItemcost8 + olist.FOneItem.FTotItemcost9 + olist.FOneItem.FTotItemcost10 + olist.FOneItem.FTotItemcost11 + olist.FOneItem.FTotItemcost12
	TotMaechul1to12		= olist.FOneItem.FTotmaechulProfit1 +  olist.FOneItem.FTotmaechulProfit2 + olist.FOneItem.FTotmaechulProfit3 + olist.FOneItem.FTotmaechulProfit4 + olist.FOneItem.FTotmaechulProfit5 + olist.FOneItem.FTotmaechulProfit6 + olist.FOneItem.FTotmaechulProfit7 + olist.FOneItem.FTotmaechulProfit8 + olist.FOneItem.FTotmaechulProfit9 + olist.FOneItem.FTotmaechulProfit10 + olist.FOneItem.FTotmaechulProfit11 + olist.FOneItem.FTotmaechulProfit12
	For i = 0 to olist.FResultCount -1
		If olist.FItemList(i).FItemcost1 = 0 OR olist.FItemList(i).FTarget1 = 0 Then itemcostper1 = "0%" Else itemcostper1 = formatnumber(olist.FItemList(i).FItemcost1 / olist.FItemList(i).FTarget1 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost2 = 0 OR olist.FItemList(i).FTarget2 = 0 Then itemcostper2 = "0%" Else itemcostper2 = formatnumber(olist.FItemList(i).FItemcost2 / olist.FItemList(i).FTarget2 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost3 = 0 OR olist.FItemList(i).FTarget3 = 0 Then itemcostper3 = "0%" Else itemcostper3 = formatnumber(olist.FItemList(i).FItemcost3 / olist.FItemList(i).FTarget3 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost4 = 0 OR olist.FItemList(i).FTarget4 = 0 Then itemcostper4 = "0%" Else itemcostper4 = formatnumber(olist.FItemList(i).FItemcost4 / olist.FItemList(i).FTarget4 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost5 = 0 OR olist.FItemList(i).FTarget5 = 0 Then itemcostper5 = "0%" Else itemcostper5 = formatnumber(olist.FItemList(i).FItemcost5 / olist.FItemList(i).FTarget5 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost6 = 0 OR olist.FItemList(i).FTarget6 = 0 Then itemcostper6 = "0%" Else itemcostper6 = formatnumber(olist.FItemList(i).FItemcost6 / olist.FItemList(i).FTarget6 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost7 = 0 OR olist.FItemList(i).FTarget7 = 0 Then itemcostper7 = "0%" Else itemcostper7 = formatnumber(olist.FItemList(i).FItemcost7 / olist.FItemList(i).FTarget7 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost8 = 0 OR olist.FItemList(i).FTarget8 = 0 Then itemcostper8 = "0%" Else itemcostper8 = formatnumber(olist.FItemList(i).FItemcost8 / olist.FItemList(i).FTarget8 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost9 = 0 OR olist.FItemList(i).FTarget9 = 0 Then itemcostper9 = "0%" Else itemcostper9 = formatnumber(olist.FItemList(i).FItemcost9 / olist.FItemList(i).FTarget9 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost10 = 0 OR olist.FItemList(i).FTarget10 = 0 Then itemcostper10 = "0%" Else itemcostper10 = formatnumber(olist.FItemList(i).FItemcost10 / olist.FItemList(i).FTarget10 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost11 = 0 OR olist.FItemList(i).FTarget11 = 0 Then itemcostper11 = "0%" Else itemcostper11 = formatnumber(olist.FItemList(i).FItemcost11 / olist.FItemList(i).FTarget11 * 100, 1) & "%" End If
		If olist.FItemList(i).FItemcost12 = 0 OR olist.FItemList(i).FTarget12 = 0 Then itemcostper12 = "0%" Else itemcostper12 = formatnumber(olist.FItemList(i).FItemcost12 / olist.FItemList(i).FTarget12 * 100, 1) & "%" End If

		If olist.FItemList(i).FMaechulProfit1 = 0 OR olist.FItemList(i).FProfit1 = 0 Then maechulper1 = "0%" Else maechulper1 = formatnumber(olist.FItemList(i).FMaechulProfit1 / olist.FItemList(i).FProfit1 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit2 = 0 OR olist.FItemList(i).FProfit2 = 0 Then maechulper2 = "0%" Else maechulper2 = formatnumber(olist.FItemList(i).FMaechulProfit2 / olist.FItemList(i).FProfit2 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit3 = 0 OR olist.FItemList(i).FProfit3 = 0 Then maechulper3 = "0%" Else maechulper3 = formatnumber(olist.FItemList(i).FMaechulProfit3 / olist.FItemList(i).FProfit3 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit4 = 0 OR olist.FItemList(i).FProfit4 = 0 Then maechulper4 = "0%" Else maechulper4 = formatnumber(olist.FItemList(i).FMaechulProfit4 / olist.FItemList(i).FProfit4 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit5 = 0 OR olist.FItemList(i).FProfit5 = 0 Then maechulper5 = "0%" Else maechulper5 = formatnumber(olist.FItemList(i).FMaechulProfit5 / olist.FItemList(i).FProfit5 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit6 = 0 OR olist.FItemList(i).FProfit6 = 0 Then maechulper6 = "0%" Else maechulper6 = formatnumber(olist.FItemList(i).FMaechulProfit6 / olist.FItemList(i).FProfit6 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit7 = 0 OR olist.FItemList(i).FProfit7 = 0 Then maechulper7 = "0%" Else maechulper7 = formatnumber(olist.FItemList(i).FMaechulProfit7 / olist.FItemList(i).FProfit7 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit8 = 0 OR olist.FItemList(i).FProfit8 = 0 Then maechulper8 = "0%" Else maechulper8 = formatnumber(olist.FItemList(i).FMaechulProfit8 / olist.FItemList(i).FProfit8 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit9 = 0 OR olist.FItemList(i).FProfit9 = 0 Then maechulper9 = "0%" Else maechulper9 = formatnumber(olist.FItemList(i).FMaechulProfit9 / olist.FItemList(i).FProfit9 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit10 = 0 OR olist.FItemList(i).FProfit10 = 0 Then maechulper10 = "0%" Else maechulper10 = formatnumber(olist.FItemList(i).FMaechulProfit10 / olist.FItemList(i).FProfit10 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit11 = 0 OR olist.FItemList(i).FProfit11 = 0 Then maechulper11 = "0%" Else maechulper11 = formatnumber(olist.FItemList(i).FMaechulProfit11 / olist.FItemList(i).FProfit11 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit12 = 0 OR olist.FItemList(i).FProfit12 = 0 Then maechulper12 = "0%" Else maechulper12 = formatnumber(olist.FItemList(i).FMaechulProfit12 / olist.FItemList(i).FProfit12 * 100, 1) & "%" End If

		If olist.FItemList(i).FMaechulProfit1 = 0 OR olist.FItemList(i).FItemcost1 = 0 Then realPer1 = "0%" Else realPer1 = formatnumber(olist.FItemList(i).FMaechulProfit1 / olist.FItemList(i).FItemcost1 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit2 = 0 OR olist.FItemList(i).FItemcost2 = 0 Then realPer2 = "0%" Else realPer2 = formatnumber(olist.FItemList(i).FMaechulProfit2 / olist.FItemList(i).FItemcost2 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit3 = 0 OR olist.FItemList(i).FItemcost3 = 0 Then realPer3 = "0%" Else realPer3 = formatnumber(olist.FItemList(i).FMaechulProfit3 / olist.FItemList(i).FItemcost3 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit4 = 0 OR olist.FItemList(i).FItemcost4 = 0 Then realPer4 = "0%" Else realPer4 = formatnumber(olist.FItemList(i).FMaechulProfit4 / olist.FItemList(i).FItemcost4 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit5 = 0 OR olist.FItemList(i).FItemcost5 = 0 Then realPer5 = "0%" Else realPer5 = formatnumber(olist.FItemList(i).FMaechulProfit5 / olist.FItemList(i).FItemcost5 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit6 = 0 OR olist.FItemList(i).FItemcost6 = 0 Then realPer6 = "0%" Else realPer6 = formatnumber(olist.FItemList(i).FMaechulProfit6 / olist.FItemList(i).FItemcost6 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit7 = 0 OR olist.FItemList(i).FItemcost7 = 0 Then realPer7 = "0%" Else realPer7 = formatnumber(olist.FItemList(i).FMaechulProfit7 / olist.FItemList(i).FItemcost7 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit8 = 0 OR olist.FItemList(i).FItemcost8 = 0 Then realPer8 = "0%" Else realPer8 = formatnumber(olist.FItemList(i).FMaechulProfit8 / olist.FItemList(i).FItemcost8 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit9 = 0 OR olist.FItemList(i).FItemcost9 = 0 Then realPer9 = "0%" Else realPer9 = formatnumber(olist.FItemList(i).FMaechulProfit9 / olist.FItemList(i).FItemcost9 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit10 = 0 OR olist.FItemList(i).FItemcost10 = 0 Then realPer10 = "0%" Else realPer10 = formatnumber(olist.FItemList(i).FMaechulProfit10 / olist.FItemList(i).FItemcost10 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit11 = 0 OR olist.FItemList(i).FItemcost11 = 0 Then realPer11 = "0%" Else realPer11 = formatnumber(olist.FItemList(i).FMaechulProfit11 / olist.FItemList(i).FItemcost11 * 100, 1) & "%" End If
		If olist.FItemList(i).FMaechulProfit12 = 0 OR olist.FItemList(i).FItemcost12 = 0 Then realPer12 = "0%" Else realPer12 = formatnumber(olist.FItemList(i).FMaechulProfit12 / olist.FItemList(i).FItemcost12 * 100, 1) & "%" End If

		If olist.FItemList(i).FSumItemcost = 0 OR olist.FItemList(i).FSumTartgetMoney = 0 Then totMaechulper = "0%" Else totMaechulper = formatnumber(olist.FItemList(i).FSumItemcost / olist.FItemList(i).FSumTartgetMoney * 100, 1) & "%" End If
		If olist.FItemList(i).FSumProfitMoney = 0 OR olist.FItemList(i).FSumMaechulProfit = 0 Then totMaechulProfitper = "0%" Else totMaechulProfitper = formatnumber(olist.FItemList(i).FSumMaechulProfit / olist.FItemList(i).FSumProfitmoney * 100, 1) & "%" End If
		If olist.FItemList(i).FSumMaechulProfit = 0 OR olist.FItemList(i).FSumItemcost = 0 Then totrealper = "0%" Else totrealper = formatnumber(olist.FItemList(i).FSumMaechulProfit / olist.FItemList(i).FSumItemcost * 100, 1) & "%" End If
%>
<!--													매출목표													-->
<tr align="center" bgcolor="FFFFFF" height="25">
	<td rowspan="5" <%= Chkiif(olist.FItemList(i).FDepth = 1, "bgcolor='#FFFFD7'","bgcolor='FFFFFF'") %> ><%= olist.FItemList(i).FCatename %></td>
	<td rowspan="5">실적</td>
	<td>구매총액</td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost1) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost2) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost3) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost4) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost5) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost6) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost7) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost8) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost9) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost10) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost11) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FItemcost12) %></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FItemList(i).FSumItemcost) %></strong></td>
	<td align="right">
		<strong><% If olist.FItemList(i).FSumItemcost = 0 or TotItemcost1to12 = 0 Then response.write "0%" Else response.write formatnumber(olist.FItemList(i).FSumItemcost / TotItemcost1to12 * 100)&"%" End If %></strong>
	</td>
</tr>
<tr align="center"  bgcolor="FFFFFF" height="25">
	<td>달성율</td>
	<td align="right"><%= itemcostper1 %></td>
	<td align="right"><%= itemcostper2 %></td>
	<td align="right"><%= itemcostper3 %></td>
	<td align="right"><%= itemcostper4 %></td>
	<td align="right"><%= itemcostper5 %></td>
	<td align="right"><%= itemcostper6 %></td>
	<td align="right"><%= itemcostper7 %></td>
	<td align="right"><%= itemcostper8 %></td>
	<td align="right"><%= itemcostper9 %></td>
	<td align="right"><%= itemcostper10 %></td>
	<td align="right"><%= itemcostper11 %></td>
	<td align="right"><%= itemcostper12 %></td>
	<td align="right"><%= totMaechulper %></td>
	<td align="right"></td>
</tr>
<tr align="center"  bgcolor="FFFFFF" height="25">
	<td>수익</td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit1) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit2) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit3) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit4) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit5) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit6) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit7) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit8) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit9) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit10) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit11) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FItemList(i).FMaechulProfit12) %></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FItemList(i).FSumMaechulProfit) %></strong></td>
	<td align="right">
		<strong><% If olist.FItemList(i).FSumMaechulProfit = 0 or TotMaechul1to12 = 0 Then response.write "0%" Else response.write formatnumber(olist.FItemList(i).FSumMaechulProfit / TotMaechul1to12 * 100)&"%" End If %></strong>
	</td>
</tr>
<tr align="center"  bgcolor="FFFFFF" height="25">
	<td>달성율</td>
	<td align="right"><%= maechulper1 %></td>
	<td align="right"><%= maechulper2 %></td>
	<td align="right"><%= maechulper3 %></td>
	<td align="right"><%= maechulper4 %></td>
	<td align="right"><%= maechulper5 %></td>
	<td align="right"><%= maechulper6 %></td>
	<td align="right"><%= maechulper7 %></td>
	<td align="right"><%= maechulper8 %></td>
	<td align="right"><%= maechulper9 %></td>
	<td align="right"><%= maechulper10 %></td>
	<td align="right"><%= maechulper11 %></td>
	<td align="right"><%= maechulper12 %></td>
	<td align="right"><%= totMaechulProfitper %></td>
	<td align="right"></td>
</tr>
<tr align="center"  bgcolor="FFFFFF" height="25">
	<td>수익율</td>
	<td align="right"><%= realPer1 %></td>
	<td align="right"><%= realPer2 %></td>
	<td align="right"><%= realPer3 %></td>
	<td align="right"><%= realPer4 %></td>
	<td align="right"><%= realPer5 %></td>
	<td align="right"><%= realPer6 %></td>
	<td align="right"><%= realPer7 %></td>
	<td align="right"><%= realPer8 %></td>
	<td align="right"><%= realPer9 %></td>
	<td align="right"><%= realPer10 %></td>
	<td align="right"><%= realPer11 %></td>
	<td align="right"><%= realPer12 %></td>
	<td align="right"><%= totrealper %></td>
	<td align="right"></td>
</tr>
<%
	Next
%>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td rowspan="3" bgcolor='#FF7E7E'>TOTAL</td>
	<td rowspan="3" >실적</td>
	<td>월별매출실적</td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost1) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost2) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost3) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost4) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost5) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost6) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost7) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost8) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost9) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost10) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost11) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(olist.FOneItem.FTotItemcost12) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(TotItemcost1to12) %></strong></td>
	<td align="right"><strong><%= Chkiif(TotItemcost1to12 = 0, "0%", "100%") %></strong></td>
</tr>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td>월별수익실적</td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit1) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit2) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit3) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit4) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit5) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit6) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit7) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit8) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit9) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit10) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit11) %></td>
	<td align="right"><%= NullOrCurrFormat(olist.FOneItem.FTotmaechulProfit12) %></td>
	<td align="right"><%= NullOrCurrFormat(TotMaechul1to12) %></td>
	<td align="right"><%= Chkiif(TotMaechul1to12 = 0, "0%", "100%") %></td>
</tr>

<tr align="center" bgcolor="FFFFFF" height="25">
	<td>월별실적수익율</td>
	<td align="right"><% If olist.FOneItem.FTotItemcost1=0 or olist.FOneItem.FTotmaechulProfit1=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit1 / olist.FOneItem.FTotItemcost1 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost2=0 or olist.FOneItem.FTotmaechulProfit2=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit2 / olist.FOneItem.FTotItemcost2 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost3=0 or olist.FOneItem.FTotmaechulProfit3=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit3 / olist.FOneItem.FTotItemcost3 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost4=0 or olist.FOneItem.FTotmaechulProfit4=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit4 / olist.FOneItem.FTotItemcost4 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost5=0 or olist.FOneItem.FTotmaechulProfit5=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit5 / olist.FOneItem.FTotItemcost5 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost6=0 or olist.FOneItem.FTotmaechulProfit6=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit6 / olist.FOneItem.FTotItemcost6 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost7=0 or olist.FOneItem.FTotmaechulProfit7=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit7 / olist.FOneItem.FTotItemcost7 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost8=0 or olist.FOneItem.FTotmaechulProfit8=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit8 / olist.FOneItem.FTotItemcost8 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost9=0 or olist.FOneItem.FTotmaechulProfit9=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit9 / olist.FOneItem.FTotItemcost9 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost10=0 or olist.FOneItem.FTotmaechulProfit10=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit10 / olist.FOneItem.FTotItemcost10 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost11=0 or olist.FOneItem.FTotmaechulProfit11=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit11 / olist.FOneItem.FTotItemcost11 * 100)&"%" End If %></td>
	<td align="right"><% If olist.FOneItem.FTotItemcost12=0 or olist.FOneItem.FTotmaechulProfit12=0 Then response.write "0%" Else response.write formatnumber(olist.FOneItem.FTotmaechulProfit12 / olist.FOneItem.FTotItemcost12 * 100)&"%" End If %></td>
	<td align="right"><% If TotTarget1to12=0 or TotProfit1to12=0 Then response.write "0%" Else response.write formatnumber(TotMaechul1to12 / TotItemcost1to12 * 100)&"%" End If %></td>
	<td align="right"></td>
</tr>
<%
End If
%>
</table>
<% SET olist = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 온라인목표매출관리
' Hieditor : 2016.05.27 김진영 생성
'###########################################################
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
Dim page, i, Depth1Code, Depth1Name, j, k
Dim yyyy, viewList, strParam, viewGubun

maxDepth = 1
dispCate	= requestCheckVar(Request("disp"),16)
page 		= requestCheckVar(Request("page"),2)
yyyy		= requestCheckVar(Request("yyyy"),4)
viewList	= requestCheckVar(Request("viewList"),1)
viewGubun	= requestCheckVar(Request("gubun"),3)

If yyyy = "" Then yyyy = LEFT(date(), 4)
If page = "" Then page = 1
If viewGubun = "" Then viewGubun = "ON"

strParam = "?menupos="&menupos&"&disp="&dispCate&"&page="&page&"&yyyy="&yyyy
SET olist = new CMDCategory
	olist.FPageSize		= 500
	olist.FCurrPage		= 1
	olist.FRectCatecode	= dispCate
	olist.FRectYyyy		= yyyy
	olist.FRectGubun	= viewGubun
	olist.getMDPurposeRegedListNew
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
			<option value="A" selected>목표+실적보기
			<option value="B" >목표보기
			<option value="C" >실적보기
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
<tr align="center" bgcolor="#DADADA">
	<td width="8%"><%= Chkiif(viewGubun= "ON", "온라인", "오프라인") %></td>
    <td colspan="14">월</td>
    <td rowspan="2" width="8%">합계</td>
    <td rowspan="2" width="8%">비중</td>
</tr>
<tr align="center" bgcolor="#DADADA">
	<td>카테고리</td>
	<td width="5%">&nbsp;</td>
	<td width="8%">구분</td>
	<% For i = 1 to 12 %>
    <td width="120"><%= i %>월</td>
	<% Next %>
</tr>
<tr><td colspan="17" bgcolor="000000" height="1"></td></tr>
<!-- #include file="./index_inc_total.asp" -->
<%
Dim isreg1, isreg2, isreg3, isreg4, isreg5, isreg6, isreg7, isreg8, isreg9, isreg10, isreg11, isreg12
Dim profitper1, profitper2, profitper3, profitper4, profitper5, profitper6, profitper7, profitper8, profitper9, profitper10, profitper11, profitper12, totProfitper
Dim itemcostper1, itemcostper2, itemcostper3, itemcostper4, itemcostper5, itemcostper6, itemcostper7, itemcostper8, itemcostper9, itemcostper10, itemcostper11, itemcostper12
Dim maechulper1, maechulper2, maechulper3, maechulper4, maechulper5, maechulper6, maechulper7, maechulper8, maechulper9, maechulper10, maechulper11, maechulper12, totMaechulper, totMaechulProfitper
Dim realPer1, realPer2, realPer3, realPer4, realPer5, realPer6, realPer7, realPer8, realPer9, realPer10, realPer11, realPer12, totrealper

If olist.FResultCount > 0 Then
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

		If olist.FItemList(i).FSumTartgetMoney = 0 OR olist.FItemList(i).FSumProfitMoney = 0 Then totProfitper = "0%" Else totProfitper = formatnumber(olist.FItemList(i).FSumProfitMoney / olist.FItemList(i).FSumTartgetMoney * 100, 1) & "%" End If
		If olist.FItemList(i).FSumItemcost = 0 OR olist.FItemList(i).FSumTartgetMoney = 0 Then totMaechulper = "0%" Else totMaechulper = formatnumber(olist.FItemList(i).FSumItemcost / olist.FItemList(i).FSumTartgetMoney * 100, 1) & "%" End If
		If olist.FItemList(i).FSumProfitMoney = 0 OR olist.FItemList(i).FSumMaechulProfit = 0 Then totMaechulProfitper = "0%" Else totMaechulProfitper = formatnumber(olist.FItemList(i).FSumMaechulProfit / olist.FItemList(i).FSumProfitmoney * 100, 1) & "%" End If
		If olist.FItemList(i).FSumMaechulProfit = 0 OR olist.FItemList(i).FSumItemcost = 0 Then totrealper = "0%" Else totrealper = formatnumber(olist.FItemList(i).FSumMaechulProfit / olist.FItemList(i).FSumItemcost * 100, 1) & "%" End If

%>
<!--													매출목표													-->
<tr align="center" bgcolor="FFFFFF" height="25">
	<td rowspan="13" <%= Chkiif(olist.FItemList(i).FDepth = 1, "bgcolor='#FFFFD7'","bgcolor='FFFFFF'") %> ><%= olist.FItemList(i).FCatename %></td>
	<td rowspan="3" >목표</td>
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
		<strong><% If olist.FItemList(i).FSumTartgetMoney = 0 or vAllTotTarget = 0 Then rw "0%" Else rw formatnumber(olist.FItemList(i).FSumTartgetMoney / vAllTotTarget * 100)&"%" End If %></strong>
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
	<td align="right"><% If olist.FItemList(i).FSumProfitMoney = 0 or vAllTotProfit = 0 Then rw "0%" Else rw formatnumber(olist.FItemList(i).FSumProfitMoney / vAllTotProfit * 100)&"%" End If %></td>
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

<!-- 전년실적 //-->
<%
	If isArray(vArr) Then
		For k = 0 To UBound(vArr,2)

		If CStr(olist.FItemList(i).FCateCode) = CStr(vArr(0,k)) Then
%>
		<tr align="center"  bgcolor="#F0F0F0" height="25">
			<td rowspan="5">전년실적</td>
			<td>구매총액</td>
			<td align="right"><%= NullOrCurrFormat(vArr(6,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(10,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(14,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(18,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(22,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(26,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(30,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(34,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(38,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(42,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(46,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(50,k)) %></td>
			<td align="right"><strong><%= NullOrCurrFormat(vArr(54,k)) %></strong></td>
			<td align="right">
				<strong><% If vArr(54,k) = 0 or vLastAllTotItemCost = 0 Then rw "0%" Else rw formatnumber(vArr(54,k) / vLastAllTotItemCost * 100)&"%" End If %></strong>
			</td>
		</tr>
		<tr align="center"  bgcolor="#F0F0F0" height="25">
			<td>달성율</td>
			<td align="right"><% If vArr(6,k)=0 OR vArr(4,k)=0 Then rw "0%" Else rw formatnumber((vArr(6,k)/vArr(4,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(10,k)=0 OR vArr(8,k)=0 Then rw "0%" Else rw formatnumber((vArr(10,k)/vArr(8,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(14,k)=0 OR vArr(12,k)=0 Then rw "0%" Else rw formatnumber((vArr(14,k)/vArr(12,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(18,k)=0 OR vArr(16,k)=0 Then rw "0%" Else rw formatnumber((vArr(18,k)/vArr(16,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(22,k)=0 OR vArr(20,k)=0 Then rw "0%" Else rw formatnumber((vArr(22,k)/vArr(20,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(26,k)=0 OR vArr(24,k)=0 Then rw "0%" Else rw formatnumber((vArr(26,k)/vArr(24,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(30,k)=0 OR vArr(28,k)=0 Then rw "0%" Else rw formatnumber((vArr(30,k)/vArr(28,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(34,k)=0 OR vArr(32,k)=0 Then rw "0%" Else rw formatnumber((vArr(34,k)/vArr(32,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(38,k)=0 OR vArr(36,k)=0 Then rw "0%" Else rw formatnumber((vArr(38,k)/vArr(36,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(42,k)=0 OR vArr(40,k)=0 Then rw "0%" Else rw formatnumber((vArr(42,k)/vArr(40,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(46,k)=0 OR vArr(44,k)=0 Then rw "0%" Else rw formatnumber((vArr(46,k)/vArr(44,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(50,k)=0 OR vArr(48,k)=0 Then rw "0%" Else rw formatnumber((vArr(50,k)/vArr(48,k))*100, 1) & "%" End If %></td>
			<td align="right"><% IF vArr(54,k)=0 OR vArr(52,k)=0 Then rw "0%" Else rw formatnumber((vArr(54,k)/vArr(52,k))*100, 1) & "%" End If %></td>
			<td align="right"></td>
		</tr>
		<tr align="center"  bgcolor="#F0F0F0" height="25">
			<td>수익</td>
			<td align="right"><%= NullOrCurrFormat(vArr(7,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(11,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(15,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(19,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(23,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(27,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(31,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(35,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(39,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(43,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(47,k)) %></td>
			<td align="right"><%= NullOrCurrFormat(vArr(51,k)) %></td>
			<td align="right"><strong><%= NullOrCurrFormat(vArr(55,k)) %></strong></td>
			<td align="right">
				<strong><% If vArr(55,k) = 0 or vLastAllTotMaechul = 0 Then rw "0%" Else rw formatnumber(vArr(55,k) / vLastAllTotMaechul * 100)&"%" End If %></strong>
			</td>
		</tr>
		<tr align="center"  bgcolor="#F0F0F0" height="25">
			<td>달성율</td>
			<td align="right"><% If vArr(7,k)=0 OR vArr(5,k)=0 Then rw "0%" Else rw formatnumber((vArr(7,k)/vArr(5,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(11,k)=0 OR vArr(9,k)=0 Then rw "0%" Else rw formatnumber((vArr(11,k)/vArr(9,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(15,k)=0 OR vArr(13,k)=0 Then rw "0%" Else rw formatnumber((vArr(15,k)/vArr(13,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(19,k)=0 OR vArr(17,k)=0 Then rw "0%" Else rw formatnumber((vArr(19,k)/vArr(17,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(23,k)=0 OR vArr(21,k)=0 Then rw "0%" Else rw formatnumber((vArr(23,k)/vArr(21,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(27,k)=0 OR vArr(25,k)=0 Then rw "0%" Else rw formatnumber((vArr(27,k)/vArr(25,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(31,k)=0 OR vArr(29,k)=0 Then rw "0%" Else rw formatnumber((vArr(31,k)/vArr(29,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(35,k)=0 OR vArr(33,k)=0 Then rw "0%" Else rw formatnumber((vArr(35,k)/vArr(33,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(39,k)=0 OR vArr(37,k)=0 Then rw "0%" Else rw formatnumber((vArr(39,k)/vArr(37,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(43,k)=0 OR vArr(41,k)=0 Then rw "0%" Else rw formatnumber((vArr(43,k)/vArr(41,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(47,k)=0 OR vArr(45,k)=0 Then rw "0%" Else rw formatnumber((vArr(47,k)/vArr(45,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(51,k)=0 OR vArr(49,k)=0 Then rw "0%" Else rw formatnumber((vArr(51,k)/vArr(49,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(53,k)=0 OR vArr(55,k)=0 Then rw "0%" Else rw formatnumber((vArr(55,k)/vArr(53,k))*100, 1) & "%" End If %></td>
			<td align="right"></td>
		</tr>
		<tr align="center"  bgcolor="#F0F0F0" height="25">
			<td>수익율</td>
			<td align="right"><% If vArr(7,k)=0 OR vArr(6,k)=0 Then rw "0%" Else rw formatnumber((vArr(7,k)/vArr(6,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(11,k)=0 OR vArr(10,k)=0 Then rw "0%" Else rw formatnumber((vArr(11,k)/vArr(10,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(15,k)=0 OR vArr(14,k)=0 Then rw "0%" Else rw formatnumber((vArr(15,k)/vArr(14,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(19,k)=0 OR vArr(18,k)=0 Then rw "0%" Else rw formatnumber((vArr(19,k)/vArr(18,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(23,k)=0 OR vArr(22,k)=0 Then rw "0%" Else rw formatnumber((vArr(23,k)/vArr(22,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(27,k)=0 OR vArr(26,k)=0 Then rw "0%" Else rw formatnumber((vArr(27,k)/vArr(26,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(31,k)=0 OR vArr(30,k)=0 Then rw "0%" Else rw formatnumber((vArr(31,k)/vArr(30,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(35,k)=0 OR vArr(34,k)=0 Then rw "0%" Else rw formatnumber((vArr(35,k)/vArr(34,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(39,k)=0 OR vArr(38,k)=0 Then rw "0%" Else rw formatnumber((vArr(39,k)/vArr(38,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(43,k)=0 OR vArr(42,k)=0 Then rw "0%" Else rw formatnumber((vArr(43,k)/vArr(42,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(47,k)=0 OR vArr(46,k)=0 Then rw "0%" Else rw formatnumber((vArr(47,k)/vArr(46,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(51,k)=0 OR vArr(50,k)=0 Then rw "0%" Else rw formatnumber((vArr(51,k)/vArr(50,k))*100, 1) & "%" End If %></td>
			<td align="right"><% If vArr(55,k)=0 OR vArr(54,k)=0 Then rw "0%" Else rw formatnumber((vArr(55,k)/vArr(54,k))*100, 1) & "%" End If %></td>
			<td align="right"></td>
		</tr>
<%
			vCateExist = "o"
			Exit For
		End If
		Next
	End If

If vCateExist = "x" Then
	Call sbCateNotExistHTML()
End IF %>
<!-- //-->

<tr align="center"  bgcolor="#FFE3E2" height="25">
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
		<strong><% If olist.FItemList(i).FSumItemcost = 0 or vAllTotItemCost = 0 Then rw "0%" Else rw formatnumber(olist.FItemList(i).FSumItemcost / vAllTotItemCost * 100)&"%" End If %></strong>
	</td>
</tr>

<tr align="center"  bgcolor="#FFE3E2" height="25">
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
<tr align="center"  bgcolor="#FFE3E2" height="25">
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
		<strong><% If olist.FItemList(i).FSumMaechulProfit = 0 or vAllTotMaechul = 0 Then rw "0%" Else rw formatnumber(olist.FItemList(i).FSumMaechulProfit / vAllTotMaechul * 100)&"%" End If %></strong>
	</td>
</tr>
<tr align="center"  bgcolor="#FFE3E2" height="25">
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
<tr align="center"  bgcolor="#FFE3E2" height="25">
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
<tr><td colspan="17" bgcolor="000000" height="1"></td></tr>
<%
		vCateExist = "x"
	Next
%>

<tr><td colspan="17" bgcolor="000000" height="1"></td></tr>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td rowspan="13" bgcolor="#D3FFFF"><strong>TOTAL</strong></td>
	<td rowspan="3" >목표</td>
	<td>월별매출목표</td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget1) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget2) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget3) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget4) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget5) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget6) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget7) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget8) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget9) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget10) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget11) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vTotalTarget12) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vAllTotTarget) %></strong></td>
	<td align="right"><strong><%= Chkiif(vAllTotTarget = 0, "0%", "100%") %></strong></td>
</tr>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td>월별수익목표</td>
	<td align="right"><%= NullOrCurrFormat(vProfit1) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit2) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit3) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit4) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit5) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit6) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit7) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit8) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit9) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit10) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit11) %></td>
	<td align="right"><%= NullOrCurrFormat(vProfit12) %></td>
	<td align="right"><%= NullOrCurrFormat(vAllTotProfit) %></td>
	<td align="right"><%= Chkiif(vAllTotProfit = 0, "0%", "100%") %></td>
</tr>
<tr align="center" bgcolor="FFFFFF" height="25">
	<td>월별수익율</td>
	<td align="right"><% If vTotalTarget1=0 or vProfit1=0 Then rw "0%" Else rw formatnumber(vProfit1 / vTotalTarget1 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget2=0 or vProfit2=0 Then rw "0%" Else rw formatnumber(vProfit2 / vTotalTarget2 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget3=0 or vProfit3=0 Then rw "0%" Else rw formatnumber(vProfit3 / vTotalTarget3 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget4=0 or vProfit4=0 Then rw "0%" Else rw formatnumber(vProfit4 / vTotalTarget4 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget5=0 or vProfit5=0 Then rw "0%" Else rw formatnumber(vProfit5 / vTotalTarget5 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget6=0 or vProfit6=0 Then rw "0%" Else rw formatnumber(vProfit6 / vTotalTarget6 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget7=0 or vProfit7=0 Then rw "0%" Else rw formatnumber(vProfit7 / vTotalTarget7 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget8=0 or vProfit8=0 Then rw "0%" Else rw formatnumber(vProfit8 / vTotalTarget8 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget9=0 or vProfit9=0 Then rw "0%" Else rw formatnumber(vProfit9 / vTotalTarget9 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget10=0 or vProfit10=0 Then rw "0%" Else rw formatnumber(vProfit10 / vTotalTarget10 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget11=0 or vProfit11=0 Then rw "0%" Else rw formatnumber(vProfit11 / vTotalTarget11 * 100)&"%" End If %></td>
	<td align="right"><% If vTotalTarget12=0 or vProfit12=0 Then rw "0%" Else rw formatnumber(vProfit12 / vTotalTarget12 * 100)&"%" End If %></td>
	<td align="right"><% If vAllTotTarget=0 or vAllTotProfit=0 Then rw "0%" Else rw formatnumber(vAllTotProfit / vAllTotTarget * 100)&"%" End If %></td>
	<td align="right"></td>
</tr>

<%
	If isArray(vArr) Then
%>
	<tr align="center" bgcolor="#F0F0F0" height="25">
		<td rowspan="5" >실적</td>
		<td>월별매출실적</td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost1) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost2) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost3) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost4) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost5) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost6) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost7) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost8) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost9) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost10) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost11) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastItemcost12) %></strong></td>
		<td align="right"><strong><%= NullOrCurrFormat(vLastAllTotItemCost) %></strong></td>
		<td align="right"><strong><%= Chkiif(vLastAllTotItemCost = 0, "0%", "100%") %></strong></td>
	</tr>
	<tr align="center" bgcolor="#F0F0F0" height="25">
		<td>월별실적매출달성율</td>
		<td align="right"><% If vLastItemcost1=0 or vLastTotalTarget1=0 Then rw "0%" Else rw formatnumber(vLastItemcost1 / vLastTotalTarget1 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost2=0 or vLastTotalTarget2=0 Then rw "0%" Else rw formatnumber(vLastItemcost2 / vLastTotalTarget2 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost3=0 or vLastTotalTarget3=0 Then rw "0%" Else rw formatnumber(vLastItemcost3 / vLastTotalTarget3 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost4=0 or vLastTotalTarget4=0 Then rw "0%" Else rw formatnumber(vLastItemcost4 / vLastTotalTarget4 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost5=0 or vLastTotalTarget5=0 Then rw "0%" Else rw formatnumber(vLastItemcost5 / vLastTotalTarget5 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost6=0 or vLastTotalTarget6=0 Then rw "0%" Else rw formatnumber(vLastItemcost6 / vLastTotalTarget6 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost7=0 or vLastTotalTarget7=0 Then rw "0%" Else rw formatnumber(vLastItemcost7 / vLastTotalTarget7 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost8=0 or vLastTotalTarget8=0 Then rw "0%" Else rw formatnumber(vLastItemcost8 / vLastTotalTarget8 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost9=0 or vLastTotalTarget9=0 Then rw "0%" Else rw formatnumber(vLastItemcost9 / vLastTotalTarget9 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost10=0 or vLastTotalTarget10=0 Then rw "0%" Else rw formatnumber(vLastItemcost10 / vLastTotalTarget10 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost11=0 or vLastTotalTarget11=0 Then rw "0%" Else rw formatnumber(vLastItemcost11 / vLastTotalTarget11 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost12=0 or vLastTotalTarget12=0 Then rw "0%" Else rw formatnumber(vLastItemcost12 / vLastTotalTarget12 * 100)&"%" End If %></td>
		<td align="right"><% If vLastAllTotItemCost=0 or vLastAllTotTarget=0 Then rw "0%" Else rw formatnumber(vLastAllTotItemCost / vLastAllTotTarget * 100)&"%" End If %></td>
		<td align="right"></td>
	</tr>
	<tr align="center" bgcolor="#F0F0F0" height="25">
		<td>월별수익실적</td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf1) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf2) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf3) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf4) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf5) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf6) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf7) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf8) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf9) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf10) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf11) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastMaechulProf12) %></td>
		<td align="right"><%= NullOrCurrFormat(vLastAllTotMaechul) %></td>
		<td align="right"><%= Chkiif(vLastAllTotMaechul = 0, "0%", "100%") %></td>
	</tr>
	<tr align="center" bgcolor="#F0F0F0" height="25">
		<td>월별실적수익달성율</td>
		<td align="right"><% If vLastProfit1=0 or vLastMaechulProf1=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf1 / vLastProfit1 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit2=0 or vLastMaechulProf2=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf2 / vLastProfit2 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit3=0 or vLastMaechulProf3=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf3 / vLastProfit3 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit4=0 or vLastMaechulProf4=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf4 / vLastProfit4 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit5=0 or vLastMaechulProf5=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf5 / vLastProfit5 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit6=0 or vLastMaechulProf6=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf6 / vLastProfit6 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit7=0 or vLastMaechulProf7=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf7 / vLastProfit7 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit8=0 or vLastMaechulProf8=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf8 / vLastProfit8 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit9=0 or vLastMaechulProf9=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf9 / vLastProfit9 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit10=0 or vLastMaechulProf10=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf10 / vLastProfit10 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit11=0 or vLastMaechulProf11=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf11 / vLastProfit11 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastProfit12=0 or vLastMaechulProf12=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf12 / vLastProfit12 * 100)&"%" End If %></td></td>
		<td align="right"><% If vLastAllTotProfit=0 or vLastAllTotMaechul=0 Then rw "0%" Else rw formatnumber(vLastAllTotMaechul / vLastAllTotProfit * 100)&"%" End If %></td>
		<td align="right"></td>
	</tr>
	<tr align="center" bgcolor="#F0F0F0" height="25">
		<td>월별실적수익율</td>
		<td align="right"><% If vLastItemcost1=0 or vLastMaechulProf1=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf1 / vLastItemcost1 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost2=0 or vLastMaechulProf2=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf2 / vLastItemcost2 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost3=0 or vLastMaechulProf3=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf3 / vLastItemcost3 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost4=0 or vLastMaechulProf4=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf4 / vLastItemcost4 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost5=0 or vLastMaechulProf5=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf5 / vLastItemcost5 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost6=0 or vLastMaechulProf6=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf6 / vLastItemcost6 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost7=0 or vLastMaechulProf7=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf7 / vLastItemcost7 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost8=0 or vLastMaechulProf8=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf8 / vLastItemcost8 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost9=0 or vLastMaechulProf9=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf9 / vLastItemcost9 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost10=0 or vLastMaechulProf10=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf10 / vLastItemcost10 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost11=0 or vLastMaechulProf11=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf11 / vLastItemcost11 * 100)&"%" End If %></td>
		<td align="right"><% If vLastItemcost12=0 or vLastMaechulProf12=0 Then rw "0%" Else rw formatnumber(vLastMaechulProf12 / vLastItemcost12 * 100)&"%" End If %></td>
		<td align="right"><% If vLastAllTotMaechul=0 or vLastAllTotItemCost=0 Then rw "0%" Else rw formatnumber(vLastAllTotMaechul / vLastAllTotItemCost * 100)&"%" End If %></td>
		<td align="right"></td>
	</tr>
<%
	End If
%>

<tr align="center" bgcolor="#FFE3E2" height="25">
	<td rowspan="5" >실적</td>
	<td>월별매출실적</td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost1) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost2) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost3) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost4) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost5) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost6) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost7) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost8) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost9) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost10) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost11) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vItemcost12) %></strong></td>
	<td align="right"><strong><%= NullOrCurrFormat(vAllTotItemCost) %></strong></td>
	<td align="right"><strong><%= Chkiif(vAllTotItemCost = 0, "0%", "100%") %></strong></td>
</tr>
<tr align="center" bgcolor="#FFE3E2" height="25">
	<td>월별실적매출달성율</td>
	<td align="right"><% If vItemcost1=0 or vTotalTarget1=0 Then rw "0%" Else rw formatnumber(vItemcost1 / vTotalTarget1 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost2=0 or vTotalTarget2=0 Then rw "0%" Else rw formatnumber(vItemcost2 / vTotalTarget2 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost3=0 or vTotalTarget3=0 Then rw "0%" Else rw formatnumber(vItemcost3 / vTotalTarget3 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost4=0 or vTotalTarget4=0 Then rw "0%" Else rw formatnumber(vItemcost4 / vTotalTarget4 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost5=0 or vTotalTarget5=0 Then rw "0%" Else rw formatnumber(vItemcost5 / vTotalTarget5 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost6=0 or vTotalTarget6=0 Then rw "0%" Else rw formatnumber(vItemcost6 / vTotalTarget6 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost7=0 or vTotalTarget7=0 Then rw "0%" Else rw formatnumber(vItemcost7 / vTotalTarget7 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost8=0 or vTotalTarget8=0 Then rw "0%" Else rw formatnumber(vItemcost8 / vTotalTarget8 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost9=0 or vTotalTarget9=0 Then rw "0%" Else rw formatnumber(vItemcost9 / vTotalTarget9 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost10=0 or vTotalTarget10=0 Then rw "0%" Else rw formatnumber(vItemcost10 / vTotalTarget10 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost11=0 or vTotalTarget11=0 Then rw "0%" Else rw formatnumber(vItemcost11 / vTotalTarget11 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost12=0 or vTotalTarget12=0 Then rw "0%" Else rw formatnumber(vItemcost12 / vTotalTarget12 * 100)&"%" End If %></td>
	<td align="right"><% If vAllTotItemCost=0 or vAllTotTarget=0 Then rw "0%" Else rw formatnumber(vAllTotItemCost / vAllTotTarget * 100)&"%" End If %></td>
	<td align="right"></td>
</tr>
<tr align="center" bgcolor="#FFE3E2" height="25">
	<td>월별수익실적</td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf1) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf2) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf3) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf4) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf5) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf6) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf7) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf8) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf9) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf10) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf11) %></td>
	<td align="right"><%= NullOrCurrFormat(vMaechulProf12) %></td>
	<td align="right"><%= NullOrCurrFormat(vAllTotMaechul) %></td>
	<td align="right"><%= Chkiif(vAllTotMaechul = 0, "0%", "100%") %></td>
</tr>
<tr align="center" bgcolor="#FFE3E2" height="25">
	<td>월별실적수익달성율</td>
	<td align="right"><% If vProfit1=0 or vMaechulProf1=0 Then rw "0%" Else rw formatnumber(vMaechulProf1 / vProfit1 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit2=0 or vMaechulProf2=0 Then rw "0%" Else rw formatnumber(vMaechulProf2 / vProfit2 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit3=0 or vMaechulProf3=0 Then rw "0%" Else rw formatnumber(vMaechulProf3 / vProfit3 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit4=0 or vMaechulProf4=0 Then rw "0%" Else rw formatnumber(vMaechulProf4 / vProfit4 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit5=0 or vMaechulProf5=0 Then rw "0%" Else rw formatnumber(vMaechulProf5 / vProfit5 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit6=0 or vMaechulProf6=0 Then rw "0%" Else rw formatnumber(vMaechulProf6 / vProfit6 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit7=0 or vMaechulProf7=0 Then rw "0%" Else rw formatnumber(vMaechulProf7 / vProfit7 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit8=0 or vMaechulProf8=0 Then rw "0%" Else rw formatnumber(vMaechulProf8 / vProfit8 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit9=0 or vMaechulProf9=0 Then rw "0%" Else rw formatnumber(vMaechulProf9 / vProfit9 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit10=0 or vMaechulProf10=0 Then rw "0%" Else rw formatnumber(vMaechulProf10 / vProfit10 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit11=0 or vMaechulProf11=0 Then rw "0%" Else rw formatnumber(vMaechulProf11 / vProfit11 * 100)&"%" End If %></td></td>
	<td align="right"><% If vProfit12=0 or vMaechulProf12=0 Then rw "0%" Else rw formatnumber(vMaechulProf12 / vProfit12 * 100)&"%" End If %></td></td>
	<td align="right"><% If vAllTotProfit=0 or vAllTotMaechul=0 Then rw "0%" Else rw formatnumber(vAllTotMaechul / vAllTotProfit * 100)&"%" End If %></td>
	<td align="right"></td>
</tr>
<tr align="center" bgcolor="#FFE3E2" height="25">
	<td>월별실적수익율</td>
	<td align="right"><% If vItemcost1=0 or vMaechulProf1=0 Then rw "0%" Else rw formatnumber(vMaechulProf1 / vItemcost1 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost2=0 or vMaechulProf2=0 Then rw "0%" Else rw formatnumber(vMaechulProf2 / vItemcost2 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost3=0 or vMaechulProf3=0 Then rw "0%" Else rw formatnumber(vMaechulProf3 / vItemcost3 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost4=0 or vMaechulProf4=0 Then rw "0%" Else rw formatnumber(vMaechulProf4 / vItemcost4 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost5=0 or vMaechulProf5=0 Then rw "0%" Else rw formatnumber(vMaechulProf5 / vItemcost5 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost6=0 or vMaechulProf6=0 Then rw "0%" Else rw formatnumber(vMaechulProf6 / vItemcost6 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost7=0 or vMaechulProf7=0 Then rw "0%" Else rw formatnumber(vMaechulProf7 / vItemcost7 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost8=0 or vMaechulProf8=0 Then rw "0%" Else rw formatnumber(vMaechulProf8 / vItemcost8 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost9=0 or vMaechulProf9=0 Then rw "0%" Else rw formatnumber(vMaechulProf9 / vItemcost9 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost10=0 or vMaechulProf10=0 Then rw "0%" Else rw formatnumber(vMaechulProf10 / vItemcost10 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost11=0 or vMaechulProf11=0 Then rw "0%" Else rw formatnumber(vMaechulProf11 / vItemcost11 * 100)&"%" End If %></td>
	<td align="right"><% If vItemcost12=0 or vMaechulProf12=0 Then rw "0%" Else rw formatnumber(vMaechulProf12 / vItemcost12 * 100)&"%" End If %></td>
	<td align="right"><% If vAllTotMaechul=0 or vAllTotItemCost=0 Then rw "0%" Else rw formatnumber(vAllTotMaechul / vAllTotItemCost * 100)&"%" End If %></td>
	<td align="right"></td>
</tr>
<tr><td colspan="17" bgcolor="000000" height="1"></td></tr>
<%
End If
%>
</table>
<% SET olist = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
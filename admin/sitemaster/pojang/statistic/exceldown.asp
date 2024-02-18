<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftStatisticCls.asp" -->
<%
	Dim i, sDate, eDate, cStat, vTab, vArrTalk, vArrDay, vArrShop
	Dim vTotPC, vTotMob, vTotApp, vTotPC1, vTotPC2, vTotMob1, vTotMob2, vTotApp1, vTotApp2, vTotTotalPC, vTotTotalMob, vTotTotalApp
	sDate = NullFillWith(request("sDate"),DateAdd("d",-10,date()))
	eDate = NullFillWith(request("eDate"),date())
	vTab = NullFillWith(request("tab"),1)
	

	Response.ContentType = "application/x-msexcel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=pojang_" & Replace(sDate,"-","") & "to" & Replace(eDate,"-","") & ".xls"

	SET cStat = New CgiftStat_list
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		cStat.sbPojangStatDaily
%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<body>
<% If vTab = "1" Then %>
<table border="1">
	<tr>
		<th rowspan="2"><div>날짜</div></th>
		<th rowspan="2"><div>채널</div></th>
		<th colspan="2"><div>포장 수량</div></th>
		<th rowspan="2"><div>포장수량 합계</div></th>
		<th rowspan="2"><div>가격</div></th>
	</tr>
	<tr>
		<th><div>상품 1개</div></th>
		<th><div>상품 2개 이상</div></th>
	</tr>
	<%
		for i=0 to cStat.FResultCount - 1
		
		vTotPC1		= vTotPC1 + cStat.FItemList(i).FPW1
		vTotPC2 	= vTotPC2 + cStat.FItemList(i).FPW2
		vTotMob1	= vTotMob1 + cStat.FItemList(i).FPM1
		vTotMob2	= vTotMob2 + cStat.FItemList(i).FPM2
		vTotApp1	= vTotApp1 + cStat.FItemList(i).FPA1
		vTotApp2	= vTotApp2 + cStat.FItemList(i).FPA2
		
		vTotPC		= cStat.FItemList(i).FPW1 + cStat.FItemList(i).FPW2
		vTotMob 	= cStat.FItemList(i).FPM1 + cStat.FItemList(i).FPM2
		vTotApp 	= cStat.FItemList(i).FPA1 + cStat.FItemList(i).FPA2
		
		vTotTotalPC		= vTotTotalPC + vTotPC
		vTotTotalMob	= vTotTotalMob + vTotMob
		vTotTotalApp	= vTotTotalApp + vTotApp
	%>
		<tr>
			<td><%= cStat.FItemList(i).FDate %></td>
			<td>PC(W)</td>
			<td><%= cStat.FItemList(i).FPW1 %></td>
			<td><%= cStat.FItemList(i).FPW2 %></td>
			<td><%= vTotPC %></td>
			<td><%= FormatNumber((vTotPC*2000),0) %></td>
		</tr>
		<tr>
			<td><%= cStat.FItemList(i).FDate %></td>
			<td>모바일웹(M)</td>
			<td><%= cStat.FItemList(i).FPM1 %></td>
			<td><%= cStat.FItemList(i).FPM2 %></td>
			<td><%= vTotMob %></td>
			<td><%= FormatNumber((vTotMob*2000),0) %></td>
		</tr>
		<tr>
			<td><%= cStat.FItemList(i).FDate %></td>
			<td>모바일앱(A)</td>
			<td><%= cStat.FItemList(i).FPA1 %></td>
			<td><%= cStat.FItemList(i).FPA2 %></td>
			<td><%= vTotApp %></td>
			<td><%= FormatNumber((vTotApp*2000),0) %></td>
		</tr>
	<%
		next
	%>
		<tr>
			<td colspan="2" class="bgGy1"><strong>합계</strong></td>
			<td class="bgGy1"><strong><%= FormatNumber((vTotPC1 + vTotMob1 + vTotApp1),0) %></strong></td>
			<td class="bgGy1"><strong><%= FormatNumber((vTotPC2 + vTotMob2 + vTotApp2),0) %></strong></td>
			<td class="bgGy1"><strong><%= FormatNumber((vTotTotalPC + vTotTotalMob + vTotTotalApp),0) %></strong></td>
			<td class="bgGy1"><strong><%= FormatNumber(((vTotTotalPC + vTotTotalMob + vTotTotalApp)*2000),0) %></strong></td>
		</tr>
</table>
<% End If %>
<% SET cStat = Nothing %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
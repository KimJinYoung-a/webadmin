<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 상담
' History : 2015.05.27 이상구 생성
'			2016.03.25 한용민 수정(문의분야 모두 DB화 시킴)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/customer_board_reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, i, nowdateStr, startdateStr, nextdateStr, fromDate, toDate, tmpDate
dim flashvar, replyUser
dim tmpSum, tmpSumArr, tmpChargeId, tmpChargeIdArr, IsNewRow
Dim trEndStart, prevRow, sumCnt, sumTotCnt, totCnt(50), divNow, divLen
dim siteGubun
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	replyUser	= req("replyUser","")
	siteGubun = request("siteGubun")

divLen = 0

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
end if

'시작일 재구성
startdateStr = CStr(DateSerial(yyyy1,mm1,dd1))
yyyy1 = year(startdateStr)
mm1 = Num2Str(month(startdateStr),2,"0","R")
dd1 = Num2Str(day(startdateStr),2,"0","R")

'종료일 재구성
nextdateStr = CStr(DateSerial(yyyy2,mm2,dd2))
yyyy2 = year(nextdateStr)
mm2 = Num2Str(month(nextdateStr),2,"0","R")
dd2 = Num2Str(day(nextdateStr),2,"0","R")

dim oreport, rs
set oreport = new CReportMaster
	oreport.FRectStart = startdateStr
	oreport.FRectEnd =  nextdateStr
	oreport.FRectSiteGubun =  siteGubun
	'rs = oreport.getQnaDivReport(replyUser)
	rs = oreport.getQnaReport(replyUser)

divLen = oreport.getQnaDivcount()

flashvar = "startdate=" + startdateStr + "&enddate=" + nextdateStr

%>
<!-- 샘플코드 : http://docs.fusioncharts.com/tutorial-getting-started-your-first-charts-building-your-first-chart.html -->
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script type="text/javascript">

function popQna(qaDiv, replyDate, replyUser){
	if (replyDate)
	{
		var replyDate1 = replyDate;
		var replyDate2 = replyDate;
	}
	else
	{
		var f = document.frm;
		var replyDate1 = f.yyyy1.value + "-" + f.mm1.value + "-" + f.dd1.value;
		var replyDate2 = f.yyyy2.value + "-" + f.mm2.value + "-" + f.dd2.value;
	}
	var url = "/cscenter/board/cscenter_qna_board_list.asp?ckReplyDateDefault=on&qaDiv=" + qaDiv + "&writeid=" + replyUser + "&replyDate1=" + replyDate1 + "&replyDate2=" + replyDate2;
	var popwin = window.open(url,"PopMyQnaList","width=1024, height=768, left=50, top=50, scrollbars=yes, resizable=yes, status=yes");
	popwin.focus();
}

FusionCharts.ready(function(){
      var revenueChart = new FusionCharts({
        "type": "pie3d",						// pie3d, column2d
        "renderAt": "chartContainer",
        "width": "800",
        "height": "400",
        "dataFormat": "json",
        "dataSource": {
          "chart": {
              	"caption": "1:1상담 통계",
              	"subCaption": "",
              	"xAxisName": "Month",
				"yAxisName": "Revenues (In USD)",
				"baseFont": "돋움",
				"baseFontSize": "12",
				"baseFontColor": "#000000",
				"outCnvBaseFont": "돋움",
				"outCnvBaseFontSize": "12",
				"outCnvBaseFontColor": "#000000",
				"theme": "fint"
			},
			"data": [<%
If IsArray(rs) Then
	tmpSum = ""
	tmpSumArr = ""
	tmpChargeId = ""
	tmpChargeIdArr = ""
	For i = 0 To UBound(rs,2)
		if (i = 0) then
			tmpChargeIdArr = rs(0,i)
			tmpChargeId = rs(0,i)
			tmpSum = CLng(rs(2,i))
		elseif (rs(0,i) <> tmpChargeId) then
			tmpChargeIdArr = tmpChargeIdArr & "," & rs(0,i)
			tmpChargeId = rs(0,i)
			if (tmpSumArr = "") then
				tmpSumArr = CStr(tmpSum)
			else
				tmpSumArr = tmpSumArr & "," & CStr(tmpSum)
			end if
			tmpSum = CLng(rs(2,i))
		else
			tmpSum = tmpSum + CLng(rs(2,i))
		end if
	next

	if (tmpSum <> "") then
		tmpSumArr = tmpSumArr & "," & CStr(tmpSum)
	end if

	tmpChargeIdArr = Split(tmpChargeIdArr, ",")
	tmpSumArr = Split(tmpSumArr, ",")

	'// { "label": "Jan", "value": "420000" }
	'// ,{ "label": "Jan", "value": "420000" }
	for i = 0 to UBound(tmpChargeIdArr)
		if (i > 0) then
			response.write ","
		end if

		response.write "{ ""label"": """ & tmpChargeIdArr(i) & """, ""value"": """ & tmpSumArr(i) & """ }"
	next
end if
				%>]
        }
    });

    revenueChart.render();
});

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간(답변일기준) : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		답변자ID : <input type="text" class="text" name="replyUser" value="<%=replyUser%>" size="12" maxlength="32">
		&nbsp;&nbsp;
		사이트 :
		<select class="select" name="siteGubun">
			<option></option>
			<option value="10x10" <%= CHKIIF(siteGubun="10x10", "selected", "") %>>10x10</option>
			<option value="extall" <%= CHKIIF(siteGubun="extall", "selected", "") %>>제휴몰</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		1시간 지연 데이터 입니다.
	</td>
	<td align="right"></td>
</tr>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100" height="25">
		<% If replyUser <> "" Then %>
			날짜
		<% Else %>
			답변자ID
		<% End If%>
	</td>
<%
If IsArray(rs) Then
	' Div명 출력
	For i=0 To UBound(rs,2)
		If i > 0 And rs(0,0) <> rs(0,i) Then
%>
		<td width="50" height="25"><b>합계</b></td>
	</tr>
	<tr>
	<%
			trEndStart = true
			Exit For
		End If
		response.write "<td align='center'>" & rs(3,i) & "</td>"
	Next

	If Not trEndStart Then
	%>
		<td width="50" height="25"><b>합계</b></td>
	</tr>
	<tr>
	<%
	End If

	' 카운트 출력
	For i=0 To UBound(rs,2)
		If prevRow <> rs(0,i) Then	' 첫번째 로우가 바뀔 때
	%>
		<td <% if (i > 0) then %>height="25"<% end if %> ><%=sumCnt%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%=rs(0,i)%></td>
		<%
			prevRow = rs(0,i)
			sumCnt	= 0
		End If

		If replyUser <> "" Then
			'/노출안함 일경우 색 변경
			if rs(4,i)="Y" then
				response.write "<td align='center'>"
			else
				response.write "<td align='center' bgcolor='"& adminColor("tabletop") &"'>"
			end if

			response.write "	<a href=""javascript:popQna('"&rs(1,i)&"','"&rs(0,i)&"','"&replyUser&"');"">" & rs(2,i) & "</a>"
			response.write "</td>"
		Else
			'/노출안함 일경우 색 변경
			if rs(4,i)="Y" then
				response.write "<td align='center'>"
			else
				response.write "<td align='center' bgcolor='"& adminColor("tabletop") &"'>"
			end if

			response.write "	<a href=""javascript:popQna('"&rs(1,i)&"','','"&rs(0,i)&"');"">" & rs(2,i) & "</a>"
			response.write "</td>"
		End If

		sumCnt	  = sumCnt + CDbl(rs(2,i))
		sumTotCnt = sumTotCnt + CDbl(rs(2,i))

		divNow = i Mod divLen
		totCnt(divNow) = totCnt(divNow) + CDbl(rs(2,i))

		Next
		%>
		<td height="25"><%=sumCnt%></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="25"><b>합계</b></td>
		<% For i=0 To divLen - 1 %>
		<td align="center"><a href="javascript:popQna('<%=rs(1,i)%>','','<%=replyUser%>');"><b><%=totCnt(i)%></b></a></td>
		<% Next %>
		<td><b><%=sumTotCnt%></b></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="25"><b>%</b></td>
		<% For i=0 To divLen - 1 %>
		<td align="center">
			<%
			If sumTotCnt > 0 Then
				response.write CInt(totCnt(i) * 100 / sumTotCnt)
			Else
				response.write "0"
			End If
			%>%
		</td>
		<% Next %>
		<td>100%</td>
	</tr>
<%
Else
	response.write "<td>검색결과가 없습니다.</td>" & vbCrLf
	response.write "</tr>" & vbCrLf
End If
%>
</table>
<br>
<table width="100%" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="left">
		<div id="chartContainer" align="center"></div>
	</td>
</tr>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  DATAMART >> Q&A 통계 
' History : 2017.06.12 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/datamart/qna/qna_summaryCls.asp"-->
<%
Dim sSdate, sEdate, sType, dayMode, summaryName, makerid, i
Dim arrList, oQna, TotalQnaAllCnt, TotalQnasecretYCnt, TotalQnasecretNCnt, TotalreplyYCnt, TotalreplyNCnt, TotalSumReplyDayCnt, Totalsnssend1Cnt, Totalsnssend2Cnt, Totalsnssend3Cnt, Totalsnssend4Cnt, Totalsnssend5Cnt
Dim TermQnaAllCnt, TermQnasecretYCnt, TermQnasecretNCnt, TermreplyYCnt, TermreplyNCnt, TermSumReplyDayCnt, Termsnssend1Cnt, Termsnssend2Cnt, Termsnssend3Cnt, Termsnssend4Cnt, Termsnssend5Cnt

sSdate			= requestCheckVar(request("iSD"),10)
sEdate			= requestCheckVar(request("iED"),10)
sType			= requestCheckVar(request("sType"),10)
dayMode			= requestCheckVar(request("dayMode"),1)
makerid			= requestCheckVar(request("makerid"),32)

If sSdate = "" Then sSdate = DateSerial(Year(Now()), Month(Now()), 1)
If sEdate = "" Then sEdate = Date()
If sType = "" Then sType = "category"
If dayMode = "" Then dayMode = "D"

If sType = "category" Then
	makerid = ""
	summaryName = "카테고리별 통계"
ElseIf sType = "brand" Then
	If makerid <> "" Then
		summaryName = "브랜드 통계"
	Else
		summaryName = "미답변 Top 20 브랜드 통계"
	End If
End If

SET oQna = new cQnaSummary
	oQna.FRectSdate			= sSdate
	oQna.FRectEdate			= sEdate
	oQna.FRectSType			= sType
	oQna.FRectDayMode		= dayMode
	oQna.FRectMakerid		= makerid
	arrList = oQna.fnQnaSummayReport

	TotalQnaAllCnt		= oQna.FTotalQnaAllCnt
	TotalQnasecretYCnt	= oQna.FTotalQnasecretYCnt
	TotalQnasecretNCnt	= oQna.FTotalQnasecretNCnt
	TotalreplyYCnt		= oQna.FTotalreplyYCnt
	TotalreplyNCnt		= oQna.FTotalreplyNCnt
	TotalSumReplyDayCnt	= oQna.FTotalSumReplyDayCnt
	Totalsnssend1Cnt	= oQna.FTotalsnssend1Cnt
	Totalsnssend2Cnt	= oQna.FTotalsnssend2Cnt
	Totalsnssend3Cnt	= oQna.FTotalsnssend3Cnt
	Totalsnssend4Cnt	= oQna.FTotalsnssend4Cnt
	Totalsnssend5Cnt	= oQna.FTotalsnssend5Cnt

	TermQnaAllCnt		= oQna.FTermQnaAllCnt		
	TermQnasecretYCnt	= oQna.FTermQnasecretYCnt	
	TermQnasecretNCnt	= oQna.FTermQnasecretNCnt	
	TermreplyYCnt		= oQna.FTermreplyYCnt		
	TermreplyNCnt		= oQna.FTermreplyNCnt		
	TermSumReplyDayCnt	= oQna.FTermSumReplyDayCnt	
	Termsnssend1Cnt		= oQna.FTermsnssend1Cnt		
	Termsnssend2Cnt		= oQna.FTermsnssend2Cnt		
	Termsnssend3Cnt		= oQna.FTermsnssend3Cnt		
	Termsnssend4Cnt		= oQna.FTermsnssend4Cnt		
	Termsnssend5Cnt		= oQna.FTermsnssend5Cnt		
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function fnAddType(v){
	$("#sTypeVal").val(v);
	$.ajax({
		type: "POST",
		url: "/admin/datamart/qna/qna_summary_ajax.asp",
		data: $("#frm").serialize(),
		dataType: "text",
		async: false,
		success : function(result){
		    $("#shopTBL").empty().html(result);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
function fnAddSumTerm(){
	$.ajax({
		type: "POST",
		url: "/admin/datamart/qna/qna_summary_sumajax.asp",
		data: $("#frm").serialize(),
		dataType: "text",
		async: false,
		success : function(result){
		    $("#shopTBL").empty().html(result);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
function fnChangeType(v){
	$("#sType").val(v);
	document.frm.submit();
}
function fnexcelDownlode(){
	document.frm2.action = "/admin/datamart/qna/qna_summary_excel.asp";
	document.frm2.submit();
}
function pop_itemqna(v){
	var popwin=window.open('/admin/datamart/qna/pop_itemqna.asp?sType=<%=sType%>&sTypeVal='+v+'&iSD=<%=sSdate%>&iED=<%=sEdate%>','notin','width=1200,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<form name="frm2" id="frm2">
<input type="hidden" name="sType2" id="sType2" value="<%= sType %>">
<input type="hidden" id="iSD2" name="iSD2" value="<%=sSdate%>"  />
<input type="hidden" id="iED2" name="iED2" value="<%=sEdate%>"  />
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" id="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sType" id="sType" value="<%= sType %>">
<input type="hidden" name="sTypeVal" id="sTypeVal" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table width="10%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td align="center" width="50%" <%= Chkiif(sType="category", "", "bgcolor='#E2E2E2'") %> onclick="fnChangeType('category');" style="cursor:pointer">카테고리</td>
			<td align="center" width="50%" <%= Chkiif(sType="brand", "", "bgcolor='#E2E2E2'") %>  onclick="fnChangeType('brand');" style="cursor:pointer">브랜드</td>
		</td>
		</table>
		&nbsp;&nbsp;
		기간 : 
		<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "iED", trigger    : "iED_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;
		<select class="select" name="dayMode">
			<option value="D" <%= Chkiif(dayMode = "D", "selected", "") %>>일별</option>
			<option value="M" <%= Chkiif(dayMode = "M", "selected", "") %>>월별</option>
			<option value="Y" <%= Chkiif(dayMode = "Y", "selected", "") %>>연별</option>	
		</select>
	<% If sType = "brand" Then %>
		&nbsp;브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
	<% End If %>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="15">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong><%= summaryName %></strong>
			</td>
			<td align="right">
				<input type="button" class="button" value="Excel 다운로드" onclick="fnexcelDownlode();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="7.1%"></td>
	<td width="7.1%">Q&A 전체(건)</td>
	<td width="7.1%">Q&A 공개(건)</td>
	<td width="7.1%">Q&A 비공개(건)</td>
	<td width="7.1%">답변(건)</td>
	<td width="7.1%">미답변(건)</td>
	<td width="7.1%">지연<br />미답변(건)</td>
	<td width="7.1%">답변율(%)</td>
	<td width="7.1%">평균<br />답변일(일)</td>
	<td width="7.1%">알림 문자<br />1차 발송(건)</td>
	<td width="7.1%">알림 문자<br />2차 발송(건)</td>
	<td width="7.1%">알림 문자<br />3차 발송(건)</td>
	<td width="7.1%">알림 문자<br />4차 발송(건)</td>
	<td width="7.1%">알림 문자<br />5차 발송(건)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td>누적 전체</td>
	<td><%= TotalQnaAllCnt %></td>
	<td><%= TotalQnasecretYCnt %></td>
	<td><%= TotalQnasecretNCnt %></td>
	<td><%= TotalreplyYCnt %></td>
	<td><%= TotalreplyNCnt %></td>
	<td><%= Totalsnssend1Cnt %></td>
	<td>
	<%
		If TotalreplyYCnt <> 0 Then
			response.write Round(TotalreplyYCnt / TotalQnaAllCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If TotalSumReplyDayCnt <> 0 Then
			response.write Round((TotalSumReplyDayCnt * 1.0 / TotalQnaAllCnt * 1.0), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= Totalsnssend1Cnt %></td>
	<td><%= Totalsnssend2Cnt %></td>
	<td><%= Totalsnssend3Cnt %></td>
	<td><%= Totalsnssend4Cnt %></td>
	<td><%= Totalsnssend5Cnt %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer;" onclick="fnAddSumTerm();">
	<td>기간 전체</td>
	<td><%= TermQnaAllCnt %></td>
	<td><%= TermQnasecretYCnt %></td>
	<td><%= TermQnasecretNCnt %></td>
	<td><%= TermreplyYCnt %></td>
	<td><%= TermreplyNCnt %></td>
	<td><%= Termsnssend1Cnt %></td>
	<td>
	<%
		If TermreplyYCnt <> 0 Then
			response.write Round(TermreplyYCnt / TermQnaAllCnt * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If TermSumReplyDayCnt <> 0 Then
			response.write Round((TermSumReplyDayCnt * 1.0 / TermQnaAllCnt * 1.0), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= Termsnssend1Cnt %></td>
	<td><%= Termsnssend2Cnt %></td>
	<td><%= Termsnssend3Cnt %></td>
	<td><%= Termsnssend4Cnt %></td>
	<td><%= Termsnssend5Cnt %></td>
</tr>
<tr align="center" bgcolor="#D2D2D2">
	<td colspan="15"></td>
</tr>
<% If IsArray(arrList) Then %>
<% For i=0 To Ubound(arrList, 2) %>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td>
		<img src="http://webadmin.10x10.co.kr/images/icon_search.jpg" onclick="pop_itemqna('<%= arrList(12, i) %>');return false;" style="cursor:pointer;"/>
		<%= arrList(0, i) %>
	</td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(1, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(2, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(3, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(4, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(5, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(7, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;">
	<%
		If arrList(4, i) <> 0 Then
			response.write Round(arrList(4, i) / arrList(1, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;">
	<%
		If arrList(6, i) <> 0 Then
			response.write Round(arrList(6, i) / arrList(1, i), 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(7, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(8, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(9, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(10, i) %></td>
	<td onclick="fnAddType('<%= arrList(12, i) %>');" style="cursor:pointer;"><%= arrList(11, i) %></td>
</tr>
<% Next %>
<% End If %>
</table>
<div id="shopTBL"></div>
<% SET oQna = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
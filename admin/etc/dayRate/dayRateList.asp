<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/dayRate/dayRateCls.asp"-->
<%
Dim i, yyyy, mm, dd, yyyymmdd, wDateNm
dd		= Day(now)

yyyy = requestCheckVar(request("yyyy"),4)
mm = requestCheckVar(request("mm"),2)

If yyyy	= "" Then yyyy	= Year(now)
If mm	= "" Then mm	= Month(now)
If dd	= "" Then dd	= Day(now)

If mm < 10 Then mm = "0" & mm

Dim year_from, year_to
year_from = Year(now) - 5
year_to = Year(now) + 1

Dim oRate, lp, weekno
Set oRate = new CRate
	oRate.FRectYear = yyyy
	oRate.FRectMonth = Num2Str(mm,2,"0","R")
	oRate.getdayRateList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function frmCheck(ymd, d){
	var USDrate, CNYrate, MYRrate, SGDrate;

	USDrate = $("#a"+d).val();
	CNYrate = $("#b"+d).val();
	MYRrate = $("#c"+d).val();
	SGDrate = $("#d"+d).val();
	
	if (USDrate == '' ){
		alert('USD 환율을 입력하세요');
		$("#a"+d).focus();
		return;
	}

	if (CNYrate == '' ){
		alert('CNY 환율을 입력하세요');
		$("#b"+d).focus();
		return;
	}

	if (MYRrate == '' ){
		alert('MYR 환율을 입력하세요');
		$("#c"+d).focus();
		return;
	}

	if (SGDrate == '' ){
		alert('SGD 환율을 입력하세요');
		$("#d"+d).focus();
		return;
	}

	if (confirm(''+ymd+' 데이터를 저장 하시겠습니까?')){
		$("#yyyymmdd").val(ymd);
		$("#USD").val(USDrate);
		$("#CNY").val(CNYrate);
		$("#MYR").val(MYRrate);
		$("#SGD").val(SGDrate);

		document.frmSvArr.target = "xLink";
		document.frmSvArr.action = "<%=apiURL%>/outmall/cjmall/actCjMallReq.asp"
		document.frmSvArr.action = "/admin/etc/dayRate/dayRateProc.asp"
		document.frmSvArr.submit();
	}
}
</script>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		년월:
		<select name="yyyy" class="select">
		<% For i = year_from to year_to %>
			<option value="<%= i %>" <%= Chkiif(CInt(yyyy) = i, "selected", "") %>><%= i %></option>
		<% Next %>
		</select>
		/
		<select name="mm" class="select">
		<% for i = 1 to 12 %>
			<option value="<%= i %>" <%= Chkiif(CInt(mm) = i, "selected", "") %>><%= i %></option>
		<% next %>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">날짜</td>
	<td>USD</td>
	<td>CNY</td>
	<td>MYR</td>
	<td>SGD</td>
	<td width="100">관리</td>
</tr>
<%
If oRate.FResultCount>0 then 
	For i = 0 To (oRate.FResultCount - 1)
%>
<% If oRate.FItemList(i).Fweek = "토" OR oRate.FItemList(i).Fweek = "일" Then %>
<tr align="center" bgcolor="#FFF0F0">
<% Else %>
<tr align="center" bgcolor="#FFFFFF">
<% End If %>
	<td width="200"><%= oRate.FItemList(i).FDate & " (" & oRate.FItemList(i).Fweek& ")" %></td>
	<td align="center"><input type="text" id="a<%=i%>" value="<%= oRate.FItemList(i).FUSD %>"></td>
	<td align="center"><input type="text" id="b<%=i%>" value="<%= oRate.FItemList(i).FCNY %>"></td>
	<td align="center"><input type="text" id="c<%=i%>" value="<%= oRate.FItemList(i).FMYR %>"></td>
	<td align="center"><input type="text" id="d<%=i%>" value="<%= oRate.FItemList(i).FSGD %>"></td>
	<td width="100"><input type="button" class="button_s" value="저장" onclick="frmCheck('<%=oRate.FItemList(i).FDate%>', <%=i%>)"></td>
</tr>
<%
	Next
End If
%>
</table>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" id="yyyymmdd" name="yyyymmdd" value="">
<input type="hidden" id="USD" name="USD" value="">
<input type="hidden" id="CNY" name="CNY" value="">
<input type="hidden" id="MYR" name="MYR" value="">
<input type="hidden" id="SGD" name="SGD" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="10"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/board/holiday/holidayCls.asp"-->
<%
Dim oHoliday, topN
Dim sSdate, sEdate, page, i, dayColor, tenholidayyn, logicholidayyn, upcheholidayyn
Dim sumWorkTenTen, sumWorkLogics, sumWorkUpche
page	= request("page")
topN	= request("topN")
sSdate	= requestCheckVar(request("iSD"),10)
sEdate	= requestCheckVar(request("iED"),10)

tenholidayyn	= requestCheckVar(request("tenholidayyn"),1)
logicholidayyn	= requestCheckVar(request("logicholidayyn"),1)
upcheholidayyn	= requestCheckVar(request("upcheholidayyn"),1)

sumWorkTenTen	= 0
sumWorkLogics	= 0
sumWorkUpche	= 0

If topN = "" Then topN = 50
If page = "" Then page = 1
If sSdate = "" Then sSdate = DateSerial(Year(Now()), Month(Now()), 1)
If sEdate = "" Then sEdate = DateSerial(Year(Now()), Month(Now()), Get_Lastday(Year(Now()), Month(Now())))

Set oHoliday = new CHoliday
	oHoliday.FCurrPage			= page
	oHoliday.FPageSize			= TopN
	oHoliday.FRectStartDate		= sSdate
	oHoliday.FRectEndDate		= sEdate
	oHoliday.FRectHoliday		= tenholidayyn
	oHoliday.FRectLogicsHoliday	= logicholidayyn
	oHoliday.FRectUpcheHoliday	= upcheholidayyn
	oHoliday.getHolidayItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function popholiday(v){
	var pCM2 = window.open("/admin/board/holiday/popholiday.asp?num="+v,"popholiday","width=500,height=300,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
</script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
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
		&nbsp;&nbsp;
		출력수 : <input type="text" class="text" name="topN" size="3" value="<%= topN %>">
		<br />
		텐바이텐 휴무 :
		<select name="tenholidayyn" class="select">
			<option value="">-선택-</option>
			<option value="Y" <%= Chkiif(tenholidayyn="Y", "selected", "") %> >Y</option>
		</select>&nbsp;&nbsp;
		물류 휴무 :
		<select name="logicholidayyn" class="select">
			<option value="">-선택-</option>
			<option value="Y" <%= Chkiif(logicholidayyn="Y", "selected", "") %> >Y</option>
		</select>&nbsp;&nbsp;
		업체 휴무 :
		<select name="upcheholidayyn" class="select">
			<option value="">-선택-</option>
			<option value="Y" <%= Chkiif(upcheholidayyn="Y", "selected", "") %> >Y</option>
		</select>&nbsp;&nbsp;
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%= FormatNumber(oHoliday.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHoliday.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="20%">양력 날짜</td>
	<td width="10%">음력 날짜</td>
	<td width="10%">텐바이텐 휴무</td>
	<td width="10%">물류 휴무</td>
	<td width="10%">업체 휴무</td>
	<td width="40%">비고</td>
</tr>
<% For i=0 to oHoliday.FResultCount - 1 %>
<%
	If oHoliday.FItemList(i).FWeek ="일" Then
		dayColor = "RED"
	ElseIf oHoliday.FItemList(i).FWeek ="토" Then
		dayColor = "BLUE"
	Else
		dayColor = "BLACK"
	End If
%>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" onclick="popholiday('<%= oHoliday.FItemList(i).FNum %>')" style="cursor:pointer;">
	<td><font size="2" color="<%= dayColor %>"><%= oHoliday.FItemList(i).FSolar_date %>&nbsp;(<%= oHoliday.FItemList(i).FWeek %>)</font></td>
	<td><%= oHoliday.FItemList(i).FLunar_date %></td>
	<td><%= Chkiif(oHoliday.FItemList(i).FHoliday<>"0", "O", "") %></td>
	<td><%= Chkiif(oHoliday.FItemList(i).FLogics_holiday="2", "O", "") %></td>
	<td><%= Chkiif(oHoliday.FItemList(i).FUpche_holiday="2", "O", "") %></td>
	<td><%= oHoliday.FItemList(i).FHoliday_name %></td>
</tr>
<%
		If oHoliday.FItemList(i).FHoliday = "0" Then
			sumWorkTenTen = sumWorkTenTen + 1
		End If

		If oHoliday.FItemList(i).FLogics_holiday = "0" Then
			sumWorkLogics = sumWorkLogics + 1
		End If

		If oHoliday.FItemList(i).FUpche_holiday = "0" Then
			sumWorkUpche = sumWorkUpche + 1
		End If
	Next
%>
<!--
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><strong>영업일 합계</strong></td>
	<td><%= sumWorkTenTen %>일</td>
	<td><%= sumWorkLogics %>일</td>
	<td><%= sumWorkUpche %>일</td>
	<td></td>
</tr>
-->
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oHoliday.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHoliday.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHoliday.StartScrollPage to oHoliday.FScrollCount + oHoliday.StartScrollPage - 1 %>
    		<% if i>oHoliday.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHoliday.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% Set oHoliday = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
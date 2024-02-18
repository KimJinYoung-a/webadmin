<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%

dim page, startDate, endDate
page = requestCheckVar(request("page"), 32)
startDate = requestCheckVar(request("startDate"), 10)
endDate = requestCheckVar(request("endDate"), 10)

if page = "" then page = 1
if startDate = "" then startDate = dateadd("d",-30,date)
if endDate = "" then endDate = dateadd("d",7,date)

dim oClass
set oClass = new CMailzineList
	oClass.FPageSize = 20
	oClass.FCurrPage = page
	oClass.FrectSDate = startDate
	oClass.FrectEDate = endDate
	oClass.MailzineClassList()

dim i

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script>
function goPage(pg) {
	document.frm.page.value=pg;
	document.frm.submit();
}

function editreg(dt) {
	var editreg = window.open('mailzine_class_detail_pop.asp?dt='+dt,'editreg','width=1024,height=520 ,scrollbars=yes,resizable=yes');
	editreg.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
		* 발송일자 :
        <input id="startDate" name="startDate" value="<%= startDate %>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startDate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="endDate" name="endDate" value="<%= endDate %>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="endDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "startDate", trigger    : "startDate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "endDate", trigger    : "endDate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goPage(1);">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type="button" value=" 등 록 " onclick="editreg('');" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= oClass.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oClass.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100" rowspan="2">발송일자</td>
	<td colspan="4">클래스 01</td>
	<td colspan="2">클래스 02</td>
	<td colspan="2">클래스 03</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">상품코드</td>
	<td width="80">할인율</td>
	<td width="280">강좌설명</td>
	<td width="280">강좌서브설명</td>
	<td width="80">상품코드</td>
	<td width="80">할인율</td>
	<td width="80">상품코드</td>
	<td width="80">할인율</td>
</tr>
<% if oClass.FresultCount>0 then %>
<% for i = 0 to oClass.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" >
	<td><a href="javascript:editreg('<%= oClass.FItemList(i).FclassDate %>')"><%= oClass.FItemList(i).FclassDate %></a></td>
	<td><%= oClass.FItemList(i).Fitemid1 %></td>
	<td><%= oClass.FItemList(i).FsalePer1 %></td>
	<td><%= oClass.FItemList(i).FclassDesc1 %></td>
	<td><%= oClass.FItemList(i).FclassSubDesc1 %></td>
	<td><%= oClass.FItemList(i).Fitemid2 %></td>
	<td><%= oClass.FItemList(i).FsalePer2 %></td>
	<td><%= oClass.FItemList(i).Fitemid3 %></td>
	<td><%= oClass.FItemList(i).FsalePer3 %></td>
	<td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
       	<% if oClass.HasPreScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%= oClass.StarScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oClass.StarScrollPage to oClass.StarScrollPage + oClass.FScrollCount - 1 %>
			<% if (i > oClass.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oClass.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:goPage(<%=i%>)" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oClass.HasNextScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%=i%>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

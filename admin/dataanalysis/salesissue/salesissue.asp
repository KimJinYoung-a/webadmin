<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 데이터분석 영업이슈
' History : 2016.01.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_salesissue_cls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim page, i, isusing, startdate, enddate, title, rowColor, department_id, reloadding
	page = requestcheckvar(request("page"),1)
	isusing = requestcheckvar(request("isusing"),10)
	startdate = requestCheckVar(Request("startdate"),10)
	enddate = requestCheckVar(Request("enddate"),10)
	title = Request("title")
	department_id = requestCheckVar(Request("department_id"),10)
	reloadding = requestCheckVar(Request("reloadding"),2)

if page = "" then page = 1
if startdate="" then
	if startdate = "" or isnull(startdate) then startdate = dateadd("m", -1, date())
end if
if enddate="" then
	if enddate = "" or isnull(enddate) then enddate = date()
end if
if reloadding="" and isusing="" then isusing="Y"

dim osales
set osales = new cdataanalysis_salesissue
	osales.FPageSize = 30
	osales.FCurrPage = page
	osales.frectisusing = isusing
	osales.frecttitle = title
	osales.frectdepartment_id = department_id
	osales.frectstartdate = startdate
	osales.frectenddate = enddate
	osales.getdataanalysis_salesissue_list()
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function goPage(page) {
	frm.page.value=page;
	frm.submit();
}

function jssalesissuereg(salesidx){
	var jssalesissuereg = window.open('/admin/dataanalysis/salesissue/salesissue_edit.asp?salesidx='+salesidx+'&menupos=<%= menupos %>','jssalesissuereg','width=1024,height=500,scrollbars=yes,resizable=yes');
	jssalesissuereg.focus();
}

</script>

<form name="frm" method="post" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<input type="hidden" name="reloadding" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간 : 
		<input id='startdate' name='startdate' value='<%= startdate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= enddate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		&nbsp;&nbsp;&nbsp;
		사용여부 : 
		<select name="isusing" value="<%= isusing %>" class="select">
			<option value="" <% if isusing = "" then response.write " selected" %>>전체</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>	
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="goPage('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		프로젝트명 : <input type="text" name="title" value="<%= title %>" size="18" class="text">
		&nbsp;&nbsp;&nbsp;&nbsp;
		부서 : <%= drawSelectBoxDepartmentALL("department_id", department_id) %>
	</td>
</tr>
</table>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button_s" value="신규등록" onClick="jssalesissuereg('');">
	</td>
</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= osales.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= osales.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>부서</td>
	<td>기간</td>
	<td>프로젝트명</td>
	<td>설명(목적/결과)</td>
	<td>등록</td>
	<td>비고</td>
</tr>
<% if osales.FresultCount>0 then %>
<% for i=0 to osales.FresultCount-1 %>
<% if osales.FItemList(i).fisusing="Y" then %>
	<tr align="center" bgcolor="#FFFFFF">
<% else %>
	<tr align="center" bgcolor="#FFCCCC">
<% end if %>
	<td align="center">
		<input type="hidden" name="salesidx" value="<%= osales.FItemList(i).fsalesidx %>">
		<%= osales.FItemList(i).fsalesidx %>
	</td>
	<td align="center">
		<%= osales.FItemList(i).fdepartmentNameFull %>
	</td>
	<td align="center">
		<%= FormatDate(osales.FItemList(i).fstartdate,"0000.00.00") %> ~ <%= FormatDate(osales.FItemList(i).fenddate,"0000.00.00") %>
	</td>
	<td align="center">
		<%= chrbyte(osales.FItemList(i).ftitle,30,"Y") %>
	</td>
	<td align="center">
		<%= chrbyte(osales.FItemList(i).fcomment,30,"Y") %>
	</td>
	<td align="center">
		<%= FormatDate(osales.FItemList(i).fregdate,"0000.00.00") %>
		<br><%= osales.FItemList(i).fusername %>
	</td>
	<td align="center">
		<input type="button" class="button_s" value="수정" onClick="jssalesissuereg('<%= osales.FItemList(i).fsalesidx %>');">
	</td>		
</tr>   
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if osales.HasPreScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%= osales.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + osales.StartScrollPage to osales.StartScrollPage + osales.FScrollCount - 1 %>
			<% if (i > osales.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(osales.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if osales.HasNextScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<script type="text/javascript">
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>

<%
set osales=nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
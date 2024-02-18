<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메인페이지
' History : 2018-03-08 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_enjoyContentsManageCls.asp" -->
<%

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun, i
dim page,strParm
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")

	If gubun = "" Then
		gubun = "index"
	End If


	if page="" then page=1

dim oMainContents
	set oMainContents = new CMainEnjoyContents
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectSelDate = prevDate
	oMainContents.GetMainEnjoyContentsList

strParm = "prevDate="&prevDate
%>
<script type="text/javascript">
<!--
function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/lib/popmainenjoycontentsedit.asp?idx=' + idx+'&<%=strParm%>','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}
//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    사용구분
		<select name="isusing" class="select">
		<option value="" <% if isusing="" then response.write "selected" %>>전체
		<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
		<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
		</select>
        &nbsp;&nbsp;
        지정일자 <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="right">
    	<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>idx</td>
    <td>이벤트명</td>
    <td>이미지</td>
    <td>시작일</td>
    <td>종료일</td>
    <td>사용여부</td>
    <td>우선순위</td>
    <td>등록자</td>
    <td>최종작업자</td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).Fidx & "</a>" %></td>
    <td align="center">[<%= oMainContents.FItemList(i).FEvt_Code %>] <%= oMainContents.FItemList(i).FEvt_Title %></td>
    <td align="center">
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).Fevt_mainimg %>" border="0" width="300"></a>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).FStartDate %></td>
    <td align="center"><%= oMainContents.FItemList(i).FEndDate %></td>
    <td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
    <td align="center"><%= oMainContents.FItemList(i).FDispOrder %></td>
    <td align="center"><%= oMainContents.FItemList(i).FRegUser %></td>
    <td align="center"><%= oMainContents.FItemList(i).FLastUser %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="12" align="center" height="30">
    <% if oMainContents.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oMainContents.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oMainContents.StarScrollPage to oMainContents.FScrollCount + oMainContents.StarScrollPage - 1 %>
		<% if i>oMainContents.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oMainContents.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<%
set oMainContents = Nothing
%>

<form name="refreshFrm" method="post">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
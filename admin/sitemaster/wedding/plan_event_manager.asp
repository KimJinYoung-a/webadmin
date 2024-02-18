<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  웨딩 기획전
' History : 2018-04-10 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/sitemaster/wedding/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/wedding_ContentsManageCls.asp" -->
<%

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun, i, DateDiv
dim page,strParm
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")
	DateDiv = request("DateDiv")

	If gubun = "" Then
		gubun = "index"
	End If

	If DateDiv="" Then DateDiv="Y"

	if page="" then page=1

dim oPlanEvent
	set oPlanEvent = new CWeddingContents
	oPlanEvent.FPageSize = 10
	oPlanEvent.FCurrPage = page
	oPlanEvent.FRectIsusing = isusing
	oPlanEvent.FRectSelDate = prevDate
	oPlanEvent.FRectDateDiv = DateDiv
	oPlanEvent.GetPlanEventList
%>
<script type="text/javascript">
<!--
function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/wedding/popweddingplaneventedit.asp?idx=' + idx+'&<%=strParm%>','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function NextPage(page){
    frm.page.value = page;
    frm.submit();
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
		<option value="" <% if isusing="" then response.write "selected" %>>전체</option>
		<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함</option>
		<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함</option>
		</select>
        &nbsp;&nbsp;
		진행상태
		<select name="DateDiv" class="select">
		<option value="A" <% if DateDiv="A" then response.write "selected" %>>전체</option>
		<option value="Y" <% if DateDiv="Y" then response.write "selected" %> >진행중</option>
		<option value="N" <% if DateDiv="N" then response.write "selected" %> >종료</option>
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
	<td colspan="10">
		검색결과 : <b><%=oPlanEvent.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oPlanEvent.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>idx</td>
    <td>이벤트명</td>
    <td>메인컷등록</td>
	<td>이미지</td>
    <td>시작일</td>
    <td>종료일</td>
    <td>사용여부</td>
    <td>우선순위</td>
    <td>등록자</td>
    <td>최종작업자</td>
</tr>
<%
	for i=0 to oPlanEvent.FResultCount - 1
%>
<% if (oPlanEvent.FItemList(i).IsEndDateExpired) or (oPlanEvent.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).Fidx %>');"  style="cursor:pointer;"><%=oPlanEvent.FItemList(i).Fidx %></td>
    <td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).Fidx %>');"  style="cursor:pointer;">
	[<%= oPlanEvent.FItemList(i).FEvt_Code %>] <%= oPlanEvent.FItemList(i).FEvt_Title %>
	</td>
	<td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).Fidx %>');"  style="cursor:pointer;">
    	<% If Instr(oPlanEvent.FItemList(i).Fevt_img_upload,"weddingban")>0 Then %>메인 이미지등록<% End If %>
    </td>
    <td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).Fidx %>');"  style="cursor:pointer;">
    	<% If oPlanEvent.FItemList(i).Fevt_img_upload<>"" Then %><img src="<%= oPlanEvent.FItemList(i).Fevt_img_upload %>" border="0" width="300"><% Else %><img src="<%= oPlanEvent.FItemList(i).Fevt_img %>" border="0" width="300"><% End If %>
    </td>
    <td align="center"><%= oPlanEvent.FItemList(i).FStartDate %></td>
    <td align="center"><%= oPlanEvent.FItemList(i).FEndDate %></td>
    <td align="center"><%= oPlanEvent.FItemList(i).FIsusing %></td>
    <td align="center"><%= oPlanEvent.FItemList(i).FDispOrder %></td>
    <td align="center"><%= oPlanEvent.FItemList(i).FRegUser %></td>
    <td align="center"><%= oPlanEvent.FItemList(i).FLastUser %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="12" align="center" height="30">
    <% if oPlanEvent.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oPlanEvent.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oPlanEvent.StarScrollPage to oPlanEvent.FScrollCount + oPlanEvent.StarScrollPage - 1 %>
		<% if i>oPlanEvent.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oPlanEvent.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<%
set oPlanEvent = Nothing
%>

<form name="refreshFrm" method="post">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
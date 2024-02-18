<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  웨딩 쇼핑리스트
' History : 2018-04-12 정태훈 생성
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
	oPlanEvent.FPageSize = 20
	oPlanEvent.FCurrPage = page
	oPlanEvent.FRectIsusing = isusing
	oPlanEvent.FRectSelDate = prevDate
	oPlanEvent.FRectDateDiv = DateDiv
	oPlanEvent.GetShoppingList
%>
<script type="text/javascript">
<!--
function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/wedding/popWeddingShoppingListedit.asp?idx=' + idx+'&<%=strParm%>','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
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


<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="right">
    	<!-- <a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a> -->
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
    <td>타이틀</td>
    <td>상품이미지</td>
    <td>최종작업자</td>
</tr>
<%
	for i=0 to oPlanEvent.FResultCount - 1
%>
<tr bgcolor="#FFFFFF">
    <td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).FWeddingStepID %>');"  style="cursor:pointer;"><%=oPlanEvent.FItemList(i).FWeddingStepID %></td>
    <td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).FWeddingStepID %>');"  style="cursor:pointer;">
	<%= oPlanEvent.FItemList(i).GetDDayTitle %> <%=oPlanEvent.FItemList(i).GetDDayImageCnt%>
	</td>
    <td align="center" onClick="AddNewMainContents('<%= oPlanEvent.FItemList(i).FWeddingStepID %>');"  style="cursor:pointer;">
    	<img src="<%= oPlanEvent.FItemList(i).Fsmallimage %>" border="0">
    </td>
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
<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : index.asp
' Discription : 모바일 카테고리 이미지관리
' History : 2013.12.12 이종화 생성
'			2013.12.15 한용민 수정
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/catetag.asp" -->

<%
Dim isusing , dispcate, page ,i, oCateTaglist, reload
	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	reload = RequestCheckVar(request("reload"),16)

if page="" then page=1
if reload="" and isusing="" then isusing="Y"

set oCateTaglist = new CMaincatetag
	oCateTaglist.FPageSize			= 20
	oCateTaglist.FCurrPage			= page
	oCateTaglist.Fisusing			= isusing
	oCateTaglist.Fcatecode			= dispcate
	oCateTaglist.GetContentsList()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>


//수정
function jsmodify(v){
	location.href = "mc_insert.asp?menupos=<%=menupos%>&idx="+v;
}

</script>

<!-- 검색 시작 -->
<table width="800" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<div style="padding-bottom:10px;">
		* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
		* 카테고리 :<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</div>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
	</td>
</tr>
</form>	
</table>
<!-- 검색 끝 -->

<table width="800" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right">
		<!-- 신규등록 -->
		<a href="mc_insert.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
	</td>
</tr>
</table>

<!--  리스트 -->
<table width="800" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		총 등록수 : <b><%=oCateTaglist.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oCateTaglist.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>카테고리</td>
	<td>키워드</td>	 
	<td>사용여부</td>
</tr>
<% 
	for i=0 to oCateTaglist.FResultCount-1 
%>
<tr height="30" align="center" bgcolor="<%=chkIIF(oCateTaglist.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
	<td onclick="jsmodify('<%=oCateTaglist.FItemList(i).fidx%>');" style="cursor:pointer;">
		<%=oCateTaglist.FItemList(i).fidx%>
	</td>
	<td><%=oCateTaglist.FItemList(i).Fcatename%></td>
	<td><%=oCateTaglist.FItemList(i).Fkword1%></td>
	<td><%=chkiif(oCateTaglist.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="800" cellpadding="0" cellspacing="0" class="a" style="padding-top:20px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if oCateTaglist.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oCateTaglist.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oCateTaglist.StartScrollPage to oCateTaglist.StartScrollPage + oCateTaglist.FScrollCount - 1 %>
				<% if (i > oCateTaglist.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oCateTaglist.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oCateTaglist.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set oCateTaglist = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
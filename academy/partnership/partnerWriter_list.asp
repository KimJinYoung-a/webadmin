<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2016.07.22 김진영 생성
'	Description : 작가 신청 리스트
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/partner_writercls.asp"-->
<%
Dim searchConfirm, searchKey, searchString
Dim oWriter, page, i, vbgcolor, vstrUsing

page    		= RequestCheckvar(request("page"),10)
searchKey 		= RequestCheckvar(request("searchKey"),16)
searchString	= request("searchString")
searchConfirm	= RequestCheckvar(request("searchConfirm"),1)
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
If page = "" Then page = 1
	
Set oWriter = new CWriter
	oWriter.FCurrPage					= page
	oWriter.FPageSize					= 20
	oWriter.FRectSearchKey				= searchKey
	oWriter.FRectsearchString			= searchString
	oWriter.FRectSearchConfirm			= searchConfirm
	oWriter.getWriterRegedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function goView(v){
	location.href='/academy/partnership/partnerWriter_view.asp?menupos=<%=menupos%>&idx='+v;	
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		답변여부
		<select name="searchConfirm" onchange="document.frm_search.submit()" class="select">
			<option value="">::선택::</option>
			<option value="Y" <%= chkiif(searchConfirm = "Y", "selected", "") %>>완료</option>
			<option value="N" <%= chkiif(searchConfirm = "N", "selected", "") %>>대기</option>
		</select>
		/ 검색
		<select name="searchKey" class="select">
			<option value="">::선택::</option>
			<option value="writername" <%= chkiif(searchKey = "writername", "selected", "") %>>작가명</option>
			<option value="bunya" <%= chkiif(searchKey = "bunya", "selected", "") %>>작품분야</option>
		</select>
		<input type="text" name="searchString" size="20" class="text" value="<%=searchString%>">	
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oWriter.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oWriter.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="50">번호</td>
	<td width="200">작품분야</td>
	<td>작품소개</td>
	<td width="200">작가명</td>
	<td width="150">전화번호</td>
	<td width="150">휴대폰번호</td>
	<td width="130">등록일</td>
	<td width="70">답변</td>
</tr>
<% 
For i = 0 to oWriter.FResultCount - 1
	If oWriter.FItemList(i).FConfirmyn = "N" Then
		vbgcolor="#FFFFFF"
		vstrUsing = "<font color=darkred>대기</font>"
	Else
		vbgcolor="#F8F8F8"
		vstrUsing = "<font color=darkblue>완료</font>"
	End if
%>
<tr align="center" bgcolor="<%= vbgcolor %>" onclick="goView('<%= oWriter.FItemList(i).FIdx %>')" style="cursor:pointer" height="25">
	<td width="50"><%= oWriter.FItemList(i).FIdx %></td>
	<td width="200"><%= oWriter.FItemList(i).FBunya %></td>
	<td><%= oWriter.FItemList(i).FIntroduce %></td>
	<td width="200"><%= oWriter.FItemList(i).FWritername %></td>
	<td width="150"><%= oWriter.FItemList(i).FUserphone %></td>
	<td width="150"><%= oWriter.FItemList(i).FUsercell %></td>
	<td width="130"><%= FormatDate(oWriter.FItemList(i).FRegdate,"0000.00.00") %></td>
	<td width="70"><%= vstrUsing %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
	<% If oWriter.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oWriter.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End if %>
	<% For i=0 + oWriter.StartScrollPage to oWriter.FScrollCount + oWriter.StartScrollPage - 1 %>
	<% if i>oWriter.FTotalpage then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
				<font color="red">[<%= i %>]</font>
			<% Else %>
				<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% End If %>
	<% Next %>
	<% If oWriter.HasNextScroll then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
		[next]
	<% End if %>
    </td>
</tr>
</table>
<% Set oWriter = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
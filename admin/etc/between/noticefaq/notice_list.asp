<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/between/noticefaqcls.asp"-->
<%
Dim page, oboard, i
Dim gubun, subject, isusing
page		= request("page")
gubun		= request("gubun")
subject		= request("subject")
isusing		= request("isusing")

If page = "" Then page = 1

SET oboard = new cNoticeFAQ
	oboard.FPageSize 			= 20
	oboard.FCurrPage			= page
	oboard.FRectGubun			= gubun
	oboard.FRectSubject			= subject
	oboard.FRectIsusing			= isusing
	oboard.getBoardList()
%>
<script language="javascript">
function goModify(idx){
	location.href = "board_write.asp?mode=U&idx="+idx;
}
function goPage(page){
    var frm = document.frmboard;
    frm.page.value=page;
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmboard" method="get"  action="<%= CurrURL %>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				구분 : 
				<select name="gubun" class="select">
					<option value="">-Choice-</option>
					<option value="1" <%= ChkIIF(gubun = "1", "selected", "") %> >공지사항</option>
					<option value="2" <%= ChkIIF(gubun = "2", "selected", "") %> >FAQ</option>
				</select>
				&nbsp;&nbsp;
				제목 : <input type="text" class="text" name="subject" value="<%= subject %>" size="95" maxlength="128">
				&nbsp;&nbsp;
				사용유무 : 
				<select name="isusing" class="select">
					<option value="">-Choice-</option>
					<option value="Y" <%= ChkIIF(isusing = "Y", "selected", "") %> >Y</option>
					<option value="N" <%= ChkIIF(isusing = "N", "selected", "") %> >N</option>
				</select>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<a href="javascript:document.frmboard.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="등록" onclick="javascript:location.href='board_write.asp?menupos=<%=menupos%>';">
	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>리스트</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		검색결과 : <b><%= FormatNumber(oboard.FTotalCount,0) %></b>&nbsp;&nbsp;페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(oboard.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="80">번호</td>
	<td width="80">구분</td>
	<td>제목</td>
	<td width="200">등록일</td>
</tr>
<%
	If oboard.FResultCount > 0 Then
		For i = 0 to oboard.FResultCount - 1
%>
<tr height="25" bgcolor="FFFFFF" onClick="goModify('<%= oboard.FItemList(i).FIdx %>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center"><%= oboard.FItemList(i).FIdx %></td>
	<td align="center">
	<%
		Select Case oboard.FItemList(i).FGubun
			Case "1"	response.write "<font color=RED><strong>공지사항</strong></font>"
			Case "2"	response.write "<font color=BLUE><strong>FAQ</strong></font>"
		End Select
	%>
	</td>
	<td><%= oboard.FItemList(i).FSubject %></td>
	<td align="center"><%= oboard.FItemList(i).FRegdate %></td>
</tr>
<%
		Next
%>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oboard.HasPreScroll then %>
		<a href="javascript:goPage('<%= oboard.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oboard.StartScrollPage to oboard.FScrollCount + oboard.StartScrollPage - 1 %>
    		<% if i>oboard.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oboard.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<%
	Else
%>
<tr height="50" align="center" bgcolor="#FFFFFF">
	<td colspan="11">등록된 내용이 없습니다.</td>
</tr>
<%	
	End If
%>
</table>
<% SET oboard = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
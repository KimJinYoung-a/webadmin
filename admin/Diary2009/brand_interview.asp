<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
dim oDiary, i , isusing, catecode
dim page
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1
		
set oDiary = new DiaryCls
	oDiary.FPageSize = 20
	oDiary.FCurrPage = page
	oDiary.getBrandInterview_List
%>

<script language="javascript">
function popBrandInterviewDetail(a){
	var popBrandInterviewDetail = window.open('/admin/diary2009/brand_interview_detail.asp?idx='+a+'','popBrandInterviewDetail','width=800,height=870,resizable=yes,scrollbars=yes')
	popBrandInterviewDetail.focus();
}
</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<input type="button" name="ins" value="신규등록" class="button_s" onclick="popBrandInterviewDetail('')">
		<input type="button" name="ins" value="닫 기" class="button_s" onclick="window.close();">
	</td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap> No</td>
	<td nowrap> 브랜드 </td>
	<td nowrap>구분</td>
	<td nowrap>리스트 타이틀 이미지</td>
	<td nowrap>내용 타이틀 이미지</td>
	<td nowrap>리스트 노출 순서</td>
	<td nowrap> 사용여부 </td>
</tr>
<% For i =0 To  oDiary.FResultCount -1 %>
<tr align="center" bgcolor="#FFFFFF" onclick="popBrandInterviewDetail('<%=oDiary.FItemList(i).FIdx%>');" style="cursor:hand;">
	<td nowrap><%= oDiary.FItemList(i).FIdx %></td>
	<td nowrap><%= oDiary.FItemList(i).fmakerid %></td>
	<td nowrap><%= cateList("",oDiary.FItemList(i).FCateCode) %></td>
	<td nowrap><img src="<%= oDiary.FItemList(i).FImage1 %>" border="0" height="50"></td>
	<td nowrap><img src="<%= oDiary.FItemList(i).ConfImg %>" border="0"></td>
	<td nowrap><%= oDiary.FItemList(i).Fsorting %></td>
	<td nowrap><%= oDiary.FItemList(i).FisUsing %></td>
</tr>
<%Next%>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>">

	<!-- 페이지 시작 -->
    	<a href="?page=1&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
		<% if oDiary.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oDiary.StartScrollPage-1 %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
		<% else %>
		&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
		<% end if %>
		<% for i = 0 + oDiary.StartScrollPage to oDiary.StartScrollPage + oDiary.FScrollCount - 1 %>
			<% if (i > oDiary.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oDiary.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
			<% end if %>
		<% next %>
		<% if oDiary.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
		<% else %>
		&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
		<% end if %>
		<a href="?page=<%= oDiary.FTotalpage %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
	<!-- 페이지 끝 -->

	</td>
</tr>
</table>

<% 'set oDiary = nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
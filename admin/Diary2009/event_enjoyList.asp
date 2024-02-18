<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/DiaryEnjoyCls.asp"-->
<%
'###############################################
' PageName : event_enjoyList.asp
' Discription : 작가따라 그려봐 목록
' History : 2009.09.30 허진원 : 생성
'###############################################

dim page, i, lp
dim makerid, isusing

page = request("page")
if page = "" then page=1
makerid = request("makerid")
isusing = request("isusing")
if isusing = "" then isusing="Y"

dim oEnjoy
set oEnjoy = New CEnjoy
oEnjoy.FCurrPage = page
oEnjoy.FPageSize=20
oEnjoy.FRectMaker = makerid
oEnjoy.FRectUsing = isusing
oEnjoy.GetDiaryEnjoyList

%>
<script language='javascript'>
<!--
// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="event_enjoyList.asp";
	document.refreshFrm.submit();
}

//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="event_enjoyList.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		브랜드:
	    <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20" >
	    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
		/ 사용여부
		<select name="isusing">
			<option value="Y">사용</option>
			<option value="N">삭제</option>
		</select>
		<script language="javascript">
		refreshFrm.isusing.value="<%=isusing%>";
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="작가 추가" onclick="self.location='event_enjoyWrite.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		검색결과 : <b><%=oEnjoy.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oEnjoy.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>브랜드</td>
	<td>이미지</td>
	<td>제목</td>
	<td>코멘트</td>
	<td>등록일</td>
</tr>
<%	if oEnjoy.FResultCount < 1 then %>
<tr>
	<td colspan="6" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to oEnjoy.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><a href="event_enjoyWrite.asp?mode=edit&menupos=<%= menupos %>&denjSn=<%= oEnjoy.FItemList(i).FdenjSn %>"><%= oEnjoy.FItemList(i).FdenjSn %></a></td>
	<td align="center"><%= oEnjoy.FItemList(i).Fbrandname %></td>
	<td align="center"><a href="event_enjoyWrite.asp?mode=edit&menupos=<%= menupos %>&denjSn=<%= oEnjoy.FItemList(i).FdenjSn %>"><img src="<%= webImgUrl & "/diary_collection/enjoy/" & oEnjoy.FItemList(i).FsmallImage %>" width="100" border="0"></a></td>
	<td align="center"><a href="event_enjoyWrite.asp?mode=edit&menupos=<%= menupos %>&denjSn=<%= oEnjoy.FItemList(i).FdenjSn %>"><%= oEnjoy.FItemList(i).Fsubject %></a></td>
	<td align="center"><%= oEnjoy.FItemList(i).FcmtCnt %>건</td>
	<td align="center"><%= left(oEnjoy.FItemList(i).Fregdate,10) %></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">
	<!-- 페이지 시작 -->
	<%
		if oEnjoy.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oEnjoy.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oEnjoy.StartScrollPage to oEnjoy.FScrollCount + oEnjoy.StartScrollPage - 1

			if lp>oEnjoy.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oEnjoy.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<%
set oEnjoy = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
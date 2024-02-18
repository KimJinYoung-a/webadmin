<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 촬영 공지사항 리스트
' History : 2012.03.09 김진영 생성
'			2018.05.14 한용민 수정
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/boardCls.asp"-->
<%
Dim search_team, detail_search, SearchString, worker
Dim wCount
search_team 	= request("search_team")
detail_search 	= request("detail_search")
SearchString 	= request("SearchString")

Dim lBoard, page, i
page = request("page")

If page = "" Then page = 1

Set worker = new CCoopUserList
worker.FMode = "BB"
worker.fnGetCoopUserList
wCount = worker.FResultCount
Set worker = nothing

set lBoard = new board
	lBoard.Fsearch_team = search_team
	lBoard.Fdetail_search = detail_search
	lBoard.Fsearchstr = SearchString
	lBoard.FPageSize = 3
	lBoard.FCurrPage = page
	
	lBoard.FAdminlsn = session("ssAdminLsn")
	lBoard.FPartpsn = session("ssAdminPsn")
	lBoard.FPositsn = session("ssAdminPOSITsn")
	lBoard.FJob_sn = session("ssAdminPOsn")
	lBoard.fnBoardlist
%>

<script language="javascript">

function searching(){
	var sform = document.fsearch;
	sform.submit();
}
function goView(bsn){
	location.href = "board_proc.asp?mode=count&brd_sn="+bsn;
}
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}

</script>

<% If wCount > 0 or C_ADMIN_AUTH Then %>
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="등록" onclick="javascript:location.href='bbs_write.asp';">
		</td>
	</tr>
	</table>
<% End If%>

<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>공지 및 전달사항</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="page">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="50">번호</td>
	<td width="70">글쓴이</td>
	<td>제목</td>
	<td width="90">등록일</td>
	<td width="60">조회수</td>
</tr>
<%
	Dim Fteam_name, re_cnt
	If lBoard.fresultcount = 0 Then
%>
<tr height="25" bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'">
	<td align="center" colspan="5">[데이터가 없습니다.]</td>
</tr>	
<%
	Else 
		For i = 0 to lBoard.fresultcount -1
%>
<tr height="25" bgcolor="FFFFFF" onClick="goView('<%=lboard.FbrdList(i).Fbrd_sn%>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center" width="50"><%=lboard.FbrdList(i).Fbrd_sn%></td>
	<td align="center" width="70"><%=lboard.FbrdList(i).Fbrd_username%></td>
	<td>
		<%
			If lboard.FbrdList(i).Fbrd_fixed = "1" Then
				response.write "<b>"&lboard.FbrdList(i).Fbrd_subject
			Else
				response.write lboard.FbrdList(i).Fbrd_subject
			End If
		%>
	</td>
	<td align="center" width="90"><%=left(lboard.FbrdList(i).Fbrd_regdate,10)%></td>
	<td align="center" width="60"><%=lboard.FbrdList(i).Fbrd_hit%></td>
</tr>
<%
		Next
	End if
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lboard.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ohistory.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lboard.StartScrollPage to lboard.StartScrollPage + lboard.FScrollCount - 1 %>
			<% If (i > lboard.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lboard.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If lboard.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 리스트
' History : 2011.02.23 김진영 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim g_MenuPos, search_team, detail_search, SearchString, search_type
	IF application("Svr_Info")="Dev" THEN
		g_MenuPos   = "1288"		'### 메뉴번호 지정.
	Else
		g_MenuPos   = "1304"		'### 메뉴번호 지정.
	End If

search_team 	= requestCheckvar(request("search_team"),20)
detail_search 	= requestCheckvar(request("detail_search"),20)
SearchString 	= requestCheckvar(request("SearchString"),20)
search_type 	= requestCheckvar(request("brd_type"),3)

Dim lBoard, page, i
page = request("page")

If page = "" Then page = 1

set lBoard = new board
	lBoard.Fsearch_team = search_team
	lBoard.Fdetail_search = detail_search
	lBoard.Fsearchstr = SearchString
	lBoard.FPageSize = 20
	lBoard.FCurrPage = page
	lBoard.Fsearch_type = search_type

	lBoard.FAdminlsn = session("ssAdminLsn")
	lBoard.FPartpsn = session("ssAdminPsn")
	lBoard.FPositsn = session("ssAdminPOSITsn")
	lBoard.FJob_sn = session("ssAdminPOsn")
	lBoard.Fdepartment_id =  GetUserDepartmentID("", session("ssBctId"))
	lBoard.fnBoardlist
%>
<script type="text/javascript" src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function seach_check(){
	var sform = document.fsearch;
	if(sform.detail_search.value != "" && sform.SearchString.value == ""){
		sform.SearchString.focus();
		alert("검색어를 입력하세요");
		return false;
	}
	if(sform.detail_search.value == "" && sform.SearchString.value != ""){
		sform.detail_search.focus();
		alert("상세 검색 조건을 선택하세요");
		return false;
	}
	sform.submit();
}
function searching(){
	var sform = document.fsearch;
	sform.submit();
}
function goView(bsn){
	location.href = "board_proc.asp?mode=count&brd_sn="+bsn+"&menupos=<%=menupos%>&brd_type=<%=search_type%>";
}
function gosubmit(page){
    var frm = document.fsearch;
    frm.page.value=page;
	frm.submit();
}
function fnSeltype(v){
	var sform = document.fsearch;
	sform.brd_type.value=v;
	sform.submit();
}
/*
function keypressed() {
	if(event.keyCode == 8) return false;
	return true;
}
document.onkeydown=keypressed;
 */

function makePagingForm(frm) {
	$form = $('<form name="frmPaging" method="' + $(frm).attr("method") + '" action="' + $(frm).attr("action") + '"></form>');
	for (var i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
        if (!e.name) continue;
		if (!e.disabled) {
			switch (e.type) {
				case 'text':
				case 'textarea':
				case 'password':
				case 'hidden':
					$form.append('<input type="hidden" name="' + e.name + '" value="' + e.value + '">');
					break;
				case 'radio':
				case 'checkbox':
					if (e.checked) {
						$form.append('<input type="hidden" name="' + e.name + '" value="' + e.value + '">');
					}
					break;
				case 'select-one':
					$form.append('<input type="hidden" name="' + e.name + '" value="' + e.value + '">');
					break;
				case 'select-multiple':
					// alert("name : " + e.name + ", value : " + e.value + "");
					break;
			}
		}
	}
	$('body').append($form);
}

function jsGotoPage(page) {
    var frm = document.frmPaging;
    frm.page.value=page;
	frm.submit();
}

function jsAddNoti() {
	self.location.href="board_write.asp?brd_type=<%=search_type%>&menupos=<%=menupos%>";
}

$(document).ready(function() {
	makePagingForm(document.fsearch);
});

</script>
<style type="text/css">
.tabArea .btnTabTop {border-bottom: 10px solid #DADADA; border-right: 10px solid transparent; height: 0; width: 100px; font-size:1px; line-height:1;}
.tabArea .btnTabBody {background: #E0E0E0; height: 18px; width: 100px; text-align:center; cursor: pointer;}

.tabArea .currentTop {border-bottom: 10px solid #BABAFF;}
.tabArea .currentBody {background: #C8C8FF;}
.tabArea .currentTop1 {border-bottom: 10px solid #B0FFB0;}
.tabArea .currentBody1 {background: #C8FFC8;}
.tabArea .currentTop2 {border-bottom: 10px solid #FFBABA;}
.tabArea .currentBody2 {background: #FFC8C8;}
.tabArea .currentTop3 {border-bottom: 10px solid #6799FF;}
.tabArea .currentBody3 {background: #8BBDFF;}
.tabArea .currentTop4 {border-bottom: 10px solid #B2EBF4;}
.tabArea .currentBody4 {background: #C4FDFF;}
.tabArea .currentTop5 {border-bottom: 10px solid #C4B73B;}
.tabArea .currentBody5 {background: #E8DB5F;}
.tabArea .currentTop6 {border-bottom: 10px solid #F29661;}
.tabArea .currentBody6 {background: #FFBA85;}
.tabArea .currentTop7 {border-bottom: 10px solid #C4B68F;}
.tabArea .currentBody7 {background: #D6C8A1;}

.bgBdgGr {background: #473; color:#FFF;}
.bgBdgRd {background: #743; color:#FFF;}
</style>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="fsearch" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		공지구분 : <%=fnBrdType("w", "Y", search_type, "onchange=""searching()""")%>&nbsp;&nbsp;&nbsp;
		열람 선택 :
		<select name="search_team" class="select" onchange="searching()">
			<option value="">--선택--</option>
			<option value="all"  <% If search_team = "all" Then response.write "selected" End If%> >전체공지</option>
			<option value="team" <% If search_team = "team" Then response.write "selected" End If%> onchange="">팀공지</option>
		</select>
		&nbsp;&nbsp;&nbsp;상세 검색 :
		<select name="detail_search" class="select">
			<option value="">--선택--</option>
			<option value="subject" <% If detail_search = "subject" Then response.write "selected" End If%> >제목</option>
			<option value="content" <% If detail_search = "content" Then response.write "selected" End If%> >내용</option>
			<option value="writer"  <% If detail_search = "writer" Then response.write "selected" End If%> >작성자</option>
			<option value="sn"      <% If detail_search = "sn" Then response.write "selected" End If%> >공지번호</option>
		</select>
		<input type="text" name="SearchString" size="30" value="<%=SearchString%>"><input type="text" style="display: none;" />
		<img src="/admin/images/search2.gif" border="0" align="absmiddle" style="cursor:hand" onclick="seach_check()">&nbsp;&nbsp;
		</td>
	</tr>
	</form>
</table>
<br>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="등록" onclick="jsAddNoti()">
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10px">
<tr>
	<td colspan="5">
		<table cellpadding="0" cellspacing="0" class="tabArea">
		<tr>
			<td><img src="/images/icon_arrow_link.gif">&nbsp;<strong>게시글 리스트</strong></td>
		</tr>
		<tr>
			<td class="btnTabTop <%=chkIIF(search_type="","currentTop7","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="1","currentTop2","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="2","currentTop3","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="3","currentTop4","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="4","currentTop6","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="5","currentTop1","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="0","currentTop","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="6","currentTop5","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="11","currentTop2","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="12","currentTop4","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="13","currentTop3","")%>">&nbsp;</td>
			<td class="btnTabTop <%=chkIIF(search_type="90","currentTop","")%>">&nbsp;</td>
		</tr>
		<tr>
			<td class="btnTabBody <%=chkIIF(search_type="","currentBody7","")%>" onclick="fnSeltype('')">전체</td>
			<td class="btnTabBody <%=chkIIF(search_type="1","currentBody2","")%>" onclick="fnSeltype('1')">인사</td>
			<td class="btnTabBody <%=chkIIF(search_type="2","currentBody3","")%>" onclick="fnSeltype('2')">경영제도</td>
			<td class="btnTabBody <%=chkIIF(search_type="3","currentBody4","")%>" onclick="fnSeltype('3')">회사내규</td>
			<td class="btnTabBody <%=chkIIF(search_type="4","currentBody6","")%>" onclick="fnSeltype('4')">경조사</td>
			<td class="btnTabBody <%=chkIIF(search_type="5","currentBody1","")%>" onclick="fnSeltype('5')">업무</td>
			<td class="btnTabBody <%=chkIIF(search_type="0","currentBody","")%>" onclick="fnSeltype('0')">일반</td>
			<td class="btnTabBody <%=chkIIF(search_type="6","currentBody5","")%>" onclick="fnSeltype('6')">보안</td>
			<td class="btnTabBody <%=chkIIF(search_type="11","currentBody2","")%>" onclick="fnSeltype('11')">인사규정</td>
			<td class="btnTabBody <%=chkIIF(search_type="12","currentBody4","")%>" onclick="fnSeltype('12')">근태</td>
			<td class="btnTabBody <%=chkIIF(search_type="13","currentBody3","")%>" onclick="fnSeltype('13')">복리후생</td>
			<td class="btnTabBody <%=chkIIF(search_type="90","currentBody","")%>" onclick="fnSeltype('90')">기타</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=lBoard.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="50">번호</td>
	<td width="70">글쓴이</td>
	<td>제목</td>
	<td width="300">열람부서</td>
	<td width="70">등록일</td>
	<td width="70">조회수</td>
</tr>
<%
	Dim Fteam_name, re_cnt
	For i = 0 to lBoard.fresultcount -1
		If lboard.FbrdList(i).Fbrd_team <> "" Then
			Fteam_name = lboard.FbrdList(i).Fbrd_team
		End If
	Fteam_name = Replace(Fteam_name, ",", "<BR>")

%>
<tr height="25" bgcolor="FFFFFF" onClick="goView('<%=lboard.FbrdList(i).Fbrd_sn%>')" style="cursor:pointer" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center" width="50"><%=lboard.FbrdList(i).Fbrd_sn%></td>
	<td align="center" width="70"><%=lboard.FbrdList(i).Fbrd_username%></td>
	<td>
		<%
			If lboard.FbrdList(i).Fbrd_fixed = "1" Then
				response.write "<b>"&lboard.FbrdList(i).Fbrd_subject
				If lboard.FbrdList(i).Fcnt = "0" Then
					response.write ""
				Else
					response.write "&nbsp;<b><font color='RED'>["&lboard.FbrdList(i).Fcnt&"]</font></b>"
				End If
			Else
				response.write lboard.FbrdList(i).Fbrd_subject
				If lboard.FbrdList(i).Fcnt = "0" Then
					response.write ""
				Else
					response.write "&nbsp;<b><font color='RED'>["&lboard.FbrdList(i).Fcnt&"]</font></b>"
				End If
			End If
		%>
	</td>
	<td width="200"><%=Fteam_name%></td>
	<td align="center" width="70"><%=left(lboard.FbrdList(i).Fbrd_regdate,10)%></td>
	<td align="center" width="70"><%=lboard.FbrdList(i).Fbrd_hit%></td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lboard.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:jsGotoPage('<%= lboard.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lboard.StartScrollPage to lboard.StartScrollPage + lboard.FScrollCount - 1 %>
			<% If (i > lboard.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lboard.FCurrPage) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:jsGotoPage('<%= i %>');">[<%= i %>]</a>
			<% End if %>
		<% Next %>
		<% If lboard.HasNextScroll Then %>
			<a href="javascript:jsGotoPage('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

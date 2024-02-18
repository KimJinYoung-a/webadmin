<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : GIFT 메인 HOT ISSUE 관리
' Hieditor : 서동석 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftmain_cls.asp" -->
<%
dim page, i
	page = requestCheckVar(getNumeric(request("page")),10)
	if page = "" then page = 1 end if
	
dim cGift
	set cGift = new Cgift_list
	cGift.FPageSize = 15
	cGift.FCurrPage = page
	cGift.FRectIsusing = "Y"
	cGift.FRectIsOpen = "Y"

	cGift.sbHotIssueList
%>
<script type='text/javascript'>

function NextPage(p){
	frm1.page.value = p;
	frm1.submit();
}

function talkhotissue(i){
	var talkhotissuepop = window.open('main_hotissue_write.asp?idx='+i+'','talkhotissuepop','width=1200,height=768,scrollbars=yes,resizable=yes');
	talkhotissuepop.focus();
}
</script>

<form name="frm1" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=request("menupos")%>">
<input type="hidden" name="page" value="">
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		정렬번호순서 : 0이 가장 위, 테마번호가 최근일수록 위
	</td>
	<td align="right">	
		<input type="button" value="새글쓰기" onClick="talkhotissue('')" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=cGift.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=cGift.FtotalPage%></b>
	</td>
</tr>
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>정렬번호</td>
    <td>테마idx</td>
    <td>테마제목</td>
    <td>오픈 ~ 종료</td>
    <td>삭제여부</td>
    <td></td>
</tr>
<%
	for i=0 to cGift.FResultCount - 1
%>
<tr bgcolor="#FFFFFF" height="30">
    <td align="center"><%=cGift.FItemList(i).Fsortno%></td>
    <td align="center"><%=cGift.FItemList(i).FthemeIdx%></td>
    <td align="center"><%= ReplaceBracket(cGift.FItemList(i).Fsubject) %></td>
    <td align="center"><%=Left(cGift.FItemList(i).Fstartdate,10)%> ~ <%=Left(cGift.FItemList(i).Fenddate,10)%></td>
    <td align="center"><%=CHKIIF(cGift.FItemList(i).Fisusing="Y","사용중","삭제처리됨")%></td>
	<td align="center">
		[<a href="<%=wwwUrl%>/gift/shop/themeView.asp?themeIdx=<%=cGift.FItemList(i).FthemeIdx%>" target="_blank">보 기</a>]&nbsp;&nbsp;&nbsp;
		[<a href="javascript:talkhotissue('<%=cGift.FItemList(i).Fidx%>');">수 정</a>]
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="15" align="center">
    <% if cGift.HasPreScroll then %>
		<a href="javascript:NextPage('<%= cGift.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + cGift.StartScrollPage to cGift.FScrollCount + cGift.StartScrollPage - 1 %>
		<% if i>cGift.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if cGift.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<% Set cGift = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프샾이용문의
' Hieditor : 2009.04.07 서동석 생성
'			 2011.05.03 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/classes/board/offshopqnacls.asp" -->
<%
dim i, j ,shopid, page ,SearchKey, SearchString, param, isNew ,boardqna
	page = Request("page")
	shopid = Request("shopid")
	isNew = Request("isNew")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	menupos = Request("menupos")

''유저구분이 가맹점인경우 박아 넣는다
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if
    

if page="" then page=1
if SearchKey="" then SearchKey="title"
if isNew="" then isNew="Y"

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&shopid=" & shopid & "&isNew=" & isNew & "&menupos=" & menupos

'나의 1:1질문답변
set boardqna = New CMyQNA
	boardqna.FPageSize = 20
	boardqna.FCurrPage = page
	boardqna.fSearchNew = isNew
	boardqna.FRectDesigner = shopid
	boardqna.FRectSearchKey = SearchKey
	boardqna.FRectSearchString = SearchString
	boardqna.list()
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		처리여부
		<select name="isNew">
			<option value="all">전체</option>
			<option value="Y">미처리</option>
			<option value="N">처리완료</option>
		</select>
		<% if fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))="" then %>
			/ 매장
			<% call printOffShopSelectBox(isNew, shopid)%>
		<% end if %>
		/ 키워드
		<select name="SearchKey">
			<option value="title">제목</option>
			<option value="userid">작성자ID</option>
			<option value="contents">내용</option>
		</select>
		<input type="text" name="SearchString" size="12" value="<%=SearchString%>">

		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.isNew.value="<%=isNew%>";
		</script>	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if boardqna.FResultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= boardqna.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= boardqna.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>고객(아이디/주문번호)</td>
    <td>구분</td>
    <td>제목</td>
    <td>처리여부</td>
    <td>작성일</td>
    <td>비고</td>
</tr>
<% for i=0 to boardqna.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			

<% if boardqna.FItemList(i).isusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>&nbsp;<%= printUserId(boardqna.FItemList(i).userid, 2, "*") %><%= boardqna.FItemList(i).orderserial %></td>
	<td align="center"><%= boardqna.FItemList(i).Fshopname %></td>
	<td align="left">&nbsp;<%= db2html(boardqna.FItemList(i).title) %></td>
	<td align="center">
	<%
		if (boardqna.FItemList(i).replyuser=""  or isnull(boardqna.FItemList(i).replyuser)) then
			Response.Write "<font color='darkred'>미처리</font>"
		else
			Response.Write "<font color='darkblue'>처리완료</font>"
		end if
	%>
	</td>
	<td align="center"><%= FormatDate(boardqna.FItemList(i).regdate, "0000-00-00") %></td>
	<td align="center">
		<a href="offshop_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).idx %>&page=<%=page & param%>">상세</a>
	</td>
</tr>   
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if boardqna.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= boardqna.StartScrollPage-1 & param %>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + boardqna.StartScrollPage to boardqna.StartScrollPage + boardqna.FScrollCount - 1 %>
			<% if (i > boardqna.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(boardqna.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i & param %>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if boardqna.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i & param%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set boardqna = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008년 11월 4일 한용민 생성
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
dim page , i
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1

dim flagdate , evt_code, CateCode
	evt_code = requestCheckVar(request("evt_codebox"),5)
	CateCode = requestCheckVar(request("cate"),2)
	flagdate = requestCheckVar(request("flagdatebox"),5)
	if flagdate = "" then flagdate = "total"

dim oDiary
set oDiary = new DiaryCls
	oDiary.FPageSize = 50
	oDiary.FCurrPage = page
	oDiary.frectflagdate = flagdate
	oDiary.frectevt_code = evt_code
	oDiary.frectcate = CateCode
	oDiary.geteventList

%>

<script language="javascript">

//신규추가
function popnew(){
	var popnew = window.open('/admin/diary2009/event_new.asp','popnew','width=600,height=600,resizable=yes,scrollbars=yes')
	popnew.focus();
}

//수정
function popedit(idx){
	var popedit = window.open('/admin/diary2009/event_edit.asp?idx='+idx,'popedit','width=600,height=600,resizable=yes,scrollbars=yes')
	popedit.focus();
}

</script>

<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		<select name="flagdatebox">
			<option value="total" <% if flagdate="total" then response.write " selected" %>>전체보기</option>
			<option value="on" <% if flagdate="on" then response.write " selected"%> >이벤트진행중</option>
		</select>
		&nbsp;&nbsp;&nbsp;
		구분 : <% SelectList "cate", CateCode %>
		&nbsp;&nbsp;&nbsp;
		이벤트코드 : <input type="text" size=10 name="evt_codebox" value="<%=evt_code%>">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="refreshFrm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="신규등록" class="button" onclick="popnew();">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
※ 해당이벤트가 사용여부Y , 이벤트진행중 일경우만 노출됩니다. 조건을 벗어나면 등록하셔도 자동으로 노출되지 않습니다.
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oDiary.FResultCount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oDiary.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oDiary.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>번호</td>
		<td>이미지</td>
		<td>구분</td>
		<td>링크구분</td>
		<td>상품코드</td>
		<td>이벤트코드</td>
		<td>이벤트명</td>
		<td>노출순서</td>
		<td>사용여부</td>
		<td>이벤트<br>시작일</td>
		<td>이벤트<br>종료일</td>
		<td>비고</td>
	</tr>

	<% For i =0 To  oDiary.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oDiary.FItemList(i).fidx %></td>
		<td><img src="<%= oDiary.FItemList(i).fevt_bannerimg %>" width=40 height=40></td>
		<td><%= cateList("",oDiary.FItemList(i).FCateCode) %></td>
		<td><%= oDiary.FItemList(i).fevent_link %></td>
		<td><%= oDiary.FItemList(i).fitemid %></td>
		<td><%= oDiary.FItemList(i).fevt_code %></td>
		<td><%= oDiary.FItemList(i).fevt_name %></td>
		<td><%= oDiary.FItemList(i).fidx_order %></td>
		<td><%= oDiary.FItemList(i).fisusing %></td>
		<td><%= oDiary.FItemList(i).fevt_startdate %></td>
		<td><%= oDiary.FItemList(i).fevt_enddate %></td>
		<td><input type="button" value="수정" class="button" onclick="popedit(<%= oDiary.FItemList(i).fidx %>);"></td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
<% End IF %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>">

		<!-- 페이지 시작 -->
	    	<a href="?page=1&flagdatebox=<%=flagdate%>&cate=<%=CateCode%>&evt_codebox=<%=evt_code%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
			<% if oDiary.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oDiary.StartScrollPage-1 %>&flagdatebox=<%=flagdate%>&cate=<%=CateCode%>&evt_codebox=<%=evt_code%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<% for i = 0 + oDiary.StartScrollPage to oDiary.StartScrollPage + oDiary.FScrollCount - 1 %>
				<% if (i > oDiary.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oDiary.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
				<% else %>
				<a href="?page=<%= i %>&flagdatebox=<%=flagdate%>&cate=<%=CateCode%>&evt_codebox=<%=evt_code%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
				<% end if %>
			<% next %>
			<% if oDiary.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&flagdatebox=<%=flagdate%>&cate=<%=CateCode%>&evt_codebox=<%=evt_code%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<a href="?page=<%= oDiary.FTotalpage %>&flagdatebox=<%=flagdate%>&cate=<%=CateCode%>&evt_codebox=<%=evt_code%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
		<!-- 페이지 끝 -->

		</td>
	</tr>
</table>
<!-- 리스트 끝 -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
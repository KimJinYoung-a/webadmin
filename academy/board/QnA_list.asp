<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/QnA_cls.asp"-->
<%
	'// 변수 선언 //
	Dim qnaId
	Dim page, searchDiv, searchKey, searchString, param, isanswer

	Dim oQnA, i, lp, bgcolor, strUsing


	'// 파라메터 접수 //
	qnaId = RequestCheckvar(request("qnaId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),10)
	searchKey = RequestCheckvar(request("searchKey"),32)
	searchString = RequestCheckvar(request("searchString"),128)
	isanswer = RequestCheckvar(request("isanswer"),1)

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"
	if isanswer="" then isanswer="N"

	param = "&menupos=" & menupos & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isanswer=" & isanswer

	'// 클래스 선언
	set oQnA = new CQnA
	oQnA.FCurrPage = page
	oQnA.FPageSize = 20
	oQnA.FRectsearchDiv = searchDiv
	oQnA.FRectsearchKey = searchKey
	oQnA.FRectsearchString = searchString
	oQnA.FRectisanswer = isanswer

	oQnA.GetQnAList
%>
<script language='javascript'>
<!--
	function chk_form(frm)
	{
		if(!frm.searchKey.value)
		{
			alert("검색 조건을 선택해주십시오.");
			frm.searchKey.focus();
			return false;
		}
		else if(!frm.searchString.value)
		{
			alert("검색어를 입력해주십시오.");
			frm.searchString.focus();
			return false;
		}

		frm.submit();
	}

	function goPage(pg)
	{
		var frm = document.frm_search;

		frm.page.value= pg;
		frm.submit();
	}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="GET" action="QnA_list.asp" onSubmit="return chk_form(this)">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="right">
		상태
		<select name="isanswer" onChange="goPage(frm_search.page.value)">
			<option value="Y">완료</option>
			<option value="N">대기</option>
		</select>
		/ 구분
		<select name="searchDiv" onChange="goPage(frm_search.page.value)">
			<option value="">선택</option>
			<%= oQnA.optCommCd("'A000' , 'C000', 'D000'", searchDiv)%>
		</select>
		/ 검색
		<select name="searchKey">
			<option value="">선택</option>
			<option value="qnaId">공지번호</option>
			<option value="qstTitle">제목</option>
			<option value="qstContents">내용</option>
			<option value="qstUserid">작성자ID</option>
			<option value="qstUsername">작성자이름</option>
			<option value="qstlecturer_id">강사ID</option>
		</select>
		<script language="javascript">
			document.frm_search.isanswer.value="<%=isanswer%>";
			document.frm_search.searchKey.value="<%=searchKey%>";
		</script>
		<input type="text" name="searchString" size="20" value="<%= searchString %>">
       	<input type="image" src="/admin/images/search2.gif" style="width:74px;height:22px;border:0px;cursor:pointer" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 상단 검색폼 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="7" align="left">검색건수 : <%= oQnA.FTotalCount %> 건 Page : <%= page %>/<%= oQnA.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">번호</td>
		<td align="center" width="80">분류</td>
		<td align="center" width="120">구분</td>
		<td align="center">제목</td>
		<td align="center" width="70">등록자</td>
		<td align="center" width="50">상태</td>
		<td align="center" width="80">등록일</td>
	</tr>
	<%
		for lp=0 to oQnA.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="QnA_view.asp?qnaId=<%= oQnA.FQnAList(lp).FqnaId %>&page=<%=page & param%>"><%= oQnA.FQnAList(lp).FqnaId %></a></td>
		<td><%= Replace(oQnA.FQnAList(lp).FgroupNm," 구분","") %></td>
		<td><%= oQnA.FQnAList(lp).FcommNm %></td>
		<td align="left"><a href="QnA_view.asp?qnaId=<%= oQnA.FQnAList(lp).FqnaId %>&page=<%=page & param%>"><%= db2html(oQnA.FQnAList(lp).FqstTitle) %></a></td>
		<td><%= oQnA.FQnAList(lp).FqstUserId %></td>
		<td><%= oQnA.FQnAList(lp).Fisanswer %></td>
		<td><%= FormatDate(oQnA.FQnAList(lp).Fregdate,"0000.00.00") %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- 페이지 시작 -->
				<%
					if oQnA.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oQnA.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oQnA.StarScrollPage to oQnA.FScrollCount + oQnA.StarScrollPage - 1
		
						if i>oQnA.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oQnA.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- 페이지 끝 -->
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oQnA = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
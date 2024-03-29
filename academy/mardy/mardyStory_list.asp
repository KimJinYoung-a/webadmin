<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_storycls.asp"-->
<%
	'// 변수 선언 //
	dim storyId
	dim page, searchKey, searchString, param

	dim oStory, i, lp, bgcolor, strUsing


	'// 파라메터 접수 //
	storyId = RequestCheckvar(request("storyId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&searchKey=" & searchKey & "&searchString=" & searchString

	'// 클래스 선언
	set oStory = new CMardyStory
	oStory.FCurrPage = page
	oStory.FPageSize = 20
	oStory.FRectsearchKey = searchKey
	oStory.FRectsearchString = searchString

	oStory.GetMardyStoryList
%>
<script language='javascript'>
<!--
	function chk_form()
	{
		var frm = document.frm_search;

		if(!frm.searchKey.value)
		{
			alert("검색 조건을 선택해주십시오.");
			frm.searchKey.focus();
			return;
		}
		else if(!frm.searchString.value)
		{
			alert("검색어를 입력해주십시오.");
			frm.searchString.focus();
			return;
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
<form name="frm_search" method="POST" action="mardyStory_list.asp" onSubmit="return false">
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
		<select name="searchKey">
			<option value="">선택</option>
			<option value="storyId">이야기번호</option>
			<option value="titleLong">제목</option>
			<option value="contents">내용</option>
		</select>
		<script language="javascript">
			document.frm_search.searchKey.value="<%=searchKey%>";
		</script>
		<input type="text" name="searchString" size="20" value="<%= searchString %>">
       	<img src="/admin/images/search2.gif" onClick="chk_form()" style="width:74px;height:22px;border:0px;cursor:pointer" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 상단 검색폼 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="7" align="left">검색건수 : <%= oStory.FTotalCount %> 건 Page : <%= page %>/<%= oStory.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">번호</td>
		<td align="center" width="52">이미지</td>
		<td align="center">제목</td>
		<td align="center" width="70">등록자</td>
		<td align="center" width="50">조회수</td>
		<td align="center" width="80">등록일</td>
		<td align="center" width="40">상태</td>
	</tr>
	<%
		for lp=0 to oStory.FResultCount - 1

			'사용유무에따른 배경색 및 상태명 지정
			if oStory.FItemList(lp).Fisusing="N" then
				bgcolor="#E0E0E0"
				strUsing = "<font color=darkred>숨김</font>"
			else
				bgcolor="#FFFFFF"
				strUsing = "<font color=darkblue>사용</font>"
			end if
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= oStory.FItemList(lp).FstoryId %></td>
		<td><img src="<%= oStory.FItemList(lp).FimgIcon_full %>" width="50"></td>
		<td align="left"><a href="/academy/mardy/mardyStory_view.asp?storyId=<%= oStory.FItemList(lp).FstoryId %>&page=<%=page & param%>"><%= oStory.FItemList(lp).FtitleLong %></a></td>
		<td><%= oStory.FItemList(lp).Fusername %></td>
		<td align="center"><%= oStory.FItemList(lp).FhitCount %></td>
		<td><%= FormatDate(oStory.FItemList(lp).Fregdate,"0000.00.00") %></td>
		<td align="center"><%=strUsing%></td>
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
					if oStory.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oStory.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oStory.StarScrollPage to oStory.FScrollCount + oStory.StarScrollPage - 1
		
						if i>oStory.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oStory.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- 페이지 끝 -->
				</td>
				<td width="80" align="right">
					<a href="mardyStory_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oStory = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.10 한용민 수정/추가
'	Description : 파트너쉽
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/partner_lecturecls.asp"-->
<%
	'// 변수 선언 //
	dim idx
	dim page, searchKey, searchString, searchConfirm, param

	dim oLecture, i, lp, bgcolor, strUsing


	'// 파라메터 접수 //
	idx = RequestCheckvar(request("idx"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	searchConfirm = RequestCheckvar(request("searchConfirm"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if page="" then searchConfirm="N"
	if page="" then page=1
	if searchKey="" then searchKey="lecname"

	param = "&searchKey=" & searchKey & "&searchString=" & server.URLencode(searchString) &_
			"&searchConfirm=" & searchConfirm & "&menupos=" & menupos

	'// 클래스 선언
	set oLecture = new CPartnerLecture
	oLecture.FCurrPage = page
	oLecture.FPageSize = 20
	oLecture.FRectsearchKey = searchKey
	oLecture.FRectsearchString = searchString
	oLecture.FRectsearchConfirm = searchConfirm

	oLecture.GetPartnerLectureList
%>

<script language='javascript'>

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

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" action="PartnerLecture_list.asp" onSubmit="return chk_form(this)">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			답변여부
			<select name="searchConfirm" onchange="document.frm_search.submit()">
				<option value="">::선택::</option>
				<option value="Y">완료</option>
				<option value="N">대기</option>
			</select>
			<script language="javascript">
				document.frm_search.searchConfirm.value="<%=searchConfirm%>";
			</script>
			/ 검색
			<select name="searchKey">
				<option value="">::선택::</option>
				<option value="idx">번호</option>
				<option value="lecname">강사이름</option>
				<option value="lectitle">강좌개요</option>
			</select>
			<script language="javascript">
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			<input type="text" name="searchString" size="20" value="<%= searchString %>">	       	
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm_search.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

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
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oLecture.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oLecture.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="40">번호</td>
		<td align="center" width="60">공예분야</td>
		<td align="center">강좌개요</td>
		<td align="center" width="70">강사명</td>
		<td align="center" width="60">생년월일</td>
		<td align="center" width="110">연락처</td>
		<td align="center" width="110">휴대폰</td>
		<td align="center" width="75">등록일</td>
		<td align="center" width="40">답변</td>
    </tr>
	<%
	if oLecture.FTotalCount > 0 then 
		
	for lp=0 to oLecture.FResultCount - 1

		'사용유무에따른 배경색 및 상태명 지정
		if oLecture.FItemList(lp).Fconfirmyn="N" then
			bgcolor="#FFFFFF"
			strUsing = "<font color=darkred>대기</font>"
		else
			bgcolor="#F8F8F8"
			strUsing = "<font color=darkblue>완료</font>"
		end if
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= oLecture.FItemList(lp).Fidx %></td>
		<td><%= oLecture.FItemList(lp).Flecarea %></td>
		<td align="left"><a href="/academy/Partnership/PartnerLecture_view.asp?idx=<%= oLecture.FItemList(lp).Fidx %>&page=<%=page & param%>"><%= oLecture.FItemList(lp).Flectitle %></a></td>
		<td><a href="/academy/Partnership/PartnerLecture_view.asp?idx=<%= oLecture.FItemList(lp).Fidx %>&page=<%=page & param%>"><%= oLecture.FItemList(lp).Flecname %></a></td>
		<td><%= oLecture.FItemList(lp).Flecbirthday %></td>
		<td><%= oLecture.FItemList(lp).Flectel %></td>
		<td><%= oLecture.FItemList(lp).Flechp %></td>
		<td><%= FormatDate(oLecture.FItemList(lp).Fregdate,"0000.00.00") %></td>
		<td align="center"><%=strUsing%></td>
	</tr>
	<%
	next
	%>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<!-- 페이지 시작 -->
			<%
				if oLecture.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oLecture.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if
	
				for i=0 + oLecture.StarScrollPage to oLecture.FScrollCount + oLecture.StarScrollPage - 1
	
					if i>oLecture.FTotalpage then Exit for
	
					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if
	
				next
	
				if oLecture.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- 페이지 끝 -->
		</td>
	</tr>
	<% end if %>
</table>

<%
	set oLecture = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
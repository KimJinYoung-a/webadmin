<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCompanyOpen.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" --><!--html-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/company/recruit_cls.asp"-->
<!-- #include virtual="/partner/lib/function/incPageFunction.asp" -->


<%
	Dim page, SearchArea, SearchKeyword, pgsize, research

	pgsize = 15
	page = requestCheckVar(request("page"),8)
	research= requestCheckvar(request("research"),10)
	SearchArea = requestCheckVar(request("SearchArea"),128)
	SearchKeyword = requestCheckVar(request("SearchKeyword"),128)
	if page="" then	page=1

	if ((research="") and (SearchArea="")) then 
	    SearchArea = "Y"
	end if

	'// 내용 접수
	dim oRecruit, lp
	Set oRecruit = new CRecruit

	oRecruit.FPagesize = 15
	oRecruit.FCurrPage = page
	oRecruit.FRectSearchArea = SearchArea
	oRecruit.FRectSearchKeyword = SearchKeyword

	oRecruit.GetRecruitList
%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="recruit_list.asp";
		document.frm.submit();
	}

	// 상세정보(수정) 페이지 이동
	function goEdit(rcbsn)
	{
		document.frm.rcbsn.value=rcbsn;
		document.frm.page.value='<%= page %>';
		document.frm.action="recruit_edit.asp";
		document.frm.submit();
	}
//-->

function searchFrm(){
//	frm.iC.value = p;
	frm.submit();
}

function recruit_url_link(lurl){
	window.open(lurl, "_blank");
//	parent.top.location.href=lurl;
}

function new_recruit(){
	location.href="recruit_write.asp?menupos=<%=menupos%>";
}

</script>

</head>
<body>
<div class="wrap"><br><br>
	<!-- search -->
	<form name="frm" method="get" action="" action="recruit_list.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="rcbsn" value="">
	<div class="searchWrap">
		<ul>
			<li>
				<label class="formTit">검색 구분 :</label>
				<select name="SearchArea">
					<option value="">::구분::</option>
					<option value="rcb_subject" <% If "rcb_subject" = cstr(SearchArea) Then%> selected <%End if%>>제목</option>
					<option value="rcb_content" <% If "rcb_content" = cstr(SearchArea) Then%> selected <%End if%>>내용</option>
				</select>
				<input type="text" name="SearchKeyword" size="12" value="<%=SearchKeyword%>">
				<input type="button" value="검색" onClick="searchFrm();" />
			 </li>
		</ul>
	</div>
	</form>

	<div class="cont">
		<div class="pad5">
			<div class="tPad15">
				<div class="overHidden">
					<div class="ftLt"><input type="button" class="btn3 btnRd" value="신규등록" onClick="new_recruit();"></div>
				</div>
				<div class="panel1 rt pad10">
					<span>검색결과 : <strong><%=oRecruit.FtotalCount%></strong></span>
				</div>
				<table class="tbType1 listTb">
					<thead>
					<tr> 
						<th><div>번호</div></th>
						<th><div>채용직무</div></th>
						<th><div>경력여부</div></th> 
						<th><div>제목</div></th>
						<th><div>기간</div></th>
						<th><div>상태</div></th>
						<th><div>지원사이트URL</div></th>
						<th><div>작성자</div></th>
						<th><div>작성일</div></th>
						<th><div>조회수</div></th>
					</tr>
					</thead>
					<tbody>
					<% if oRecruit.FResultCount > 0 then %>
						<% for lp = 0 to (oRecruit.FResultCount - 1) %>
						<tr>
							<td><%=oRecruit.FitemList(lp).Frcb_sn%></td> <% '번호(idx) %>

							<td><%= oRecruit.FitemList(lp).Frcb_jobtype %></td>
							
							<td>
							<%
								if oRecruit.FitemList(lp).Frcb_career=1 then
									response.write "신입"
								elseif oRecruit.FitemList(lp).Frcb_career=2 then
									response.write "경력"
								elseif oRecruit.FitemList(lp).Frcb_career=3 then
									response.write "신입/경력"
								else
								end if
							 %>
							</td>
							
							<td><a href="javascript:goEdit(<%=oRecruit.FitemList(lp).Frcb_sn%>)"><%=oRecruit.FitemList(lp).Frcb_subject%></a></td>

							<td><%=Replace(left(oRecruit.FitemList(lp).Frcb_startdate,10),"-",".") & " ~ " & Replace(left(oRecruit.FitemList(lp).Frcb_enddate,10),"-",".")%></td>

							<td><%=getRecruitState(oRecruit.FitemList(lp).Frcb_state, oRecruit.FitemList(lp).Frcb_startdate, oRecruit.FitemList(lp).Frcb_enddate)%></td>

							<td>
								<% if oRecruit.FitemList(lp).Frcb_recruit_url <> "" then %>
									<input type="button" class="btn3 btnDkGy" value="바로가기"  onclick="recruit_url_link('<%=oRecruit.FitemList(lp).Frcb_recruit_url%>');" />		
								<% end if %>
							</td>

							<td><%=oRecruit.FitemList(lp).Fuserid%></td>

							<td><%=Replace(left(oRecruit.FitemList(lp).Frcb_regdate,10),"-",".")%></td>

							<td><%=oRecruit.FitemList(lp).Frcb_hit%></td>
						</tr>
						<% next %>
					<% else %>
						<tr>
							<td colspan="9">등록(검색)된 공고가 없습니다.</td>
						</tr>
					<% end if %>
					</tbody>
				</table>
				<div class="ct tPad20 cBk1">
					<% sbDisplayPaging "page", page, oRecruit.FTotalCount , pgsize , "10", menupos %>
				</div>
			</div>
		</div>
	</div>
</div>
</body>
</html>
<!-- 페이지 끝 -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCompanyClose.asp" -->
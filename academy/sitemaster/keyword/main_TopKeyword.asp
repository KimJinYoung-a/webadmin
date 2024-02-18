<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/academy/lib/classes/sitemaster/main_TopKeywrdCls.asp"-->
<%
'###############################################
' PageName : 탑메뉴검색어지정
' Discription : 메인 탑 키워드 목록
' History : 2009.09.16 한용민 10x10어드민 이전후 변경
'###############################################

dim page, SearchString, strUse, lp , i ,oKeyword , keyword_gubun
	page = RequestCheckvar(request("page"),10)
	SearchString = request("SearchString")
	keyword_gubun = RequestCheckvar(request("keyword_gubun"),10)
	strUse = RequestCheckvar(request("strUse"),1)
	if page = "" then page=1
	if strUse = "" then strUse="Y"

set oKeyword = New CSearchKeyWord
	oKeyword.FCurrPage = page
	oKeyword.FPageSize=20
	oKeyword.frectkeyword_gubun = keyword_gubun
	oKeyword.FRectUsing = strUse
	oKeyword.FRectSearch = SearchString
	oKeyword.GetSearchKeyWord()
%>
<script language='javascript'>

function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmBuyPrc;

	if(frm.selIdx.length) {
		for (var i=0;i<frm.selIdx.length;i++){
			frm.selIdx[i].checked = bool;
		}
	} else {
		frm.selIdx.checked = bool;
	}
}

function CheckSelected(){
	var pass = false;
	var frm = document.frmBuyPrc;

	if(frm.selIdx.length) {
		for (var i=0;i<frm.selIdx.length;i++){
			pass = ((pass)||(frm.selIdx[i].checked));
		}
	} else {
		pass = ((pass)||(frm.selIdx.checked));
	}

	if (!pass) {
		return false;
	}
	return true;
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	
	if (upfrm.allusing.value=='Y'){
		var ret = confirm('선택 아이템을 사용함으로 변경합니다');
	} else {
		var ret = confirm('선택 아이템을 사용안함으로 변경합니다');
	}

	if (ret) {
		upfrm.mode.value="changeUsing";
		upfrm.action="doMainTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
	}
}


function changeSort(upfrm){
	var arrSort="";
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	if(confirm('선택 아이템에 입력하신 순서번호대로 저장합니다.')) {

		if(upfrm.selIdx.length) {
			for (var i=0;i<upfrm.selIdx.length;i++){
				if(upfrm.selIdx[i].checked) arrSort = arrSort + upfrm.SortNo[i].value + ",";
			}
		} else {
			if(upfrm.selIdx.checked) arrSort=upfrm.SortNo.value;
		}
		upfrm.arrSort.value = arrSort;

		upfrm.mode.value="changeSort";
		upfrm.action="doMainTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
	}
}

function frm_search()
{
	refreshFrm.target = "";
	refreshFrm.action = "main_TopKeyword.asp";
}

	// 페이지 이동
	function goPage(pg)
	{
		document.refreshFrm.page.value=pg;
		document.refreshFrm.action="main_TopKeyword.asp";
		document.refreshFrm.submit();
	}

// 카테고리 변경시 명령
function changecontent() {
}

function AssignReal(upfrm , keyword_gubun, device){
	var idxarr; 
	var tmp;
	tmp =0;
	idxarr = "";
	
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	if(confirm('적용하시겠습니까?')) {
		if(upfrm.selIdx.length) {
			for (var i=0;i<upfrm.selIdx.length;i++){
				if(upfrm.selIdx[i].checked){
					idxarr = idxarr + upfrm.selIdx[i].value + ",";
					tmp = tmp + 1
				}	
			}
		}
	}else{
		return;
	}
	
	if (keyword_gubun == '0'){
		if (tmp > 3){
			alert('구분[메인]은 검색어 지정 3개까지만 가능합니다.');
			return;
		}	
	}else if(keyword_gubun == '1'){
		if (tmp > 7){
			alert('구분[검색]은 검색어 지정 7개까지만 가능합니다.');
			return;
		}	
	}else if(keyword_gubun == '3'){
		if (tmp > 1){
			alert('헤더 텍스트는 1개까지만 가능합니다.');
			return;
		}
		idxarr = upfrm.selIdx.value+",";
	}
	if(device == "W"){
		AssignbestReal = window.open("<%=www1Fingers%>/chtml/make_keyword.asp?idxarr=" +idxarr+ "&keyword_gubun="+keyword_gubun, "AssignbestReal","width=400,height=300,scrollbars=yes,resizable=yes");
	}else{
		AssignbestReal = window.open("<%=mob1Fingers%>/chtml/make_keyword.asp?idxarr=" +idxarr+ "&keyword_gubun="+keyword_gubun, "AssignbestReal","width=400,height=300,scrollbars=yes,resizable=yes");
	}
	AssignbestReal.focus();
}

</script>

<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				구분 : <% drawkeyword_gubun "keyword_gubun",keyword_gubun %>
				사용여부
				<select class="select" name="strUse">
					<option value="all">전체</option>
					<option value="Y">사용</option>
					<option value="N">삭제</option>
				</select>
				/ 키워드 검색
				<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">
				<script language="javascript">
					document.refreshFrm.strUse.value="<%=strUse%>";
				</script>
			</td>
		</tr>
		</table>
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
<form name="frmBuyPrc" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td>
		<% if keyword_gubun <> "" then %>
			<input type="button" onclick="AssignReal(frmBuyPrc,'<%=keyword_gubun%>','W')" value="PC실서버적용" class="button">
			&nbsp;<input type="button" onclick="AssignReal(frmBuyPrc,'<%=keyword_gubun%>','M')" value="Mobile실서버적용" class="button">
		<% end if %>
	</td>
	<td align="right">
		<select class="select" name="allusing">
			<option value="Y">선택 -> Y</option>
			<option value="N">선택 -> N</option>
		</select>
		<input type="button" class="button" value="적용" onclick="changeUsing(frmBuyPrc);">
		/
		<input type="button" class="button" value="순서변경" onclick="changeSort(frmBuyPrc);">
		/
		<input type="button" value="아이템 추가" onclick="self.location='main_TopKeyword_write.asp?menupos=<%= menupos %>'" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=oKeyword.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oKeyword.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>번호</td>
	<td>구분</td>
	<td>키워드</td>
	<td>링크정보</td>
	<td>사용유무</td>
	<td>순서</td>
	<td>등록일</td>
</tr>
<%	if oKeyword.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to oKeyword.FResultCount-1
%>
<tr align="center" bgcolor="<% if oKeyword.FItemList(i).Fisusing = "Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><input type="checkbox" name="selIdx" value="<%= oKeyword.FItemList(i).Fidx %>"></td>
	<td><%= oKeyword.FItemList(i).Fidx %></td>
	<td><%= drawkeyword_gubunname(oKeyword.FItemList(i).fkeyword_gubun) %></td>
	<td><a href="main_TopKeyword_write.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Fkeyword %></a></td>
	<td align="left" style="word-break : break-all;">&nbsp;<a href="main_TopKeyword_write.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Flinkinfo %></a></td>
	<td><%=oKeyword.FItemList(i).Fisusing%></td>
	<td><input type="text" class="text" name="SortNo" value="<%=oKeyword.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td><%= FormatDate(oKeyword.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- 페이지 시작 -->
	<%
		if oKeyword.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oKeyword.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oKeyword.StartScrollPage to oKeyword.FScrollCount + oKeyword.StartScrollPage - 1

			if lp>oKeyword.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oKeyword.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</form>
</table>

<%
set oKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

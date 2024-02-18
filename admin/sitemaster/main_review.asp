<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TopReviewCls.asp"-->
<%
'###############################################
' PageName : main_topKeyword.asp
' Discription : 메인 탑 키워드 목록
' History : 2008.04.18 허진원 : 생성
'           2012.01.09 허진원 : 사이트구분 추가
'###############################################

'// 변수 선언 //
dim page, SearchString, strUse, siteDiv, lp

page = request("page")
SearchString = request("SearchString")
strUse = request("strUse")
if page = "" then page=1
if strUse = "" then strUse="Y"

dim oKeyword
set oKeyword = New CSearchKeyWord
oKeyword.FCurrPage = page
oKeyword.FPageSize=20
oKeyword.FRectUsing = strUse
oKeyword.FRectSearch = SearchString

oKeyword.GetSearchreview

dim i
%>
<script language='javascript'>
<!--
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
		alert('선택코멘트가 없습니다.');
		return;
	}

	if (upfrm.allusing.value=='Y'){
		var ret = confirm('선택 코멘트를 사용함으로 변경합니다');
	} else {
		var ret = confirm('선택 코멘트를 사용안함으로 변경합니다');
	}

	if (ret) {
		upfrm.mode.value="changeUsing";
		upfrm.action="doMainReview.asp";
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
		upfrm.action="doMainReview.asp";			<!--                                 -->
		upfrm.submit();
	} else {
		return;
	}
}


function RefreshCaFavKeyWordRec(){
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_review_JS.asp";
			refreshFrm.submit();
}

function frm_search()
{
	refreshFrm.target = "";
	refreshFrm.action = "main_review.asp";
}

	// 페이지 이동
	function goPage(pg)
	{
		document.refreshFrm.page.value=pg;
		document.refreshFrm.action="main_review.asp";
		document.refreshFrm.submit();
	}

// 카테고리 변경시 명령
function changecontent() {
}
//-->
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
				<td align="right">

					 사용여부
					<select class="select" name="strUse">
						<option value="all">전체</option>
						<option value="Y">사용</option>
						<option value="N">삭제</option>
					</select>
					/ ItemID 검색
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
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>(신규등록, 수정 후 꼭!! 옆의 버튼을 눌러주세요)</td>
	<td align="right">
		<select class="select" name="allusing">
			<option value="Y">선택 -> Y</option>
			<option value="N">선택 -> N</option>
		</select>
		<input type="button" class="button" value="적용" onclick="changeUsing(frmBuyPrc);">
		/
		<input type="button" class="button" value="순서변경" onclick="changeSort(frmBuyPrc);">
		/
		<input type="button" value="상품후기 추가" onclick="self.location='main_review_write.asp?menupos=<%= menupos %>'" class="button">
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
	<td width="50"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>카테고리</td>
	<td width="100">ItemID</td>
	<td>Comment</td>
	<td width="60">Item 가격</td>
	<td width="50">사용유무</td>
	<td width="50">순서</td>
	<td width="100">등록일</td>
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
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><input type="checkbox" name="selIdx" value="<%= oKeyword.FItemList(i).Fidx %>"></td>
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><%= oKeyword.FItemList(i).Fcate_large %></td>
	<td style="padding:10px;" <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><a href="<%=wwwurl%>/shopping/category_prd.asp?itemid=<%=oKeyword.FItemList(i).FItemid  %>"><%= oKeyword.FItemList(i).FItemid %></a></td>
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %> align="left" style="word-break : break-all;">&nbsp;<a href="main_review_write.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Fcomment %></a></td>
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><%= oKeyword.FItemList(i).Fsellcash %> 원</td>
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><%=oKeyword.FItemList(i).Fisusing%></td>
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><input type="text" class="text" name="SortNo" value="<%=oKeyword.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td <% If i < 10 Then response.write "bgcolor='#FFFFF0'" End If %>><%= FormatDate(oKeyword.FItemList(i).FRegdate,"0000.00.00") %></td>
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


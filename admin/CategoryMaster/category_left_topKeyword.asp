<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : Category_left_topKeyword.asp
' Discription : 카테고리 좌측 탑키워드 목록
' History : 2008.03.29 허진원 : 생성
'         : 2008.10.27 중카테고리 처리 추가(허진원)
'         : 2009.04.15 이미지 추가(허진원)
'###############################################

'// 변수 선언 //
dim page,cdl,cdm, SearchString, strUse, lp

page = request("page")
SearchString = request("SearchString")
strUse = request("strUse")
if page = "" then page=1
if strUse = "" then strUse="Y"
cdl = request("cdl")
cdm = request("cdm")

dim ocate
set ocate = New CCategoryKeyWord
ocate.FCurrPage = page
ocate.FPageSize=20
ocate.FRectCDL = cdl
ocate.FRectCDM = cdm
ocate.FRectUsing = strUse
ocate.FRectSearch = SearchString

ocate.GetCaFavKeyWord

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(iid,frm){
	window.open("/admin/pop/viewitemlist.asp?designerid=" + iid + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
}

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
		upfrm.action="doCateTopKeyword.asp";
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
		upfrm.action="doCateTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
	}
}


function RefreshCaFavKeyWordRec(){
	if (document.refreshFrm.cdl.value==""){
		alert("적용을 원하시는 카테고리를 선택해주세요!!");
	}
	else{
	var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_TopKeyword_JS.asp";
		 refreshFrm.submit();
	}
}

function RefreshChannelKeyWordRec() {
	if (document.refreshFrm.cdl.value==""){
		alert("적용을 원하시는 대카테고리를 선택해주세요!!");
	}
	else if (document.refreshFrm.cdm.value==""){
		alert("적용을 원하시는 중카테고리를 선택해주세요!!");
	}
	else{
	var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_TopKeyword_JS.asp";
		 refreshFrm.submit();
	}
}

function frm_search()
{
	refreshFrm.target = "";
	refreshFrm.action = "category_left_topKeyword.asp";
}

	// 페이지 이동
	function goPage(pg)
	{
		document.refreshFrm.page.value=pg;
		document.refreshFrm.action="category_left_topKeyword.asp";
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
				<td align="left">
					카테고리 <% DrawSelectBoxCategoryLarge "cdl", cdl %>
					<% if cdl="110" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
				</td>
				<td align="right">
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
		<%
			if cdl<>"" then
				if cdl<>"110" then
		%>
		<a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>(신규등록, 수정 후 꼭!! 카테고리 선택 후 옆의 버튼을 눌러주세요)
		<%
				elseif cdm<>"" then
		%>
		<a href="javascript:RefreshChannelKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>(신규등록, 수정 후 꼭!! 카테고리 선택 후 옆의 버튼을 눌러주세요)
		<%
				end if
			end if
		%>
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
		<input type="button" value="아이템 추가" onclick="self.location='category_left_topKeyword_write.asp?menupos=<%= menupos %>'" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=ocate.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=ocate.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>카테고리</td>
	<td>이미지</td>
	<td>키워드</td>
	<td>링크정보</td>
	<td>사용유무</td>
	<td>순서</td>
	<td>등록일</td>
</tr>
<%	if ocate.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to ocate.FResultCount-1
%>
<tr align="center" bgcolor="<% if ocate.FItemList(i).Fisusing = "Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><input type="checkbox" name="selIdx" value="<%= ocate.FItemList(i).Fidx %>"></td>
	<td><%
		Response.Write ocate.FItemList(i).FCDL_Nm
		if Not(ocate.FItemList(i).FCDM_Nm="" or isNull(ocate.FItemList(i).FCDM_Nm)) then
			Response.Write "<br>/" & ocate.FItemList(i).FCDM_Nm
		end if
	%></td>
	<td>
	<% if Not(ocate.FItemList(i).FImageSmall="" or isNull(ocate.FItemList(i).FImageSmall)) then %>
		<img src="<%=ocate.FItemList(i).FImageSmall%>" border="0" width="50">
	<% else %>
		<img src="http://fiximage.10x10.co.kr/web2008/category/blank.gif" border="0" width="50">
	<% end if %>
	</td>
	<td><a href="category_left_topKeyword_write.asp?idx=<%= ocate.FItemList(i).Fidx %>&page=<%=page%>"><%= ocate.FItemList(i).Fkeyword %></a></td>
	<td align="left" style="word-break : break-all;">&nbsp;<a href="category_left_topKeyword_write.asp?idx=<%= ocate.FItemList(i).Fidx %>&page=<%=page%>"><%= ocate.FItemList(i).Flinkinfo %></a></td>
	<td><%=ocate.FItemList(i).Fisusing%></td>
	<td><input type="text" class="text" name="SortNo" value="<%=ocate.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td><%= FormatDate(ocate.FItemList(i).FRegdate,"0000.00.00") %></td>
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
		if ocate.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & ocate.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + ocate.StartScrollPage to ocate.FScrollCount + ocate.StartScrollPage - 1

			if lp>ocate.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if ocate.HasNextScroll then
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
set ocate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


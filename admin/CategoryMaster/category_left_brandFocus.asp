<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp" -->
<%
'###############################################
' PageName : Category_left_BrandFocus.asp
' Discription : 카테고리 좌측 브랜드 포커스 목록
' History : 2008.04.04 허진원 : 생성
'###############################################

'// 변수 선언
dim cdl, page, i, lp
cdl = request("cdl")
page = request("page")

if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=10
omd.FRectCDL = cdl
omd.GetBrandFocusList

%>
<script language='javascript'>
<!--
function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmarr;

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
	var frm = document.frmarr;

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

function delitems(upfrm){
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	if (confirm('선택 아이템을 삭제하시겠습니까?')) {
		upfrm.mode.value="del";
		upfrm.action="doCategoryLeftbrandFocus.asp";
		upfrm.submit();
	}
}

function RefreshLeftbrandFocusRec(){
	if (document.refreshFrm.cdl.value == ""){
		alert("카테고리를 선택해주세요");
		document.refreshFrm.cdl.focus();
	}
	else{
		 var popwin = window.open('','refreshPop','');
		 popwin.focus();
		 refreshFrm.target = "refreshPop";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_left_brandFocus_JS.asp";
		 refreshFrm.submit();
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
		upfrm.action="doCategoryLeftbrandFocus.asp";
		upfrm.submit();
	} else {
		return;
	}
}

// 브랜드 추가 처리
function addbrandFocus(upfrm)
{
	if(!upfrm.cdl.value) {
		alert("추가할 대상 카테고리를 선택해주세요.");
		return;
	}

	if(!upfrm.makerid.value) {
		alert("브랜드ID를 입력해주세요.");
		return;
	}

	if (confirm('지정하신 브랜드를 추가하시겠습니까?')) {
		upfrm.mode.value="add";
		upfrm.action="doCategoryLeftbrandFocus.asp";
		upfrm.submit();
	}
}

// 브랜드 검색 팝업
function popBrandSearch(fm,tg){
	var popup_item = window.open("/admin/member/popBrandSearch.asp?frmName=" + fm + "&compName=" + tg, "popup_brand", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="category_left_brandFocus.asp";
	document.refreshFrm.submit();
}

// 카테고리 변경시 명령
function changecontent(){
	document.frmarr.cdl.value=refreshFrm.cdl.value;
}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="category_left_brandFocus.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		카테고리 <% DrawSelectBoxCategoryLarge "cdl", cdl %>
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
<form name="frmarr" method="get" action="doCategoryLeftbrandFocus.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td><input type="button" value="선택아이템삭제" onclick="delitems(frmarr);" class="button"></td>
	<td align="right">
		<img src="/images/icon_reload.gif" onClick="RefreshLeftbrandFocusRec()" style="cursor:pointer" align="absmiddle" alt="html만들기">
		프론트에 적용 /
		<input type="button" class="button" value="순서변경" onclick="changeSort(frmarr);">
		/
		브랜드ID
		<input type="text" class="text" name="makerid" value="" onClick="popBrandSearch('frmarr','makerid')" style="cursor:pointer">
		<input type="button" value="아이템 추가" onclick="addbrandFocus(frmarr)" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		검색결과 : <b><%=omd.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=omd.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>번호</td>
	<td>카테고리명</td>
	<td>업체명</td>
	<td>이미지</td>
	<td>순서</td>
</tr>
<%	if omd.FResultCount < 1 then %>
<tr>
	<td colspan="6" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to omd.FResultCount-1
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="selIdx" value="<%= omd.FItemList(i).Fidx %>"></td>
	<td><%= omd.FItemList(i).Fidx %></td>
	<td><%= omd.FItemList(i).Fcode_nm %></td>
	<td><%= omd.FItemList(i).Fmakerid %></td>
	<td><img src="<%= omd.FItemList(i).FImageSmall %>"><img src="<%= omd.FItemList(i).Ftitleimgurl %>" ></td>
	<td><input type="text" class="text" name="SortNo" value="<%=omd.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
</tr>
<%
		next
	end if
%>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">
	<!-- 페이지 시작 -->
	<%
		if omd.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & omd.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1

			if lp>omd.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if omd.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
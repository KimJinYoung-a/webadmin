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
' PageName : Category_left_bestBrand.asp
' Discription : 카테고리 좌측 베스트 브랜드 목록
' History : 2008.04.02 허진원 : 생성
'			2008.05.06 김정인 수정 [ fnSearch함수,실서버적용조건 추가(isUsing='Y'))
'###############################################

'// 변수 선언
dim cdl, page, isusing, i, lp
cdl = request("cdl")
page = request("page")
isusing = request("isusing")

if page="" then page=1
if isusing="" then isusing="Y"

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=10
omd.FRectCDL = cdl
omd.FRectIsusing = isusing
omd.GetBestBrandList

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
		upfrm.action="doCategoryLeftBestBrand.asp";
		upfrm.submit();
	}
}

function RefreshLeftBestBrandRec(){
	if (document.refreshFrm.cdl.value == ""){
		alert("카테고리를 선택해주세요");
		document.refreshFrm.cdl.focus();
	}
	else{
		 var popwin = window.open('','refreshPop','');
		 popwin.focus();
		 refreshFrm.target = "refreshPop";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_left_bestBrand_JS.asp";
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
		upfrm.action="doCategoryLeftBestBrand.asp";
		upfrm.submit();
	} else {
		return;
	}
}

// 브랜드 추가 페이지로 이동
function addBestBrand()
{
	document.frmarr.cdl.value = document.refreshFrm.cdl.value;
	document.frmarr.mode.value = "add";
	document.frmarr.action="category_left_bestBrand_write.asp";
	document.frmarr.submit();
}

// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="category_left_bestBrand.asp";
	document.refreshFrm.submit();
}
function fnSearch()
{
	document.refreshFrm.action='category_left_BestBrand.asp';
	document.refreshFrm.target='';
	document.refreshFrm.submit();
}

// 카테고리 변경시 명령
function changecontent(){ }
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="category_left_BestBrand.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		카테고리 <% DrawSelectBoxCategoryLarge "cdl", cdl %> /
		사용유무 <select name="isusing" class="select"><option value="Y">Yes</option><option value="N">No</option></select>
		<script language="javascript">
			document.refreshFrm.isusing.value="<%=isusing%>";
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" onclick="fnSearch();" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmarr" method="get" action="doCategoryLeftBestBrand.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="cdl" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td><input type="button" value="선택아이템삭제" onclick="delitems(frmarr);" class="button"></td>
	<td align="right">
		<img src="/images/icon_reload.gif" onClick="RefreshLeftBestBrandRec()" style="cursor:pointer" align="absmiddle" alt="html만들기">
		프론트에 적용 /
		<input type="button" class="button" value="순서변경" onclick="changeSort(frmarr);">
		/
		<input type="button" value="아이템 추가" onclick="addBestBrand()" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		검색결과 : <b><%=omd.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=omd.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>카테고리명</td>
	<td>업체명</td>
	<td>이미지</td>
	<td>사용유무</td>
	<td>순서</td>
	<td>등록일</td>
</tr>
<%	if omd.FResultCount < 1 then %>
<tr>
	<td colspan="7" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to omd.FResultCount-1
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="selIdx" value="<%= omd.FItemList(i).Fidx %>"></td>
	<td><%= omd.FItemList(i).Fcode_nm %></td>
	<td><%= omd.FItemList(i).Fmakerid %></td>
	<td>
		<a href="/admin/categorymaster/category_left_bestbrand_write.asp?mode=edit&idx=<%= omd.FItemList(i).Fidx %>&page=<%=page%>">
		<img src="<%= staticImgUrl & "/left/bestbrand/" & omd.FItemList(i).Fimage %>" border="0" height="60"></a>
	</td>
	<td><%= omd.FItemList(i).Fisusing %></td>
	<td><input type="text" class="text" name="SortNo" value="<%=omd.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td><%= FormatDateTime(omd.FItemList(i).Fregdate,2) %></td>
</tr>
<%
		next
	end if
%>
<tr bgcolor="#FFFFFF">
	<td colspan="7" align="center">
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
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/categoryCls.asp"-->
<%
'###############################################
' PageName : Category_left_Bigchance.asp
' Discription : 카테고리 좌측 빅찬스 목록
' History : 2008.03.31 허진원 : 생성
'           2008.07.25 허진원 수정 : 상품 정렬순서 추가
'###############################################

dim cdl, cdm, page, lp
cdl = request("cdl")
cdm = request("cdm")
page = request("page")

if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectCDL = cdl
omd.FRectCDM = cdm
omd.GetCategoryBigChanceList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(tgf){
	<% if cdl<>"110" then %>
		if (document.refreshFrm.cdl.value == ""){
			alert("카테고리를 선택해 주세요!");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdl.value=="110") {
			alert("감성채널은 검색을 실행하여 중카테고리를 선택하셔야합니다.");
		} else {
			var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.refreshFrm.cdl.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
	<% else %>
		if (document.refreshFrm.cdl.value == ""){
			alert("카테고리를 선택해 주세요!");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdm.value == ""){
			alert("중카테고리를 선택해 주세요!");
			document.refreshFrm.cdm.focus();
		} else {
			var popup_item = window.open("/common/pop_CateItemList.asp?cdl=" + document.refreshFrm.cdl.value + "&cdm=" + document.refreshFrm.cdm.value + "&target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
			popup_item.focus();
		}
	<% end if %>
}

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
			if(frm.selIdx[i].checked) frm.arrSort.value = frm.arrSort.value + frm.sortNo[i].value + ",";
		}
	} else {
		pass = ((pass)||(frm.selIdx.checked));
		frm.arrSort.value = frm.sortNo.value;
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
		upfrm.action="doCategoryLeftBigchance.asp";
		upfrm.submit();
	}
}

// 선택아이템의 정렬번호 적용(2008.07.25; 허진원 추가)
function submitSortNo(upfrm) {
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	if (confirm('선택 아이템의 정렬번호를 적용하시겠습니까?')) {
		upfrm.mode.value="sort";
		upfrm.action="doCategoryLeftBigchance.asp";
		upfrm.submit();
	}	
}


function AddIttems(){
	<% if cdl<>"110" then %>
		if (document.refreshFrm.cdl.value == ""){
			alert("카테고리를 선택해주세요");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdl.value == "110"){
			alert("감성채널은 검색을 실행하여 중카테고리를 선택하셔야합니다.");
		} else if (confirm(frmarr.itemidarr.value + '아이템을 추가하시겠습니까?')){
			frmarr.itemid.value = frmarr.itemidarr.value;
			frmarr.cdl.value = refreshFrm.cdl.value;
			frmarr.mode.value="add";
			frmarr.submit();
		}
	<% else %>
		if (document.refreshFrm.cdl.value == ""){
			alert("카테고리를 선택해주세요");
			document.refreshFrm.cdl.focus();
		} else if(document.refreshFrm.cdm.value == ""){
			alert("중카테고리를 선택해주세요");
			document.refreshFrm.cdm.focus();
		} else if (confirm(frmarr.itemidarr.value + '아이템을 추가하시겠습니까?')){
			frmarr.itemid.value = frmarr.itemidarr.value;
			frmarr.cdl.value = refreshFrm.cdl.value;
			frmarr.cdm.value = refreshFrm.cdm.value;
			frmarr.mode.value="add";
			frmarr.submit();
		}
	<% end if %>
}

function RefreshMainRotateEventRec(){
	<% if cdl<>"110" then %>
		if (document.refreshFrm.cdl.value == ""){
			alert("카테고리를 선택해주세요");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdl.value == "110"){
			alert("감성채널은 검색을 실행하여 중카테고리를 선택하셔야합니다.");
		} else{
			 var popwin = window.open('','refreshPop','');
			 popwin.focus();
			 refreshFrm.target = "refreshPop";
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_category_left_bigchance_JS.asp";
			 refreshFrm.submit();
		}
	<% else %>
		if (document.refreshFrm.cdl.value == ""){
			alert("카테고리를 선택해주세요");
			document.refreshFrm.cdl.focus();
		} else if (document.refreshFrm.cdm.value == ""){
			alert("중카테고리를 선택해주세요");
			document.refreshFrm.cdm.focus();
		} else {
			 var popwin = window.open('','refreshPop','');
			 popwin.focus();
			 refreshFrm.target = "refreshPop";
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_channel_left_bigchance_JS.asp";
			 refreshFrm.submit();
		}
	<% end if %>
	location.reload();
}

// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="category_left_Bigchance.asp";
	document.refreshFrm.submit();
}

// 카테고리 변경시 명령
function changecontent(){}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" action="category_left_Bigchance.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		카테고리 <% DrawSelectBoxCategoryLarge "cdl", cdl %>
		<% if cdl="110" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
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
<form name="frmarr" method="post" action="doCategoryLeftBigchance.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="cdl" value="">
<input type="hidden" name="cdm" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>
				<input type="text" name="itemidarr" value="" size="80" class="text">
				<input type="button" value="아이템 직접 추가" onclick="AddIttems()" class="button">
			</td>
			<td align="right">
				<img src="/images/icon_reload.gif" onClick="RefreshMainRotateEventRec()" style="cursor:pointer" align="absmiddle" alt="html만들기">
				프론트에 적용
			</td>
		</tr>
		<tr>
			<td><input type="button" value="선택아이템 삭제" onClick="delitems(frmarr)" class="button"></td>
			<td align="right">
				<input type="button" value="정렬번호 적용" onClick="submitSortNo(frmarr)" class="button">
				<input type="button" value="아이템 추가" onclick="popItemWindow('frmarr.itemidarr')" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=omd.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=omd.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>카테고리명</td>
	<td>Image</td>
	<td>ItemID</td>
	<td>제품명</td>
	<td>할인률</td>
	<td>상태</td>
	<td>정렬번호</td>
</tr>
<%	if omd.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to omd.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="selIdx" value="<%= omd.FItemList(i).Fidx %>"></td>
	<td align="center"><%
		Response.Write omd.FItemList(i).Fcode_nm
		if Not(omd.FItemList(i).FCDM_Nm="" or isNull(omd.FItemList(i).FCDM_Nm)) then
			Response.Write "<br>/" & omd.FItemList(i).FCDM_Nm
		end if
	%>
	</td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center"><% if omd.FItemList(i).FsailYn="Y" then Response.Write formatPercent(1-omd.FItemList(i).FsailPrice/omd.FItemList(i).ForgPrice,1) %></td>
	<td align="center"><% if omd.FItemList(i).FsellYn<>"Y" then Response.Write "품절" %></td>
	<td align="center"><input type="text" name="sortNo" value="<%=omd.FItemList(i).FsortNo %>" size="3" style="text-align:right"></td>
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
</form>
</table>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

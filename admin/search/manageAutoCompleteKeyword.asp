<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim orgkeyword, relatedKeyword, modiType, useYN, page, isEnginMayAssign
dim i
dim research

research 		= requestCheckvar(request("research"),32)
orgkeyword 		= Trim(requestCheckvar(request("orgkeyword"),32))
relatedKeyword 	= Trim(requestCheckvar(request("relatedKeyword"),32))
modiType 		= Trim(requestCheckvar(request("modiType"),32))
useYN 			= Trim(requestCheckvar(request("useYN"),32))
page			= requestCheckvar(request("page"),10)
isEnginMayAssign	= requestCheckvar(request("isEnginMayAssign"),10)

if (research = "") then
	useYN = "Y"
end if

if (page="") then page = 1


'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.FPageSize = 100
osearchKeyword.FCurrPage = page

osearchKeyword.FRectOrgKeyword		= orgkeyword
osearchKeyword.FRectRelatedKeyword	= relatedKeyword
osearchKeyword.FRectModiType		= modiType
osearchKeyword.FRectUseYN			= useYN
osearchKeyword.FRectIsEnginMayAssign	= isEnginMayAssign

osearchKeyword.GetAutoCompleteKeywordList

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsPopAutoCompleteKeywordAdd() {
    var popwin = window.open('popAutoCompleteKeywordModiNew.asp','popAutoCompleteKeywordModiNew','width=330,height=220,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsPopAutoCompleteKeywordModi(rect) {
    var popwin = window.open('popAutoCompleteKeywordModiNew.asp?rect=' + rect,'popAutoCompleteKeywordModiNew','width=330,height=220,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsDelAutoCompleteKeyword(rect) {
	var ret = confirm("삭제하시겠습니까?");
	if(ret){
		var frm = document.frmAct;
		frm.mode.value = "delevt";
		frm.rect.value = rect;
		frm.submit();
	}
}

function jsUseAutoCompleteKeyword(rect) {
	var ret = confirm("사용전환 하시겠습니까?");
	if(ret){
		var frm = document.frmAct;
		frm.mode.value = "useevt";
		frm.rect.value = rect;
		frm.submit();
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30" >
			키워드 : <input type="text" class="text" name="orgkeyword" value="<%= orgkeyword %>">
			&nbsp;
			사용여부 :
			<select class="select" name="useYN">
				<option value=""></option>
				<option value="Y" <% if (useYN = "Y") then %>selected<% end if %> >Y</option>
				<option value="N" <% if (useYN = "N") then %>selected<% end if %> >N</option>
			</select>
			<!--
			&nbsp;
			<input type="checkbox" name="isEnginMayAssign" value="1" <%= CHKIIF(isEnginMayAssign="1", "checked", "") %>> 검색엔진 적용 검색어만 표시
			-->
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<input type="button" class="button" value=" 키워드등록 " onClick="jsPopAutoCompleteKeywordAdd()">

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			검색결과 : <b><%= osearchKeyword.FTotalcount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50" height="30">IDX</td>
		<td width="150">키워드</td>
		<td width="80">검색횟수<br />(3일)</td>
		<td width="80">검색횟수<br />(7일)</td>
		<td width="80">검색횟수<br />(누적)</td>
		<td width="80">검색결과<br />(연관)</td>
		<td width="50">수정여부</td>
		<td width="50">사용여부</td>
		<td width="50">가중치<br />부여</td>
		<td width="150">등록일</td>
		<td width="150">최종수정</td>
		<td>비고</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).FRowNum %>
		</td>
		<td align="center"><%= osearchKeyword.FItemList(i).Frect %></td>
		<td><%= osearchKeyword.FItemList(i).FrecentAcctCNT %></td>
		<td><%= osearchKeyword.FItemList(i).FrecentAcctCNT2 %></td>
		<td><%= osearchKeyword.FItemList(i).FacctCNT %></td>
		<td><%= osearchKeyword.FItemList(i).FlastrectCNT %></td>
		<td><%= osearchKeyword.FItemList(i).GetIsAutoTypeName %></td>
		<td><%= osearchKeyword.FItemList(i).GetIsUsingTypeName %></td>
		<td><%= osearchKeyword.FItemList(i).FUserAddCNT %></td>
		<td><%= osearchKeyword.FItemList(i).Fregdate %></td>
		<td><%= osearchKeyword.FItemList(i).Flastupdate %></td>
		<td align="left">
			<% if (osearchKeyword.FItemList(i).FisUsingType = 0) then %>
			<input type="button" class="button" value=" 사용 " onClick="jsUseAutoCompleteKeyword('<%= osearchKeyword.FItemList(i).Frect %>')">
			<% else %>
			<input type="button" class="button" value=" 삭제 " onClick="jsDelAutoCompleteKeyword('<%= osearchKeyword.FItemList(i).Frect %>')">
			<input type="button" class="button" value=" 가중치부여 " onClick="jsPopAutoCompleteKeywordModi('<%= osearchKeyword.FItemList(i).Frect %>')">
			<% end if %>
		</td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="15">
			검색결과가 없습니다.
		</td>
	</tr>
	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if osearchKeyword.HasPreScroll then %>
			<a href="javascript:NextPage('<%= osearchKeyword.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for i=0 + osearchKeyword.StartScrollPage to osearchKeyword.FScrollCount + osearchKeyword.StartScrollPage - 1 %>
				<% if i>osearchKeyword.FTotalPage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if osearchKeyword.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<% end if %>
</table>
<%
set osearchKeyword = Nothing
%>
<form name="frmAct" method="post" action="manageAutoCompleteKeyword_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="rect" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->

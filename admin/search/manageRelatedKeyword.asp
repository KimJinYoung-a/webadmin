<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim orgkeyword, relatedKeyword, modiType, useYN, page
dim i
dim research

research 		= request("research")
orgkeyword 		= Trim(request("orgkeyword"))
relatedKeyword 	= Trim(request("relatedKeyword"))
modiType 		= Trim(request("modiType"))
useYN 			= Trim(request("useYN"))
page			= requestCheckvar(request("page"),10)

if (research = "") then
	useYN = "Y"
end if

if (page="") then page = 1


'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.FPageSize = 50
osearchKeyword.FCurrPage = page

osearchKeyword.FRectOrgKeyword		= orgkeyword
osearchKeyword.FRectRelatedKeyword	= relatedKeyword
osearchKeyword.FRectModiType		= modiType
osearchKeyword.FRectUseYN			= useYN

osearchKeyword.getRelatedKeywordModi_Paging

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsPopRelatedKeywordAdd() {
    var popwin = window.open('popRelatedKeywordAdd.asp','jsPopRelatedKeywordAdd','width=330,height=220,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsDelRelatedKeyword(idx) {
	var ret = confirm("삭제하시겠습니까?");
	if(ret){
		var frm = document.frmAct;
		frm.mode.value = "del";
		frm.idx.value = idx;
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
			원검색어 : <input type="text" class="text" name="orgkeyword" value="<%= orgkeyword %>">
			&nbsp;
			연관검색어 : <input type="text" class="text" name="relatedKeyword" value="<%= relatedKeyword %>">
			&nbsp;
			구분 :
			<select class="select" name="modiType">
				<option value=""></option>
				<option value="A" <% if (modiType = "A") then %>selected<% end if %> >추가</option>
				<option value="D" <% if (modiType = "D") then %>selected<% end if %> >제외</option>
			</select>
			사용여부 :
			<select class="select" name="useYN">
				<option value=""></option>
				<option value="Y" <% if (useYN = "Y") then %>selected<% end if %> >Y</option>
				<option value="N" <% if (useYN = "N") then %>selected<% end if %> >N</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30">
			* 검색엔진 반영은 07, 11, 15, 19 시에 이루어집니다.
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<input type="button" class="button" value=" 등록 " onClick="jsPopRelatedKeywordAdd()">

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			검색결과 : <b><%= osearchKeyword.FTotalcount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= osearchKeyword.FTotalPage %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50" height="30">IDX</td>
		<td width="150">원검색어</td>
		<td width="150">연관검색어</td>
		<td width="80">가중치</td>
		<td width="50">구분</td>
		<td width="100">등록자</td>
		<td width="50">사용여부</td>
		<td width="150">등록일</td>
		<td>비고</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).Fidx %>
		</td>
		<td align="center"><%= osearchKeyword.FItemList(i).ForgKeyword %></td>
		<td align="center"><%= osearchKeyword.FItemList(i).FrelatedKeyword %></td>
		<td align="center">
			<% if (osearchKeyword.FItemList(i).FmodiType = "A") then %>
			<%= osearchKeyword.FItemList(i).FsearchCount %>
			<% end if %>
		</td>
		<td align="center">
			<% if (osearchKeyword.FItemList(i).FmodiType = "D") then %><font color="red"><% end if %>
			<%= osearchKeyword.FItemList(i).GetModiTypeName %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).Freguserid %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).FuseYN %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).Fregdate %>
		</td>
		<td align="left">
			<input type="button" class="button" value=" 삭제 " onClick="jsDelRelatedKeyword(<%= osearchKeyword.FItemList(i).Fidx %>)">
		</td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="9">
			검색결과가 없습니다.
		</td>
	</tr>
	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21" align="center">
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
<form name="frmAct" method="post" action="manageRelatedKeyword_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

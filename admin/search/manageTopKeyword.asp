<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim topKeyword, modiType, useYN
dim i
dim research

research 		= request("research")
topKeyword 		= Trim(request("topKeyword"))
modiType 		= Trim(request("modiType"))
useYN 			= Trim(request("useYN"))

if (research = "") then
	useYN = "Y"
end if


'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.FRectKeyword			= topKeyword
osearchKeyword.FRectModiType		= modiType
osearchKeyword.FRectUseYN			= useYN

osearchKeyword.getTopKeywordModi

%>

<script language='javascript'>

function jsPopTopKeywordAdd() {
    var popwin = window.open('popTopKeywordAdd.asp','jsPopTopKeywordAdd','width=330,height=220,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsDelTopKeyword(idx) {
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
			검색어 : <input type="text" class="text" name="topKeyword" value="<%= topKeyword %>">
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
			* 검색엔진 반영은 매 2시간마다 이루어집니다.
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<input type="button" class="button" value=" 등록 " onClick="jsPopTopKeywordAdd()">

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50" height="30">IDX</td>
		<td width="150">검색어</td>
		<td width="80">가중치</td>
		<td width="50">구분</td>
		<td width="100">등록자</td>
		<td width="50">사용여부</td>
		<td width="150">등록일</td>
		<td>비고</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FTotalCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).Fidx %>
		</td>
		<td align="center"><%= osearchKeyword.FItemList(i).FtopKeyword %></td>
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
			<input type="button" class="button" value=" 삭제 " onClick="jsDelTopKeyword(<%= osearchKeyword.FItemList(i).Fidx %>)">
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
	<% end if %>
</table>
<%
set osearchKeyword = Nothing
%>
<form name="frmAct" method="post" action="manageTopKeyword_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

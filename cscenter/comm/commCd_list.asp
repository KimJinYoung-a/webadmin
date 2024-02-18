<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [cs]공통코드관리
' Hieditor : 이상구 생성
'			 2023.08.28 한용민 수정(고객노출여부 추가, 소스표준코드로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/CsCommCdcls.asp"-->
<%
dim comm_cd, page, groupCd, searchKey, searchString, comm_isDel, sortType, oComm, i, lp, bgcolor, strUsing
dim dispyn
	comm_cd     = requestCheckVar(request("comm_cd"),32)
	page        = requestCheckVar(getNumeric(request("page")),9)
	groupCd     = requestCheckVar(request("groupCd"),32)
	searchKey   = requestCheckVar(request("searchKey"),32)
	searchString = requestCheckVar(request("searchString"),32)
	comm_isDel  = requestCheckVar(request("comm_isDel"),32)
	sortType	= requestCheckVar(request("sortType"),2)
	dispyn	= requestCheckVar(request("dispyn"),2)

if page="" then page=1
if searchKey="" then searchKey="comm_name"
if sortType="" then sortType="sa"

set oComm = new CCommCd
	oComm.FCurrPage = page
	oComm.FPageSize = 50
	oComm.FRectgroupCd = groupCd
	oComm.FRectsearchKey = searchKey
	oComm.FRectsearchString = searchString
	oComm.FSortType = sortType
	oComm.FRectisDel = comm_isDel
	oComm.FRectdispyn = dispyn
	oComm.GetCommList
%>
<script type='text/javascript'>

function popCsAsGubunHelpEdit(icomm_cd){
	var popwin = window.open('popCsAsGubunHelpEdit.asp?comm_cd=' + icomm_cd,'popCsAsGubunHelpEdit','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function chk_form(frm){
	if(!frm.searchKey.value)
	{
		alert("검색 조건을 선택해주십시오.");
		frm.searchKey.focus();
		return false;
	}
	else if(!frm.searchString.value)
	{
		alert("검색어를 입력해주십시오.");
		frm.searchString.focus();
		return false;
	}

	frm.page.value= 1;
	frm.submit();
}

function goPage(pg){
	var frm = document.frm_search;

	frm.page.value= pg;
	frm.submit();
}

function chgSort(t,s) {
	var frm = document.frm_search;
	frm.sortType.value= t+s;
	frm.submit();
}

function popCommCdReg(){
	var popwin = window.open('/cscenter/comm/commCd_write.asp?menupos=<%=menupos%>','popCommCdReg','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popCommCdEdit(comm_cd){
	var popwin = window.open('/cscenter/comm/CommCd_modi.asp?comm_cd='+comm_cd+'&menupos=<%=menupos%>','popCommCdEdit','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm_search" method="GET" action="CommCd_list.asp" onSubmit="return chk_form(this)" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>" />
<input type="hidden" name="menupos" value="<%=menupos%>" />
<input type="hidden" name="sortType" value="<%=sortType%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			* 상태:
			<select class="select" name="comm_isDel" onChange="goPage(1)">
				<option value="">전체</option>
				<option value="N">사용</option>
				<option value="Y">삭제</option>
			</select>
			&nbsp;
			* 그룹:
			<select class="select" name="groupCd" onChange="goPage(1)">
				<option value="">전체</option>
				<%
				' 아주 예전(2010년이전)에 만들어진 코드가 꼬여서 코드를 개선해 볼려고 했으나 cs는 계속 운영중이고, 저장값이 있어서 수정이 불가능함.
				' comm_group "z999" 코드에 "공통" 구분자 코드를 추가할 방법이 없어서 검색이라도 되게 공통값은 우선 수기로 박음.
				' comm_cd 값과.. comm_group 코드값이 순환적으로 계속 물고 물리는 구조로 되어 있음.
				' comm_cd 값은 db_cs.dbo.tbl_new_as_list 테이블 gubun01 필드에 입력되는 구조임.
				%>
				<option value="C004" <% if groupCd="C004" then response.write "selected" %>>공통</option>
				<%= oComm.optGroupCd(groupCd)%>
			</select>
			&nbsp;
			* 노출여부 : <% drawSelectBoxisusingYN "dispyn", dispyn,"" %>
			&nbsp;
			* 검색:
			<select class="select" name="searchKey">
				<option value="comm_cd">공통코드</option>
				<option value="comm_name">코드명</option>
			</select>
			<script language="javascript">
				document.frm_search.comm_isDel.value="<%=comm_isDel%>";
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			&nbsp;
			<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="goPage(1)" />
		</td>
	</tr>
</table>
</form>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">공통코드가 추가되고, 공통코드의 하위코드는 추가되지 않습니다.</font>
	</td>
	<td align="right">
		<input type="button" value="신규등록" onclick="popCommCdReg();" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oComm.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oComm.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="60">순서 <span onclick="chgSort('s','<%=chkIIF(left(sortType,1)="s",chkIIF(right(sortType,1)="d","a","d"),"a")%>')" style="cursor:pointer;"><%=chkIIF(left(sortType,1)="s",chkIIF(right(sortType,1)="d","▲","▼"),"▽")%></span></td>
	<td align="center" width="140">그룹 <span onclick="chgSort('g','<%=chkIIF(left(sortType,1)="g",chkIIF(right(sortType,1)="d","a","d"),"a")%>')" style="cursor:pointer;"><%=chkIIF(left(sortType,1)="g",chkIIF(right(sortType,1)="d","▲","▼"),"▽")%></span></td>
	<td align="center" width="80">공통코드 <span onclick="chgSort('c','<%=chkIIF(left(sortType,1)="c",chkIIF(right(sortType,1)="d","a","d"),"a")%>')" style="cursor:pointer;"><%=chkIIF(left(sortType,1)="c",chkIIF(right(sortType,1)="d","▲","▼"),"▽")%></span></td>
	<td align="center">코드명</td>
	<td align="center" width="90">프론트노출여부</td>
	<td align="center" width="50">Color</td>
	<td align="center" width="50">상태</td>
	<td align="center" width="100">비고</td>
</tr>
<%
for lp=0 to oComm.FResultCount - 1
	if oComm.FItemList(lp).Fcomm_isDel="<font color=darkblue>사용</font>" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#E0E0E0"
	end if
%>
<tr align="center" bgcolor="<%=bgcolor%>">
	<td><%= oComm.FItemList(lp).Fsortno %></td>
	<td><%= oComm.FItemList(lp).Fgroup_name %></td>
	<td><%= oComm.FItemList(lp).Fcomm_cd %></td>
	<td align="left"><%= db2html(oComm.FItemList(lp).Fcomm_name) %></td>
	<td ><%= oComm.FItemList(lp).fdispyn %></td>
	<td ><%= oComm.FItemList(lp).Fcomm_color %></td>
	<td><%= oComm.FItemList(lp).Fcomm_isDel %></td>
	<td>
		<input type="button" value="수정" onclick="popCommCdEdit('<%= oComm.FItemList(lp).Fcomm_cd %>');" class="button">

		<% if Left(oComm.FItemList(lp).Fcomm_cd,1)="A" then %>
			<br><input type="button" value="수정(안내설명)" onclick="popCsAsGubunHelpEdit('<%= oComm.FItemList(lp).Fcomm_cd %>');" class="button">
		<% end if %>
	</td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<!-- 페이지 시작 -->
		<%
			if oComm.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oComm.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if

			for i=0 + oComm.StartScrollPage to oComm.FScrollCount + oComm.StartScrollPage - 1

				if i>oComm.FTotalpage then Exit for

				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if

			next

			if oComm.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
		%>
		<!-- 페이지 끝 -->
	</td>
</tr>
</table>

<%
set oComm = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
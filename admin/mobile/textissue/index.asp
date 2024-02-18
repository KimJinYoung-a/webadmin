<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TextIssueCls.asp"-->
<%
'###############################################
' Discription : 텍스트 이슈
' History : 2013.12.14 이종화
'###############################################

'// 변수 선언 //
dim page, SearchString, strUse, siteDiv, lp

page = request("page")
SearchString = request("SearchString")
strUse = request("strUse")
if page = "" then page=1
if strUse = "" then strUse="Y"
if siteDiv = "" then siteDiv="T"

dim oKeyword
set oKeyword = New CSearchKeyWord
oKeyword.FCurrPage = page
oKeyword.FPageSize=20
oKeyword.FRectUsing = strUse
oKeyword.FRectSearch = SearchString

oKeyword.GetSearchKeyWord

dim i
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
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
		upfrm.action="dotextissue.asp";
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
		upfrm.action="dotextissue.asp";
		upfrm.submit();
	} else {
		return;
	}
}


function RefreshCaFavKeyWordRec(){
	if(confirm("모바일- 텍스트이슈에 적용하시겠습니까?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_textissue_xml.asp";
			refreshFrm.submit();
	}
}

function frm_search()
{
	refreshFrm.target = "";
	refreshFrm.action = "index.asp";
}

	// 페이지 이동
	function goPage(pg)
	{
		document.refreshFrm.page.value=pg;
		document.refreshFrm.action="index.asp";
		document.refreshFrm.submit();
	}

// 카테고리 변경시 명령
function changecontent() {
}

$(function(){
	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='SortNo']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='SortNo']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});
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
<input type="hidden" name="siteDiv" value="<%=siteDiv%>">
<input type="hidden" name="arrSort" value="">
<tr>
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>(신규등록, 수정 후 꼭!! 옆의 버튼을 눌러주세요) ※5개 까지만 저장 됩니다.※</td>
	<td align="right">
		<select class="select" name="allusing">
			<option value="Y">선택 -> Y</option>
			<option value="N">선택 -> N</option>
		</select>
		<input type="button" class="button" value="적용" onclick="changeUsing(frmBuyPrc);">
		/
		<input type="button" class="button" value="순서변경" onclick="changeSort(frmBuyPrc);">
		/
		<input type="button" value="아이템 추가" onclick="self.location='text_insert.asp?menupos=<%= menupos %>&siteDiv=<%=siteDiv%>'" class="button">
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
	<td>텍스트이슈</td>
	<td>링크정보</td>
	<td>종료예상일</td>
	<td>사용유무</td>
	<td>순서</td>
	<td>등록일</td>
</tr>
<%	if oKeyword.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	Else
%>
<tbody id="subList">
<%	
		for i=0 to oKeyword.FResultCount-1
%>
<tr align="center" bgcolor="<% if oKeyword.FItemList(i).Fisusing = "Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><input type="checkbox" name="selIdx" value="<%= oKeyword.FItemList(i).Fidx %>"></td>
	<td><a href="text_insert.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Ftextname %></a></td>
	<td align="left" style="word-break : break-all;">&nbsp;<a href="text_insert.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Flinkinfo %></a></td>
	<td><%=oKeyword.FItemList(i).Fenddate%></td>
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
</tbody>
</form>
</table>
<%
set oKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


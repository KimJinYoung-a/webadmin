<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TextIssueCls.asp"-->
<%
'###############################################
' Discription : 텍스트 이슈
' History : 2013.12.14 이종화
'###############################################

	Dim idx, siteDiv

	idx = Request("idx")

	'// 내용 접수
	dim oKeyword
	Set oKeyword = new CSearchKeyWord
	oKeyword.FRectIdx = idx

	if idx<>"" then
		oKeyword.GetSearchKeyWord
	end if
%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="javascript">
<!--
	// 아이템 저장 실행
	function goSubmit()
	{
		// 키워드 입력여부 검사
		if(!document.frm.keyword.value) {
			alert("관련 키워드를 입력해주세요.");
			document.frm.keyword.focus();
			return;
		}
		// 링크 입력여부 검사
		if(!document.frm.linkinfo.value) {
			alert("키워드 클릭시 이동할 링크를 입력해주세요.");
			document.frm.linkinfo.focus();
			return;
		}
		// 순서 입력여부 검사
		if(!document.frm.sortNo.value) {
			alert("표시 순서를 입력해주세요.\n※ 순서는 숫자이며 적을수록 순번이 높습니다.");
			document.frm.sortNo.focus();
			return;
		}

		<% if idx="" then %>
		if(confirm("작성하신 내용을 등록하시겠습니까?")) {
			document.frm.mode.value="add";
			document.frm.action="dotextissue.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("수정하신 내용을 저장하시겠습니까?")) {
			document.frm.mode.value="modify";
			document.frm.action="dotextissue.asp";
			document.frm.submit();
		}
		<% end if %>
	}

	function putLinkText(key) {
		var frm = document.frm;
		switch(key) {
			case 'search':
				frm.linkinfo.value='/search/search_item.asp?rect=' + document.frm.keyword.value;
				break;
			case 'event':
				frm.linkinfo.value='/event/eventmain.asp?eventid=이벤트번호';
				break;
			case 'itemid':
				frm.linkinfo.value='/category/category_itemprd.asp?itemid=상품코드';
				break;
			case 'category':
				frm.linkinfo.value='/category/category_list.asp?disp=카테고리';
				break;
			case 'brand':
				frm.linkinfo.value='/street/street_brand.asp?makerid=브랜드아이디';
				break;
		}
	}

	function fnChangeDiv(val) {
		if(val=="T") {
			document.getElementById("lyrExMCate").style.display="none";
		} else {
			document.getElementById("lyrExMCate").style.display="";
		}
	}
//-->
</script>
<!-- 폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="dotextissue.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>텍스트이슈 등록</b></font>
		<% else %>
		<font color="red"><b>텍스트이슈 수정</b></font>
		<% end if%>
	</td>
</tr>
<% if idx<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">일련번호</td>
	<td align="left"><input type="text" name="idx" value="<%=idx%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">텍스트이슈</td>
	<td align="left"><input type="text" name="keyword" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Ftextname%>" size="32" maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">링크</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkinfo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Flinkinfo%>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">예)</font></td>
			<td valign="top">
				<font color="#707070">
				- <span style="cursor:pointer" onClick="putLinkText('search')">검색결과 링크 : /search/search_item.asp?rect=<font color="darkred">검색어</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">마감일자</td>
	<td align="left"><input id="prevDate" name="prevDate" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Fenddate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	<script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "prevDate", trigger    : "prevDate_trigger",
			onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
	</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">표시순서</td>
	<td align="left"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FsortNo: else Response.Write "99" %>" size="3" class="text"></td></td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="저장" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="취소" onClick="self.history.back()">
	</td>
</tr>
</form>
<!-- 폼 끝 -->
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

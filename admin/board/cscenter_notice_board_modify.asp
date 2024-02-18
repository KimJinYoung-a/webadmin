<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.04.29 한용민 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/boardnoticecls.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim i, j

'==============================================================================
'공지사항
dim boardnotice

dim SearchKey, SearchString, param, page, noticetype, menupos,oldyn
page = Request("page")
noticetype = Request("noticetype")
SearchKey = Request("SearchKey")
SearchString = Request("SearchString")
menupos = Request("menupos")
oldyn = request("oldyn")
param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&noticetype=" & noticetype & "&menupos=" & menupos

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&oldyn="& oldyn &"&noticetype=" & noticetype & "&menupos=" & menupos



set boardnotice = New CBoardNotice

boardnotice.read(request("id"))

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 10pt; }
-->
</STYLE>
<link rel=stylesheet type="text/css" href="/bct.css">
고객센타 - 공지사항<br><br>
<script>
function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function SubmitForm()
{
        if (document.f.title.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }
        if (document.f.yuhyostart.value == "") {
                alert("유효시작일을 입력하세요.");
                return;
        }
        if (document.f.yuhyoend.value == "") {
                alert("유효종료일을 입력하세요.");
                return;
        }

		if (confirm("수정하시겠습니까?") == true) {
			document.f.submit();
		}
}

function SubmitDelete()
{
        if (confirm("삭제하시겠습니까?") == true) {
                document.f.mode.value = "delete";
                document.f.submit();
        }
}
</script>


<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="cscenter_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= request("id") %>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="listtype" value="<%=noticetype%>">
<input type="hidden" name="oldyn" value="<%=oldyn%>">
<input type="hidden" name="menupos" value="<%=menupos%>">

<tr bgcolor="#FFFFFF">
	<td>공지유형</td>
	<td>
		  <select name="noticetype">
				<option value="" <% if boardnotice.results(0).Fnoticetype = "" then response.write "selected" %>>선택</option>
				<!--<option value="01" <% if boardnotice.results(0).Fnoticetype = "01" then response.write "selected" %>>전체공지</option> 2015리뉴얼에서 빠짐. 이상준대리.//-->
				<option value="02" <% if boardnotice.results(0).Fnoticetype = "02" then response.write "selected" %>>안내</option>
				<option value="03" <% if boardnotice.results(0).Fnoticetype = "03" then response.write "selected" %>>이벤트공지</option>
				<option value="04" <% if boardnotice.results(0).Fnoticetype = "04" then response.write "selected" %>>배송공지</option>
				<option value="05" <% if boardnotice.results(0).Fnoticetype = "05" then response.write "selected" %>>당첨자공지</option>
				<option value="06" <% if boardnotice.results(0).Fnoticetype = "06" then response.write "selected" %>>CultureStation</option>
		  </select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>카테고리</td>
	<td>
		<% DrawSelectBoxCategoryOnlyLarge "malltype", boardnotice.results(0).Fmalltype, ""%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>제목</td>
	<td><input type="text" name="title" size="60" value="<%= boardnotice.results(0).Ftitle %>" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>내용</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"><%= db2html(boardnotice.results(0).Fcontents) %></textarea><br><font color="red">(실제적용사이즈 입니다. 적당히 엔터키로 줄맞춤해주세요!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>유효시작일</td>
	<td><input type="text" size="10" name="yuhyostart" value="<%= boardnotice.results(0).Fyuhyostart %>" onClick="jsPopCal('f','yuhyostart');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>유효종료일</td>
	<td><input type="text" size="10" name="yuhyoend" value="<%= boardnotice.results(0).Fyuhyoend %>" onClick="jsPopCal('f','yuhyoend');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>고정글유무</td>
	<td><input type="radio" name="fixyn" value="Y" <% if boardnotice.results(0).Ffixyn = "Y" then response.write "checked" %>>사용 <input type="radio" name="fixyn" value="N" <% if boardnotice.results(0).Ffixyn = "N" then response.write "checked" %>>사용안함</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>모바일 중요공지 노출</td>
	<td><input type="radio" name="importantnotice" value="Y" <% if boardnotice.results(0).FImportantNotice = "Y" then response.write "checked" %>>사용 <input type="radio" name="importantnotice" value="N" <% if boardnotice.results(0).FImportantNotice = "N" then response.write "checked" %>>사용안함</td>
</tr>
</form>
</table>
<br>
<a href="javascript:SubmitForm()" onfocus="this.blur()"><img src="/images/icon_modify.gif" border="0" align="absmiddle"></a>
<a href="javascript:SubmitDelete()" onfocus="this.blur()"><img src="/images/icon_delete.gif" border="0" align="absmiddle"></a>
<a href="cscenter_notice_board_list.asp?page=<%=page & param%>" onfocus="this.blur()"><img src="/images/icon_list.gif" border="0" align="absmiddle"></a>
<br><br>
<table cellpadding="5" cellspacing="0" border="0" class="a">
<tr>
	<td bgcolor="#F8F8FA" style="border:1px solid #D8D8DA">
		<b>(링크 양식)</b><br>
		&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;소냐 이벤트 바로가기&lt;/a&gt;
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
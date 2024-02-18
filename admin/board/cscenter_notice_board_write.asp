<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.04.29 한용민 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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

		if (confirm("등록하시겠습니까?") == true) {
			document.f.submit();
		}
}
</script>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="cscenter_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<tr bgcolor="#FFFFFF">
	<td>공지유형</td>
	<td>
		  <select name="noticetype">
				<option value="">선택</option>
				<!--<option value="01">전체공지</option> 2015리뉴얼에서 빠짐. 이상준대리.//-->
				<option value="02">안내</option>
				<option value="03">이벤트공지</option>
				<option value="04">배송공지</option>
				<option value="05">당첨자공지</option>
				<option value="06">CultureStation</option>
		  </select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>카테고리</td>
	<td>
		<%DrawSelectBoxCategoryOnlyLarge"malltype", "","" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>제목</td>
	<td><input type="text" name="title" size="60" value="" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>내용</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"></textarea><br><font color="red">(실제적용사이즈 입니다. 적당히 엔터키로 줄맞춤해주세요!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>유효시작일</td>
	<td><input type="text" size="10" name="yuhyostart" value="" onClick="jsPopCal('f','yuhyostart');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>유효종료일</td>
	<td><input type="text" size="10" name="yuhyoend" value="" onClick="jsPopCal('f','yuhyoend');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>고정글유무</td>
	<td><input type="radio" name="fixyn" value="Y">사용 <input type="radio" name="fixyn" value="N" checked>사용안함</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>모바일 중요공지 노출</td>
	<td><input type="radio" name="importantnotice" value="Y">사용 <input type="radio" name="importantnotice" value="N" checked>사용안함</td>
</tr>
</form>
</table>
<br>
<input type="button" value=" 등록 " onclick="SubmitForm()">
<br><br>
(링크 양식)<br>
&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;소냐 이벤트 바로가기&lt;/a&gt;

<!-- #include virtual="/lib/db/dbclose.asp" -->
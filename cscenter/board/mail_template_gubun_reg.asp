<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim mode
dim idx, mastergubun, gubun, gubunname, title, contents, disporder, isusing


mode = request("mode")

mastergubun = "20"		'// MAIL
idx = request("idx")


dim omailtemplate
set omailtemplate = New CCSTemplate
omailtemplate.FCurrPage = 1
omailtemplate.FPageSize = 1
omailtemplate.FRectIdx = idx
omailtemplate.FRectMasterGubun = mastergubun

if (mode <> "addgubun") then
	omailtemplate.GetCSTemplateList

	gubun		= omailtemplate.FItemList(0).Fgubun
	gubunname	= omailtemplate.FItemList(0).Fgubunname
	title		= omailtemplate.FItemList(0).GetTitle
	contents	= omailtemplate.FItemList(0).GetContents
	disporder	= omailtemplate.FItemList(0).Fdisporder
	isusing		= omailtemplate.FItemList(0).Fisusing
end if

%>
<script language="JavaScript">
<!--

function SubmitAction(frm) {
	/*
	if (frm.gubun.value == "") {
		alert("구분을 입력하세요");
		frm.gubun.focus();
		return;
	}

	if ((frm.gubun.value.length != 2) || (frm.gubun.value*0 != 0)) {
		alert("구분은 2글자의 숫자만 가능합니다.");
		frm.gubun.focus();
		return;
	}
	*/

	if (frm.gubunname.value == "") {
		alert("구분명을 입력하세요");
		frm.gubunname.focus();
		return;
	}

	if (frm.gubunname.value.length > 15) {
		alert("구분명은 15글자까지 가능합니다.");
		frm.gubunname.focus();
		return;
	}

	if (frm.disporder.value == "") {
		alert("표시순서를 입력하세요");
		frm.disporder.focus();
		return;
	}

	if (frm.disporder.value*0 != 0) {
		alert("표시순서는 숫자만 가능합니다.");
		frm.disporder.focus();
		return;
	}

	if (confirm('저장 하시겠습니까?') == true) {
		var v = frm.title.value + "__|__" + frm.contents.value;
		frm.contents.value = v;

		frm.submit();
	}
}

//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frm" action="sms_template_process.asp">
<input type="hidden" name="menupos" value="<% = menupos %>">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="mastergubun" value="<% = mastergubun %>">
<input type="hidden" name="idx" value="<% = idx %>">

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		구분
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text_ro" name="gubun" size="4" value="<%= gubun %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		구분명
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="gubunname" size="30" value="<%= gubunname %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		이메일제목
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="title" size="45" value="<%= title %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		내용
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="contents" class="textarea" cols="80" rows="25"><%= contents %></textarea>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		표시순서
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="disporder" size="4" value="<%= disporder %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		사용
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" class="select">
			<option value="Y">사용함</option>
			<option value="N" <% if (isusing = "N") then %>selected<% end if %>>사용않함</option>
		</select>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="40">
	<td class="a" width="80">
		자동변환
	</td>
	<td bgcolor="#FFFFFF" align="left">
		* 일반정보<br>
		[아이디] : 고객 아이디<br>
		[이름] : 작성자 이름<br><br>

		*업체정보<br>
		[업체반품주소] : 업체 반품주소<br>
		[업체반품담당자] : 업체 반품담당자<br>
		[업체반품전화] : 업체반품전화<br>
		[업체거래택배사] : 업체 거래택배사<br>
		[업체스트리트명] : 업체 스트리트명
	</td>
</tr>

<tr align="left" bgcolor="<%= adminColor("tabletop") %>" height="35">
	<td colspan="2" bgcolor="#FFFFFF">
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="저장하기" onclick="SubmitAction(frm);" class="button">
		<input type="button" value="취소하기" onclick="history.back();" class="button">
	</td>
</tr>
</form>
</table>

<% set omailtemplate = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
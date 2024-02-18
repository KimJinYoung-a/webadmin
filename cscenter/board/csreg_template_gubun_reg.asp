<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim mode
dim idx, mastergubun, gubun, gubunname, contents, disporder, isusing


mode = request("mode")
mastergubun = request("mastergubun")

if (mastergubun = "") then
	mastergubun = "30"		'// CS접수
end if

idx = request("idx")


dim ocsregtemplate
set ocsregtemplate = New CCSTemplate
ocsregtemplate.FCurrPage = 1
ocsregtemplate.FPageSize = 1
ocsregtemplate.FRectIdx = idx
ocsregtemplate.FRectMasterGubun = mastergubun

if (mode <> "addgubun") then
	ocsregtemplate.GetCSTemplateList

	gubun		= ocsregtemplate.FItemList(0).Fgubun
	gubunname	= ocsregtemplate.FItemList(0).Fgubunname
	contents	= ocsregtemplate.FItemList(0).Fcontents
	disporder	= ocsregtemplate.FItemList(0).Fdisporder
	isusing		= ocsregtemplate.FItemList(0).Fisusing
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

	if (confirm("저장 하시겠습니까?") == true) {
		frm.submit();
	}
}

//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frm" action="csreg_template_process.asp">
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
		내용
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="contents" class="textarea" cols="52" rows="10"><%= contents %></textarea>
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
		[이름] : 작성자 이름<br>
		[직통전화] : 작성자 직통전화
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

<% set ocsregtemplate = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

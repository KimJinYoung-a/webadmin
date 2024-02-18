<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx, gubun, title, mode
idx = request("idx")
gubun = request("gubun")

title = "승인(취소)일자"
mode = "ModiAppDate"
if (gubun = "ipkumDate") then
    title = "입금일자"
    mode = "ModiIpkumDate"
end if

%>

<script language="javascript">

function jsSubmitReg(frm) {
	if (frm.appdate.value == "") {
		alert("<%= title %>를 입력하세요");
		return;
	}

	if (frm.appdate.value.length != 10) {
		alert("YYYY-MM-DD 형식으로 입력하세요");
		return;
	}

	if (confirm("등록 하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="<%= mode %>">
	<input type="hidden" name="logidx" value="<%= idx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td><%= title %></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
            <input type="text" class="text" name="appdate" value="" size="10">
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="입력하기" onClick="jsSubmitReg(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

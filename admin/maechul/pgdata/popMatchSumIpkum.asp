<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx
idx = request("idx")

%>

<script language="javascript">

function jsSubmitMatch(frm) {
	if (frm.arrOrderSerial.value == "") {
		alert("내용을 입력하세요.");
		return;
	}

	if (confirm("매칭하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="matchsumipkum">
	<input type="hidden" name="logidx" value="<%= idx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>합계입금 주문번호</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<textarea class="textarea" name="arrOrderSerial" cols="25" rows="12"></textarea>
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="입력하기" onClick="jsSubmitMatch(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

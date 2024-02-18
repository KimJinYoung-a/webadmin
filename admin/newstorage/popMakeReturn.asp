<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

dim reguser, divcode,baljuname,regname,comment
dim shopid, masteridx, mastercode, socid
reguser = session("ssBctid")
regname = session("ssBctCname")
masteridx = request("idx")
mastercode = request("code")
socid = request("socid")

dim yyyymmdd
''yyyymmdd = Left(Now(), 10)

%>
<script>

function monthDiff(d1, d2) {
    var months;
    months = (d2.getFullYear() - d1.getFullYear()) * 12;
    months -= d1.getMonth();
    months += d2.getMonth();
    return months <= 0 ? 0 : months;
}

function jsSubmit() {
    var frm = document.frm;

    if (frm.executedt.value == '') {
        alert('반품일자를 입력하세요.');
        return;
    }

    var nowDate = new Date('<%= Left(Now(), 10) %>');
    var inpDate = new Date(frm.executedt.value);

    if (nowDate.getDate() <= 5) {
        // 전월까지 허용
        if (monthDiff(inpDate, nowDate) > 1) {
            alert('반품일자는 전월까지만 가능합니다.');
            return;
        }
    } else {
        // 현재달만 허용
        if (monthDiff(inpDate, nowDate) > 0) {
            alert('반품일자는 현재월까지만 가능합니다.');
            return;
        }
    }

    if (confirm('반품등록 하시겠습니까?')) {
        frm.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="2">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>반품</strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	<form name="frm" method="post" action="ipchuledit_process.asp">
	<input type="hidden" name="mode" value="regchulgoreturn">
    <input type="hidden" name="masterid" value="<%= masteridx %>">
    <input type="hidden" name="code" value="<%= mastercode %>">
    <input type="hidden" name="socid" value="<%= socid %>">
	<tr bgcolor="#FFFFFF"ccc height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">출고코드</td>
		<td><%= mastercode %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">반품일자</td>
		<td>
            <input type="text" class="text" name="executedt" value="<%= yyyymmdd %>" size=11 readonly ><a href="javascript:calendarOpen(frm.executedt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">불량등록</td>
		<td>
		<input type="radio" name="regbad" value="Y" > 불량등록
		<input type="radio" name="regbad" value="N" checked > 등록안함
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="35">
		<td colspan="2" align="center">
            <input type="button" class="button" value="반품등록" onclick="jsSubmit()">
		</td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

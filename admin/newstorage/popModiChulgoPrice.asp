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

dim yyyymmdd
yyyymmdd = Left(Now(), 7) + "-01"
if (Day(Now()) <= 7) then
    yyyymmdd = Left(DateAdd("m", -1, yyyymmdd), 10)
end if

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

    if (confirm('출고가 일괄수정 하시겠습니까?')) {
        frm.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="2">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>출고가 일괄수정</strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	<form name="frm" method="post" action="ipchuledit_process.asp">
	<input type="hidden" name="mode" value="modichulgoprc">
    <input type="hidden" name="yyyymmdd" value="<%= yyyymmdd %>">
	<tr bgcolor="#FFFFFF"ccc height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="120">수정가능 출고일자</td>
		<td>
            <%= yyyymmdd %>
            * 이전 출고내역은 수정되지 않습니다.
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">수정내역</td>
		<td>
            출고코드 &lt;TAB&gt; 상품구분 &lt;TAB&gt; 상품코드 &lt;TAB&gt; 옵션 &lt;TAB&gt; 출고가<br />
            <textarea name="modlst" cols="55" rows="42"></textarea>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="35">
		<td colspan="2" align="center">
            <input type="button" class="button" value="수정하기" onclick="jsSubmit()">
		</td>
	</tr>
	</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

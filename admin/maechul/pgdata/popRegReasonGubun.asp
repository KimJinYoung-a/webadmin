<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx, gubun
idx = request("idx")
gubun = request("gubun")

%>

<script language="javascript">

function jsSubmitReg(frm) {
	if (frm.reasonGubun.value == "") {
		alert("사유를 선택하세요");
		return;
	}

	if (confirm("등록 하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="regReasonGubun<%= CHKIIF(gubun="off", "Off", "")%>">
	<input type="hidden" name="logidx" value="<%= idx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>상세사유</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<select class="select" name="reasonGubun">
				<option></option>
				<option value="001">선수금(매출)</option>
				<option value="002">선수금(제휴사 매출)</option>
                <option value="003">선수금(이니랜탈)</option>
				<option value="020">선수금(예치금)</option>
				<option value="025">선수금(예치금환급)</option>
				<option value="030">선수금(기프트)</option>
				<option value="035">선수금(기프트환급)</option>
                <option value="004">선수금(B2B 매출)</option>
				<option value="">---------------</option>
				<option value="040">CS서비스</option>
				<option value="">---------------</option>
				<option value="950">무통장미확인</option>
				<option value="999">취소매칭</option>
				<option value="901">핑거스현금매출</option>
				<option value="800">이자수익</option>
				<option value="900">기타</option>
				<option value="">---------------</option>
				<option value="XXX">입력이전</option>
			</select>
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

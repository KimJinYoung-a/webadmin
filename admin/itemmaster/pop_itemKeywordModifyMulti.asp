<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

%>

<script language='javascript'>

function trim(value) {
 return value.replace(/^\s+|\s+$/g,"");
}

function SaveItem(frm) {
	if (frm.keywords.value == "") {
		alert("저장할 내용이 없습니다.");
		return;
	}

	//var errorFound = false;
	var rows, row, itemid, keywords, totalCount;
	rows = frm.keywords.value.split("\n");

	totalCount = 0;
	for (var i = 0; i < rows.length; i++) {
		row = trim(rows[i]);
		if (row == "") {
			continue;
		}

		row = row.split("\t");
		if (row.length != 2) {
			alert("올바른 형식이 아닙니다.[TAB]");
			return;
		}

		itemid = row[0];
		keywords = row[1];

		if ((itemid == "") || (keywords == "")) {
			alert("에러\n\n" + itemid + "\n" + keywords);
			return;
		}

		if (itemid*0 != 0) {
			alert("에러\n\n" + itemid);
			return;
		}

		totalCount = totalCount + 1;
	}

	if (totalCount > 200) {
		alert("최대 200건까지만 적용가능합니다.");
		return;
	}

	var ret = confirm('저장 하시겠습니까?(' + totalCount + '건)');

	if(ret) {
		frm.submit();
	}
}

function CloseWindow() {
    window.close();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">

	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="itemKeyword_process.asp">
<input type=hidden name=mode value="editmulti">
<tr>
	<td colspan="2" bgcolor="#FFFFFF">
		<table width="100%" cellspacing=1 cellpadding=1 border="0" class=a bgcolor=#BABABA>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">상품코드+키워드</td>
				<td bgcolor="#FFFFFF">
					<textarea class="textarea" name="keywords" rows="25" cols="128"></textarea>
				</td>
			</tr>
		</table>
	</td>
</tr>
</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<input type="button" class="button" value="저장하기" onclick="SaveItem(frm2)">
			&nbsp;
			<input type="button" class="button" value=" 닫 기 " onclick="CloseWindow()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

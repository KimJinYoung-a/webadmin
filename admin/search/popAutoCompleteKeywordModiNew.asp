<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim prect, rect
dim mode

rect = requestCheckVar(request("rect"), 32)

if (rect = "") then
	mode = "add"
else
	mode = "modievtUserAddCNT"
end if

%>

<script language='javascript'>

function jsSubmit(frm) {
	if (frm.UserAddCNT.value == "") {
		alert('가중치를 입력하세요.');
		return;
    }

	if (frm.UserAddCNT.value*0 != 0) {
		alert('가중치는 숫자만 가능합니다.');
		return;
    }

	var ret = confirm("수정 하시겠습니까?");
	if(ret){
		frm.submit();
	}
}

function jsSubmitAdd(frm) {
	if (frm.rect.value == "") {
		alert('키워드를 입력하세요.');
		return;
    }

	var ret = confirm("등록 하시겠습니까?");
	if(ret){
		frm.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>키워드 등록</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method=post action="manageAutoCompleteKeyword_process.asp">
	<input type="hidden" name="mode" value="<%= mode %>">
	<% if (mode = "add") then %>
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">키워드</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" class="text" name="rect" value="" size="10">
    	</td>
    </tr>
	<% else %>
	<input type="hidden" name="rect" value="<%= rect %>">
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">키워드</td>
    	<td bgcolor="#FFFFFF">
    		<%= rect %>
    	</td>
    </tr>
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">가중치</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" class="text" name="UserAddCNT" value="" size="10">
    	</td>
    </tr>
	<% end if %>
    </form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<% if (mode = "add") then %>
            <input type="button" class="button" value="추가" onclick="jsSubmitAdd(frm);">
			<% else %>
			<input type="button" class="button" value="수정" onclick="jsSubmit(frm);">
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

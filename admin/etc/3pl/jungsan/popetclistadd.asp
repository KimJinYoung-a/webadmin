<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/jungsanCls.asp"-->
<%
dim gubun,masteridx

gubun   	= requestCheckvar(request("gubun"),16)
masteridx	= requestCheckvar(request("idx"),10)


dim otpljungsanmaster
set otpljungsanmaster = new CTplJungsan
otpljungsanmaster.FRectIdx = masteridx
otpljungsanmaster.GetTPLJungsanMasterList

if (otpljungsanmaster.FResultCount<1) then
    dbget_TPL.Close : dbget.Close(): response.end
end if

%>

<script language='javascript'>
function adddata(frm){
	if (frm.itemname.value.length<1){
		alert('내용을 입력하세요.');
		frm.itemname.focus();
		return;
	}

	if (frm.itemno.value.length<1){
		alert('갯수를 입력하세요.');
		frm.itemno.focus();
		return;
	}

	if (frm.itemno.value*0 != 0){
		alert('갯수는 숫자만 가능합니다.');
		frm.itemno.focus();
		return;
	}

	if (frm.cbmX.value.length<1){
		alert('CBM X 를 입력하세요.');
		frm.cbmX.focus();
		return;
	}

	if (!IsDigit(frm.cbmX.value)){
		alert('CBM X 는 숫자만 가능합니다.');
		frm.cbmX.focus();
		return;
	}

	if (frm.cbmY.value.length<1){
		alert('CBM Y 를 입력하세요.');
		frm.cbmY.focus();
		return;
	}

	if (!IsDigit(frm.cbmY.value)){
		alert('CBM Y 는 숫자만 가능합니다.');
		frm.cbmY.focus();
		return;
	}

	if (frm.cbmZ.value.length<1){
		alert('CBM Z 를 입력하세요.');
		frm.cbmZ.focus();
		return;
	}

	if (!IsDigit(frm.cbmZ.value)){
		alert('CBM Z 는 숫자만 가능합니다.');
		frm.cbmZ.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>기타내역추가</strong></font>
        </td>
        <td align="right">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmadd" method="post" action="dotpljungsan.asp">
    <input type="hidden" name="mode" value="etcadd">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="masteridx" value="<%= masteridx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>내용</td>
		<td width="40">수량</td>
        <td width="80">CBM X(mm)</td>
        <td width="80">CBM Y(mm)</td>
        <td width="80">CBM Z(mm)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
		<td><input type="text" name="itemname" value="" size="55"></td>
		<td><input type="text" name="itemno" value="1" size="3" style="text-align:center"></td>
		<td><input type="text" name="cbmX" value="" size="8" style="text-align:right"></td>
		<td><input type="text" name="cbmY" value="" size="8" style="text-align:right"></td>
		<td><input type="text" name="cbmZ" value="" size="8" style="text-align:right"></td>
    </tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value="내역 추가" onclick="adddata(frmadd)" class="button"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 traffic analysis  수동 입력 페이지
' History : 2007.09.04 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/traffic/traffic_class.asp"-->

<script language="javascript">
function sudongsubmit()
{
document.frm.action = "traffic_analysis_sudong_submit.asp";
document.frm.submit();
}

function back()
{
history.back();
}

</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>※ 텐바이텐 traffic analysis 수동입력</strong></font>
			</td>
			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
</table>
<!--표 헤드끝-->

<!--본문 시작-->
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<form action="" name="frm" method="get">
	<tr bgcolor=#FFFFFF>
		<td align="right" colspan=6><input type="button" value="텐바이텐 DB에 저장" onclick="sudongsubmit()"> <input type="button" value="전페이지로" onclick="back()"></td>
	</tr>
		<tr bgcolor=#DDDDFF>
			<td align="center">날짜</td>
		   <td align="center">페이지뷰</td>
		   <td align="center">방문자수</td>
		   <td align="center">신규방문자수</td>
		   <td align="center">재방문자수</td>
		   <td align="center">실제방문자수</td>
		</tr>
		<tr bgcolor=#FFFFFF>
			<td align="center"><input type="text" maxsize="10" name="yyyymmdd"></td>
			<td align="center"><input type="text" maxsize="10" name="pageview"></td>
			<td align="center"><input type="text" maxsize="10" name="totalcount"></td>
			<td align="center"><input type="text" maxsize="10" name="newcount"></td>
			<td align="center"><input type="text" maxsize="10" name="recount"></td>
			<td align="center"><input type="text" maxsize="10" name="realcount"></td>
		</tr>
	</form>	
	</table>
<!-- 본문 끝 -->

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
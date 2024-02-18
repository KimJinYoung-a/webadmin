<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->


<br>
<img src="/images/icon_search.jpg" border="0">
<img src="/images/icon_star.gif" border="0">
<img src="/images/icon_plus.gif" border="0">
<img src="/images/icon_minus.gif" border="0">
<img src="/images/cs_icon_user.gif" border="0">

<img src="/images/icon_arrow_down.gif" border="0">
<img src="/images/icon_arrow_up.gif" border="0">
<img src="/images/icon_arrow_left.gif" border="0">
<img src="/images/icon_arrow_right.gif" border="0">
<img src="/images/icon_arrow_link.gif" border="0">


<img src="/images/exclam.gif" border="0">
<img src="/images/calicon.gif" border="0">


<br><br>
<img src="/images/iexplorer.gif" border="0">
<img src="/images/iexcel.gif" border="0">
<img src="/images/icon_print02.gif" border="0">
<img src="/images/icon_search.jpg" border="0">
<img src="/images/icon_new.gif" border="0">


<br><br>
<img src="/images/icon_confirm.gif" border="0">
<img src="/images/icon_cancel.gif" border="0">
<img src="/images/icon_delete.gif" border="0">
<img src="/images/icon_hide.gif" border="0">
<img src="/images/icon_issue.gif" border="0">
<img src="/images/icon_modify.gif" border="0">
<img src="/images/icon_help.gif" border="0">
<img src="/images/icon_list.gif" border="0">
<img src="/images/icon_next.gif" border="0">
<img src="/images/icon_reply.gif" border="0">
<img src="/images/icon_save.gif" border="0">
<img src="/images/icon_search.gif" border="0">
<img src="/images/icon_use.gif" border="0">
<img src="/images/icon_change.gif" border="0">
<img src="/images/icon_detail.gif" border="0">
<input type="button" value="버튼">

<br><br>
<img src="/images/icon_new_registration.gif" border="0">
<img src="/images/button_reload.gif" border="0">
<img src="/images/search2.gif" border="0" valign="absbottom">
<input type="button" value="버튼">


<br><br>
<img src="/images/btn_excel.gif" border="0">
<img src="/images/btn_word.gif" border="0">

<br><br>
<img src="/images/bu_2.gif" border="0">
<img src="/images/icon_ok.gif" border="0">
<img src="/images/page_2_4.gif" border="0">
<img src="/images/icon_1.gif" border="0">
<img src="/images/page_2_3.gif" border="0">

<br><br>

<img src="/images/icon_num01.gif" border="0">
<img src="/images/icon_num02.gif" border="0">
<img src="/images/icon_num03.gif" border="0">
<img src="/images/icon_num04.gif" border="0">
<img src="/images/icon_num05.gif" border="0">
<img src="/images/icon_num06.gif" border="0">
<img src="/images/icon_num07.gif" border="0">
<img src="/images/icon_num08.gif" border="0">
<img src="/images/icon_num09.gif" border="0">
<img src="/images/icon_num10.gif" border="0">


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>어드민 아이콘리스트</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">&nbsp;</td>
	        <td valign="top" align="right">&nbsp;</td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="50">이미지</td>
    	<td width="200">파일명</td>
      	<td width="200">사이즈</td>
      	<td>비고</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>이미지</td>
    	<td>파일명</td>
      	<td>사이즈</td>
      	<td>비고</td>
	</tr>
</table>


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


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->
<%
dim mode,diaryid ,itemid
dim sql
diaryid= request("diaryid")


dim objLink ,intLoop
set objLink = new clsDiary
objLink.getDiaryLinkedItem diaryid

%>
<script language="javascript">

function subchk(){

	if(document.regfrm.itemid.value.length<1){
		document.regfrm.itemid.focus();
		alert('상품 번호를 입력하셔야 합니다.');
		return false;
	}

	document.regfrm.submit();
}
window.resizeTo(500,300);
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="center">
        	<b>관련상품 등록 </b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<% If objLink.FResultCount>0 Then %>
		<% For intLoop =0 to objLink.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"> <%= objLink.FItemList(intLoop).FItemid %></td><td><%= objLink.FItemList((intLoop)).FItemName %><a href="proc_diary_LInkItem.asp?mode=del&diaryid=<%= diaryid %>&itemid=<%= objLink.FItemList(intLoop).FItemid %>">[삭제]</a></td>
	</tr>
		<% Next %>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2" height="40"> [등록된 상품이 없습니다.상품을 추가해 주세요.] </td>
	</tr>
	<% end if %>

	<form name="regfrm" method="post" action="proc_diary_LInkItem.asp">
	<input type="hidden" name="diaryid" value="<%= diaryid %>">
	<input type="hidden" name="mode" value="write">
	<tr bgcolor="#FFFFFF">
		<td align="center" class="required">상품번호</td>
		<td class="input_1"><input type="text" name="itemid" size="30" value="" /><br>ex>(123,456,789,...123) [마지막에 콤마(,)넣지 마세요]</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<input type="button" class="button" value="확인" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="취소" onclick="window.close();"/>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<% set objLink = nothing %>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
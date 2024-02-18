<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' History : 2007.11.12 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/othermall_main_event_rotationcls.asp"-->

<%
dim idx,mode
	idx = request("idx")
	mode = request("mode")
%>

<script language='javascript'>
function SubmitForm(){

	if (document.SubmitFrm.linkinfo.value.length < 1){
		alert('링크 정보를 입력 하세요');
		document.SubmitFrm.linkinfo.focus();
		return;
	}

	if (document.SubmitFrm.disporder.value.length < 1){
		alert('전시 순서를 입력 하세요');
		document.SubmitFrm.disporder.focus();
		return;
	}


	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		document.SubmitFrm.submit();
	}
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
			<font color="red"><strong>외부몰 엠디 추천 상품 입력</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
</table>
<!--표 헤드끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
  <form name="SubmitFrm" method="post" action="<%=staticImgUrl%>/chtml/othermall_doMainMdChoiceRotate.asp" onsubmit="return false;" enctype="multipart/form-data">
    <input type="hidden" name="mode" value="<%= mode %>">
<%
if mode = "modify" then
dim mdchoicerotate
set mdchoicerotate = new CMainMdChoiceRotate
mdchoicerotate.FCurrPage = 1
mdchoicerotate.FPageSize = 1
mdchoicerotate.read idx
%>
	<input type="hidden" name="idx" value="<% = idx %>">
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">이미지</td>
	  <td><input type="file" name="photoimg" value="" size="32" maxlength="32" class="file">
	  <br>
	  <img src="<%= mdchoicerotate.FItemList(0).Fphotoimg %>" >
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">전시순서</td>
	  <td><input type="text" name="disporder" value="<% = mdchoicerotate.FItemList(0).Fdisporder  %>" size="2" class="input_b">
	  <font color="red">(1~12 사이의 값.)</font>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">상품코드</td>
	  <td><input type="text" name="linkitemid" value="<% = mdchoicerotate.FItemList(0).Flinkitemid  %>" size="6" class="input_b">
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">link정보</td>
	  <td><input type="text" name="linkinfo" value="<% = mdchoicerotate.FItemList(0).Flinkinfo  %>" size="70" class="input_b">
	  <br>
	  <font color="red">(상대경로로 입력하세요 /shopping/category_prd.asp?itemid=72367)</font>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">사용여부</td>
	  <td>
	  	<input type="radio" name="isusing" value="Y" <% if mdchoicerotate.FItemList(0).FIsUsing="Y" then response.write "checked" %> >Y
	  	<input type="radio" name="isusing" value="N" <% if mdchoicerotate.FItemList(0).FIsUsing="N" then response.write "checked" %> >N
	  </td>
	</tr>
	</form>
</table>
<%
set mdchoicerotate = Nothing
else
%>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">이미지</td>
	  <td><input type="file" name="photoimg" value="" size="32" maxlength="32" class="file"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">전시순서</td>
	  <td><input type="text" name="disporder" value="6" size="2" class="input_b">
	  <font color="red">(1~12 사이의 값.)</font>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">상품코드</td>
	  <td><input type="text" name="linkitemid" value="" size="6" class="input_b">
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">link정보</td>
	  <td><input type="text" name="linkinfo" size="70"  class="input_b">
	  <br>
	  <font color="red">(상대경로로 입력하세요 /shopping/category_prd.asp?itemid=72367)</font>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="100" align="center">사용여부</td>
	  <td>
	  	<input type="radio" name="isusing" value="Y" checked >Y
	  	<input type="radio" name="isusing" value="N" >N
	  </td>
	</tr>
	</form>
</table>
<% end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="left">
	  		<input type="button" value="저장" onClick="SubmitForm()">
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

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
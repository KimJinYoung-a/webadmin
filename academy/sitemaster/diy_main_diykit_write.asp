<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/diy_main_diykitcls.asp"-->
<%
dim idx,mode
idx = RequestCheckvar(request("idx"),10)
mode = RequestCheckvar(request("mode"),16)
%>

<script language='javascript'>
function SubmitForm(){

	if (document.SubmitFrm.disporder.value.length < 1){
		alert('전시 순서를 입력 하세요');
		document.SubmitFrm.disporder.focus();
		return;
	}

	if (document.SubmitFrm.linkitemid.value.length < 1){
		alert('상품코드를 입력 하세요');
		document.SubmitFrm.linkitemid.focus();
		return;
	}

	if (document.SubmitFrm.linkinfo.value.length < 1){
		alert('링크 정보를 입력 하세요');
		document.SubmitFrm.linkinfo.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}

</script>
<br><br>

※전시순서는 숫자가 작을수록 먼저 나오는 순서 입니다. 같은 순서인 경우 신상품 순서입니다.<br><br>
<table width="700" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black">
  <form name="SubmitFrm" method="post" action="<%=imgFingers%>/linkweb/sitemaster/doDiyMainDIYKit.asp" onsubmit="return false;" enctype="multipart/form-data">
    <input type="hidden" name="mode" value="<% = mode %>">
<%
if mode = "modify" then
dim mdchoicerotate
set mdchoicerotate = new CMainMdChoiceRotate
mdchoicerotate.FCurrPage = 1
mdchoicerotate.FPageSize = 1
mdchoicerotate.read idx
%>
	<input type="hidden" name="idx" value="<% = idx %>">
	<!--
	<tr>
	  <td width="100">이미지</td>
	  <td><input type="file" name="photoimg" value="" size="32" maxlength="32" class="file">
	  <br>
	  <img src="<%= mdchoicerotate.FItemList(0).Fphotoimg %>" >
	  	<font color="red">(119px × 135px GIF 혹은 JPG 이미지)</font>
	  </td>
	</tr>
	//-->
	<tr>
	  <td width="100">전시순서</td>
	  <td><input type="text" name="disporder" value="<% = mdchoicerotate.FItemList(0).Fdisporder  %>" size="2" class="input_b">
	  <font color="red">(2자리 숫자)</font>
	  </td>
	</tr>
	<tr>
	  <td width="100">상품코드</td>
	  <td><input type="text" name="linkitemid" value="<% = mdchoicerotate.FItemList(0).Flinkitemid  %>" size="6" class="input_b">
	  </td>
	</tr>
	<tr>
	  <td width="100">link정보</td>
	  <td><input type="text" name="linkinfo" value="<% = mdchoicerotate.FItemList(0).Flinkinfo  %>" size="70" class="input_b">
	  <br>
	  <font color="red">(상대경로로 입력하세요 /diyshop/shop_prd.asp?itemid=1001)</font>
	  </td>
	</tr>
	<tr>
	  <td width="100">사용여부</td>
	  <td>
	  	<input type="radio" name="isusing" value="Y" <% if mdchoicerotate.FItemList(0).FIsUsing="Y" then response.write "checked" %> >Y
	  	<input type="radio" name="isusing" value="N" <% if mdchoicerotate.FItemList(0).FIsUsing="N" then response.write "checked" %> >N
	  </td>
	</tr>
	<tr>
	  <td colspan="2" align="center">
	  	<input type="button" value="저 장" onClick="SubmitForm()">
	  	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  	<input type="button" value="리스트로" onClick="location.href='/academy/sitemaster/diy_main_diykit.asp?menupos=1229';">
	  </td>
	</tr>
	</form>
</table>
<%
set mdchoicerotate = Nothing
else
%>
	<!--
	<tr>
	  <td width="100">이미지</td>
	  <td>
	  	<input type="file" name="photoimg" value="" size="32" maxlength="32" class="file">
	  	<font color="red">(119px × 135px GIF 혹은 JPG 이미지)</font>
	  </td>
	</tr>
	//-->
	<tr>
	  <td width="100">전시순서</td>
	  <td><input type="text" name="disporder" value="99" size="2" class="input_b">
	  <font color="red">(2자리 숫자)</font>
	  </td>
	</tr>
	<tr>
	  <td width="100">상품코드</td>
	  <td><input type="text" name="linkitemid" value="" size="6" class="input_b">
	  </td>
	</tr>
	<tr>
	  <td width="100">link정보</td>
	  <td><input type="text" name="linkinfo" size="70"  class="input_b">
	  <br>
	  <font color="red">(상대경로로 입력하세요 /diyshop/shop_prd.asp?itemid=1001)</font>
	  </td>
	</tr>
	<tr>
	  <td width="100">사용여부</td>
	  <td>
	  	<input type="radio" name="isusing" value="Y" checked >Y
	  	<input type="radio" name="isusing" value="N" >N
	  </td>
	</tr>
	<tr>
	  <td colspan="2" align="center">
	  	<input type="button" value="저 장" onClick="SubmitForm()">
	  	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  	<input type="button" value="리스트로" onClick="location.href='/academy/sitemaster/diy_main_diykit.asp?menupos=1229';">
	  </td>
	</tr>
	</form>
</table>
<%
end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
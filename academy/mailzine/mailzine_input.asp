<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim yyyy1,mm1,dd1
dim nowdate

nowdate = Left(CStr(now()),10)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

%>
<script language="JavaScript">
<!--

function checkok(frm){
      frm.submit();
}

//-->
</script>
<form method="post" name="monthly" action="http://110.93.128.113/ftp/academy_mailzine_input_ok.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="write">
<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
<tr class="a">
	<td align="center" height="35" colspan="2"><b>메일진 작성</b></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">메일제목</td>
	<td>&nbsp;<input type="text" name="title" class="input" size="55"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">메일진 등록일</td>
	<td>&nbsp;<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">1번이미지</td>
	<td>&nbsp;<input type="file" name="img1" class="input" size="40"></td>
</tr>
<tr class="a">
	<td colspan="2" align="center">
	&nbsp;&nbsp;이미지맵 코드넣기
	</td>
</tr>
<tr>
	<td colspan="2">
	   <table border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<textarea name="Imagemap1" rows="10" cols="75" class="textarea"><map name="ImgMap1">

</map></textarea>
		</td>
	   </tr>
	   <tr>
		<td>
			ex) &lt;area shape="rect" coords="8,8,590,347"<br> href="http://www.10x10.co.kr/designfingers/designfingers.asp"<br> <font  color="red">target="_top"</font> onFocus='this.blur()'&gt;<br>
			<font color="#330099">위의 예제처럼 target은 탑으로 주시고 이미지맵 이름은 고치지 말아주세요~~!!</font>
		</td>
	   </tr>

	   </table>
	</td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">2번이미지</td>
	<td>&nbsp;<input type="file" name="img2" class="input" size="40"></td>
</tr>
<tr class="a">
	<td colspan="2" align="center">
	&nbsp;&nbsp;이미지맵 코드넣기
	</td>
</tr>
<tr>
	<td colspan="2">
	   <table border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<textarea name="Imagemap2" rows="10" cols="75" class="textarea"><map name="ImgMap2">

</map></textarea>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr>
	<td align="right" colspan="2" height="30"><input type="button" value="메일진 등록" onclick="checkok(this.form);" class="button">&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
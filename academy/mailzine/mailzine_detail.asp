<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_mailzinecls.asp"-->
<%
dim idx,code,omail
dim yyyy1,mm1,dd1


idx = RequestCheckvar(request("idx"),10)

set omail = new CUploadMaster
omail.MailzineDetail idx

yyyy1 = omail.Fcode1
mm1 = omail.Fcode2
dd1 = omail.Fcode3

code = mm1 & dd1

%>
<script language="JavaScript">
<!--

function checkok(frm){
   frm.submit();
}

//-->
</script>
<form method="post" name="modify" action="http://110.93.128.113/ftp/academy_mailzine_input_ok.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="idx" value="<% = idx %>">
<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
<tr class="a">
	<td align="center" height="35" colspan="2"><b>메일진 작성</b></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">메일제목</td>
	<td>&nbsp;<input type="text" name="title" class="input" size="55" value="<% = omail.Ftitle %>"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">메일진 등록일</td>
	<td>&nbsp;<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">1번이미지</td>
	<td>&nbsp;<input type="file" name="img1" class="input" size="40"><br><% = omail.Fimg1 %></td>
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
			<textarea name="imagemap1" rows="10" cols="75" class="textarea"><% = omail.Fimgmap1 %></textarea>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">2번이미지</td>
	<td>&nbsp;<input type="file" name="img2" class="input" size="40"><br><% = omail.Fimg2 %></td>
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
			<textarea name="imagemap2" rows="10" cols="75" class="textarea"><% = omail.Fimgmap2 %></textarea>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">디스플레이여부</td>
	<td>&nbsp;<input type="radio" name="isusing" value="Y" <% if omail.Fisusing = "Y" then response.write "checked" %>>사용 <input type="radio" name="isusing" value="N" <% if omail.Fisusing = "N" then response.write "checked" %>>미사용</td>
</tr>
<tr>
	<td align="right" colspan="2" height="30"><input type="button" value="메일진 수정" onclick="checkok(this.form);" class="button">&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
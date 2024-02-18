<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<%

dim intPosNo , intID , strIMG , strURL , strMode
dim StrSQL



intPosNo = request("pn")
strMode = request("md")
IF intPosNo="" Then
	Alert_close("오류")
End IF

IF intPosNo<>"" and strMode<>"" Then
	StrSQL =" SELECT TOP 1 id,PosNo,Img,Url  "&_
			" FROM db_diary2010.dbo.tbl_diaryMain "&_
			" WHERE PosNo="& intPosNo &" ORDER BY id desc "
	rsget.open StrSQL,dbget,2

	IF not rsget.Eof then
		intID 		= rsget("id")
		intPosNo	= rsget("PosNo")
		strIMG		= rsget("Img")
		strURL		= rsget("Url")
	End IF

	rsget.close
End IF
%>
<script language="javascript">

window.onload = function(){
	window.resizeTo(350,250);
}

</script>
<table  border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="rfrm" method="post" action="<%= uploadImgUrl %>/linkWeb/diary/DiaryMainReg_Proc.asp" target="regframe" enctype="multipart/form-data" >
<input type="hidden" name="md" value="<%= strMode %>">
<input type="hidden" name="id" value="<%= intID%>">
<input type="hidden" name="pn" value="<%= intPosNo %>">
<tr bgcolor="#FFFFFF">
	<td>이미지</td>
	<td>
		<input type="file" name="img" value=""><br>
		<%
		SELECT CASE intPosNo
			CASE 1 '// 투데이
				response.write "(384x400)"
			CASE 6 '//2층 이벤트
				response.write "(284x200)"
			CASE 14,18 '//3,4층 이벤트
				response.write "(221x200)"
			CASE ELSE
				response.write "(142x200)"
		END SELECT
		%>
	</td>
</tr>

<% IF intPosNo=1 Then %>
<tr bgcolor="#FFFFFF">
	<td>상품코드</td>
	<td><input type="text" name="url" value="<%= strURL %>"></td>
</tr>
<% ELSE %>
<tr bgcolor="#FFFFFF">
	<td>링크</td>
	<td><input type="text" name="url" value="<%= strURL %>"></td>
</tr>
<% End IF %>
<tr bgcolor="#FFFFFF">
	<td colspan="2"><input type="submit" value="저장"></td>
</tr>
</form>
</table>

<iframe name="regframe" src="" frameborder="0" width="0" height="0"></iframe>

<!-- #include virtual="/lib/db/dbclose.asp" -->
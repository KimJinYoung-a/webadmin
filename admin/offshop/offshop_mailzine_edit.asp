<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 메일진
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_mailzinecls.asp"-->
<%
dim yyyy1,mm1,dd1
dim nowdate

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)

Dim omail,idx
idx = requestCheckVar(request("idx"),10)

set omail = new CUploadMaster
omail.MailzineDetail idx

nowdate = Left(CStr(omail.Fregdate),10)
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


<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black">
<form method="post" name="monthly" action="<%=uploadUrl%>/ftp/offshop_mailzine_ok.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="idx" value="<% = idx %>">
<tr class="a">
	<td align="center" height="35" colspan="2"><b>메일진 작성</b></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">메일진 등록일</td>
	<td>&nbsp;<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">메일제목</td>
	<td>&nbsp;<input type="text" name="title" class="input" size="55" value="<% = omail.Ftitle %>"></td>
</tr>
<tr class="a">
	<td align="center" height="25" colspan="2">공지사항<br><font color="#FF7D7D"><font color="#FF3737">(실제 보여지는 내용의 길이만큼 적용됩니다. 줄바꿈을 잘 적용시켜주세요)</font></font></td>
</tr>
<tr>
	<td colspan="2">
	   <table border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<textarea name="news" rows="10" cols="75" class="textarea"><% = omail.Fnews %></textarea>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">핑거스 이미지</td>
	<td>&nbsp;<input type="file" name="img1" class="input" size="40"><br>&nbsp;<% = omail.Fimg1 %><br>&nbsp;<input type="text" name="url1" class="input" size="60" value="<% = omail.Furl1 %>"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">MD추천상품배너</td>
	<td>&nbsp;<input type="file" name="img2" class="input" size="40"><br>&nbsp;<% = omail.Fimg2 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">이벤트배너1</td>
	<td>&nbsp;<input type="file" name="img3" class="input" size="40"><br>&nbsp;<% = omail.Fimg3 %><br>&nbsp;<input type="text" name="url2" class="input" size="60" value="<% = omail.Furl2 %>"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">이벤트배너2</td>
	<td>&nbsp;<input type="file" name="img4" class="input" size="40"><br>&nbsp;<% = omail.Fimg4 %><br>&nbsp;<input type="text" name="url3" class="input" size="60" value="<% = omail.Furl3 %>"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="50">이벤트배너3</td>
	<td>&nbsp;<input type="file" name="img5" class="input" size="40"><br>&nbsp;<% = omail.Fimg5 %><br>&nbsp;<input type="text" name="url4" class="input" size="60" value="<% = omail.Furl4 %>"></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">POP이미지</td>
	<td>&nbsp;<input type="file" name="img6" class="input" size="40"><br>&nbsp;<% = omail.Fimg6 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">POP이미지Big</td>
	<td>&nbsp;<input type="file" name="img7" class="input" size="40"><br>&nbsp;<% = omail.Fimg7 %></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">추천브랜드 6</td>
	<td>&nbsp;<input type="text" name="brand" class="input" size="60" value="<% = omail.Fbrand %>">&nbsp;<input type="button" class="button" value="이미지올리기" onclick="TnFtpUpload('D:/home/cube1010/imgstatic/main/brand/','/main/brand/');"><br><font color="#FF3737">(마지막에 콤마(,)를 넣어주세요 ex:mmmg,ia,heewoo,)</font></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">받는사람메일주소</td>
	<td>&nbsp;<textarea name="sendmailer" rows="5" cols="58" class="textarea"><% = omail.Fsendmailer %></textarea><br><font color="#FF3737">(주소사이에 세미콜론(;)을 넣어주시고 띄어쓰기나 엔터키로<br> 줄바꿈하지 마세요. ex:corpse2@10x10.co.kr;gundolly@10x10.co.kr)</font></td>
</tr>
<tr class="a">
	<td align="center" width="100" height="35">오픈여부</td>
	<td>&nbsp;<input type="radio" name="isusing" value="Y" <% if omail.Fisusing = "Y" then response.write "checked" %>> 오픈 &nbsp;<input type="radio" name="isusing" value="N" <% if omail.Fisusing = "N" then response.write "checked" %>> 오픈안함</td>
</tr>
<tr>
	<td align="right" colspan="2" height="30"><input type="button" value="메일진 등록" onclick="checkok(this.form);" class="button">&nbsp;&nbsp;&nbsp;</td>
</tr>
</form>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
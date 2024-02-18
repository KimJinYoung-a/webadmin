<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script language='javascript'>
<!--
	// 이미지 종류 검사
	function checkImageSuffix (fileInput)
	{
	   var suffixPattern = /(gif|jpg|jpeg|png)$/i;
	   if (!suffixPattern.test(fileInput.value)) {
	     alert('GIF,JEPG,PNG 파일만 가능합니다.');
	     fileInput.focus();
	     fileInput.select();
	     return false;
	   }
	   return true;
	}

	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.imgTitle.value)
		{
			alert("타이틀 이미지를 선택해주십시오.\n\n※ 이미지는 JPG, GIF형식으로 선택해주십시요.");
			return false;
		}
		else if(!checkImageSuffix(frm.imgTitle))
		{
			return false;
		}

		if(!frm.title.value)
		{
			alert("제목을 입력해주십시오.");
			frm.title.focus();
			return false;
		}

		if(!frm.imgProd.value)
		{
			alert("완성된 작품의 이미지를 선택해주십시오.\n\n※ 이미지는 JPG, GIF형식으로 선택해주십시요.");
			return false;
		}
		else if(!checkImageSuffix(frm.imgProd))
		{
			return false;
		}

		if(!frm.scrName.value)
		{
			alert("작품이름을 입력해주십시오.");
			frm.scrName.focus();
			return false;
		}

		if(!frm.scrTime.value)
		{
			alert("소요시간을 입력해주십시오.");
			frm.scrTime.focus();
			return false;
		}
		
		if(frm.summ.value.length>150){
			alert("메인요약 내용은 150자 이내료 작성해 주십시오.");
			frm.summ.focus();
			return false;
		}

		// 폼 전송
		return true;

	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyScrap.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="wirte_main">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>마디 스크랩 신규 등록</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>타이틀 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgTitle" size="60"><br>
		<font color=darkred>※ 너비 612px 크기의 JPG/GIF 이미지 파일을 업로드하십시오.</font>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="80" maxlength="120"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>완성품 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgProd" size="60">
		<font color=darkred>※ 10:7(가로:세로) 비율의 JPG/GIF 파일입니다.</font>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>작품명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrName" size="60" maxlength="100"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>난이도</td>
	<td bgcolor="#FFFFFF">
		<select name="scrDef">
			<option value="1">[1]★☆☆☆☆</option>
			<option value="2" selected>[2]★★☆☆☆</option>
			<option value="3">[3]★★★☆☆</option>
			<option value="4">[4]★★★★☆</option>
			<option value="5">[5]★★★★★</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>소요시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrTime" size="30" maxlength="60"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">재료</td>
	<td bgcolor="#FFFFFF"><textarea name="scrSource" rows="2" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">도구</td>
	<td bgcolor="#FFFFFF"><textarea name="scrTool" rows="2" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>스크랩 형태</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="center"><input type="radio" name="printType" value="A" checked> Type A</td>
			<td width="50%" align="center"><input type="radio" name="printType" value="B"> Type B</td>
		</tr>
		<tr>
			<td align="center"><img src="/images/tpl_typeA.gif" alt="내용이 오른쪽의 이미지를 설명하는 방식입니다."></td>
			<td align="center"><img src="/images/tpl_typeB.gif" alt="왼쪽의 이미지를 설명하는 방식입니다."></td>
		</tr>
		<tr>
			<td align="center"><input type="radio" name="printType" value="C"> Type C</td>
			<td align="center"><input type="radio" name="printType" value="D"> Type D</td>
		</tr>
		<tr>
			<td align="center"><img src="/images/tpl_typeC.gif" alt="위쪽의 이미지를 상세 설명하는 방식입니다."></td>
			<td align="center"><img src="/images/tpl_typeD.gif" alt="HTML로 외부의 이미지 혹은 표를 사용하여 자유롭게 작성합니다."></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">TIP</td>
	<td bgcolor="#FFFFFF"><textarea name="scrTip" rows="4" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">메인요약</td>
	<td bgcolor="#FFFFFF"><textarea name="summ" rows="2" cols="80"></textarea> (150자 이내)</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_next.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
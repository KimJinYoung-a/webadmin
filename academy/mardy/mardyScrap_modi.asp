<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_Scrapcls.asp"-->
<%
	'// 변수 선언 //
	dim ScrapId
	dim page, searchKey, searchString, param
	dim oScrap, i, lp

	'// 파라메터 접수 //
	ScrapId = RequestCheckvar(request("ScrapId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	'// 메인 정보 접수
	set oScrap = new CMardyScrap
	oScrap.FRectScrapId = ScrapId

	oScrap.GetMardyScrapView
%>
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
		if(frm.imgTitle.value && !checkImageSuffix(frm.imgTitle))
		{
			return false;
		}

		if(!frm.title.value)
		{
			alert("제목을 입력해주십시오.");
			frm.title.focus();
			return false;
		}

		if(frm.imgProd.value && !checkImageSuffix(frm.imgProd))
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
<input type="hidden" name="mode" value="modify_main">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="ScrapId" value="<%=ScrapId%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>마디 스크랩 정보 수정</b></td>
			<td height="26" align="right">
				<font color=gray><b>사용여부</b></font>
				<select name="isusing">
					<option value="Y" <% if oScrap.FItemList(0).Fisusing="Y" then Response.Write "selected" %>>사용</option>
					<option value="N" <% if oScrap.FItemList(0).Fisusing="N" then Response.Write "selected" %>>숨김</option>
				</select>&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>타이틀 이미지</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<% if oScrap.FItemList(0).FimgTitle<>"" then %>
			<td width="124" align="center">
				<img src="<%=oScrap.FItemList(0).FimgTitle_full%>" style="width:120px;border:1px solid #C0C0C0">
			</td>
			<% end if %>
			<td>
				<% if oScrap.FItemList(0).FimgTitle<>"" then %>
					(현재 : <%= oScrap.FItemList(0).FimgTitle%>)
				<% end if %>
				<input type="file" name="imgTitle" size="50"><br>
				<font color=darkred>※ 너비 612px 크기의 JPG/GIF 이미지 파일을 업로드하십시오.</font>
				<input type="hidden" name="orgTitle" value="<%=oScrap.FItemList(0).FimgTitle%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="80" maxlength="120" value="<%=oScrap.FItemList(0).Ftitle%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>완성품 이미지</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<% if oScrap.FItemList(0).FimgProd<>"" then %>
			<td width="124" align="center">
				<img src="<%=oScrap.FItemList(0).FimgProd_full%>" style="width:120px;border:1px solid #C0C0C0">
			</td>
			<% end if %>
			<td>
				<% if oScrap.FItemList(0).FimgProd<>"" then %>
					(현재 : <%= oScrap.FItemList(0).FimgProd%>)
				<% end if %>
				<input type="file" name="imgProd" size="50"><br>
				<font color=darkred>※ 10:7(가로:세로) 비율의 JPG/GIF 파일입니다.</font>
				<input type="hidden" name="orgProd" value="<%=oScrap.FItemList(0).FimgProd%>">
				<input type="hidden" name="orgIcon" value="<%=oScrap.FItemList(0).FimgIcon%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>작품명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrName" size="60" maxlength="100" value="<%=oScrap.FItemList(0).FscrName%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>난이도</td>
	<td bgcolor="#FFFFFF">
		<select name="scrDef">
			<option value="1" <% if oScrap.FItemList(0).FscrDef="1" then Response.Write "selected" %>>[1]★☆☆☆☆</option>
			<option value="2" <% if oScrap.FItemList(0).FscrDef="2" then Response.Write "selected" %>>[2]★★☆☆☆</option>
			<option value="3" <% if oScrap.FItemList(0).FscrDef="3" then Response.Write "selected" %>>[3]★★★☆☆</option>
			<option value="4" <% if oScrap.FItemList(0).FscrDef="4" then Response.Write "selected" %>>[4]★★★★☆</option>
			<option value="5" <% if oScrap.FItemList(0).FscrDef="5" then Response.Write "selected" %>>[5]★★★★★</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>소요시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="scrTime" size="30" maxlength="60" value="<%=oScrap.FItemList(0).FscrTime%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">재료</td>
	<td bgcolor="#FFFFFF"><textarea name="scrSource" rows="2" cols="80"><%=oScrap.FItemList(0).FscrSource%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">도구</td>
	<td bgcolor="#FFFFFF"><textarea name="scrTool" rows="2" cols="80"><%=oScrap.FItemList(0).FscrTool%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>스크랩 형태</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="center"><input type="radio" name="printType" value="A" <% if oScrap.FItemList(0).FprintType="A" then Response.Write "checked" %>> Type A</td>
			<td width="50%" align="center"><input type="radio" name="printType" value="B" <% if oScrap.FItemList(0).FprintType="B" then Response.Write "checked" %>> Type B</td>
		</tr>
		<tr>
			<td align="center"><img src="/images/tpl_typeA.gif" alt="내용이 오른쪽의 이미지를 설명하는 방식입니다."></td>
			<td align="center"><img src="/images/tpl_typeB.gif" alt="왼쪽의 이미지를 설명하는 방식입니다."></td>
		</tr>
		<tr>
			<td align="center"><input type="radio" name="printType" value="C" <% if oScrap.FItemList(0).FprintType="C" then Response.Write "checked" %>> Type C</td>
			<td align="center"><input type="radio" name="printType" value="D" <% if oScrap.FItemList(0).FprintType="D" then Response.Write "checked" %>> Type D</td>
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
	<td bgcolor="#FFFFFF"><textarea name="scrTip" rows="4" cols="80"><%=oScrap.FItemList(0).FscrTip%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">메인요약</td>
	<td bgcolor="#FFFFFF"><textarea name="summ" rows="2" cols="80"><%=oScrap.FItemList(0).Fsummary%></textarea> (150자 이내)</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_modify.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
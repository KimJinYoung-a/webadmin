<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_tipcls.asp"-->
<%
	'// 변수 선언 //
	dim tipId
	dim page, searchKey, searchString

	dim oTip, oTipImage, i, lp, maxAddImg

	maxAddImg=10		'추가 이미지 최대수 지정

	'// 파라메터 접수 //
	tipId = RequestCheckvar(request("tipId"),10)
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
	if page="" then page=1
	if searchKey="" then searchKey="tipName"

	'// 내용 접수
	set oTip = new CMardyTip
	oTip.FRectTipId = tipId

	oTip.GetMardyTipView
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
		if(!frm.title.value)
		{
			alert("제목을 입력해주십시오.");
			frm.title.focus();
			return false;
		}

		if(!frm.tipName.value)
		{
			alert("작품이름을 입력해주십시오.");
			frm.tipName.focus();
			return false;
		}

		if(!frm.tipUsage.value)
		{
			alert("사용도를 입력해주십시오.");
			frm.tipUsage.focus();
			return false;
		}

		if(!frm.tipTime.value)
		{
			alert("소요시간을 입력해주십시오.");
			frm.tipTime.focus();
			return false;
		}

		for(var i=0;i<frm.imgFile.length;i++)
		{
			if ((frm.imgFile[i].value.length>0)&&(!checkImageSuffix(frm.imgFile[i]))){
				return false;
			}
		}

		// 폼 전송
		return true;

	}

	// 추가 이미지폼 제어
	function addImgControl(lid, md, mx)
	{
		var frm = document.all;
		var btext = "";

		if(md=="add")
		{
			// 레이어 보이기 및 내용 활성화
			frm["addimg"+(lid+1)].style.display="";
			frm["btn"+(lid)].innerHTML = "";
			if(lid<mx)
				frm.filedelete[lid].value = "";

			lid++;

			// 버튼처리
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='이미지 삭제' onClick=addImgControl(" + lid + ",'del'," + mx + ") style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='이미지 추가' onClick=addImgControl(" + lid + ",'add'," + mx + ") style='cursor:pointer'>";

			frm["btn"+lid].innerHTML = btext;
		}
		else if(md=="del")
		{
			// 레이어 숨기기 및 내용 삭제
			frm.imgFile[lid-1].select();
			document.execCommand('Delete');

			if(lid>mx)
				frm.imgCont[lid-1].value = "";
			else
				frm.filedelete[lid-1].value = "Y";

			frm["addimg"+lid].style.display="none";
			frm["btn"+lid].innerHTML = "";

			lid--;

			// 버튼처리
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='이미지 삭제' onClick=addImgControl(" + lid + ",'del'," + mx + ") style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='이미지 추가' onClick=addImgControl(" + lid + ",'add'," + mx + ") style='cursor:pointer'>";

			frm["btn"+lid].innerHTML = btext;
		}
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyTip.asp" enctype="multipart/form-data">
<input type="hidden" name="tipId" value="<%=tipId%>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>마디수첩 내용 수정</b></td>
			<td height="26" align="right">
				<font color=gray><b>사용여부</b></font>
				<select name="isusing">
					<option value="Y">사용</option>
					<option value="N">숨김</option>
				</select>&nbsp;
				<script language="javascript">
					document.frm_write.isusing.value="<%=oTip.FItemList(0).Fisusing%>";
				</script>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="80" maxlength="120" value="<%=oTip.FItemList(0).Ftitle%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">목록 이미지</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="52"><img src="<%=oTip.FItemList(0).FimgIcon_full%>" width="50" height="50"></td>
			<td>
				<input type="file" name="imgIcon" size="60"><br>
				<font color=darkred>※ 50px * 50px 크기의 JPG/GIF 이미지 파일을 업로드하십시오.</font>
				<input type="hidden" name="orgIcon" value="<%=oTip.FItemList(0).FimgIcon%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>작품명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipName" size="60" maxlength="100" value="<%=oTip.FItemList(0).FtipName%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>사용도</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipUsage" size="40" maxlength="60" value="<%=oTip.FItemList(0).FtipUsage%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>난이도</td>
	<td bgcolor="#FFFFFF">
		<select name="tipDef">
			<option value="1">[1]★☆☆☆☆</option>
			<option value="2">[2]★★☆☆☆</option>
			<option value="3">[3]★★★☆☆</option>
			<option value="4">[4]★★★★☆</option>
			<option value="5">[5]★★★★★</option>
		</select>
		<script language="javascript">
			document.frm_write.tipDef.value="<%=oTip.FItemList(0).FtipDef%>";
		</script>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>소요시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipTime" size="30" maxlength="60" value="<%=oTip.FItemList(0).FtipTime%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">비용</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tipPrice" size="40" maxlength="100" value="<%=oTip.FItemList(0).FtipPrice%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">주의사항</td>
	<td bgcolor="#FFFFFF"><textarea name="tipAttent" rows="3" cols="80"><%=oTip.FItemList(0).FtipAttent%></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">Tip</td>
	<td bgcolor="#FFFFFF"><textarea name="tipCont" rows="3" cols="80"><%=oTip.FItemList(0).FtipCont%></textarea></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<%
	'// 추가 이미지 폼 작성
	set oTipImage = new CMardyTip
	oTipImage.FRectTipId = tipId

	oTipImage.GetMardyTipImageList

	'// 서브 목록
	for i=0 to oTipImage.FTotalCount - 1
		lp = lp+1
%>
<tr id="addimg<%=lp%>">
	<td align="center" width="120" bgcolor="#DDDDFF"><% if lp=1 then %><font color="darkred">* </font><% end if %>이미지 #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="76" rowspan="2"><img src="<%=oTipImage.FItemList(i).FimgFile_full%>" width="74" height="74"></td>
			<td width="30" align="center">파일</td>
			<td colspan="2">
				<input type="file" name="imgFile" size="50">
				<input type="hidden" name="orgImgId" value="<%=oTipImage.FItemList(i).FimgId%>">
				<input type="hidden" name="orgFile" value="<%=oTipImage.FItemList(i).FimgFile%>">
				<input type="hidden" name="filedelete" value="">
			</td>
		</tr>
		<tr>
			<td width="30" align="center">설명</td>
			<td><textarea name="imgCont" rows="3" cols="70"><%=oTipImage.FItemList(i).FimgCont%></textarea></td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>">
				<% if lp>1 then %><img src="/images/icon_minus.gif" alt="이미지 삭제" onClick="addImgControl(<%=lp%>,'del',<%=oTipImage.FTotalCount%>)" style="cursor:pointer"> <% end if %>
				<% if lp=oTipImage.FTotalCount then %><img src="/images/icon_plus.gif" alt="이미지 추가" onClick="addImgControl(<%=lp%>,'add', <%=oTipImage.FTotalCount%>)" style="cursor:pointer"><% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<%
	next

	'// 여유 빈칸
	for lp=i+1 to maxAddImg
%>
<tr id="addimg<%=lp%>" style='display:none;'>
	<td align="center" width="120" bgcolor="#DDDDFF"><% if lp=1 then %><font color="darkred">* </font><% end if %>이미지 #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="30" align="center">파일</td>
			<td colspan="2">
				<input type="file" name="imgFile" size="50">
				<input type="hidden" name="orgImgId" value="">
				<input type="hidden" name="orgFile" value="">
				<input type="hidden" name="filedelete" value="">
			</td>
		</tr>
		<tr>
			<td width="30" align="center">설명</td>
			<td><textarea name="imgCont" rows="3" cols="70"></textarea></td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>"></td>
		</tr>
		</table>
	</td>
</tr>
<%	next %>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
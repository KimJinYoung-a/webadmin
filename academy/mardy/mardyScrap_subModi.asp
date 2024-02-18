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
	dim ScrapId, subId
	dim page, searchKey, searchString, param
	dim oScrap, oScrapSub, i, lp

	'// 파라메터 접수 //
	ScrapId = RequestCheckvar(request("ScrapId"),10)
	subId = RequestCheckvar(request("subId"),10)
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

	'// 서브 정보 접수
	set oScrapSub = new CMardyScrapSub
	oScrapSub.FRectSubId = subId

	oScrapSub.GetMardyScrapImageView
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.subCont.value)
		{
			alert("내용을 입력해주십시오.");
			frm.subCont.focus();
			return false;
		}

		// 폼 전송
		return true;
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" colspan="2" align="left">마디 스크랩 기본 정보</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">제목</td>
	<td bgcolor="#FFFFFF"><%=oScrap.FItemList(0).Ftitle%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">출력 형태</td>
	<td bgcolor="#FFFFFF">Type <%=oScrap.FItemList(0).FprintType%></td>
</tr>
</table>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyScrapSub.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="modify_sub">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="ScrapId" value="<%=ScrapId%>">
<input type="hidden" name="subId" value="<%=subId%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="printType" value="<%=oScrap.FItemList(0).FprintType%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>마디 스크랩 만드는법 수정</b></td>
</tr>
<%
	'형태별 분기
	Select Case oScrap.FItemList(0).FprintType
		Case "A"
			'///// Type A /////
%>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">제목</td>
		<td bgcolor="#FFFFFF"><input type="text" name="subName" size="80" maxlength="120" value="<%=oScrapSub.FItemView(0).FsubName%>"></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="4" cols="80"><%=oScrapSub.FItemView(0).FsubCont%></textarea></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지 #1</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile1<>"" then %>
				<td width="124" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile1_full%>" style="width:120px;height:90px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile1<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile1%>
						- 삭제 <input type="checkbox" name="filedelete1" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile1" size="60"><br>
					<font color=darkred>※ 4:3(가로:세로) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile1" value="<%=oScrapSub.FItemView(0).FimgFile1%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지 #2</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile2<>"" then %>
				<td width="124" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile2_full%>" style="width:120px;height:90px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile2<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile2%>
						- 삭제 <input type="checkbox" name="filedelete2" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile2" size="60"><br>
					<font color=darkred>※ 4:3(가로:세로) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile2" value="<%=oScrapSub.FItemView(0).FimgFile2%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지 #3</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile3<>"" then %>
				<td width="124" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile3_full%>" style="width:120px;height:90px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile3<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile3%>
						- 삭제 <input type="checkbox" name="filedelete3" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile3" size="60"><br>
					<font color=darkred>※ 4:3(가로:세로) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile3" value="<%=oScrapSub.FItemView(0).FimgFile3%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
<%

		Case "B"
			'///// Type B /////
%>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="4" cols="80"><%=oScrapSub.FItemView(0).FsubCont%></textarea></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지 #1</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile1<>"" then %>
				<td width="104" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile1_full%>" style="width:100px;height:100px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile1<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile1%>
						- 삭제 <input type="checkbox" name="filedelete1" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile1" size="60"><br>
					<font color=darkred>※ 1:1(정사각형) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile1" value="<%=oScrapSub.FItemView(0).FimgFile1%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지 #2</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile2<>"" then %>
				<td width="104" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile2_full%>" style="width:100px;height:100px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile2<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile2%>
						- 삭제 <input type="checkbox" name="filedelete2" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile2" size="60"><br>
					<font color=darkred>※ 1:1(정사각형) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile2" value="<%=oScrapSub.FItemView(0).FimgFile2%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지 #3</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile3<>"" then %>
				<td width="104" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile3_full%>" style="width:100px;height:100px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile3<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile3%>
						- 삭제 <input type="checkbox" name="filedelete3" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile3" size="60"><br>
					<font color=darkred>※ 1:1(정사각형) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile3" value="<%=oScrapSub.FItemView(0).FimgFile3%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
<%
		Case "C"
			'///// Type C /////
%>

	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="5" cols="80"><%=oScrapSub.FItemView(0).FsubCont%></textarea></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">이미지</td>
		<td bgcolor="#FFFFFF">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<% if oScrapSub.FItemView(0).FimgFile1<>"" then %>
				<td width="174" align="center">
					<img src="<%=oScrapSub.FItemView(0).FimgFile1_full%>" style="width:170px;height:120px;border:1px solid #C0C0C0">
				</td>
				<% end if %>
				<td>
					<% if oScrapSub.FItemView(0).FimgFile1<>"" then %>
						(현재 : <%= oScrapSub.FItemView(0).FimgFile1%>
						- 삭제 <input type="checkbox" name="filedelete1" value="Y">)<br>
					<% end if %>
					<input type="file" name="imgFile1" size="60"><br>
					<font color=darkred>※ 17:12(가로:세로) 비율의 JPG/GIF 파일입니다.</font>
					<input type="hidden" name="orgFile1" value="<%=oScrapSub.FItemView(0).FimgFile1%>">
				</td>
			</tr>
			</table>
		</td>
	</tr>
<%
		Case "D"
			'///// Type D /////
%>

	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="12" cols="80"><%=oScrapSub.FItemView(0).FsubCont%></textarea></td>
	</tr>
<%
	End Select
%>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_confirm.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
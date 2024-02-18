<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_storycls.asp"-->
<%
	'// 변수 선언 //
	dim storyId
	dim page, searchKey, searchString, param

	dim oStory, oStoryImage, i, lp

	'// 파라메터 접수 //
	storyId = RequestCheckvar(request("storyId"),10)
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
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

	'// 내용 접수
	set oStory = new CMardyStory
	oStory.FRectstoryId = storyId

	oStory.GetMardyStoryView

	'// 서브 이미지 접수
	set oStoryImage = new CMardyStory
	oStoryImage.FRectstoryId = storyId

	oStoryImage.GetMardyStoryImageList
%>
<script language="javascript">
<!--
	// 글삭제
	function GotoStoryDel(){
		if (confirm('본 게시물을 영구히 삭제 하시겠습니까?\n\n※ 삭제 후에는 다시 복구 할 수 없습니다.')){
			document.frm_trans.submit();
		}
	}


	// 사용 상태 변경
	function GotoUse(md)
	{
		switch(md)
		{
			case "use" :
				if (confirm('사이트 목록에 출력되도록 상태를 "사용"으로 변경하시겠습니까?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=storyId%>&mode=StoryUse";
				}
				break;

			case "del" :
				if (confirm('사이트에서 볼 수 없도록 상태를 "숨김"으로 변경하시겠습니까?')){
					FrameCHK.location="inc_Mardy_Use.asp?Idx=<%=storyId%>&mode=StoryDel";
				}
				break;
		}
	}

//-->
</script>
<!-- 보기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>마디 이야기 상세 정보</b></td>
			<td height="26" align="right">
				<font color=gray>사용여부 - </font>
				<%
					if oStory.FItemList(0).Fisusing="N" then
						Response.Write "<font color=darkred><b>숨김</b></font>"
					else
						Response.Write "<font color=darkblue><b>사용</b></font>"
					end if
				%>&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">짧은 제목</td>
	<td bgcolor="#FFFFFF"><%=oStory.FItemList(0).FtitleShort%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">상세 제목</td>
	<td bgcolor="#FFFFFF"><%=oStory.FItemList(0).FtitleLong%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">이미지</td>
	<td bgcolor="#FFFFFF">
		<table border="0" class="a" cellpadding="0" cellspacing="0">
		<%
			for lp=0 to oStoryImage.FTotalCount - 1
		%>
		<tr>
			<td align="center">
				<img src="<%=oStoryImage.FItemList(lp).FimgFile_full%>" ><br><br>
			</td>
		</tr>
		<% next %>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">내용 설명</td>
	<td bgcolor="#FFFFFF"><%=nl2br(oStory.FItemList(0).Fcontents)%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.gif" onClick="self.location='mardyStory_modi.asp?menupos=<%=menupos%>&storyId=<%=storyId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% if oStory.FItemList(0).Fisusing="N" then %>
		<img src="/images/icon_use.gif" onClick="GotoUse('use')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoStoryDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% else %>
		<img src="/images/icon_hide.gif" onClick="GotoUse('del')" style="cursor:pointer" align="absmiddle"> &nbsp;
		<% end if %>
		<img src="/images/icon_list.gif" onClick="self.location='mardyStory_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="http://image.thefingers.co.kr/linkweb/doMardyStory.asp" enctype="multipart/form-data">
<input type="hidden" name="storyId" value="<%=storyId%>">
<input type="hidden" name="mode" value="delete">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/DiaryEnjoyCls.asp"-->
<%
'###############################################
' PageName : evnet_enjoyWrite.asp
' Discription : 작가따라 그려봐 등록/수정
' History : 2009.09.30 허진원 생성
'###############################################

dim denjSn,mode,i
mode=request("mode")
denjSn=request("denjSn")
%>
<script language="javascript">
<!--
function subcheck(){
	var frm=document.inputfrm;

	if(!frm.makerid.value) {
		alert("브랜드ID를 입력해주세요!");
		frm.makerid.focus();
		return;
	}

	if(!frm.subject.value) {
		alert("제목을 입력해주세요!");
		frm.subject.focus();
		return;
	}

	if(!frm.videoSn.value) {
		alert("연결할 동영상을 [동영상 검색]을 이용하여 선택해주세요!");
		return;
	}

	if(!frm.smallImage.value&&!frm.denjSn.value) {
		alert("메뉴에 표시할 작은 이미지를 선택해주세요!");
		frm.smallImage.focus();
		return;
	}

	if(!frm.listImage.value&&!frm.denjSn.value) {
		alert("목록에 표시할 이미지를 선택해주세요!");
		frm.listImage.focus();
		return;
	}

	if(!frm.introImage.value&&!frm.denjSn.value) {
		alert("소개글 이미지를 선택해주세요!");
		frm.introImage.focus();
		return;
	}

	//if(!frm.bestImage.value&&!frm.denjSn.value) {
	//	alert("베스트 목록에 표시할 이미지를 선택해주세요!");
	//	frm.bestImage.focus();
	//	return;
	//}

	frm.submit();
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('본 게시물을 삭제하시겠습니까?')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="<%= uploadImgUrl %>/linkweb/Diary/doDiaryEnjoyProcess.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>작가따라 그려봐 등록/수정</b></font>
	</td>
</tr>
<% if mode="add" then %>
<input type="hidden" name="denjSn" value="">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
	<td bgcolor="#FFFFFF">
	    <input type="text" class="text" name="makerid" value="" size="20" >
	    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="subject" value="" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="videoSn" value="" size="3" readonly>
		<input type="button" class="button" value="동영상 검색" onclick="jsSearchVideoSn(this.form.name,'videoSn','dia');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 기간 및 발표</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="eventday" value="기간:2009.11.16 ~ 11.18 / 발표:11.19" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">작은 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="smallImage" value="" size="40"> (※ JPG,GIF 이미지, 154px × 104px, 최대 200KB 이하)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">목록 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="listImage" value="" size="40"> (※ JPG,GIF 이미지, 200px × 134px, 최대 300KB 이하)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">소개글 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="introImage" value="" size="40"> (※ JPG,GIF 이미지, 276px × 239x, 최대 500KB 이하)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">베스트 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="bestImage" value="" size="40"> (※ JPG,GIF 이미지, 120px × 120px, 최대 200KB 이하)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ver.2 메인 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="v2mainImage" value="" size="40"> (※ JPG,GIF 이미지, 186px × 195px, 최대 300KB 이하)
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New CEnjoy
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectEnSN=denjSn
	fmainitem.GetDiaryEnjoyList
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">번호</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FdenjSn%></b>
		<input type="hidden" name="denjSn" value="<%=fmainitem.FItemList(0).FdenjSn%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
	<td bgcolor="#FFFFFF">
	    <input type="text" class="text" name="makerid" value="<%=fmainitem.FItemList(0).Fmakerid%>" size="20" >
	    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="subject" value="<%=fmainitem.FItemList(0).Fsubject%>" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="videoSn" value="<%=fmainitem.FItemList(0).FvideoSn%>" size="3" readonly>
		<input type="button" class="button" value="동영상 검색" onclick="jsSearchVideoSn(this.form.name,'videoSn','dia');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 기간 및 발표</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="eventday" value="<%=fmainitem.FItemList(0).Feventday%>" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">작은 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="smallImage" value="" size="40"> (※ JPG,GIF 이미지, 165px × 115px, 최대 200KB 이하)
		<%
			if Not(fmainitem.FItemList(0).FsmallImage="" or isNull(fmainitem.FItemList(0).FsmallImage)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).FsmallImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">목록 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="listImage" value="" size="40"> (※ JPG,GIF 이미지, 200px × 134px, 최대 300KB 이하)
		<%
			if Not(fmainitem.FItemList(0).FlistImage="" or isNull(fmainitem.FItemList(0).FlistImage)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).FlistImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">소개글 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="introImage" value="" size="40"> (※ JPG,GIF 이미지, 270px × 246px, 최대 500KB 이하)
		<%
			if Not(fmainitem.FItemList(0).FintroImage="" or isNull(fmainitem.FItemList(0).FintroImage)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).FintroImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">베스트 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="bestImage" value="" size="40"> (※ JPG,GIF 이미지, 120px × 120px, 최대 200KB 이하)
		<%
			if Not(fmainitem.FItemList(0).FbestImage="" or isNull(fmainitem.FItemList(0).FbestImage)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).FbestImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ver.2 메인 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="v2mainImage" value="" size="40"> (※ JPG,GIF 이미지, 186px × 195px, 최대 300KB 이하)
		<%
			if Not(fmainitem.FItemList(0).Fv2mainimage="" or isNull(fmainitem.FItemList(0).Fv2mainimage)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).Fv2mainimage & ")"
			end if
		%>
	</td>
</tr>
<input type="hidden" name="imgname" value="<%=fmainitem.FItemList(0).FsmallImage%>">
<input type="hidden" name="imgname_1" value="<%=fmainitem.FItemList(0).Fv2mainimage%>">
<% end if %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isUsing" value="Y" <% If mode = "add" Then Response.Write "checked" Else If fmainitem.FItemList(0).Fisusing = "Y" Then Response.Write "checked" End If End If %>> Y
		<input type="radio" name="isUsing" value="N" <% If mode = "edit" Then If fmainitem.FItemList(0).Fisusing = "N" Then Response.Write "checked" End If End If %>> N
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><!--<input type="button" value=" 삭제 " class="button" onclick="delitems();"> &nbsp;&nbsp;//--><% end if %>
		<input type="button" value=" 취소 " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

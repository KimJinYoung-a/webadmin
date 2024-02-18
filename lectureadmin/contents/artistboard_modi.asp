<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/artistboard_cls.asp"-->
<%
Dim CBCont
Dim lecuserid, idx, userid, title, content, regdate, imgurl1, imgurl2,ithread, idepth
idx = requestCheckVar(request("idx"),10)
lecuserid = Session("ssBctId")

Set CBCont = new CArtistRoomBoard
	CBCont.Fidx = idx
	CBCont.FLecuserid = lecuserid
	CBCont.fnGetContent
	userid = CBCont.FUserid
	title = CBCont.FTitle
	content = CBCont.FContent
	imgurl1 = CBCont.FImgUrl1
	imgurl2 = CBCont.FImgUrl2
	regdate = CBCont.FRegdate
	
Set CBCont = nothing
%>
<script language='javascript'>
<!--
	
	function jsSubmit(frm){
	if(fnChkBlank(frm.sT.value)) {
	 	alert("제목을 입력해 주세요");
	 	frm.sT.focus();
	 	return false;
	}
	
	 if(fnChkBlank(frm.tC.value)) {
		alert("내용을 입력해 주세요");
		frm.tC.focus();
		return false;
	}
	
	if (frm.sImg1.value) {
		 if (!checkFile(frm.sImg1))  return false;
	 }
	 
	if (frm.sImg2.value) {
		 if (!checkFile(frm.sImg2))  return false;
	 }
}

function fnChkBlank(str)
{
    if (str == "" || str.split(" ").join("") == ""){
        return true;
	}
    else{
        return false;
	}
}	
	
function checkFile(sfile){
  //파일 확장자 확인
   var suffixPattern = /(gif|jpg|jpeg|png)$/i;
	if (!suffixPattern.test(sfile.value)) {
     alert('GIF,JPG 파일만 가능합니다.');
     sfile.focus();
     sfile.select();
     return false;
   }
   
  //파일 사이즈 확인 
  var img = new Image();
  img.dynsrc = sfile.value;
  var filesize = img.fileSize;
	if(filesize > 1024000){
	 alert('파일크기 초과입니다. 최대 1MB까지 가능합니다.');
     sfile.focus();
     sfile.select();
     return false;
	}
	
	return true;
}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmReg" method="post" action="http://image.thefingers.co.kr/linkweb/artist/procboard.asp" enctype="multipart/form-data" onSubmit="return jsSubmit(this);">      	
<input type="hidden" name="sM" value="U">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sUID" value="<%=Session("ssBctId")%>">
<input type="hidden" name="lecuserid" value="<%=Session("ssBctId")%>">
<input type="hidden" name="retUrl" value="http://webadmin.10x10.co.kr/lectureadmin/contents/artistboard_list.asp?menupos=979">
<input type="hidden" name="sfImg1" value="<%=imgurl1%>">
<input type="hidden" name="sfImg2" value="<%=imgurl2%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">아이디</td>
	<td bgcolor="#FFFFFF"><%=Session("ssBctId")%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sT" size="40" maxlength="50" value="<%=title%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">내용</td>
	<td bgcolor="#FFFFFF">
	<textarea name="tC" rows="14" cols="80"><%=db2html(content)%></textarea>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">이미지첨부1</td>
	<td bgcolor="#FFFFFF">파일크기는 1MB이하,JPG또는 GIF형식의 파일만 가능합니다.<br>
		사이즈는 WIDTH - 400이하로 설정해 주시기 바랍니다.<br>
			<% dim arrimg 
          			IF imgurl1 <> "" THEN
          			 arrimg = split(imgurl1,"/")%>
          			<p class="text2"><%=arrimg(ubound(arrimg))%> <input type="checkbox" name="chkimg1">삭제<br>
          			<%END IF%>
		<input type="file" name="sImg1">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">이미지첨부2</td>
	<td bgcolor="#FFFFFF">파일크기는 1MB이하,JPG또는 GIF형식의 파일만 가능합니다.<br>
		사이즈는 WIDTH - 400이하로 설정해 주시기 바랍니다.<br>
	<% dim arrimg2
    	IF imgurl2 <> "" THEN
    	 arrimg2 = split(imgurl2,"/")%>
    <p class="text2"><%=arrimg(ubound(arrimg2))%> <input type="checkbox" name="chkimg2">삭제<br>
    <%END IF%>
		<input type="file" name="sImg2">
	</td>
</tr>
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

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

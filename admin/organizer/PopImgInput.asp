<%@ language=vbscript %>
<% option explicit %>
<% Response.ChaRset = "EUC-KR" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/admin/lib/popheader.asp"-->

<%
dim YearUse,orgImgName ,newImgName ,ImagePath ,maxFileSize ,maxFileWidth ,makeThumbYn,divName,inputname
YearUse ="2009"
orgImgName 		= request("orgImgName")
ImagePath 		= request("ImagePath")
maxFileSize 	= request("maxFileSize")
maxFileWidth	= request("maxFileWidth")
makeThumbYn 	= request("makeThumbYn")
divName = request("divName")
inputname = request("inputname")

newImgName 		= year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())

%>
<script language="javascript">

function subchk(){
	if (document.regfrm.imgFile.value.length<1){
		alert('이미지를 선택해 주세요');
		return false;
	}
	document.regfrm.submit();
}
window.resizeTo(400,250);
</script>
<!-- 상단 메뉴 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="center">
        	<b>이미지삽입</b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" target="prcframe" action="http://upload.10x10.co.kr/linkweb/organizer/organizerImg_InputProc.asp" enctype="multipart/form-data">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="divName" value="<%= divName %>">
	<input type="hidden" name="orgImgName" value="<%= orgImgName %>">
	<input type="hidden" name="newImgName" value="<%= newImgName %>">
	<input type="hidden" name="inputname" value="<%= inputname %>">

	<input type="hidden" name="ImagePath" value="<%= ImagePath %>">
	<input type="hidden" name="maxFileSize" value="<%= maxFileSize %>">
	<input type="hidden" name="maxFileWidth" value="<%= maxFileWidth %>">
	<input type="hidden" name="makeThumbYn" value="<%= makeThumbYn %>">

	<tr bgcolor="#FFFFFF">
		<td>
			<input type="file" name="imgFile" size="35" value="">
			<br>
			이미지를 저장하실 경우 기존 이미지는 삭제 됩니다.
		</td>
	</tr>

	</form>
</table>
<!-- 하단 페이징 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<input type="button" class="button" value="저장" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="취소" onclick="self.close();"/>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
</body>
</html>
<iframe name="prcframe" src="" frameborder="0" width="400" height="200"></iframe>

<!-- #include virtual="/lib/db/dbclose.asp" -->
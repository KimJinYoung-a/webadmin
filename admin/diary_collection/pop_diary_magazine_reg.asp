<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%
dim YearUse
YearUse = request("YearUse")
%>
<script language="javascript">

function subchk(){

	if(document.regfrm.imagename.value.length<1){
		alert('이미지를 입력해 주세요');
		return false;
	}
	document.regfrm.submit();
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes,scrollbars=yes');
}
function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='diary_img_input.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='diary_img_input.asp';
	document.imginputfrm.submit();
}

document.domain = "10x10.co.kr";
window.resizeTo(600,600);
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
        	<b>다이어리 매거진 등록 </b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 중단 내용 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="proc_diary_magazine.asp">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="mode" value="reg">
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>TITLE</b></td>
		<td align="left">
			<input type="text" name="title" size="60" maxlength="120">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>사용 유무</b></td>
		<td>
			<label><input type="radio" name="isusing" value="Y" /> 사용 </label>
			<label><input type="radio" name="isusing" value="N" checked /> 사용안함 </label>
		</td>

	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>이미지 -1</b><br></td>
		<td>
			<input type="button" class="button" size="30" value="이미지 추가" onclick="jsImgInput('imgdiv','imagename','magazine','400','750','false');"/>
			(<b><font color="red">width :750 </font></b>|
			<b><font color="red">400 </font></b>KB|
			<b><font color="red">JPG,GIF</font></b>만가능)
			<input type="hidden" name="imagename" value="">
			<div align="right" id="imgdiv"></div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>TEXT & HTML -1</b></td>
		<td align="left">
			<textarea name="magazinetxt" cols="50" rows="10"></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>이미지 -2 </b><br></td>
		<td>
			<input type="button" class="button" size="30" value="이미지 추가" onclick="jsImgInput('imgdiv2','imagename2','magazine','400','750','false');"/>
			(<b><font color="red">width :750 </font></b>|
			<b><font color="red">400 </font></b>KB|
			<b><font color="red">JPG,GIF</font></b>만가능)
			<input type="hidden" name="imagename2" value="">
			<div align="right" id="imgdiv2"></div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>TEXT & HTML -2</b></td>
		<td align="left">
			<textarea name="magazinetxt2" cols="50" rows="10"></textarea>
		</td>
	</tr>
	</form>
</table>
<!-- 하단  시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<input type="button" class="button" value="확인" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="취소" onclick="history.go(-1);"/>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="<%= YearUse %>">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->

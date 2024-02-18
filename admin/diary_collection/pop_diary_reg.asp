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

	if(document.regfrm.itemid.value.length<1){
		document.regfrm.itemid.focus();
		alert('상품 번호를 입력하셔야 합니다.');
		return false;
	}
	if(document.regfrm.basicimgName.value.length<1){
		alert('이미지를 입력해 주세요');
		return false;
	}
	document.regfrm.submit();
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
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
        	<b>다이어리 기초정보 등록 </b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 중단 내용 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="proc_diary_reg.asp">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="100" bgcolor="<%= adminColor("topbar") %>"><b>구분</b></td>
		<td>
			<select name="diaryType">
				<option value="illust" selected >일러스트</option>
				<option value="photo">포토/명화</option>
				<option value="system">시스템</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>상품번호</b></td>
		<td><input type="text" name="itemid" maxlength="10" value="" /></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>기본 이미지</b><br></td>
		<td>
			<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','basicimgName','basic','200','400','true');"/>
			(<b><font color="red">400x400</font></b>,<b><font color="red">JPG</font></b>만가능)
			<input type="hidden" name="basicimgName" value="">
			<div align="right" id="imgdiv"></div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>기타</b></td>
		<td align="left">
			<label><input type="checkbox" name="hitYn"/>HIT</label><font color="orange" size="3"> | </font>
			<label><input type="checkbox" name="giftYn"/>Gift</label><font color="orange" size="3"> | </font>
			<label><input type="checkbox" name="onlyyearYn"/>2008년 전용 </label><font color="orange" size="3"> | </font>
			<label><input type="checkbox" name="freeBaeSongYn" disabled onclick="return false;"/>무료 배송 </label>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>사용 유무</b></td>
		<td>
			<label><input type="radio" name="isusing" value="Y" /> 사용 </label>
			<label><input type="radio" name="isusing" value="N" checked /> 사용안함 </label>
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
			<input type="button" class="button" value="취소" onclick="window.close();"/>
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

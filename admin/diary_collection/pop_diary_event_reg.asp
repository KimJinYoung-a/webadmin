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

	if(document.regfrm.bannerUrl.value.length<1){
		document.regfrm.bannerUrl.focus();
		alert('링크를 입력하셔야 합니다.');
		return false;
	}
	if(document.regfrm.imagename.value.length<1){
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

function jsSelevtType(v){

	var aa = document.getElementsByTagName("div");
	
	for (var i=0;i<aa.length;i++){
		if (aa[i].id==v||aa[i].id=='imgdiv'){
			aa[i].style.display='';
		}else {
			aa[i].style.display='none';
		}
	}
	if (v=='dzone'||v=='evtmain'){
		document.regfrm.mapusing.value='Y';
		fnmapusing('Y');
	} else {
		document.regfrm.mapusing.value='N';
		fnmapusing('N');
	}
}

function fnmapusing(v){
	var ipturl = document.getElementById('bannerUrl');
	var udiv = document.getElementById('urldiv');
	var tTxt = ipturl.value;
	
	if(v=='Y'){
		//udiv.innerHTML='<input type="text" name="bannerUrl" size="60" maxlength="100" value="">';
		udiv.innerHTML='<textarea name="bannerUrl" cols="60" rows="10" style="overflow=hidden" wrap="hard">'+ tTxt + '</textarea>';
	}else{
		udiv.innerHTML='<input type="text" name="bannerUrl" size="60" maxlength="100" value="' + tTxt + '">';
	}
}

document.domain = "10x10.co.kr";
window.resizeTo(700,500);
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
        	<b>다이어리 이벤트 배너등록</b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 중단 내용 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="proc_diary_event.asp">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="mapusing" value="N">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="100" bgcolor="<%= adminColor("topbar") %>"><b>배너위치</b></td>
		<td>
			<select name="bannerType" onchange="jsSelevtType(this.value)">
				<option value="multi">메인 멀티</option>
				<option value="left">좌측 메뉴</option>
				<option value="power">파워이벤트</option>
				<option value="today">Today`s Diary</option>
				<option value="quiz">Diary Quiz</option>
				<option value="dzone">디자인존</option>
				<option value="tdayitem">Today`s Item</option>
				<option value="evtmain">이벤트메인 배너</option>
				<option value="othermall_left">[외부몰]좌측메뉴</option>
				<option value="othermall_multi">[외부몰]메인멀티</option>
				<option value="othermall_right">[외부몰]우측메뉴</option>
			</select>

		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>이벤트 코드</b></td>
		<td><input type="text" name="evtcode" size="10" maxlength="10" value="" />(없을시 공란)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>URL</b></td>
		<td><!--<input type="checkbox" name="mapusing" onclick="fnmapusing(this.checked);">이미지맵 사용<br>-->
			(http://www.10x10.co.kr)생략 / 이미지맵 &lt;map name="design_zone"&gt; <br>
			<span id="urldiv"><input type="text" name="bannerUrl" size="60" maxlength="100" value=""></span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>배너 이미지</b><br></td>
		<td>
			<!-- 메인멀티 -->
			<div id="multi" style="display:none;">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','400','600','false');"/>
				(<b><font color="red">Width 600</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 좌측배너 -->
			<div id="left" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','194','false');"/>
				(<b><font color="red">Width 194</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 파워이벤트 -->
			<div id="power" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','230','false');"/>
				(<b><font color="red">Width 230</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 투데이 다이어리 -->
			<div id="today" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','192','false');"/>
				(<b><font color="red">192 X 124</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 퀴즈 이벤트 -->
			<div id="quiz" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','207','false');"/>
				(<b><font color="red">207 X 133</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 메인 디자인 존 -->
			<div id="dzone" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','285','false');"/>
				(<b><font color="red">285 X 133</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 투데이 아이템 -->
			<div id="tdayitem" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','153','false');"/>
				(<b><font color="red">153 X 145</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 이벤트 메인 배너 -->
			<div id="evtmain" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','400','750','false');"/>
				(<b><font color="red">width 750</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 제휴몰 좌측 -->
			<div id="othermall_left" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','194','false');"/>
				(<b><font color="red">Width 194</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 제휴몰 멀티 -->
			<div id="othermall_multi" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','400','600','false');"/>
				(<b><font color="red">Width 600</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<!-- 제휴몰 우측 -->
			<div id="othermall_right" style="display:none">
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','imagename','eventbanner','121','153','false');"/>
				(<b><font color="red">Width 153 height 121</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
			</div>
			<input type="hidden" name="imagename" value="">
			<div align="right" id="imgdiv"></div>
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
<script>jsSelevtType(document.regfrm.bannerType.value);fnmapusing(document.regfrm.mapusing.checked)</script>
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

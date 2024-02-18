<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
	CONST CMMS_IMG_MAXSIZE = 198   'KB
	CONST CeMail_IMG_MAXSIZE = 420   'KB

	dim oGiftcard, mode
	dim designId, cardItemid, groupDiv, cardDesignName, MMSThumb, MMSImage, MMSText, emailThumb, emailImage, emailText, isUsing, sortNo

	cardItemid = request("cardid")
	designId = request("designid")
	mode = "add"
	isUsing = "Y" : groupDiv = "1" : sortNo="1"

	if cardItemid<>"" then
		Set oGiftcard = new cGiftCard
		oGiftcard.FRectDesignId=designId
		oGiftcard.fGiftcard_oneDesign
		if oGiftcard.FResultCount>0 then
			cardItemid		= oGiftcard.FOneItem.FcardItemid
			groupDiv		= oGiftcard.FOneItem.FgroupDiv
			cardDesignName	= oGiftcard.FOneItem.FcardDesignName
			MMSThumb		= oGiftcard.FOneItem.FMMSThumb
			MMSImage		= oGiftcard.FOneItem.FMMSImage
			MMSText			= oGiftcard.FOneItem.FMMSText
			emailThumb		= oGiftcard.FOneItem.FemailThumb
			emailImage		= oGiftcard.FOneItem.FemailImage
			emailText		= oGiftcard.FOneItem.FemailText
			isUsing			= oGiftcard.FOneItem.FisUsing
			sortNo			= oGiftcard.FOneItem.FsortNo

			mode = "modi"
		end if

		Set oGiftcard = Nothing
	end if
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript">
<!--
// 저장하기
function SubmitSave() {
    if (validate(itemreg)==false) {
        return;
    }
    
    //디자인명 길이체크 64Byte
	if (getByteLength(itemreg.cardDesignName.value)>64){
	    alert("디자인명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.cardDesignName.focus();
		return;
	}

    if (itemreg.MMSImage.value != "") {
        if (CheckImage(itemreg.MMSImage, <%= CMMS_IMG_MAXSIZE %>, 640, 640, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.emailImage.value != "") {
        if (CheckImage(itemreg.emailImage, <%= CMMS_IMG_MAXSIZE %>, 800, 1000, 'jpg',40) != true) {
            return;
        }
    }

	if(confirm("디자인정보를 <%=chkIIF(mode="add","등록","수정")%>하시겠습니까?")){
		itemreg.action = "<%= ItemUploadUrl %>/linkweb/giftCard/doGiftcardDesignReg.asp";
		itemreg.target = "FrameCKP";
		itemreg.submit();
	}
}

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
 }

// 이미지표시
function ClearImage(img,fsize,wd,ht,filesize) {
	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", "+filesize+", "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";
}

function CheckExtension(imgname, allowext) {
    var ext = imgname.lastIndexOf(".");
    if (ext < 0) {
        return false;
    }

    ext = imgname.toLowerCase().substring(ext + 1);
    allowext = "," + allowext + ",";
    if (allowext.indexOf(ext) < 0) {
        return false;
    }

    return true;
}

function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
{
    var ext;
    var filename;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지파일은 다음의 파일만 사용하세요.[" + extname + "]");
        ClearImage(img,fsize, imagewidth, imageheight,filesize);
        return false;
    }

    return true;
}

// 그룹/소트번호 변경 > 디자인번호 변경
function chgGrpSrt(gcd,sno) {
	if(sno<10) {
		itemreg.designId.value= gcd + "0" + sno;
	} else {
		itemreg.designId.value= gcd + sno;
	}
}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>기프트카드 디자인 <%=chkIIF(mode="add","등록","수정")%></strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<b><%=chkIIF(mode="add","신규 기프트카드 디자인정보를 등록합니다.","등록된 기프트카드 디자인정보를 수정합니다.")%></b>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>기본정보</strong></td>
	<td align="right">상품코드 : <strong><%=cardItemId%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<form name="itemreg" method="post" onsubmit="return false;" style="margin:0px;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="cardItemId" value="<%=cardItemId%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">디자인코드 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="designId" readonly size="10" class="text_ro" value="<%=designId%>" id="[on,off,off,off][디자인코드]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">디자인그룹 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<select name="groupDiv" <%=chkIIF(mode="modi","disabled","")%> class="select" onchange="chgGrpSrt(this.value,sortNo.value)">
			<option value="1" <%=chkIIF(groupDiv="1","selected","")%>>기본-Basic</option>
			<option value="2" <%=chkIIF(groupDiv="2","selected","")%>>생일-Birthday</option>
			<option value="3" <%=chkIIF(groupDiv="3","selected","")%>>감사-Thanks</option>
			<option value="4" <%=chkIIF(groupDiv="4","selected","")%>>축하-Congratulation</option>
			<option value="5" <%=chkIIF(groupDiv="5","selected","")%>>사랑-Love</option>
		</select>
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">디자인명 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardDesignName" maxlength="64" size="40" class="text" value="<%=cardDesignName%>" id="[on,off,off,off][디자인명]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">표시순서 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="sortNo" <%=chkIIF(mode="modi","readonly","")%> maxlength="2" size="3" class="text" value="<%=sortNo%>" id="[on,on,off,off][표시순서]" onkeyup="chgGrpSrt(groupDiv.value,this.value)">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">사용여부 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="radio" name="isUsing" value="Y" <%=chkIIF(isUsing="Y","checked","")%>>사용
		<input type="radio" name="isUsing" value="N" <%=chkIIF(isUsing="N","checked","")%>>사용안함
	</td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>MMS 정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="30" align="left">
	<td width="20%" bgcolor="#DDDDFF">MMS이미지 :</td>
	<td width="80%" bgcolor="#FFFFFF">
	  <% if (MMSImage <> "") then %>
		<div id="divMMSImage"><img src="<%= MMSImage %>" width="300"></div>
	  <% end if %>
	  <input type="file" name="MMSImage" onchange="CheckImage(this, <%= CMMS_IMG_MAXSIZE %>, 640, 640, 'jpg', 40);" class="text" size="40">
	  <br>(640 X 640px 이내, <b><font color="red">jpg</font></b>, <b><font color=red><%= CMMS_IMG_MAXSIZE %>kb</b> 이하</font>)
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">MMS기본문구 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="MMSText" rows="5" class="textarea" style="width:100%" id="[off,off,off,off][MMS기본문구]"><%=MMSText%></textarea>
	</td>
</tr>
</table>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>이메일 정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="30" align="left">
	<td width="20%" bgcolor="#DDDDFF">이메일이미지 :</td>
	<td width="80%" bgcolor="#FFFFFF">
	  <% if (emailImage <> "") then %>
		<div id="divemailImage"><img src="<%= emailImage %>" width="300"></div>
	  <% end if %>
	  <input type="file" name="emailImage" onchange="CheckImage(this, <%= CeMail_IMG_MAXSIZE %>, 800, 1000, 'jpg', 40);" class="text" size="40">
	  <br>(800 X 1000px 이내, <b><font color="red">jpg</font></b>, <font color=red><%= CeMail_IMG_MAXSIZE %>kb 이하</font>)
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">이메일기본문구 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="emailText" rows="5" class="textarea" style="width:100%" id="[off,off,off,off][이메일기본문구]"><%=emailText%></textarea>
	</td>
</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
      <input type="button" value=" <%=chkIIF(mode="add","등 록","수 정")%> " class="button" onclick="SubmitSave();">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</p>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
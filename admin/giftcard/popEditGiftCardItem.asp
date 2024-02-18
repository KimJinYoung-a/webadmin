<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 기프트카드 상품목록
' History : 이상구 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
	CONST CBASIC_IMG_MAXSIZE = 260   'KB

	dim oGiftcard, mode
	dim cardItemid, cardItemName, cardInfo, cardDesc, cardSellYn, basicImage

	cardItemid = request("cardid")
	mode = "add"
	cardSellYn = "Y"

	if cardItemid<>"" then
		Set oGiftcard = new cGiftCard
		oGiftcard.FRectCardItemid=cardItemid
		oGiftcard.fGiftcard_oneItem
		if oGiftcard.FResultCount>0 then
			cardItemid		= oGiftcard.FOneItem.FcardItemid
			cardItemName	= ReplaceBracket(oGiftcard.FOneItem.FcardItemName)
			cardInfo		= ReplaceBracket(oGiftcard.FOneItem.FcardInfo)
			cardDesc		= ReplaceBracket(oGiftcard.FOneItem.FcardDesc)
			cardSellYn		= oGiftcard.FOneItem.FcardSellYn
			basicImage		= oGiftcard.FOneItem.FbasicImage

			mode = "modi"
		end if

		Set oGiftcard = Nothing
	end if
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type='text/javascript'>
<!--
// 저장하기
function SubmitSave() {
    if (validate(itemreg)==false) {
        return;
    }
    
    //상품명 길이체크 64Byte
	if (getByteLength(itemreg.cardItemName.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.cardItemName.focus();
		return;
	}

    //간략설명 길이체크 600Byte
	if (getByteLength(itemreg.cardInfo.value)>600){
	    alert("간략설명은 최대 600byte 이하로 입력해주세요.(한글300자 또는 영문600자)");
		itemreg.cardInfo.focus();
		return;
	}

    if (itemreg.basicImage.value != "") {
        if (CheckImage(itemreg.basicImage, <%= CBASIC_IMG_MAXSIZE %>, 600, 600, 'jpg',32) != true) {
            return;
        }
    } else {
        <% if mode="add" then %>
        alert("기본이미지는 필수입니다.");
        return;
        <% end if %>
    }

	if(confirm("상품을 <%=chkIIF(mode="add","등록","수정")%>하시겠습니까?")){
		itemreg.action = "<%= ItemUploadUrl %>/linkweb/giftCard/doGiftcardItemReg.asp";
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
function ClearImage(img,fsize,wd,ht) {
	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";
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
        ClearImage(img,fsize, imagewidth, imageheight);
        return false;
    }

    return true;
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
	<font color="red"><strong>기프트카드 상품 <%=chkIIF(mode="add","등록","수정")%></strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<b><%=chkIIF(mode="add","신규 기프트카드 상품정보를 등록합니다.","등록된 기프트카드 상품정보를 수정합니다.")%></b>
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
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>상품 기본정보</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<form name="itemreg" method="post" onsubmit="return false;" style="margin:0px;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<%=mode%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<% if mode="modi" then %>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">상품코드 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardItemId" readonly size="10" class="text_ro" value="<%=cardItemId%>" id="[on,off,off,off][상품코드]">
	</td>
</tr>
<% end if %>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">상품명 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardItemName" maxlength="64" size="40" class="text" value="<%=cardItemName%>" id="[on,off,off,off][상품명]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">간략설명 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="cardInfo" rows="3" class="textarea" style="width:100%" id="[on,off,off,off][간략설명]"><%=cardInfo%></textarea>
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">상세설명 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="cardDesc" rows="7" class="textarea" style="width:100%" id="[on,off,off,off][상세설명]"><%=cardDesc%></textarea>
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">판매여부 :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="radio" name="cardSellYn" value="Y" <%=chkIIF(cardSellYn="Y","checked","")%>>판매
		<input type="radio" name="cardSellYn" value="N" <%=chkIIF(cardSellYn="N","checked","")%>>품절
	</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>이미지정보</strong>
		<br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="30" align="left">
	<td width="20%" bgcolor="#DDDDFF">기본이미지 :</td>
	<td width="80%" bgcolor="#FFFFFF">
	  <% if (basicImage <> "") then %>
		<div id="basicImage"><img src="<%= basicImage %>" width="300" height="300"></div>
	  <% end if %>
	  <input type="file" name="basicImage" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 600, 600, 'jpg', 32);" class="text" size="32">
	  (<font color=red>필수</font>,600X600,<b><font color="red">jpg</font></b>)
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

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����Ʈī�� ��ǰ���
' History : �̻� ����
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
// �����ϱ�
function SubmitSave() {
    if (validate(itemreg)==false) {
        return;
    }
    
    //��ǰ�� ����üũ 64Byte
	if (getByteLength(itemreg.cardItemName.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.cardItemName.focus();
		return;
	}

    //�������� ����üũ 600Byte
	if (getByteLength(itemreg.cardInfo.value)>600){
	    alert("���������� �ִ� 600byte ���Ϸ� �Է����ּ���.(�ѱ�300�� �Ǵ� ����600��)");
		itemreg.cardInfo.focus();
		return;
	}

    if (itemreg.basicImage.value != "") {
        if (CheckImage(itemreg.basicImage, <%= CBASIC_IMG_MAXSIZE %>, 600, 600, 'jpg',32) != true) {
            return;
        }
    } else {
        <% if mode="add" then %>
        alert("�⺻�̹����� �ʼ��Դϴ�.");
        return;
        <% end if %>
    }

	if(confirm("��ǰ�� <%=chkIIF(mode="add","���","����")%>�Ͻðڽ��ϱ�?")){
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

// �̹���ǥ��
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
        alert("�̹��������� ������ ���ϸ� ����ϼ���.[" + extname + "]");
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
	<font color="red"><strong>����Ʈī�� ��ǰ <%=chkIIF(mode="add","���","����")%></strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<b><%=chkIIF(mode="add","�ű� ����Ʈī�� ��ǰ������ ����մϴ�.","��ϵ� ����Ʈī�� ��ǰ������ �����մϴ�.")%></b>
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

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>��ǰ �⺻����</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<form name="itemreg" method="post" onsubmit="return false;" style="margin:0px;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<%=mode%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<% if mode="modi" then %>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardItemId" readonly size="10" class="text_ro" value="<%=cardItemId%>" id="[on,off,off,off][��ǰ�ڵ�]">
	</td>
</tr>
<% end if %>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">��ǰ�� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardItemName" maxlength="64" size="40" class="text" value="<%=cardItemName%>" id="[on,off,off,off][��ǰ��]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�������� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="cardInfo" rows="3" class="textarea" style="width:100%" id="[on,off,off,off][��������]"><%=cardInfo%></textarea>
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�󼼼��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="cardDesc" rows="7" class="textarea" style="width:100%" id="[on,off,off,off][�󼼼���]"><%=cardDesc%></textarea>
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="radio" name="cardSellYn" value="Y" <%=chkIIF(cardSellYn="Y","checked","")%>>�Ǹ�
		<input type="radio" name="cardSellYn" value="N" <%=chkIIF(cardSellYn="N","checked","")%>>ǰ��
	</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>�̹�������</strong>
		<br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="30" align="left">
	<td width="20%" bgcolor="#DDDDFF">�⺻�̹��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
	  <% if (basicImage <> "") then %>
		<div id="basicImage"><img src="<%= basicImage %>" width="300" height="300"></div>
	  <% end if %>
	  <input type="file" name="basicImage" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 600, 600, 'jpg', 32);" class="text" size="32">
	  (<font color=red>�ʼ�</font>,600X600,<b><font color="red">jpg</font></b>)
	</td>
</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
      <input type="button" value=" <%=chkIIF(mode="add","�� ��","�� ��")%> " class="button" onclick="SubmitSave();">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
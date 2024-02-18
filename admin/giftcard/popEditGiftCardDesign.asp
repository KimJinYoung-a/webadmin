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
// �����ϱ�
function SubmitSave() {
    if (validate(itemreg)==false) {
        return;
    }
    
    //�����θ� ����üũ 64Byte
	if (getByteLength(itemreg.cardDesignName.value)>64){
	    alert("�����θ��� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
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

	if(confirm("������������ <%=chkIIF(mode="add","���","����")%>�Ͻðڽ��ϱ�?")){
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

// �̹���ǥ��
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
        alert("�̹��������� ������ ���ϸ� ����ϼ���.[" + extname + "]");
        ClearImage(img,fsize, imagewidth, imageheight,filesize);
        return false;
    }

    return true;
}

// �׷�/��Ʈ��ȣ ���� > �����ι�ȣ ����
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
	<font color="red"><strong>����Ʈī�� ������ <%=chkIIF(mode="add","���","����")%></strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<b><%=chkIIF(mode="add","�ű� ����Ʈī�� ������������ ����մϴ�.","��ϵ� ����Ʈī�� ������������ �����մϴ�.")%></b>
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
<tr height="5">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>�⺻����</strong></td>
	<td align="right">��ǰ�ڵ� : <strong><%=cardItemId%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<form name="itemreg" method="post" onsubmit="return false;" style="margin:0px;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="cardItemId" value="<%=cardItemId%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�������ڵ� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="designId" readonly size="10" class="text_ro" value="<%=designId%>" id="[on,off,off,off][�������ڵ�]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�����α׷� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<select name="groupDiv" <%=chkIIF(mode="modi","disabled","")%> class="select" onchange="chgGrpSrt(this.value,sortNo.value)">
			<option value="1" <%=chkIIF(groupDiv="1","selected","")%>>�⺻-Basic</option>
			<option value="2" <%=chkIIF(groupDiv="2","selected","")%>>����-Birthday</option>
			<option value="3" <%=chkIIF(groupDiv="3","selected","")%>>����-Thanks</option>
			<option value="4" <%=chkIIF(groupDiv="4","selected","")%>>����-Congratulation</option>
			<option value="5" <%=chkIIF(groupDiv="5","selected","")%>>���-Love</option>
		</select>
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�����θ� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardDesignName" maxlength="64" size="40" class="text" value="<%=cardDesignName%>" id="[on,off,off,off][�����θ�]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">ǥ�ü��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="sortNo" <%=chkIIF(mode="modi","readonly","")%> maxlength="2" size="3" class="text" value="<%=sortNo%>" id="[on,on,off,off][ǥ�ü���]" onkeyup="chgGrpSrt(groupDiv.value,this.value)">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">��뿩�� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="radio" name="isUsing" value="Y" <%=chkIIF(isUsing="Y","checked","")%>>���
		<input type="radio" name="isUsing" value="N" <%=chkIIF(isUsing="N","checked","")%>>������
	</td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>MMS ����</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="30" align="left">
	<td width="20%" bgcolor="#DDDDFF">MMS�̹��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
	  <% if (MMSImage <> "") then %>
		<div id="divMMSImage"><img src="<%= MMSImage %>" width="300"></div>
	  <% end if %>
	  <input type="file" name="MMSImage" onchange="CheckImage(this, <%= CMMS_IMG_MAXSIZE %>, 640, 640, 'jpg', 40);" class="text" size="40">
	  <br>(640 X 640px �̳�, <b><font color="red">jpg</font></b>, <b><font color=red><%= CMMS_IMG_MAXSIZE %>kb</b> ����</font>)
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">MMS�⺻���� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="MMSText" rows="5" class="textarea" style="width:100%" id="[off,off,off,off][MMS�⺻����]"><%=MMSText%></textarea>
	</td>
</tr>
</table>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>�̸��� ����</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="30" align="left">
	<td width="20%" bgcolor="#DDDDFF">�̸����̹��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
	  <% if (emailImage <> "") then %>
		<div id="divemailImage"><img src="<%= emailImage %>" width="300"></div>
	  <% end if %>
	  <input type="file" name="emailImage" onchange="CheckImage(this, <%= CeMail_IMG_MAXSIZE %>, 800, 1000, 'jpg', 40);" class="text" size="40">
	  <br>(800 X 1000px �̳�, <b><font color="red">jpg</font></b>, <font color=red><%= CeMail_IMG_MAXSIZE %>kb ����</font>)
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�̸��ϱ⺻���� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<textarea name="emailText" rows="5" class="textarea" style="width:100%" id="[off,off,off,off][�̸��ϱ⺻����]"><%=emailText%></textarea>
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
</p>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
	dim oGiftcard, cardItemid, cardOption, mode
	dim cardOptionName, cardSellCash, optSellYn

	cardItemid = request("cardid")
	cardOption = request("cardOption")
	mode = "add"
	optSellYn = "Y"

	if cardOption<>"" then
		Set oGiftcard = new cGiftCard
		oGiftcard.FRectCardItemid=cardItemid
		oGiftcard.FRectCardOption=cardOption
		oGiftcard.fGiftcard_oneOption
		if oGiftcard.FResultCount>0 then
			cardOptionName	= oGiftcard.FOneItem.FcardOptionName
			cardSellCash	= oGiftcard.FOneItem.FcardSellCash
			optSellYn		= oGiftcard.FOneItem.FoptSellYn

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
    
    //�ɼǸ� ����üũ 64Byte
	if (getByteLength(itemreg.cardOptionName.value)>64){
	    alert("�ɼǸ��� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.cardOptionName.focus();
		return;
	}

	if(confirm("�ɼ��� <%=chkIIF(mode="add","���","����")%>�Ͻðڽ��ϱ�?")){
		itemreg.action = "doGiftcardOptionProc.asp";
		itemreg.target = "FrameCKP";
		itemreg.mode.value = "<%=mode%>";
		itemreg.submit();
	}
}

function delOption() {
	if(confirm("�ɼ��� �����Ͻðڽ��ϱ�?")){
		itemreg.action = "doGiftcardOptionProc.asp";
		itemreg.target = "FrameCKP";
		itemreg.mode.value = "del";
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
	<font color="red"><strong>����Ʈī�� �ɼ� <%=chkIIF(mode="add","���","����")%></strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<b><%=chkIIF(mode="add","�ű� ����Ʈī�� �ɼ������� ����մϴ�.","��ϵ� ����Ʈī�� �ɼ������� �����մϴ�.")%></b>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> ��ǰ�ڵ� : <strong><%=cardItemId%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<form name="itemreg" method="post" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="cardItemId" value="<%=cardItemId%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<% if mode="modi" then %>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�ɼ��ڵ� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardOption" readonly size="5" class="text_ro" value="<%=cardOption%>" id="[on,off,off,off][�ɼ��ڵ�]">
	</td>
</tr>
<% end if %>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�ɼǸ� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardOptionName" maxlength="64" size="40" class="text" value="<%=cardOptionName%>" id="[on,off,off,off][�ɼǸ�]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�ǸŰ� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="text" name="cardSellCash" size="8" class="text" value="<%=cardSellCash%>" id="[on,on,off,off][�ǸŰ�]">
	</td>
</tr>
<tr align="left" height="30">
	<td width="20%" bgcolor="#DDDDFF">�Ǹſ��� :</td>
	<td width="80%" bgcolor="#FFFFFF">
		<input type="radio" name="optSellYn" value="Y" <%=chkIIF(optSellYn="Y","checked","")%>>�Ǹ�
		<input type="radio" name="optSellYn" value="N" <%=chkIIF(optSellYn="N","checked","")%>>ǰ��
	</td>
</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
      <input type="button" value=" <%=chkIIF(mode="add","�� ��","�� ��")%> " class="button" onclick="SubmitSave();">
      <% if mode="modi" then %>
      &nbsp; &nbsp;<input type="button" value=" �� �� " class="button" onclick="delOption();" style="background-color:#FFDDDD;">
      <% end if %>
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
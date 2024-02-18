<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mainWCMSCls.asp" -->
<%
'###############################################
' PageName : popSubItemEdit.asp
' Discription : ���������� ��ǰ�ڵ� �ϰ� ���
' History : 2013.05.14 ������ : �ű� ����
'###############################################

'// ���� ����
Dim mainIdx, i
Dim oTemplate, oMain
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isExtDataUse, isImgDescUse, tplinfoDesc, tplSortNo
Dim itemname, smallImage

Dim mainStartDate, mainEndDate, mainTitle, mainSubNum

Dim subImage1, subImage2, subLinkUrl, subText1, subText2, subItemid, subVideoUrl, subBGColor, subImageDesc
Dim subSortNo, subRegUserid, subRegDate, subLastModiUserid, subLastModiDate, subIsUsing

'// �Ķ���� ����
mainIdx = request("mainIdx")

if mainIdx="" then
	Call Alert_Close("�Ķ���� ����(Err:01)")
	dbget.Close: Response.End
end if

'// ���� ����
set oMain = new CCMSContent
	oMain.FRectMainIdx = MainIdx
	oMain.GetOneMainPage
	if oMain.FResultCount>0 then
		tplIdx = oMain.FOneItem.FtplIdx
		mainStartDate = oMain.FOneItem.FmainStartDate
		mainEndDate = oMain.FOneItem.FmainEndDate
		mainTitle = oMain.FOneItem.FmainTitle
		mainSubNum = oMain.FOneItem.FmainSubNum
	end if
set oMain = Nothing

if tplIdx="" then
	Call Alert_Close("�������� �ʰų� ������ �����Դϴ�. (Err:02)")
	dbget.Close: Response.End
end if

'// ���ø� ����
set oTemplate = new CCMSContent
oTemplate.FRectTplIdx = tplIdx
oTemplate.GetOneTemplate
if oTemplate.FResultCount>0 then
	tplType			= oTemplate.FOneItem.FtplType
	tplName			= oTemplate.FOneItem.FtplName
	isTimeUse		= oTemplate.FOneItem.FisTimeUse
	isIconUse		= oTemplate.FOneItem.FisIconUse
	isSubNumUse		= oTemplate.FOneItem.FisSubNumUse
	isTopImgUse		= oTemplate.FOneItem.FisTopImgUse
	isTopLinkUse	= oTemplate.FOneItem.FisTopLinkUse
	isImageUse		= oTemplate.FOneItem.FisImageUse
	isTextUse		= oTemplate.FOneItem.FisTextUse
	isLinkUse		= oTemplate.FOneItem.FisLinkUse
	isItemUse		= oTemplate.FOneItem.FisItemUse
	isVideoUse		= oTemplate.FOneItem.FisVideoUse
	isBGColorUse	= oTemplate.FOneItem.FisBGColorUse
	isExtDataUse	= oTemplate.FOneItem.FisExtDataUse
	isImgDescUse	= oTemplate.FOneItem.FisImgDescUse
	tplinfoDesc		= oTemplate.FOneItem.FtplinfoDesc
	tplSortNo		= oTemplate.FOneItem.FtplSortNo
end if
set oTemplate = Nothing

if isExtDataUse="Y" then
	Call Alert_Close("�ܺε����͸� ����ϴ� ���� ���ø��Դϴ�.\n���縦 ����� �� �����ϴ�.")
	dbget.Close: Response.End
end if

if Not(isItemUse="Y") then
	Call Alert_Close("��ǰ�ڵ带 ������� �ʴ� ���� ���ø��Դϴ�.\n���縦 ����� �� �����ϴ�.")
	dbget.Close: Response.End
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//������ư
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

// ���˻�
function SaveForm(frm) {
	var selChk=true;
	if(frm.subItemidArray.value=="") {
		alert("�ϰ� ����Ͻ� ��ǰ�ڵ带 �Է����ּ���");
		frm.subItemidArray.focus();
		return;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<center>
<form name="frmSub" method="post" action="doSubRegItemCdArray.asp" style="margin:0px;">
<input type="hidden" name="mainIdx" value="<%=mainIdx%>" />
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>���� ���� - ��ǰ�ڵ� �ϰ� ���</b></td>
</tr>
<colgroup>
	<col width="100" />
	<col width="*" />
	<col width="100" />
	<col width="*" />
</colgroup>
<tr height="26" bgcolor="#FFFFFF">
    <td rowspan="2" bgcolor="#DDDDFF">���ø�</td>
    <td colspan="3">
        [<%=tplName %>]
        <b>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="3" style="padding:5px;">
        <%=nl2br(tplinfoDesc)%>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ǰ�ڵ�</td>
    <td colspan="3">
        <textarea name="subItemidArray" class="textarea" title="��ǰ�ڵ�" style="width:100%; height:80px;"></textarea>
        <p>�� ��ǰ�ڵ带 ��ǥ(,) �Ǵ� ���ͷ� �����Ͽ� �Է�</p>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ļ���</td>
    <td>
        <input type="text" name="subSortNo" class="text" size="4" value="0" />
    </td>
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="subIsUsing" id="rdoUsing1" value="Y" checked /><label for="rdoUsing1">���</label>
		<input type="radio" name="subIsUsing" id="rdoUsing2" value="N" /><label for="rdoUsing2">����</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveForm(this.form);"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
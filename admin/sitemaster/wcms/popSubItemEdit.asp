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
' Discription : ���������� ���/����
' History : 2013.04.04 ������ : �ű� ����
'###############################################

'// ���� ����
Dim mainIdx, subIdx, i
Dim oTemplate, oMain, oSub
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isExtDataUse, isImgDescUse, tplinfoDesc, tplSortNo
Dim itemname, smallImage

Dim mainStartDate, mainEndDate, mainTitle, mainSubNum

Dim subImage1, subImage2, subLinkUrl, subText1, subText2, subItemid, subVideoUrl, subBGColor, subImageDesc
Dim subSortNo, subRegUserid, subRegDate, subLastModiUserid, subLastModiDate, subIsUsing

'// �Ķ���� ����
mainIdx = request("mainIdx")
subIdx = request("subIdx")

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

'// ���� ������ ����
set oSub = new CCMSContent
oSub.FRectSubIdx = subIdx
if subIdx<>"" then
	oSub.GetOneSubItem

	if oSub.FResultCount>0 then
		subImage1			= oSub.FOneItem.getImageUrl(1)
		subImage2			= oSub.FOneItem.getImageUrl(2)
		subLinkUrl			= oSub.FOneItem.FsubLinkUrl
		subText1			= oSub.FOneItem.FsubText1
		subText2			= oSub.FOneItem.FsubText2
		subItemid			= oSub.FOneItem.FsubItemid
		subVideoUrl			= oSub.FOneItem.FsubVideoUrl
		subBGColor			= oSub.FOneItem.FsubBGColor
		subImageDesc		= oSub.FOneItem.FsubImageDesc
		subSortNo			= oSub.FOneItem.FsubSortNo
		subRegUserid		= oSub.FOneItem.FsubRegUserid
		subRegDate			= oSub.FOneItem.FsubRegDate
		subLastModiUserid	= oSub.FOneItem.FsubLastModiUserid
		subLastModiDate		= oSub.FOneItem.FsubLastModiDate
		subIsUsing			= oSub.FOneItem.FsubIsUsing
		itemName			= oSub.FOneItem.FitemName
		smallImage			= oSub.FOneItem.FsmallImage
	else
		Call Alert_Close("�������� �ʰų� ������ �����Դϴ�. (Err:03)")
		dbget.Close: Response.End
	end if
else
	subSortNo = "0"
end if
set oSub = Nothing
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
	//�÷���Ŀ
	$("input[name='subBGColor']").colorpicker();
});

// ���˻�
function SaveForm(frm) {
	var selChk=true;
	$("input[type='text'],input[type='file']").each(function(){
		if($(this).val()==""&&$(this).attr("require")!="N") {
			alert($(this).attr("title")+"��(��) �Է����ּ���");
			$(this).focus();
			selChk=false;
			return false;
		}
	});

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}

// ��ǰ���� ����
function fnGetItemInfo(iid) {
	$.ajax({
		type: "GET",
		url: "act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='50' />"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo").fadeIn();
				$("#lyItemInfo").html(rst);
			} else {
				$("#lyItemInfo").fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo").fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}
</script>
<center>
<form name="frmSub" method="post" action="<%=uploadUrl%>/linkweb/wcms/doSubContentsReg.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mainIdx" value="<%=mainIdx%>" />
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>���� ���� ���/����</b></td>
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
<% if subIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���� ��ȣ</td>
    <td colspan="3">
        <%=subIdx %>
        <input type="hidden" name="subIdx" value="<%=subIdx%>" />
    </td>
</tr>
<% end if %>
<% if isImageUse="Y" or isTopImgUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�̹��� #1</td>
    <td>
		<input type="file" name="subImage1" class="file" title="�̹��� #1" require="N" style="width:100%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		<% end if %>
    </td>
    <td bgcolor="#DDDDFF">�̹��� #2 (�ɼ�)</td>
    <td>
		<input type="file" name="subImage2" class="file" title="�̹��� #2" require="N" style="width:100%;" />
		<% if subImage2<>"" then %>
		<br>
		<img src="<%= subImage2 %>" width="100" /><br><%= subImage2 %>
		<% end if %>
    </td>
</tr>
<% end if %>
<% if isImgDescUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�̹��� ����</td>
    <td colspan="3">
        <textarea name="subImageDesc" class="textarea" style="width:100%; height:60px;" title="�̹��� ����(for 'alt' Tag)"><%=subImageDesc%></textarea>
    </td>
</tr>
<% end if %>
<% if isTextUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�ؽ�Ʈ #1</td>
    <td>
		<input type="text" name="subText1" maxlength="128" value="<%=subText1%>" title="�ؽ�Ʈ #1" require="N" class="text" style="width:100%;" />
    </td>
    <td bgcolor="#DDDDFF">�ؽ�Ʈ #2 (�ɼ�)</td>
    <td>
		<input type="text" name="subText2" maxlength="128" value="<%=subText2%>" title="�ؽ�Ʈ #2" require="N" class="text" style="width:100%;" />
    </td>
</tr>
<% end if %>
<% if isBGColorUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">������</td>
    <td colspan="3">
		<input type="text" name="subBGColor" value="<%=subBGColor%>" class="text" require="N" style="width:80px;" />
    </td>
</tr>
<% end if %>
<% if isLinkUse="Y" or isTopLinkUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ũURL</td>
    <td colspan="3">
        <input type="text" name="subLinkUrl" value="<%= subLinkUrl %>" maxlength="256" title="��ũURL" <%=chkIIF(isTopLinkUse="Y","require='N'","")%> class="text" style="width:100%;" />
    </td>
</tr>
<% end if %>
<% if isItemUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ǰ�ڵ�</td>
    <td colspan="3">
        <input type="text" name="subItemid" value="<%= subItemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="��ǰ�ڵ�" />
        <div id="lyItemInfo" style="display:<%=chkIIF(subItemid="","none","")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='50' />"
        		Response.Write itemName
        	end if
        %>
        </div>
    </td>
</tr>
<% end if %>
<% if isVideoUse="Y" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">������URL</td>
    <td colspan="3">
        <input type="text" name="subVideoUrl" value="<%= subVideoUrl %>" maxlength="256" class="text" title="������URL" style="width:100%;" />
    </td>
</tr>
<% end if %>
<% if SubIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�����</td>
    <td><%=getStaffUserName(subRegUserId)%></td>
    <td bgcolor="#DDDDFF">�����</td>
    <td><%=subRegDate%></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�۾���</td>
    <td><%=getStaffUserName(subLastModiUserid)%></td>
    <td bgcolor="#DDDDFF">�۾���</td>
    <td><%=subLastModiDate%></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ļ���</td>
    <td>
        <input type="text" name="subSortNo" class="text" size="4" value="<%=subSortNo%>" />
    </td>
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="subIsUsing" id="rdoUsing1" value="Y" <%=chkIIF(subIsUsing="Y" or subIsUsing="","checked","")%> /><label for="rdoUsing1">���</label>
		<input type="radio" name="subIsUsing" id="rdoUsing2" value="N" <%=chkIIF(subIsUsing="N","checked","")%> /><label for="rdoUsing2">����</label>
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
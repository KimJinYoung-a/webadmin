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
' PageName : popTemplateEdit.asp
' Discription : ���ø� ���/����
' History : 2013.04.01 ������ : �ű� ����
'###############################################

'// ���� ����
Dim siteDiv, pageDiv, i, page
Dim oTemplate
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isExtDataUse, isImgDescUse, tplinfoDesc, tplSortNo

'// �Ķ���� ����
siteDiv = request("site")
pageDiv = request("pDiv")
tplIdx = request("tplIdx")
page = request("page")

if siteDiv="" then siteDiv="P"		'�⺻�� PC��(P:PC��, M:�����)
if pageDiv="" then pageDiv="10"		'�⺻�� ����Ʈ����(10:����Ʈ����, 20:�̺�Ʈ����...)
if page="" then page="1"

'// ���ø� ����
	set oTemplate = new CCMSContent
	oTemplate.FRectTplIdx = tplIdx
    if tplIdx<>"" then
    	oTemplate.GetOneTemplate
		if oTemplate.FResultCount>0 then
			tplType			= oTemplate.FOneItem.FtplType
			tplName			= oTemplate.FOneItem.FtplName
			siteDiv			= oTemplate.FOneItem.FsiteDiv
			pageDiv			= oTemplate.FOneItem.FpageDiv
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
    else
    	tplSortNo = "0"
    end if
    set oTemplate = Nothing

'// ���ø� ���
	set oTemplate = new CCMSContent
	oTemplate.FRectSiteDiv = siteDiv
	oTemplate.FRectPageDiv = pageDiv
	oTemplate.FCurrPage = page
    oTemplate.GetTemplateList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
//���ø� ����(preset,group)
function chgTplType(v) {
	switch(v) {
		case "A" :
			$("#tplTpDesc").html("�����+Ű���� ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isImageUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isImgDescUse']").val("Y");
			break;
		case "B" :
			$("#tplTpDesc").html("�ؽ�Ʈ ��ũ ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isBGColorUse']").val("Y");
			break;
		case "C" :
			$("#tplTpDesc").html("�̹��� ��ũ ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isImageUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isBGColorUse']").val("Y");
			$("select[name='isImgDescUse']").val("Y");
			break;
		case "D" :
			$("#tplTpDesc").html("ī��/������ ��ǰ��ũ ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isIconUse']").val("Y");
			$("select[name='isSubNumUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isItemUse']").val("Y");
			break;
		case "E" :
			$("#tplTpDesc").html("����Ʈ ��ǰ ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isExtDataUse']").val("Y");
			break;
		case "F" :
			$("#tplTpDesc").html("�̹���, ��ǰ ��ũ ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isSubNumUse']").val("Y");
			$("select[name='isTopImgUse']").val("Y");
			$("select[name='isTopLinkUse']").val("Y");
			$("select[name='isItemUse']").val("Y");
			break;
		case "G" :
			$("#tplTpDesc").html("������ ��ũ ����");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isVideoUse']").val("Y");
			break;
		default :
			$("#tplTpDesc").html("");
			$("select:not(select[name='tplType'])").val("");
	}
}

// ���˻�
function SaveTemplate(frm) {
	var selChk=true;
	$("select").each(function(){
		if($(this).val()=="") {
			alert($(this).attr("title")+"��(��) �������ּ���");
			$(this).focus();
			selChk=false;
			return false;
		}
	});
	if(!selChk) return;

	if($("input[name='tplName']").val()=="") {
		alert("���ø����� �Է����ּ���.");
		$("input[name='tplName']").focus();
		selChk=false;
	}
	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<center>
<form name="frmTemplate" method="post" action="doTemplate.asp" style="margin:0px;">
<input type="hidden" name="page" value="" />
<table width="690" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>���ø� ���/����</b></td>
</tr>
<% if tplIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">���ø���ȣ</td>
    <td width="610" colspan="3">
        <%=tplIdx %>
        <input type="hidden" name="tplIdx" value="<%=tplIdx %>" />
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">���ø�����</td>
    <td width="610" colspan="3">
        <select name="tplType" class="select" onchange="chgTplType(this.value)" title="���ø�����">
        	<option value="">::����::</option>
        	<option value="A" <%=chkIIF(tplType="A","selected","")%>>A Type</option>
        	<option value="B" <%=chkIIF(tplType="B","selected","")%>>B Type</option>
        	<option value="C" <%=chkIIF(tplType="C","selected","")%>>C Type</option>
        	<option value="D" <%=chkIIF(tplType="D","selected","")%>>D Type</option>
        	<option value="E" <%=chkIIF(tplType="E","selected","")%>>E Type</option>
        	<option value="F" <%=chkIIF(tplType="F","selected","")%>>F Type</option>
        	<option value="G" <%=chkIIF(tplType="G","selected","")%>>G Type</option>
        </select>
        &nbsp;<span id="tplTpDesc"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">����Ʈ����</td>
    <td width="230">
    	<%=chkIIF(siteDiv="P","PC��","�����")%>
    	<input type="hidden" name="site" value="<%=siteDiv%>" />
    </td>
    <td width="100" bgcolor="#DDDDFF">���ó</td>
    <td width="230">
        <select name="pageDiv" class="select" title="���ó">
        	<option value="">::����::</option>
        	<option value="10" <%=chkIIF(pageDiv="10","selected","")%>>����Ʈ ����</option>
        	<option value="20" <%=chkIIF(pageDiv="20","selected","")%>>�̺�Ʈ ����</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">���ø���</td>
    <td width="610" colspan="3">
        <input type="text" name="tplName" value="<%= tplName %>" maxlength="64" size="64" title="���ø���">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�ð�ǥ�� ����</td>
    <td>
        <select name="isTimeUse" class="select" title="�ð�ǥ�� ����">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isTimeUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isTimeUse="N","selected","")%>>������</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">������ ���</td>
    <td>
        <select name="isIconUse" class="select" title="������ ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isIconUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isIconUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���簳�� ����</td>
    <td>
        <select name="isSubNumUse" class="select" title="���簳�� ���� ����">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isSubNumUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isSubNumUse="N","selected","")%>>������</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">�ܺ��ڷ� ���</td>
    <td>
        <select name="isExtDataUse" class="select" title=�ܺ��ڷ� ��� ����">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isExtDataUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isExtDataUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">ž�̹��� ����</td>
    <td>
        <select name="isTopImgUse" class="select" title="ž�̹��� ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isTopImgUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isTopImgUse="N","selected","")%>>������</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">ž��ũ ����</td>
    <td>
        <select name="isTopLinkUse" class="select" title="ž�̹��� ��ũ ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isTopLinkUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isTopLinkUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�̹��� ���</td>
    <td>
        <select name="isImageUse" class="select" title="�̹��� ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isImageUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isImageUse="N","selected","")%>>������</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">�̹������� ���</td>
    <td>
        <select name="isImgDescUse" class="select" title="�̹������� ��뿩��(�̹����� �ִ� ���)">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isImgDescUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isImgDescUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�ؽ�Ʈ ���</td>
    <td>
        <select name="isTextUse" class="select" title="�ؽ�Ʈ ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isTextUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isTextUse="N","selected","")%>>������</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">��ũ ���</td>
    <td>
        <select name="isLinkUse" class="select" title="��ũ ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isLinkUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isLinkUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ǰ�ڵ� ���</td>
    <td>
        <select name="isItemUse" class="select" title="��ǰ�ڵ� ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isItemUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isItemUse="N","selected","")%>>������</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">������ ���</td>
    <td>
        <select name="isVideoUse" class="select" title="������URL ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isVideoUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isVideoUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���� ���</td>
    <td colspan="3">
        <select name="isBGColorUse" class="select" title="���� ��뿩��">
        	<option value="">::����::</option>
        	<option value="Y" <%=chkIIF(isBGColorUse="Y","selected","")%>>���</option>
        	<option value="N" <%=chkIIF(isBGColorUse="N","selected","")%>>������</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ø� ���ȳ�</td>
    <td colspan="3">
        <textarea name="tplinfoDesc" class="textarea" style="width:100%; height:60px;" title="���ø� ���ȳ�"><%=tplinfoDesc%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ļ���</td>
    <td colspan="3">
        <input type="text" name="tplSortNo" class="text" size="4" value="<%=tplSortNo%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveTemplate(this.form);"></td>
</tr>
</table>
</form>
<br>
<!-- // ��ϵ� ���ø� ��� --------->
<table width="690" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td colspan="5" align="right"><a href="?site=<%=siteDiv%>"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td width="60">code</td>
    <td width="80">���ø�����</td>
    <td width="100">���ø�����</td>
    <td>���ø���</td>
    <td width="80">����</td>
</tr>
<% for i=0 to oTemplate.FResultCount-1 %>
<% if (CStr(oTemplate.FItemList(i).FtplIdx)=tplIdx) then %>
<tr bgcolor="#9999CC">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><a href="?site=<%=siteDiv%>&tplIdx=<%= oTemplate.FItemList(i).FtplIdx %>&page=<%= page %>"><%= oTemplate.FItemList(i).FtplIdx %></a></td>
    <td align="center"><%= oTemplate.FItemList(i).FtplType %> Type</td>
    <td align="center"><%= oTemplate.FItemList(i).getPageDiv %></td>
    <td ><a href="?site=<%=siteDiv%>&tplIdx=<%= oTemplate.FItemList(i).FtplIdx %>&page=<%= page %>"><%= oTemplate.FItemList(i).FtplName %></a></td>
    <td align="center"><%= oTemplate.FItemList(i).FtplSortNo %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="5" align="center">
    <% if oTemplate.HasPreScroll then %>
		<a href="?site=<%=siteDiv%>&page=<%= oTemplate.StartScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oTemplate.StartScrollPage to oTemplate.FScrollCount + oTemplate.StartScrollPage - 1 %>
		<% if i>oTemplate.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?site=<%=siteDiv%>&page=<%= i %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oTemplate.HasNextScroll then %>
		<a href="?site=<%=siteDiv%>&page=<%= i %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</center>
<%	set oTemplate = Nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̹��� ����
' History : 2016.07.28 ������ ����
'			2016.08.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcImageMngCls.asp"-->
<%
dim folderidx, etcimgIdx, menupos, foldertitle, i, imagename, realPath, subfolder, lastuserid, lastupdate
folderidx = requestCheckVar(request("folderidx"),10)
etcimgIdx = requestCheckVar(request("etcimgIdx"),10)
menupos   = requestCheckVar(request("menupos"),10)

if (LEN(folderidx)<1) then
    response.write "������ �������� ����."
    dbget.close() : response.end
end if
foldertitle = getImgEtcFolderTitleByFolderIdx(folderidx)

dim oEtcone
SET oEtcone = new CEtcImageManage
	oEtcone.FRectetcimgIdx = etcimgIdx

	if (etcimgIdx<>"") then
		oEtcone.getEtcImage_one

		if oEtcone.FResultCount > 0 then
			imagename = oEtcone.FOneItem.Fimagename
			realPath = oEtcone.FOneItem.FrealPath
			subfolder = oEtcone.FOneItem.Fsubfolder
			lastuserid = oEtcone.FOneItem.Flastuserid
			lastupdate = oEtcone.FOneItem.Flastupdate
		end if
	end if

%>
<script type="text/javascript">

document.domain = '10x10.co.kr';

function jsUpload(){
	if(!document.frmImg.folderidx.value){
		alert("������ ������ �ּ���.");			
		return false;
	}

	frmImg.mode.value='regetc';
	frmImg.target='view';
	frmImg.submit();
}

function jsdel(){
	frmImg.mode.value='deletc';
	frmImg.target='view';
	frmImg.submit();
}

function openerreloads(){
	opener.location.reload();
	self.close();
}

</script>

<iframe id="view" name="view" height=0 width=0 frameborder="0" scrolling="no"></iframe>

<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/etcimage/etcimage_upload.asp" enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<input type="hidden" name="mode" value="">

<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=150>�ϷĹ�ȣ</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="etcimgIdx" value="<%= etcimgIdx %>">
		<%= etcimgIdx %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="folderidx" value="<%= folderidx %>">
		<%= foldertitle %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹�������</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="sfImg1">

		<% if imagename <> "" then %>
			<Br>
			<img src="<%= webImgUrl %>\<%= realPath %>\<%= subfolder %>\<%= imagename %>" width=100 height=100>
			<input type="checkbox" name="delimg" value="ON">�̹�������
		<% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<% if lastuserid<>"" then %>
			<%= lastupdate %><Br>(<%= lastuserid %>)
		<% end if %>
	</td>
</tr>
<tr align="center" >
	<td bgcolor="#FFFFFF" colspan=2>
		<input type="button" onclick="jsUpload();" value="����" class="button">
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" onclick="jsdel();" value="�ۻ���" class="button">
	</td>
</tr>
</table>

</form>

<%
set oEtcone = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp"-->
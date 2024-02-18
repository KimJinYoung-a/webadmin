<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이미지 관리
' History : 2016.07.28 서동석 생성
'			2016.08.12 한용민 수정
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
    response.write "구분이 지정되지 않음."
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
		alert("구분을 지정해 주세요.");			
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
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=150>일렬번호</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="etcimgIdx" value="<%= etcimgIdx %>">
		<%= etcimgIdx %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">구분</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="folderidx" value="<%= folderidx %>">
		<%= foldertitle %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이미지선택</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="sfImg1">

		<% if imagename <> "" then %>
			<Br>
			<img src="<%= webImgUrl %>\<%= realPath %>\<%= subfolder %>\<%= imagename %>" width=100 height=100>
			<input type="checkbox" name="delimg" value="ON">이미지삭제
		<% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">최종수정</td>
	<td bgcolor="#FFFFFF">
		<% if lastuserid<>"" then %>
			<%= lastupdate %><Br>(<%= lastuserid %>)
		<% end if %>
	</td>
</tr>
<tr align="center" >
	<td bgcolor="#FFFFFF" colspan=2>
		<input type="button" onclick="jsUpload();" value="저장" class="button">
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" onclick="jsdel();" value="글삭제" class="button">
	</td>
</tr>
</table>

</form>

<%
set oEtcone = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp"-->
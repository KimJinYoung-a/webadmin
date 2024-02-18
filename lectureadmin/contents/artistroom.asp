<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/artistsroomcls.asp"-->
<%
dim lp
dim maxAddImg
maxAddImg = 30

dim lecuserid
lecuserid = Session("ssBctId")

dim oartistroom
set oartistroom = new CArtistsRoom
oartistroom.FRectUserid =  lecuserid

if lecuserid<>"" then
	oartistroom.GetOneArtistRoom
end if

dim oartistroommajorlec
set oartistroommajorlec = new CArtistsRoom
oartistroommajorlec.FRectUserid =  lecuserid

if lecuserid<>"" then
	oartistroommajorlec.GetMajorLec
end if

dim oartistroomimage
set oartistroomimage = new CArtistsRoom
oartistroomimage.FRectUserid =  lecuserid
oartistroomimage.GetImageList

dim i
%>

<script language='javascript'>
<!--
function CheckDel(comp){
	var frm = comp.form;
	if (comp.checked){
		frm.filedelete[comp.value].value = "Y";
	}else{
		frm.filedelete[comp.value].value = "";
	}
}

// �ɼ��� �߰��Ѵ�
function InsertOption(comp, ft, fv)
{
	comp.options[comp.options.length] = new Option(ft, fv);
}

// ���õ� �ɼ� ����
function delItemOptionAdd(comp)
{
	var sidx = comp.options.selectedIndex;

	if(sidx<0)
		alert("������ �ɼ��� �������ֽʿ�.");
	else
	{
		comp.options[sidx] = null;
	}
}


function SelectLecture(lecid,lecname,lecuserid){
	if (lecuserid!='<%= lecuserid %>'){
		alert('���� ���縸 �߰� �����մϴ�.');
		return;
	}

	InsertOption(frm_modi.majorlec,lecid + ' - ' + lecname,lecid);
}

function SelectLec(compname){
	var lecturer = eval(compname).form.lecuserid.value;
	var popwin = window.open('/lectureadmin/lib/popselectlec.asp?lecturer=' + lecturer + '&parentcomp=' + compname,'popLecSelct','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}


// �̹��� ���� �˻�
function checkImageSuffix (fileInput)
{
   var suffixPattern = /(gif|jpg|jpeg|png)$/i;
   if (!suffixPattern.test(fileInput.value)) {
     alert('GIF,JEPG,PNG ���ϸ� �����մϴ�.');
     fileInput.focus();
     fileInput.select();
     return false;
   }
   return true;
}

// �Է��� �˻�
function chkSubmit(frm)
{
	if(frm.lecuserid.value.length<1)	{
		alert("���� ���̵� �Է����ֽʽÿ�.");
		frm.lecuserid.focus();
		return;
	}

	if(frm.summarycontents.value.length<1)	{
		alert("���°��並 �Է����ֽʽÿ�.");
		frm.summarycontents.focus();
		return;
	}

	if ((frm.summaryimage.value.length>0)&&(!checkImageSuffix(frm.summaryimage))){
		return ;
	}

<% if oartistroom.FResultCount>0 then %>
	for(var i=0;i<frm.imgFile.length;i++){
		if ((frm.imgFile[i].value.length>0)&&(!checkImageSuffix(frm.imgFile[i]))){
			return ;
		}
	}

	for(var i=0;i<frm.imagetype.length;i++){
		if ((frm.imgFile[i].value.length>0)&&(frm.imagetype[i].value.length<1)){
			alert('�̹��� Ÿ�� ������ �����ϼ���.');
			frm.imagetype[i].focus();
			return ;
		}
	}

	frm.majorlecarr.value = "";

	for(var i=0; i<frm.majorlec.options.length; i++) {
		frm.majorlecarr.value += (frm.majorlec.options[i].value + ",")
	 }
<% end if %>

	// �� ����
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

// �߰� �̹����� ����
function addImgControl(lid, md, mx)
{
	var frm = document.all;
	var btext = "";

	if(md=="add")
	{
		// ���̾� ���̱� �� ���� Ȱ��ȭ
		frm["addimg"+(lid+1)].style.display="";
		frm["btn"+(lid)].innerHTML = "";
		if(lid<mx)
			frm.filedelete[lid].value = "";

		lid++;

		// ��ưó��
		if(lid>1)
			btext += "<img src='/images/icon_minus.gif' alt='�̹��� ����' onClick=addImgControl(" + lid + ",'del'," + mx + ") style='cursor:pointer'> ";
		if(lid<<%=maxAddImg%>)
			btext += "<img src='/images/icon_plus.gif' alt='�̹��� �߰�' onClick=addImgControl(" + lid + ",'add'," + mx + ") style='cursor:pointer'>";

		frm["btn"+lid].innerHTML = btext;
	}
	else if(md=="del")
	{
		// ���̾� ����� �� ���� ����
//		frm.imgFile[lid-1].select();
//		document.execCommand('Delete');

//		if(lid>mx)
//			frm.imgCont[lid-1].value = "";
//		else
//			frm.filedelete[lid-1].value = "Y";

//		frm["addimg"+lid].style.display="none";
//		frm["btn"+lid].innerHTML = "";

//		lid--;

		// ��ưó��
//		if(lid>1)
//			btext += "<img src='/images/icon_minus.gif' alt='�̹��� ����' onClick=addImgControl(" + lid + ",'del'," + mx + ") style='cursor:pointer'> ";
//		if(lid<<%=maxAddImg%>)
//			btext += "<img src='/images/icon_plus.gif' alt='�̹��� �߰�' onClick=addImgControl(" + lid + ",'add'," + mx + ") style='cursor:pointer'>";

//		frm["btn"+lid].innerHTML = btext;
	}
}
//-->
</script>
<% if oartistroom.FResultCount<1 then %>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" action="http://image.thefingers.co.kr/linkweb/doArtistRoom.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="write">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�۰��ǹ� �ű� ���</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�۰� ID</td>
	<td bgcolor="#FFFFFF">
		<b><%=lecuserid%></b>
		<input type="hidden" name="lecuserid" value="<%=lecuserid%>">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���°���</td>
	<td bgcolor="#FFFFFF"><textarea name="summarycontents" rows="4" cols="80"></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred"> </font>���°��� �̹���</td>
	<td bgcolor="#FFFFFF"><input type="file" name="summaryimage" size="50"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred"> </font>�ϴ� ����</td>
	<td bgcolor="#FFFFFF"><textarea name="contents1" rows="4" cols="80"></textarea></td>
</tr>
<tr>
	<td colspan="2" align="center" bgcolor="#FFFFFF"><input type="button" value="�ű� ���" onclick="chkSubmit(frm_write)"></td>
</tr>
</form>
</table>
<% else %>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_modi" method="POST" action="http://image.thefingers.co.kr/linkweb/doArtistRoom.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="modi">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�۰��ǹ� ����</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�۰� ID</td>
	<td bgcolor="#FFFFFF">
		<b><%=lecuserid%></b>
		<input type="hidden" name="lecuserid" value="<%=lecuserid%>">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���°���</td>
	<td bgcolor="#FFFFFF"><textarea name="summarycontents" rows="6" cols="80"><%= oartistroom.FOneItem.Fsummarycontents %></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���°��� �̹���</td>
	<td bgcolor="#FFFFFF">
	<input type="file" name="summaryimage" size="50">
	<br>
	<img src="<%= oartistroom.FOneItem.Fsummaryimage %>">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��Ÿ����</td>
	<td bgcolor="#FFFFFF"><textarea name="contents1" rows="6" cols="80"><%= oartistroom.FOneItem.Fcontents1 %></textarea></td>
</tr>
<tr>
	<td colspan="2" height="2" bgcolor="#FFFFFF"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��ǥ����</td>
	<td bgcolor="#FFFFFF">
	<input type="hidden" name="majorlecarr" value="">
		<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<select name="majorlec" size="4" style='width:300;'>
				<% for i=0 to oartistroommajorlec.FResultCount-1 %>
				<option value="<%= oartistroommajorlec.FItemList(i).Flec_idx %>"><%= oartistroommajorlec.FItemList(i).Flec_idx %> - <%= oartistroommajorlec.FItemList(i).Flec_title %>
				<% next %>
				</select>
			</td>
			<td>
				<input type="button" value="��ǥ���� ����" onclick="SelectLec('frm_modi.majorlec');">
				<br>
				<input type="button" value="��ǥ���� ����" onclick="delItemOptionAdd(frm_modi.majorlec);">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" height="20" bgcolor="#FFFFFF">

	</td>
</tr>
<%
	'// ���� �̹��� �� �ۼ�
	for lp=1 to maxAddImg
%>
<tr id="addimg<%=lp%>" <% if ((lp>1) and (lp>oartistroomimage.FResultCount)) then Response.Write "style='display:none;'"%>>
	<td align="center" width="120" bgcolor="#DDDDFF"><% if lp=1 then %><% end if %>�̹��� #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="100" align="center">
			<% if oartistroomimage.FResultCount>=lp then %>
			<input type="hidden" name="imagetype"  value="<%= oartistroomimage.FItemList(lp-1).Fimagetype %>">
			<%= oartistroomimage.FItemList(lp-1).GetimageName %>
			<% else %>
			<select name="imagetype">
			<option value=""> ����
			<option value="10"> �����̹���
			<option value="20"> �����̹���
			<option value="50"> ��ǰ�Ұ�
			</select>

			<% end if %>
			</td>
			<td >
				<input type="file" name="imgFile" size="50">
				<% if oartistroomimage.FResultCount>=lp then %>
				<br><img src="<%= oartistroomimage.FItemList(lp-1).Fimagevalue %>" width="100" height="100">
				<input type="checkbox" name="checkdelete" value="<%= lp-1 %>" onclick="CheckDel(this);"><font color="red">����<%= oartistroomimage.FItemList(lp-1).Fimgidx %></font>
				<input type="hidden" name="orgImgId" value="<%= oartistroomimage.FItemList(lp-1).Fimgidx %>">
				<input type="hidden" name="orgFile" value="<%= oartistroomimage.FItemList(lp-1).FOrgimagevalue %>">
				<input type="hidden" name="orgIconFile" value="<%= oartistroomimage.FItemList(lp-1).FOrgimageicon %>">
				<input type="hidden" name="filedelete" value="">
				<textarea name="imgcontents" cols="80" rows="2"><%= oartistroomimage.FItemList(lp-1).Fimgcontents %></textarea>
				<% else %>
				<input type="hidden" name="orgImgId" value="">
				<input type="hidden" name="orgFile" value="">
				<input type="hidden" name="orgIconFile" value="">
				<input type="hidden" name="filedelete" value="">
				<textarea name="imgcontents" cols="80" rows="2"></textarea>
				<% end if %>
			</td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>">
				<% if oartistroomimage.FResultCount<1 then %>
				<img src="/images/icon_plus.gif" alt="�̹��� �߰�" onClick="addImgControl(<%=lp%>,'add')" style="cursor:pointer">
				<% elseif (lp=oartistroomimage.FResultCount) then %>
				<img src="/images/icon_plus.gif" alt="�̹��� �߰�" onClick="addImgControl(<%=lp%>,'add')" style="cursor:pointer">
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="2" align="center" bgcolor="#FFFFFF"><input type="button" value="�� ��" onclick="chkSubmit(frm_modi)"></td>
</tr>
</form>
</table>
<% end if %>
<%
set oartistroomimage = Nothing
set oartistroommajorlec = Nothing
set oartistroom = Nothing
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
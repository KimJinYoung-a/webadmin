<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ��������
' Hieditor : 2009.10.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i 
dim photoid, photoword, mainimage, regdate, isusing , detailimage
dim ingimage , wordimage ,wordovimage
	photoid = requestcheckvar(request("photoid"),8)

'// �̺�Ʈ ����Ʈ
set ocontents = new cphoto_list
	ocontents.frectphotoid = photoid
	
	'//�����ϰ�쿡�� ����
	if photoid <> "" then
	ocontents.fphoto_oneitem()
	end if
	
	if ocontents.ftotalcount > 0 then
		photoid = ocontents.FOneItem.fphotoid
		photoword = ocontents.FOneItem.fphotoword
		mainimage = ocontents.FOneItem.fmainimage
		regdate = ocontents.FOneItem.fregdate
		isusing = ocontents.FOneItem.fisusing		
		detailimage = ocontents.FOneItem.fdetailimage
		wordimage = ocontents.FOneItem.fwordimage		
		ingimage = ocontents.FOneItem.fingimage
		wordovimage = ocontents.FOneItem.fwordovimage		
	end if
%>

<script language="javascript">

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

//����
function reg(){
	if (frm.photoword.value==''){
	alert('Ű���带 �Է����ּ���');
	frm.photoword.focus();
	return;
	}
	if (frm.isusing.value==''){
	alert('��뿩�θ� �������ּ���');
	return;
	}
	
	frm.action='/admin/momo/photo/photo_process.asp';
	frm.mode.value='edit';	
	frm.submit();
}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>Ű����ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= photoid %><input type="hidden" name="photoid" value="<%= photoid %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>Ű����</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="photoword" value="<%= photoword %>" size=20>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�������̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('ingimgdiv','ingimg','ing','2000','800','true');"/>		
		<input type="hidden" name="ingimg" value="<%= ingimage %>">
		<div align="right" id="ingimgdiv"><% IF ingimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/ing/<%= ingimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('mainimgdiv','mainimg','main','2000','800','true');"/>		
		<input type="hidden" name="mainimg" value="<%= mainimage %>">
		<div align="right" id="mainimgdiv"><% IF mainimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/main/<%= mainimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����Ʈ�̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('wordimgdiv','wordimg','word','2000','250','true');"/>		
		<input type="hidden" name="wordimg" value="<%= wordimage %>">
		<div align="right" id="wordimgdiv"><% IF wordimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/word/<%= wordimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����Ʈ�����̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('wordovimgdiv','wordovimg','wordov','2000','250','true');"/>		
		<input type="hidden" name="wordovimg" value="<%= wordovimage %>">
		<div align="right" id="wordovimgdiv"><% IF wordovimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/wordov/<%= wordovimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('detailimgdiv','detailimg','detail','2000','800','true');"/>		
		<input type="hidden" name="detailimg" value="<%= detailimage %>">
		<div align="right" id="detailimgdiv"><% IF detailimage<>"" THEN %><img src="<%=webImgUrl%>/momo/photo/detail/<%= detailimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��뿩��</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>��뿩��</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="����"></td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="2009">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<%
	set ocontents = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
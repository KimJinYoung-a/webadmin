<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� �츮��������
' Hieditor : 2009.11.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i , mainimage
dim vote_num, title, question, startdate, enddate, isusing
	vote_num = requestcheckvar(request("vote_num"),8)

'//��
set ocontents = new cvote_list
	ocontents.frectvote_num = vote_num
	
	'//�����ϰ�쿡�� ����
	if vote_num <> "" then
	ocontents.fvote_oneitem()
	end if
		
	if ocontents.ftotalcount > 0 then
		vote_num = ocontents.FOneItem.fvote_num
		title = ocontents.FOneItem.ftitle
		question = ocontents.FOneItem.fquestion
		startdate = ocontents.FOneItem.fstartdate
		enddate = ocontents.FOneItem.fenddate		
		isusing = ocontents.FOneItem.fisusing
		mainimage = ocontents.FOneItem.fmainimage
	end if
%>

<script language="javascript">

	//����
	function reg(){
		if (frm.title.value==''){
		alert('������ �Է����ּ���');
		frm.title.focus();
		return;
		}
		if (frm.question.value==''){
		alert('������ �Է����ּ���');
		frm.question.focus();
		return;
		}
		if (frm.startdate.value==''){
		alert('�������� �Է����ּ���');
		frm.startdate.focus();
		return;
		}		
		if (frm.enddate.value==''){
		alert('�������� �Է����ּ���');
		frm.enddate.focus();
		return;
		}						
		if (frm.isusing.value==''){
		alert('��뿩�θ� �������ּ���');
		return;
		}
		
		frm.action='/admin/momo/vote/vote_process.asp';
		frm.mode.value='add';
		frm.submit();
	}

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
	
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>voteid</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= vote_num %><input type="hidden" name="vote_num" value="<%= vote_num %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="title" value="<%= title %>" size=64 maxlength=35>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="question" style="width:450px; height:100px;"><%=question%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>�Ⱓ</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="startdate" size=10 value="<%= startdate %>">			
		<a href="javascript:calendarOpen3(frm.startdate,'������',frm.startdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
		<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>">
		<a href="javascript:calendarOpen3(frm.enddate,'��������',frm.enddate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����̹���<br>326x95</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('mainimgdiv','mainimg','main','2000','252','true');"/>		
		<input type="hidden" name="mainimg" value="<%= mainimage %>">
		<div align="right" id="mainimgdiv"><% IF mainimage<>"" THEN %><img src="<%=webImgUrl%>/momo/vote/main/<%= mainimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
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
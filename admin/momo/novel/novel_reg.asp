<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���ټҼ�
' Hieditor : 2009.11.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i , wordimage , winner
dim novelid,startdate,enddate,regdate,prolog,title,genre,isusing
	novelid = requestcheckvar(request("novelid"),8)

'//��
set ocontents = new cnovel_list
	ocontents.frectnovelid = novelid
	
	'//�����ϰ�쿡�� ����
	if novelid <> "" then
	ocontents.fnovel_oneitem()
	end if
		
	if ocontents.ftotalcount > 0 then
		novelid = ocontents.FOneItem.fnovelid
		startdate = ocontents.FOneItem.fstartdate
		enddate = ocontents.FOneItem.fenddate
		regdate = ocontents.FOneItem.fregdate
		prolog = ocontents.FOneItem.fprolog
		title = ocontents.FOneItem.ftitle
		genre = ocontents.FOneItem.fgenre
		isusing = ocontents.FOneItem.fisusing
		wordimage = ocontents.FOneItem.fwordimage
		winner = ocontents.FOneItem.fwinner
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
		if (frm.title.value==''){
		alert('������ �Է����ּ���');
		frm.title.focus();
		return;
		}
		if (frm.genre.value==''){
		alert('�帣�� �Է����ּ���');
		frm.genre.focus();
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
		if (frm.prolog.value==''){
		alert('���ѷα׸� �Է����ּ���');
		frm.prolog.focus();
		return;
		}						
		if (frm.isusing.value==''){
		alert('��뿩�θ� �������ּ���');
		return;
		}
		
		frm.action='/admin/momo/novel/novel_process.asp';
		frm.mode.value='edit';
		frm.submit();
	}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>���ټҼ�ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= novelid %><input type="hidden" name="novelid" value="<%= novelid %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>���ǹ�����</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= regdate %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="title" value="<%= title %>" size=64 maxlength=35>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�帣</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="genre" value="<%= genre %>" size=20 maxlength=10>		
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
	<td>���ѷα�</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="prolog" style="width:450px; height:100px;"><%=prolog%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�Ҽ������̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('wordimgdiv','wordimg','word','2000','800','true');"/>		
		<input type="hidden" name="wordimg" value="<%= wordimage %>">
		<div align="right" id="wordimgdiv"><% IF wordimage<>"" THEN %><img src="<%=webImgUrl%>/momo/novel/word/<%= wordimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��÷��ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="winner" value="<%=winner%>"> �س����� �ʿ��� ��츸 �Է��ϼ���	
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
	<td colspan=2><input type="button" onclick="reg();" value="����" class="button"></td>
</tr>
</form>
</table>
<%
	set ocontents = nothing
%>
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
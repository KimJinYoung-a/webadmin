<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� �Բ��ؿ�
' Hieditor : 2010.11.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim idx ,withid ,withgubun ,userid ,comment ,regdate ,isusing ,withimage_small ,withimage_large ,orderno
dim ocontents , i
	withid = requestcheckvar(request("withid"),10)
	idx = requestcheckvar(request("idx"),10)

if withid = "" then response.write "<script>alert('ID���� �����ϴ�'); self.close();</script>"

'//��
set ocontents = new cwith_list
	ocontents.frectwithid = withid
	ocontents.frectidx = idx
	
	'//�����ϰ�쿡�� ����
	if idx <> "" and withid <> "" then
	ocontents.fwith_snsoneitem()
	end if
		
	if ocontents.ftotalcount > 0 then		
		withid = ocontents.FOneItem.fwithid
		withgubun = ocontents.FOneItem.fwithgubun
		userid = ocontents.FOneItem.fuserid
		comment = ocontents.FOneItem.fcomment
		regdate = ocontents.FOneItem.fregdate
		isusing = ocontents.FOneItem.fisusing
		withimage_small = ocontents.FOneItem.fwithimage_small
		withimage_large = ocontents.FOneItem.fwithimage_large
		orderno = ocontents.FOneItem.forderno
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
		if (frm.withgubun.value==''){
		alert('������ �Է����ּ���');
		frm.withgubun.focus();
		return;
		}
		if (frm.userid.value==''){
		alert('���̵� �Է����ּ���');
		frm.userid.focus();
		return;
		}
		if (frm.comment.value==''){
		alert('�ڸ�Ʈ�� �Է����ּ���');
		frm.comment.focus();
		return;
		}		
		if (frm.orderno.value==''){
		alert('�켱������ �Է����ּ���');
		frm.orderno.focus();
		return;
		}						
		if (frm.isusing.value==''){
		alert('��뿩�θ� �������ּ���');
		return;
		}
		
		frm.action='with_process.asp';
		frm.mode.value='editsns';
		frm.submit();
	}
	
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�Բ��ؿ��ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= withid %><input type="hidden" name="withid" value="<%=withid%>">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>SNS��ȣ</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= idx %><input type="hidden" name="idx" value="<%= idx %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="withgubun">
			<option value="" <% if withgubun = "" then response.write " selected"%>>����</option>
			<option value="0" <% if withgubun = "0" then response.write " selected"%>>Ʈ����</option>
			<option value="1" <% if withgubun = "1" then response.write " selected"%>>��������</option>
			<option value="2" <% if withgubun = "2" then response.write " selected"%>>����</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>�����</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="regdate" size=10 value="<%= regdate %>">			
		<a href="javascript:calendarOpen3(frm.regdate,'�����',frm.regdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="userid" value="<%= userid %>" size=20 maxlength=10>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>�ڸ�Ʈ</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="comment" rows="14" cols="80"><%= comment %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>���ļ���</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="orderno">			
			<% for i = 0 to 50 %>			
			<option value="<%=i%>" <% if orderno = i then response.write " selected"%>><%=i%></option>
			<% next %>
		</select>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�̹���</td>
	<td bgcolor="#FFFFFF" align="left">
		(500*333 �̹��� ����Ͻø� 194*133 �̹����� �ڵ� �����˴ϴ�)<br>
		<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('withimage_largediv','withimage_large','wi_large','2000','500','true');"/>		
		<input type="hidden" name="withimage_large" value="<%= withimage_large %>">
		<input type="hidden" name="withimage_small" value="<%= withimage_small %>">
		<div align="right" id="withimage_largediv">
			<% IF withimage_large<>"" THEN %><img src="<%=webImgUrl%>/momo/with/wi_large/<%= withimage_large %>"><% End IF %>
			<% if withimage_small <> "" then %><br><img src="<%=webImgUrl%>/momo/with/wi_large/<%= withimage_small %>"><% end if %>
		</div>
		
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
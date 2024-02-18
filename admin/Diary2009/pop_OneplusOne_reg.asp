<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%
'####################################################
' Description : ���̾ ���丮 1+1,1:1 ����
' History : 2015.09.24 ���¿� ����(olorcodeleft, colorcoderight �߰�)
' 			2018-08-20 ����ȭ ���� ()
'####################################################
%>
<%
Dim idx, mode, itemid, image1, image2, imageEnd, startdate, isusing, ii, endLink, explain, m_image1, m_image2
dim m_image1link, m_image2link, GiftSu, plustype, topimage1, topimage2, topimage3
dim colorcodeleft, colorcoderight , swipertext , eventid
	idx = request("idx")
	Mode = request("mode")

GiftSu=0
IF Mode = "" THEN Mode = "add"
IF Mode = "edit" Then
	
dim oDiary
set oDiary = new DiaryCls
	oDiary.FIdx = idx
	oDiary.getDiaryOneplusOne_View

	itemid = oDiary.FItem.FItemid
	image1 = oDiary.FItem.FImage1
	image2 = oDiary.FItem.FImage2
	imageEnd = oDiary.FItem.FImageEnd
	endLink = oDiary.FItem.FendLink
	explain = oDiary.FItem.Fexplain
	startdate = oDiary.FItem.Fstartdate
	isusing = oDiary.FItem.FIsusing
	m_image1 = oDiary.FItem.FMImage1
	m_image2 = oDiary.FItem.FMImage2
	m_image1link = oDiary.FItem.FMImage1Link
	m_image2link = oDiary.FItem.FMImage2Link
	ii = idx
	plustype = oDiary.FItem.fplustype
	topimage1 = oDiary.FItem.ftopimage1
	topimage2 = oDiary.FItem.ftopimage2
	topimage3 = oDiary.FItem.ftopimage3
	colorcodeleft = oDiary.FItem.Fcolorcodeleft
	colorcoderight = oDiary.FItem.Fcolorcoderight
	swipertext = oDiary.FItem.Fswipertext
	eventid	= oDiary.FItem.Feventid
	

	GiftSu = oDiary.getGiftDiaryExists(itemid) '����ǰ ��
End If

if plustype="" then plustype="1"
%>

<script language="javascript">

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function form_check(){
	var frm = document.frmreg;
	if(frm.itemid.value == ""){
		alert("��ǰ�ڵ带 �Է��ϼ���");
		frm.itemid.focus();
		return false;
	}
	if(frm.startdate.value == ""){
		alert("�������� �Է��ϼ���");
		frm.startdate.focus();
		return false;
	}
	if(!frm.isusing[0].checked && !frm.isusing[1].checked){
		alert("��������� �����ϼ���");
		return false;
	}

	frm.submit();
}

function onlyNumberInput(){
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}

//�̹��� ��â Ȯ�뺸��
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmreg" method="post" action="<%=uploadUrl%>/linkweb/diary/image_proc2.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<%= Mode %>">
<input type="hidden" name="idx" value="<%= ii %>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" height="70">
			<td colspan="2" align="center"><b><font size="5" color="red">���������� �̹���, ��ũ�� ��� �Էµ����� ������ ������ �� ����.</font></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ��ǰ�ڵ�</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="itemid" value="<%=itemid%>" _onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> �̺�Ʈ�ڵ�</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="eventid" value="<%=eventid%>" style="IME-MODE: disabled" />
			</td>
		</tr>
		<!-- ����ǰ ���� ���� -->
		<% If  Mode = "edit" Then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����ǰ ���� ����</td>
			<td bgcolor="#FFFFFF" align="left">
				<% If GiftSu > 0 Then %>
				<input type="text" value="<%=GiftSu%>" style="IME-MODE: disabled" readOnly/>�� ����
				<% Else %>
				����ǰ ����
				<% End If %>
			</td>
		</tr>
		<% End If %>
		<!-- ����ǰ ���� ���� -->
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ������ </td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" name="startdate" value="<%=startdate%>" size="10" maxlength="10" readonly onClick="jsPopCal('startdate');"  style="cursor:hand;" class="input_b">
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����ǥ��</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="radio" name="plustype" value="1" <% If plustype = "1" Then response.write "checked"%>>1+1
				<input type="radio" name="plustype" value="2" <% If plustype = "2" Then response.write "checked"%>>1:1
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> ��뿩��</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="radio" name="isusing" value="Y" <% If isusing = "Y" Then response.write "checked"%>>���
				<input type="radio" name="isusing" value="N" <% If isusing = "N" Then response.write "checked"%>>������
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����ž�̹���</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="file1" value="<%= image1 %>" size="32" maxlength="32" class="file">
			  <% If image1 <> "" Then %>
			  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= image1 %>" onClick="jsImgView('<%=uploadUrl%>/diary/oneplusone/<%= image1 %>')" border="0" height="100">
			  <!--<br><%=uploadUrl%>/diary/oneplusone/<%= image1 %>//-->
			  <% End If %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">�������� �ϴ� �ؽ�Ʈ</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="swipertext" value="<%=swipertext%>"/><span style="color:red">��ex) ���� Ź�� �÷��ʡ�</span>
			</td>
		</tr>
		<!-- ���ι�� ��,�� �÷��� �߰� ���¿� 150921 -->
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> 1+1,1:1,���� ������ ����<br>�⺻left<br>center �Է½� ��� ����</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="colorcodeleft" value="<%=colorcodeleft%>"  style="IME-MODE: disabled" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ������</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="text" class="text" name="colorcoderight" value="<%=colorcoderight%>"  style="IME-MODE: disabled" />
			</td>
		</tr>

		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����ž�̹���1</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="topimage1" value="<%= topimage1 %>" size="32" maxlength="32" class="file">
			  <% If topimage1 <> "" Then %>
				  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= topimage1 %>" border="0" height="100">
			  <% End If %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����ž�̹���2</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="topimage2" value="<%= topimage2 %>" size="32" maxlength="32" class="file">
			  <% If topimage2 <> "" Then %>
				  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= topimage2 %>" border="0" height="100">
			  <% End If %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����ž�̹���3</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="topimage3" value="<%= topimage3 %>" size="32" maxlength="32" class="file">
			  <% If topimage3 <> "" Then %>
				  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= topimage3 %>" border="0" height="100">
			  <% End If %>
			</td>
		</tr>		
	<!--	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">1+1 �̹���(��)-ǰ�� ��</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="file2" value="<%= image1 %>" size="32" maxlength="32" class="file">
			  <% If image2 <> "" Then %>
			  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= image2 %>" border="0" height="100">
			  <br><%=uploadUrl%>/diary/oneplusone/<%= image2 %>
			  <% End If %>
			</td>
		</tr>	-->
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">�����</td>
			<td bgcolor="#FFFFFF" align="left">
				<textarea name="explain" cols="60" rows="5"><%=explain%></textarea>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">1+1 �̹���-ǰ�� ��<br> (1140x420)</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="imageEnd" value="<%= imageEnd %>" size="32" maxlength="32" class="file">
			  <% If imageEnd <> "" Then %>
			  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= imageEnd %>" border="0" height="100">
			  <!--<br><%=uploadUrl%>/diary/oneplusone/<%= imageEnd %>//-->
			  <% End If %>
			  <br>�̹����� �� ��ũ�� : <input type="text" name="endLink" value="<%=endLink%>" size="60">
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����Ͽ� �̹��� 1</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="m_image1" value="<%= m_image1 %>" size="32" maxlength="32" class="file">
			  <% If m_image1 <> "" Then %>
			  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= m_image1 %>" onClick="jsImgView('<%=uploadUrl%>/diary/oneplusone/<%= m_image1 %>')" border="0" height="100">
			  <% End If %>
			  <br>�̹����� �� ��ũ�� : <input type="text" name="m_image1link" value="<%=m_image1link%>" size="60">
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150">����Ͽ� �̹��� 2</td>
			<td bgcolor="#FFFFFF" align="left"><input type="file" name="m_image2" value="<%= m_image2 %>" size="32" maxlength="32" class="file">
			  <% If m_image2 <> "" Then %>
			  <br><img src="<%=uploadUrl%>/diary/oneplusone/<%= m_image2 %>" border="0" height="100">
			  <% End If %>
			  <br>�̹����� �� ��ũ�� : <input type="text" name="m_image2link" value="<%=m_image2link%>" size="60">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center" bgcolor="<%=adminColor("green")%>"><br>
		<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="form_check();" style="cursor:pointer">
		<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="frmreg.reset();" style="cursor:pointer">
	</td>
</tr>
</form>
</table>

<%set oDiary = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 ���
' History : 2009.04.17 �̻� ����
'			2016.03.25 �ѿ�� ����(���Ǻо� ��� DBȭ ��Ŵ)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim i, j, reffrom, orderinfo
	reffrom = request("reffrom")

'���� 1:1�����亯
dim boardqna
set boardqna = New CMyQNA

boardqna.read(request("id"))

if boardqna.results(0).userid <> "" then
	set orderinfo = New CMyQNAOrderInfo
	'orderinfo.UserOrderInfo (boardqna.results(0).userid)
	'orderinfo.UserMinusOrderInfo (boardqna.results(0).userid)
end if

if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then
	dim myqnalist
	set myqnalist = New CMyQNA
	if boardqna.results(0).userid <> "" then
	    myqnalist.SearchUserID = boardqna.results(0).userid
	end if
	if boardqna.results(0).orderserial <> "" then
	    myqnalist.SearchOrderSerial = boardqna.results(0).orderserial
	end if

    myqnalist.PageSize = 100
    myqnalist.CurrPage = 1

	'// ���� ��� ���
    myqnalist.list
end if

dim useridForShow : useridForShow = "��"
if boardqna.results(0).userid <> "" then
	useridForShow = boardqna.results(0).userid
end if

dim extItemURL

%>

<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript">

function SubmitForm(){
	var replyuser = "<%= session("ssBctID") %>";
	var userid = "<%= boardqna.results(0).userid %>";

	// �� ������ ����
	/* ����, skyer9, 2016-04-29
	if (replyuser != "bseo") {
		if (userid == "majorblue") {
			if (confirm("�ȳ�!!\n\n�̼��� ������� �����ϱ�� �Ǿ� �ִ� �����Դϴ�.\n\n��� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		}
	}
	*/

	if (document.frm.replytitle.value == "") {
		alert("������ �Է��ϼ���.");
		return;
	}
	if (document.frm.replycontents.value == "") {
		alert("������ �Է��ϼ���.");
		return;
	}

	if (confirm("�Է��� ��Ȯ�մϱ�?") == true) {
		var btnSubmit = document.getElementById('btnSubmit');
		btnSubmit.disabled = true;

		document.frm.submit();
	}
}

function updateqadiv(){
	if (confirm("�����Ͻðڽ��ϱ�?")){
		document.updateform.submit();
	}
}

function updateitemid() {
	var itemid = document.frm.itemid.value;

	if (itemid == "") {
		alert("��ǰ�ڵ带 �Է��ϼ���.");
		return;
	}

	if (itemid*0 != 0) {
		alert("�߸��� ��ǰ�ڵ��Դϴ�.");
		return;
	}

	if (confirm("��ǰ�ڵ带 �����Ͻðڽ��ϱ�?")) {
		document.updateform.mode.value = "CGHITEMID";
		document.updateform.itemid.value = itemid;

		document.updateform.submit();
	}
}

function updateorderserial() {
	var orderserial = document.frm.orderserial.value;

	if (orderserial == "") {
		alert("�ֹ���ȣ�� �Է��ϼ���.");
		return;
	}

	if (confirm("�ֹ���ȣ�� �����Ͻðڽ��ϱ�?")) {
		document.updateform.mode.value = "CGHORDSERIAL";
		document.updateform.orderserial.value = orderserial;

		document.updateform.submit();
	}
}

function delqadiv(){
	if (confirm("�����Ͻðڽ��ϱ�?")){
		document.delform.submit();
	}
}

function popMyQNAUTF8(idx) {
    var window_width = 600;
    var window_height = 400;
	var popwin = window.open("popMyQNAUTF8.asp?idx=" + idx,"popMyQNAUTF8","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function popMyQNA_IMAGE(idx, imgidx) {
    var window_width = 1200;
    var window_height = 800;
	var popwin = window.open("popMyQNA_IMAGE.asp?idx=" + idx + "&imgidx=" + imgidx,"popMyQNA_IMAGE","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function jsSetMakerID(delmakerid) {
	var frm = document.frm;

	if (delmakerid == true) {
		frm.targetMakerID.value = '';
	} else {
		if (frm.targetMakerID.value == '') {
			alert('�귣�带 �Է��ϼ���.');
			frm.targetMakerID.focus();
			return;
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = "setmakerid";
		frm.submit();
	}
}

document.title = "1:1 ��㸮��Ʈ";

function resizeTextArea(textarea, textareawidth) {
	var lines = textarea.value.split("\n");

	if (lines.length < 10) {
		return;
	}

	var textareaheight = 1;
	for (x = 0; x < lines.length; x++) {
		c = lines[x].length;

		if (c >= textareawidth) {
			textareaheight += (Math.ceil(c / textareawidth) - 1);
		}
	}
	textareaheight += (lines.length - 1);

	textarea.rows = textareaheight;
}

window.onload = function() {
	if (document.getElementById("replycontents")) {
		resizeTextArea(document.getElementById("replycontents"), 90);
	}

	// ÷������
	var idAttachFile = document.getElementById('idAttachFile');
	if (idAttachFile && idAttachFile.style) {
		if (idAttachFile.clientWidth > 800) {
			idAttachFile.style.width = 800;
		}
	}

	var idAttachFile2 = document.getElementById('idAttachFile2');
	if (idAttachFile2 && idAttachFile2.style) {
		if (idAttachFile2.clientWidth > 800) {
			idAttachFile2.style.width = 800;
		}
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_star.gif" align="absbottom">
	    <font color="red"><strong>1:1 ��� �亯</strong></font>
	</td>
</tr>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
<form method=post name="updateform" action="cscenter_qna_board_act.asp">
<input type="hidden" name="mode" value="CHG">
<input type="hidden" name="id" value="<% = boardqna.results(0).id %>">
<input type="hidden" name="sitename" value="<% = boardqna.results(0).Fsitename %>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="orderserial" value="">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
	    <font color="red"><b>���ǳ���</b></font>
	    &nbsp;&nbsp;
	    ������������ :
	    <% drawSelectBoxqadiv "qadiv", boardqna.results(0).qadiv, "", "Y", "N", "Y" %>

		<!-- ��Ʈ�� �̻� -->
	    <input type="button" class="button" value="����" onclick="updateqadiv();" <% if Not (C_ADMIN_AUTH or C_CSPowerUser) then %>disabled<% end if %>>
	    <% if Not (C_ADMIN_AUTH or C_CSPowerUser) then %><font color=gray>(�����Ұ� : ��Ʈ�幮��)</font><% end if %>
	</td>
</tr>
</form>

<form method="post" name="frm" action="cscenter_qna_board_act.asp" onsubmit="return false">
<!--
<%' if boardqna.results(0).replyuser<>"" then %>
<input type="hidden" name="mode" value="reply">
<%' else %>
<input type="hidden" name="mode" value="firstreply">
<%' end if %>
-->
<input type="hidden" name="mode" value="REP">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="username" value="<%= boardqna.results(0).username %>">
<input type="hidden" name="userphone" value="<%= boardqna.results(0).userphone %>">
<input type="hidden" name="regdate" value="<%= boardqna.results(0).regdate %>">
<input type="hidden" name="title" value="<%= boardqna.results(0).title %>">
<input type="hidden" name="contents" value='<%= replace(html2db(boardqna.results(0).contents),"'","") %>'> <!-- -.- -->
<input type="hidden" name="replydate" value="<%= boardqna.results(0).replydate %>">
<input type="hidden" name="email" value="<%= Replace(boardqna.results(0).usermail, " ", "") %>">
<input type="hidden" name="emailok" value="<%= boardqna.results(0).emailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.results(0).Fextsitename %>">
<input type="hidden" name="sitename" value="<%= boardqna.results(0).Fsitename %>">
<input type="hidden" name="replyuser" value="<%= session("ssBctID") %>">
<input type="hidden" name="imsitxt">
<tr>
	<td width="90" align="center" bgcolor="#FFFFFF"><b>�ۼ���</b></td>
	<td width="570" bgcolor="#FFFFFF">
	    <font color="#464646"><%= boardqna.results(0).username %>(<%= boardqna.results(0).userid %>/<%= boardqna.results(0).orderserial %>)</font>
	    &nbsp;&nbsp;
	    [ <font color="<%= getUserLevelColorByDate(boardqna.results(i).fUserLevel, left(boardqna.results(0).regdate,10)) %>">
		<b><%= getUserLevelStrByDate(boardqna.results(i).fUserLevel, left(boardqna.results(0).regdate,10)) %></b></font> ]
	    <%
	    	if boardqna.results(0).Frealnamecheck="Y" then
	    		Response.Write " / �Ǹ�Ȯ��ȸ��"
	    	end if
			if boardqna.results(0).Fsitename <> "10x10" then
				response.write " / <b>" & boardqna.results(0).Fsitename & "</b>"
				if (boardqna.results(0).FuserGubun = "M") then
					response.write " / ����"
				else
					response.write " / ��"
				end if
			end if
	    %>
	    <% if boardqna.results(0).userid<>"" then %>
    	    <a href="javascript:PopOrderMasterWithCallRingUserid('<%= boardqna.results(0).userid %>');"> >> [ID �� �ֹ��˻�]</a>
		<% end if %>
	</td>
	<td width="90" align="center" bgcolor="#FFFFFF"><b>�����ֹ���ȣ</b></td>
	<td bgcolor="#FFFFFF">
	    <% if boardqna.results(0).orderserial<>"" then %>
    	    <a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= boardqna.results(0).orderserial %>');"><%= boardqna.results(0).orderserial %> >>�󼼺���</a>
		<% end if %>
		<% if boardqna.results(0).orderserial = "" and boardqna.results(0).Fsitename <> "10x10" then %>
			<input type="text" class="text" name="orderserial" size="20" value="">
			<input type="button" class="button" value="����" onclick="updateorderserial()">
		<% end if %>
	</td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>�ۼ��Ͻ�</b></td>
	<td bgcolor="#FFFFFF"><font color="#464646"><%= boardqna.results(i).regdate %></font></td>
	<td align="center" bgcolor="#FFFFFF"><b>���ǻ�ǰ</b></td>
	<td bgcolor="#FFFFFF">
	    <%= boardqna.results(0).itemid %>
	    &nbsp;&nbsp;
	    <% if boardqna.results(0).itemid<>"" and boardqna.results(0).itemid>0 then %>
	    	<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= boardqna.results(0).itemid %>" target="_blank">>>��ǰ����</a>
			<%
			if boardqna.results(0).Fsitename <> "10x10" then
				extItemURL = GetExtItemURL(boardqna.results(0).Fsitename, boardqna.results(0).itemid)
				if extItemURL = "" then
					rw "&nbsp;&nbsp; >> <a href='javascript:alert(""�۾�����"")'>���޸� ��ǰ����</a>"
				else
					rw "&nbsp;&nbsp; >> <a href='" & extItemURL & "' target=_blank>���޸� ��ǰ����</a>"
				end if
			end if
			%>
	    <% elseif (boardqna.results(0).orderserial<>"") then %>
	    	<input type="text" class="text" name="itemid" size="6" value="">
	    	<input type="button" class="button" value="����" onclick="updateitemid()">
		<% else %>
	    	<input type="text" class="text" name="itemid" size="6" value="">
	    	<input type="button" class="button" value="����" onclick="updateitemid()">
	    <% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>�亯 ������</b></td>
	<td bgcolor="#FFFFFF" height="25"><font color="#464646"><%= boardqna.results(0).FExpectReplyDate %></font></td>


	<td align="center" bgcolor="#FFFFFF"><b>�� ����ó</b></td>
	<td bgcolor="#FFFFFF">
	    <%= boardqna.results(0).userphone %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>��������</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.results(0).title)) %></font> <input type="button" class="button" value="UTF8����" onClick="popMyQNAUTF8(<% = boardqna.results(0).id %>)"></td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>�ֹ���ǰ����</b></td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if Not IsNull(boardqna.results(i).Fitemname) and boardqna.results(i).Fitemname <> "" then %>
			<%= boardqna.results(i).Fitemname %>
			<% if (boardqna.results(i).Fitemoption <> "0000") then %>
				<font color="blue">[<%= boardqna.results(i).Fitemoptionname %>]</font>
			<% end if %>
		<% end if %>
	</td>
</tr>
<% if Not IsNull(boardqna.results(i).Fdevice) and boardqna.results(i).Fdevice <> "" then %>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>�ý��� ȯ��</b></td>
	<td bgcolor="#FFFFFF" colspan="3">
			<% If boardqna.results(i).Fdevice="P" Then %>PC : <% Else %>Mobile : <% End If %>
			<%= boardqna.results(i).FOS%>
			<% if (boardqna.results(i).FOSetc <> "") then %>
				 [<%= boardqna.results(i).FOSetc %>]
			<% end if %>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>���ǳ���</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.results(0).contents)) %></font></td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>÷�λ���</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25">
		<% if boardqna.results(0).FattachFile <> "" then %>
			<a href="javascript:popMyQNA_IMAGE(<%= boardqna.results(0).id %>, 0)"><img id="idAttachFile" src="<%= uploadUrl %><%= boardqna.results(0).FattachFile %>" border="0"></a>
		<% end if %>
		<% if boardqna.results(0).FattachFile2 <> "" then %>
			<a href="javascript:popMyQNA_IMAGE(<%= boardqna.results(0).id %>, 1)"><img id="idAttachFile2" src="<%= uploadUrl %><%= boardqna.results(0).FattachFile2 %>" border="0"></a>
		<% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>�亯������</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25">
		<% if (boardqna.results(0).FEvalPoint > 0) then %>
			<% for i = 1 to boardqna.results(0).FEvalPoint %><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/star_red.gif"><% next %>
		<% end if %>
	</td>
</tr>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
<tr height="25" valign="top" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="4" valign="middle">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><b>�亯�ۼ�</b></font>
    </td>
</tr>
<tr>
    <td width="90" align="center" bgcolor="#FFFFFF" height="30">��������</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="radio" name="replyqadiv" value="01" <% if boardqna.results(0).Freplyqadiv = "01" then response.write "checked" %> > �ܼ�����
		<input type="radio" name="replyqadiv" value="02" <% if boardqna.results(0).Freplyqadiv = "02" then response.write "checked" %> > ��ü�Ҹ�
		<input type="radio" name="replyqadiv" value="03" <% if boardqna.results(0).Freplyqadiv = "03" then response.write "checked" %> > ���(CJ)�Ҹ�
		<input type="radio" name="replyqadiv" value="10" <% if boardqna.results(0).Freplyqadiv = "10" then response.write "checked" %> > �ý��۰�����û
		<input type="radio" name="replyqadiv" value="99" <% if boardqna.results(0).Freplyqadiv = "99" then response.write "checked" %> > ��Ÿ�Ҹ�
	</td>
</tr>

<% if boardqna.results(0).replyuser<>"" then %>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">�亯����</td>
		<td colspan="1" bgcolor="#FFFFFF"><input type="text" class="text" name="replytitle" size="55" value="<%= boardqna.results(0).replytitle %>"></td>
		<td colspan="2" bgcolor="#FFFFFF">��ü 1���亯</td>
	</tr>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">�亯����</td>
		<td width="800" bgcolor="#FFFFFF">
			<textarea class="textarea" name="replycontents" cols="90" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea>
		</td>
		<td valign="top" colspan="2" bgcolor="#FFFFFF">
			�귣�� : <input type="text" class="text" name="targetMakerID" value="<%= boardqna.results(0).FtargetMakerID %>" size="20">
			<input type="button" class="button" value="����" onClick="jsSetMakerID(false)">
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
			<input type="button" class="button" value="����" onClick="jsSetMakerID(true)">
			<% end if %>
			<br /><br />
			���û�ǰ �귣�� : <%= boardqna.results(0).Fmakerid %><%= CHKIIF(boardqna.results(0).Fisupchebeasong = "Y", "<font color='red'>(����)</font>", "")%><br /><br />
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
				�亯�� : <%= boardqna.results(0).Fupchereplyuser %><br />
				�亯�Ͻ� : <%= boardqna.results(0).Fupchereplydate %><br /><br />
				�亯���� :<br />
				<%= nl2br(db2html(boardqna.results(0).Fupchereplycontents)) %>
			<% end if %>
		</td>
	</tr>
<% Else %>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">�亯����</td>
		<td colspan="1" bgcolor="#FFFFFF">
			<input type="text" class="text" name="replytitle" value="[�ٹ�����] �ȳ��ϼ���. ���� ���ǿ� ���� �亯�帳�ϴ�." size="55">&nbsp;
			<!-- #include virtual="/cscenter/board/cs_reply_xml_selectbox.asp"-->
		</td>
		<td colspan="2" bgcolor="#FFFFFF">��ü 1���亯</td>
	</tr>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">�亯����</td>
		<td width="800" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="90" rows="20"></textarea></td>
		<td valign="top" colspan="2" bgcolor="#FFFFFF">
			�귣�� : <input type="text" class="text" name="targetMakerID" value="<%= boardqna.results(0).FtargetMakerID %>" size="20">
			<input type="button" class="button" value="����" onClick="jsSetMakerID(false)">
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
			<input type="button" class="button" value="����" onClick="jsSetMakerID(true)">
			<% end if %>
			<br /><br />
			���û�ǰ �귣�� : <%= boardqna.results(0).Fmakerid %><%= CHKIIF(boardqna.results(0).Fisupchebeasong = "Y", "<font color='red'>(����)</font>", "")%><br /><br />
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
				�亯�� : <%= boardqna.results(0).Fupchereplyuser %><br />
				�亯�Ͻ� : <%= boardqna.results(0).Fupchereplydate %><br /><br />
				�亯���� :<br />
				<%= nl2br(db2html(boardqna.results(0).Fupchereplycontents)) %>
			<% end if %>
		</td>
	</tr>
<% End If %>

<tr>
	<td colspan="15" align="center" bgcolor="#FFFFFF">
	    <input type="button" class="button" value=" �亯���� " onclick="SubmitForm()" id="btnSubmit">
	    <input type="button" class="button" value=" ������� " onclick="PopMyQnaList('', '', 'N')">
	</td>
</tr>
</form>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<font color="red"><strong>���� ��� ���</strong></font>
	</td>
</tr>

<% if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="60">����</td>
	    <td width="80">�ֹ���ȣ</td>
	    <td width="50">��ǰ</td>
	    <td>����</td>
	    <td width="200">����</td>
	    <td width="90">�ۼ���</td>
	    <td width="70">�亯����</td>
	    <td width="70">�亯��</td>
	    <td width="70">�亯��</td>
	    <td width="40">����</td>
	</tr>

	<% if myqnalist.ResultCount < 0 then %>
	<% else %>
		<% for i = 0 to (myqnalist.ResultCount - 1) %>
		<tr align="center" <% if (myqnalist.results(i).id <> CLng(request("id"))) then %>bgcolor="#FFFFFF"<% else %> class="tr_select" bgcolor="#AFEEEE"<% end if %>>
		    <td><b><%= myqnalist.results(i).GetUserLevelStr %></b></td>
		    <td><%= myqnalist.results(i).orderserial %></td>
		    <td><%= myqnalist.results(i).itemid %></td>
		    <td align="left"><a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.results(i).title %></a></td>
		    <td>
		    	<a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>">
		    	<%= myqnalist.results(i).fqadivname %></a>
		    </td>
		    <td><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
		    <td><% if (myqnalist.results(i).replyuser<>"") then %>�亯�Ϸ�<% end if %></td>
		    <td><% if (myqnalist.results(i).replyuser<>"") then %><%= myqnalist.results(i).replyuser %><% end if %></td>
		    <td><acronym title="<%= myqnalist.results(i).replydate %>"><%= Left(myqnalist.results(i).replydate,10) %></acronym></td>
		    <td><% if (myqnalist.results(i).dispyn="N") then %><font color="red">����</font><% end if %></td>
		</tr>
		<% next %>
	<% end if %>
<% end if %>
</table>

<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="mode" value="del">
</form>

<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">

function TnChangePrefaceNew(SelectGubun){
	PrefaceFrame.location.href="/cscenter/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.results(0).userid %>&masterid=01";
}

function TnChangeText(str){
	var basictext;
	basictext = "�ȳ��ϼ���. <%= useridForShow %> ��\n"
	basictext = basictext + "�ٹ����� ���ູ���� <%= session("ssBctCname") %>�Դϴ�.\n"
	basictext = basictext + "(����)\n"
	basictext = basictext + "���������亯�� �Ǽ̴�����\n\n"

	if(str == ''){
		document.frm.replycontents.value = basictext;
	}
	else{
		document.frm.replycontents.value = str;
	}
}

</script>
<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">

document.onload = getOnload();

function getOnload() {
	<% if IsNull(boardqna.results(0).replyuser) then %>
	// �亯 ���ø� ����
	requestSelectBoxMaster();
	// �⺻ �λ縻 ����
	TnChangePrefaceNew("<%= CHKIIF(boardqna.results(0).Fsitename <> "10x10", "55", "00") %>");
	<% end if %>
}

function fnSelectBoxDetailSelected(v) {
	TnChangePrefaceNew("<%= CHKIIF(boardqna.results(0).Fsitename <> "10x10", "55", "00") %>");
	setTimeout(function(){
		document.frm.replycontents.value = document.frm.replycontents.value.replace('�� ����', v)
	}, 150);
}

function fnCopyToClipBoard() {
	document.frm.replycontentstr.focus();
	document.frm.replycontentstr.select();

	/*
	if (window.clipboardData && clipboardData.setData) {
		// IE
		clipboardData.setData('text', document.frm.replycontentstr.value);

		alert("����Ǿ����ϴ�.");
	} else {
		alert("��Ʈ��C �� �����ϼ���.");
	}
	*/
}

</script>

<%
set myqnalist = Nothing
set boardqna = Nothing
set orderinfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

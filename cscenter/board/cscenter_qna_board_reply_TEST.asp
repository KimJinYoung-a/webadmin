<%@ language=vbscript %>
<% option explicit %>
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

dim i, j
dim reffrom
reffrom = request("reffrom")

'==============================================================================
'���� 1:1�����亯
dim boardqna
set boardqna = New CMyQNA

boardqna.read(request("id"))


if boardqna.results(0).userid <> "" then
dim orderinfo
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
    myqnalist.list
end if


dim useridForShow : useridForShow = "��"
if boardqna.results(0).userid <> "" then
	useridForShow = boardqna.results(0).userid
end if

%>

<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script>
function SubmitForm()
{
	var replyuser = "<%= session("ssBctID") %>";
	var userid = "<%= boardqna.results(0).userid %>";

	// �� ������ ����
	if (replyuser != "bseo") {
		if (userid == "majorblue") {
			if (confirm("�ȳ�!!\n\n�̼��� ������� �����ϱ�� �Ǿ� �ִ� �����Դϴ�.\n\n��� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		}
	}

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

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">

	<form method=post name="updateform" action="cscenter_qna_board_act.asp">
	<input type="hidden" name="mode" value="CHG">
	<input type="hidden" name="id" value="<% = boardqna.results(0).id %>">
	<input type="hidden" name="itemid" value="">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="15">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
		    <font color="red"><b>���ǳ���</b></font>
		    &nbsp;&nbsp;
		    ������������ :
		    <select name="qadiv" class="select">
		    <option>����</option>
		        <option value="00" <% if boardqna.results(0).qadiv = "00" then response.write "selected" %>>��۹���</option>
		        <option value="01" <% if boardqna.results(0).qadiv = "01" then response.write "selected" %>>�ֹ�����</option>
		        <option value="02" <% if boardqna.results(0).qadiv = "02" then response.write "selected" %>>��ǰ����</option>
		        <option value="03" <% if boardqna.results(0).qadiv = "03" then response.write "selected" %>>�����</option>
		        <option value="04" <% if boardqna.results(0).qadiv = "04" then response.write "selected" %>>��ҹ���</option>
		        <option value="05" <% if boardqna.results(0).qadiv = "05" then response.write "selected" %>>ȯ�ҹ���</option>
		        <option value="06" <% if boardqna.results(0).qadiv = "06" then response.write "selected" %>>��ȯ����</option>
		        <option value="07" <% if boardqna.results(0).qadiv = "07" then response.write "selected" %>>AS����</option>
		        <option value="08" <% if boardqna.results(0).qadiv = "08" then response.write "selected" %>>�̺�Ʈ����</option>
		        <option value="09" <% if boardqna.results(0).qadiv = "09" then response.write "selected" %>>������������</option>
		        <option value="10" <% if boardqna.results(0).qadiv = "10" then response.write "selected" %>>�ý��۹���</option>
		        <option value="11" <% if boardqna.results(0).qadiv = "11" then response.write "selected" %>>ȸ����������</option>
		        <option value="12" <% if boardqna.results(0).qadiv = "12" then response.write "selected" %>>ȸ����������</option>
		        <option value="13" <% if boardqna.results(0).qadiv = "13" then response.write "selected" %>>��÷����</option>
		        <option value="14" <% if boardqna.results(0).qadiv = "14" then response.write "selected" %>>��ǰ����</option>
		        <option value="15" <% if boardqna.results(0).qadiv = "15" then response.write "selected" %>>�Աݹ���</option>
		        <option value="16" <% if boardqna.results(0).qadiv = "16" then response.write "selected" %>>�������ι���</option>
		        <option value="17" <% if boardqna.results(0).qadiv = "17" then response.write "selected" %>>����/���ϸ�������</option>
		        <option value="18" <% if boardqna.results(0).qadiv = "18" then response.write "selected" %>>�����������</option>
		        <option value="20" <% if boardqna.results(0).qadiv = "20" then response.write "selected" %>>��Ÿ����</option>
                <option value="21" <% if boardqna.results(0).qadiv = "21" then response.write "selected" %>>���̶�ҹ���</option>
                <option value="23" <% if boardqna.results(0).qadiv = "23" then response.write "selected" %>>����ǰ����</option>
                <option value="24" <% if boardqna.results(0).qadiv = "24" then response.write "selected" %>>POINT1010����</option>
                <option value="25" <% if boardqna.results(0).qadiv = "25" then response.write "selected" %>>�������幮��</option>
		    </select>
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
    <input type="hidden" name="replyuser" value="<%= session("ssBctID") %>">
    <input type="hidden" name="imsitxt">
    <tr>
    	<td width="90" align="center" bgcolor="#FFFFFF"><b>�ۼ���</b></td>
    	<td width="570" bgcolor="#FFFFFF">
    	    <font color="#464646"><%= boardqna.results(0).username %>(<%= boardqna.results(0).userid %>/<%= boardqna.results(0).orderserial %>)</font>
    	    &nbsp;&nbsp;
    	    [ <b><%= boardqna.results(0).GetUserLevelStr %></b> ]
    	    <%
    	    	if boardqna.results(0).Frealnamecheck="Y" then
    	    		Response.Write " / �Ǹ�Ȯ��ȸ��"
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
<!--     	        <a href="/admin/ordermaster/viewordermaster.asp?orderSerial=<%= boardqna.results(0).orderserial %>" target="_blank">>>�󼼺���</a> -->
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
    	    <% elseif (boardqna.results(0).orderserial<>"") then %>
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
    <tr>
    	<td align="center" bgcolor="#FFFFFF"><b>���ǳ���</b></td>
    	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.results(0).contents)) %></font></td>
    </tr>
    <tr>
    	<td align="center" bgcolor="#FFFFFF"><b>÷�λ���</b></td>
    	<td colspan="3" bgcolor="#FFFFFF" height="25">
			<% if boardqna.results(0).FattachFile <> "" then %>
				<img src="<%= uploadUrl %><%= boardqna.results(0).FattachFile %>">
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

<p>

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
    	<td colspan="3" bgcolor="#FFFFFF"><input type="text" class="text" name="replytitle" size="55" value="<%= boardqna.results(0).replytitle %>"></td>
    </tr>
    <tr>
        <td align="center" bgcolor="#FFFFFF">�亯����</td>
    	<td width="570" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="90" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea></td>
		<td valign="top" colspan="2" bgcolor="#FFFFFF">
			<p>
			&nbsp;
			<!-- #include virtual="/cscenter/board/cs_reply_xml_selectbox.asp"-->
			&nbsp;
			<input class="button" type="button" value="�����ϱ�" onClick="fnCopyToClipBoard()">
			<p>
			<textarea class="textarea" name="replycontentstr" cols="75" rows="8"></textarea>
		</td>
    </tr>
    <% Else %>
    <tr>
        <td align="center" bgcolor="#FFFFFF">�亯����</td>
    	<td colspan="1" bgcolor="#FFFFFF">
    		  <input type="text" class="text" name="replytitle" value="[�ٹ�����] �ȳ��ϼ���. ���� ���ǿ� ���� �亯�帳�ϴ�." size="55">&nbsp;
    		  <% SelectBoxQnaPreface "01" %>&nbsp;
    		  <% SelectBoxQnaCompliment "" %>
    	</td>
    	<td colspan="2" bgcolor="#FFFFFF">�ȳ�����</td>
    </tr>
    <tr>
        <td align="center" bgcolor="#FFFFFF">�亯����</td>
    	<td bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="90" rows="20"></textarea></td>
		<td valign="top" colspan="2" bgcolor="#FFFFFF">
			<p>
			&nbsp;
			<!-- #include virtual="/cscenter/board/cs_reply_xml_selectbox.asp"-->
			&nbsp;
			<input class="button" type="button" value="�����ϱ�" onClick="fnCopyToClipBoard()">
			<p>
			<textarea class="textarea" name="replycontentstr" cols="75" rows="8"></textarea>
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

<p>


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
        <td width="70">����</td>
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
        <td><a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.code2name(myqnalist.results(i).qadiv) %></a></td>
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
<!--

 function TnChangePreface(SelectGubun){
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
<% if boardqna.results(0).replyuser <> "" then %>
<% else %>
TnChangeText('');
<% end if %>
//-->
</script>
<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangeCompliment(SelectGubun){
	ComplimentFrame.location.href="/cscenter/board/compliment_select.asp?masterid=01&gubun=" + SelectGubun;
 }

 function TnChangeText2(str){

	if(str == ''){
	}
	else{
		document.frm.replycontents.value = document.frm.imsitxt.value + "\n" + str;
	}
}

document.onload = getOnload();

function getOnload(){
	requestSelectBoxMaster();

	if (document.frm.preface) {
		document.frm.preface.value = "00";
		TnChangePreface("00")
	}
}

function fnSelectBoxDetailSelected(v) {
	document.frm.replycontentstr.value = v;
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

//-->
</script>
<%
set myqnalist = Nothing
set boardqna = Nothing
set orderinfo = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

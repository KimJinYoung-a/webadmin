<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [��ü]�Խ���
' Hieditor : 2015.05.27 �̻� ����
'			 2020.03.12 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim i, j, reffrom, page, SearchKey, SearchString, gubun, replyYn, usingYn, param, sWorkerId, sWorkerName, selDate, sDate, eDate
dim idx, replytitle
	reffrom = request("reffrom")
	page = requestcheckvar(getNumeric(Request("page")),10)
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	gubun = requestcheckvar(getNumeric(Request("gubun")),2)
	replyYn = requestcheckvar(Request("replyYn"),1)
	usingYn = requestcheckvar(Request("usingYn"),1)
	selDate		= requestCheckVar(Request("selDate"),1)
	sDate 		= requestCheckVar(Request("sDate"),10)
	eDate 		= requestCheckVar(Request("eDate"),10)
	idx = requestcheckvar(getNumeric(Request("idx")),10)

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&gubun=" & gubun & "&replyYn=" & replyYn & "&usingYn=" & usingYn & "&selDate=" & selDate & "&sDate=" & sDate & "&eDate=" & eDate

'==============================================================================
dim boardqna
set boardqna = New CUpcheQnADetail
	boardqna.FRectIdx = idx
	boardqna.read()

sWorkerId = boardqna.Fworkerid
replytitle = boardqna.freplytitle
if replytitle="" or isnull(replytitle) then replytitle="�ȳ��ϼ���"
If sWorkerId <> "" Then
	sWorkerName = fnGetMemberName(sWorkerId)
End IF
%>

<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script type="text/javascript">

function workerlist(){
	var worker = document.frm.workerid.value;
	window.open('/designer/board/PopWorkerList.asp?workerid='+worker+'&idx=<%= idx %>','worker','width=590,height=527,scrollbars=yes');
}

function SubmitForm(){
	if (document.frm.workerid) {
        if (document.frm.workerid.value == "") {
                alert("����ڸ� �������ּ���.");
                return;
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

    if (confirm("�Է��� ��Ȯ�մϱ�?") == true) { document.frm.submit(); }
}

function updateqadiv(){
	if (confirm("�����Ͻðڽ��ϱ�?")){
		document.updateform.submit();
	}
}

function delqadiv(){
	if (confirm("�����Ͻðڽ��ϱ�?")){
		document.delform.submit();
	}
}

function TnCSTemplateGubunChanged(gubun) {

	CSTemplateFrame.location.href="/cscenter/board/cs_template_select_process.asp?mastergubun=31&gubun=" + gubun;
}

function TnCSTemplateGubunProcess(v, errMSG) {

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if(v == "") {
		//
	} else {
		document.frm.replycontents.value = v;
		// alert(v);
	}
}

function workerchange()
{
    if (document.frm.workerid.value == "") {
            alert("����ڸ� �������ּ���.");
            return;
    }

	if(confirm("�����Ͻ� ����ڷ� �����Ͻðڽ��ϱ�?") == true) {
		document.frm.mode.value = "edit";
		document.frm.submit();
		return true;
     } else {
     	return false;
     }
}
</script>

<form method="post" name="frm" action="/admin/board/upche_qna_board_act.asp" onsubmit="return false" style="margin:0px;" >
<input type="hidden" name="mode" value="reply">
<input type="hidden" name="idx" value="<%= boardqna.Fidx %>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="replyYn" value="<%=replyYn%>">
<input type="hidden" name="selDate" value="<%=selDate%>">
<input type="hidden" name="sDate" value="<%=sDate%>">
<input type="hidden" name="eDate" value="<%=eDate%>">
<input type="hidden" name="imsitxt">
<table width="800" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<font color="red"><b>���ǳ���</b></font>

	</td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#FFFFFF"><b>�ۼ���</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Fusername %>(<%= boardqna.fuserid %>)</td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>�ۼ��Ͻ�</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Fregdate %></td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>��������</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= nl2br(ReplaceBracket(boardqna.Ftitle)) %></td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>���ǳ���</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= nl2br(ReplaceBracket(boardqna.Fcontents)) %></td>
</tr>

<tr height="25" valign="top" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="4" valign="middle">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<font color="red"><b>�亯�ۼ�</b></font>
	</td>
</tr>

<% if not isnull(boardqna.Freplyuser) then %>
	<tr height="25">
		<td align="center" bgcolor="#FFFFFF"><b>�亯��</b></td>
		<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Freplyuser %></td>
	</tr>
	<tr height="25">
		<td align="center" bgcolor="#FFFFFF"><b>�亯�Ͻ�</b></td>
		<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Freplydate %></td>
	</tr>
<% else %>
	<tr height="25">
		<td align="center" bgcolor="#FFFFFF"><b>�����</b></td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="workername" value="<%=sWorkerName%>" size="10" readonly>
			<input type="hidden" name="workerid" value="<%=sWorkerId%>">
			<input type="button" class="button_s" value="����ڸ���Ʈ" onClick="workerlist()">
			<input type="button" class="button_s" value="����ں����ϱ�" onClick="workerchange()">
		</td>
	</tr>
<% end if %>

<tr>
	<td align="center" bgcolor="#FFFFFF"><b>�亯����</b></td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="replytitle" size="30" value="<%= replytitle %>">&nbsp;

		<% if boardqna.freplycontents = "" then  %>
			<% SelectBoxCSTemplateGubunNew "31", "csreg_template", "" %>
			*[CS]�������� >> [��ü�Խ���]���ø����� ����
			<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
		<% end if  %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>�亯����</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="80" rows="10"><%= (boardqna.freplycontents) %></textarea></td>
</tr>
<tr>
	<td colspan="15" align="center" bgcolor="#FFFFFF">
		<% if boardqna.fisusing="Y" then %>
			�亯�̸��Ϲ߼�
			<input type="checkbox" name="csmailsend" >
			<input type="button" class="button" value="�亯����" onclick="SubmitForm()">
		<% end if %>

		<input type="button" class="button" value="�������" onclick="window.location='upche_qna_board_list.asp?page=<%=page & Param%>'">
	</td>
</tr>
</table>
</form>

<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script type="text/JavaScript">

function TnChangePreface(SelectGubun){
	PrefaceFrame.location.href="/cscenter/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.fuserid %>&masterid=03";
}

 function TnChangeText(str){
	var basictext;
	basictext = "";
	//basictext = "�ȳ��ϼ���?\n�ٹ����� <%= session("ssBctCname") %> �Դϴ�.\n\n\n�����ϼ���~\n�ñ��Ͻ� ������ �Ʒ��� �����ּ���^^\n";
	//basictext = basictext + "��� : <%= session("ssBctCname") %>";

	if(str == ''){
		document.frm.replycontents.value = basictext;
	} else {
		document.frm.replycontents.value = str;
	}
 }

<% if boardqna.freplycontents="" or isNull(boardqna.freplycontents) then %>
	TnChangeText('');
<% end if %>

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

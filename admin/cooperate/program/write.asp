<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/programchangeCls.asp"-->

<%
	Dim vIdx, vTitle, vContent, iCurrentpage, vRegUserID, vSign1, vSign2, vSign1Date, vSign2Date, vFileName, vRegdate, FUsername, vDoc_Idx
	Dim vParam, vChkList, vSign1Chk, vSign2Chk
	vIdx 			= requestCheckVar(Request("pidx"),10)
	vDoc_Idx		= requestCheckVar(Request("didx"),10)
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)

	Dim cPrCh
	Set cPrCh = New CProgramChange
	cPrCh.FPIdx = vIdx
	cPrCh.fnGetPrChView

	vTitle = cPrCh.FTitle
	vContent = cPrCh.FContent
	vRegUserID = cPrCh.FReguserid
	FUsername = cPrCh.FUsername
	vFileName = cPrCh.FFileName
	vSign1 = cPrCh.FSign1
	vSign2 = cPrCh.FSign2
	vSign1Date = cPrCh.FSign1date
	vSign2Date = cPrCh.FSign2date
	vRegdate = cPrCh.FRegdate
	If vIdx <> "" Then
		vDoc_Idx = cPrCh.FDocIdx
		If vDoc_Idx = "0" Then vDoc_Idx = "" End If
	End If
	vChkList = cPrCh.FChkList
	vSign1Chk = cPrCh.FSign1Chk
	vSign2Chk = cPrCh.FSign2Chk
	Set cPrCh = Nothing

	vParam = "&menupos="&request("menupos")&"&reguserid="&Request("reguserid")&"&title="&Request("title")&"&1check="&Request("1check")&"&2check="&Request("2check")&""
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="Javascript">
function checkform()
{
	if (frm.title.value == "")
	{
		alert("������ �Է��ϼ���!");
		frm.title.focus();
		return;
	}
	if (frm.contents.value == "")
	{
		alert("������ �Է��ϼ���!");
		return;
	}
	if($('input[name="programchk"]').is(":checked") == false)
	{
		alert("üũ����Ʈ�� ������ üũ�ϼ���!");
		return;
	}
	<% If session("ssBctId") = "tozzinet" Then %>
	if($('input[name="sign1chk"]').is(":checked") == false)
	{
		alert("1�� ���� Ȯ���� üũ�ϼ���!");
		return;
	}
	<% ElseIf session("ssBctId") = "kobula" Then %>
	if($('input[name="sign2chk"]').is(":checked") == false)
	{
		alert("2�� ���� Ȯ���� üũ�ϼ���!");
		return;
	}
	<% End If %>
	frm.submit();
}

function goSign(){
	frm.gubun.value = "sign";
	frm.submit();
}
</script>
<form name="frm" action="proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="pidx" value="<%=vIdx%>">
<input type="hidden" name="gubun" value="">
<input type="hidden" name="menupos" value="<%=request("menupos")%>">
<input type="hidden" name="iC" value="<%=request("iCurrentpage")%>">
<input type="hidden" name="didx" value="<%=vDoc_Idx%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<% If vIdx <> "" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ȣ</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=vIdx%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�ۼ���</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=FUsername%> (����� : <%=vRegdate%>)</td>
		</tr>
		<% End If %>
		<% If vDoc_Idx <> "" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����������ȣ</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=vDoc_Idx%></td>
		</tr>
		<% End If %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="title" value="<%=vTitle%>" size="110" maxlength="74">
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���ϸ�</td>
			<td bgcolor="#FFFFFF" style="padding: 5 0 5 5">
				<textarea class="textarea" name="filename" cols="110" rows="6"><%=vFileName%></textarea>
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
			<td bgcolor="#FFFFFF" style="padding: 5 5 5 5">
				<input type="text" class="text" name="contents" value="<%=vContent%>" size="110" maxlength="198">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">
				<% If session("ssBctId") = "tozzinet" OR session("ssBctId") = "kobula" Then %>
					�ۼ���<br>üũ����Ʈ
				<% Else %>
					������<br>üũ����Ʈ
				<% End If %>
			</td>
			<td bgcolor="#FFFFFF" style="padding: 5 5 5 5">
				<table class="a" width="100%">
				<tr>
					<td style="padding:3px;"><label id="programchk1" style="cursor:pointer;"><input type="checkbox" name="programchk" value="1" id="programchk1" <%=fnCheckBoxCheck(vChkList,"1")%>> �Ķ���� üũ(����, �Ӽ�(������,������), ���� ����, �ҿ����� ���� ó�� ��)</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk2" style="cursor:pointer;"><input type="checkbox" name="programchk" value="2" id="programchk2" <%=fnCheckBoxCheck(vChkList,"2")%>> ������ ��������(ID, PW ���� �߿�����)�� ����ִ��� üũ(ȸ������, �α��� ����)</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk3" style="cursor:pointer;"><input type="checkbox" name="programchk" value="3" id="programchk3" <%=fnCheckBoxCheck(vChkList,"3")%>> �α��� ������ �ݵ�� �ʿ��� �������� IDüũ include ������ �ִ��� üũ</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk4" style="cursor:pointer;"><input type="checkbox" name="programchk" value="4" id="programchk4" <%=fnCheckBoxCheck(vChkList,"4")%>> "(��)�ٹ����� ���� ǥ�� �� ���� �ڵ� ���̵�" �� �´� �ڵ����� üũ</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk5" style="cursor:pointer;"><input type="checkbox" name="programchk" value="5" id="programchk5" <%=fnCheckBoxCheck(vChkList,"5")%>> ���ε� ������ �ִ� ��� MIME TYPE, �뷮, ���Ἲ ���� üũ�� �Ǿ����� üũ</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk6" style="cursor:pointer;"><input type="checkbox" name="programchk" value="6" id="programchk6" <%=fnCheckBoxCheck(vChkList,"6")%>> ��ȹ�� ���� ��� ����� ���� �׽�Ʈ�� �Ͽ����� üũ</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk7" style="cursor:pointer;"><input type="checkbox" name="programchk" value="7" id="programchk7" <%=fnCheckBoxCheck(vChkList,"7")%>> �Ǽ����� �ø��� ���� �غ� �Ͽ����� üũ(�׽�Ʈ ������ ����, �ҽ����� ��)</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk8" style="cursor:pointer;"><input type="checkbox" name="programchk" value="8" id="programchk8" <%=fnCheckBoxCheck(vChkList,"8")%>> ���� ����ڰ� �ִ� ��� ���������� ������� ������ �̷�������� üũ</label></td>
				</tr>
				<% If session("ssBctId") = "tozzinet" Then %>
				<tr>
					<td height="70" style="padding:15px;" bgcolor="<%= adminColor("tabletop") %>"><label id="sign1chk" style="cursor:pointer;"><input type="checkbox" name="sign1chk" value="1" id="sign1chk" <%=CHKIIF(vSign1Chk=True,"checked","")%>> 1�� ���� Ȯ��</label></td>
				</tr>
				<% ElseIf session("ssBctId") = "kobula" Then %>
				<tr>
					<td height="70" style="padding:15px;" bgcolor="<%= adminColor("tabletop") %>"><label id="sign2chk" style="cursor:pointer;"><input type="checkbox" name="sign2chk" value="1" id="sign2chk" <%=CHKIIF(vSign2Chk=True,"checked","")%>> 2�� ���� Ȯ��</label></td>
				</tr>
				<% End If %>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<table width="810" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left">
	<% If vDoc_Idx = "" Then %>
		<input type="button" value="����Ʈ" onClick="location.href='index.asp?iC=<%=iCurrentpage%><%=vParam%>';">
	<% End If %>
	</td>
	<td width="50%" align="right">
		<% If vRegUserID = session("ssBctId") OR vIdx = "" Then %>
			<input type="button" value="�� ��" onClick="checkform();">
		<% End If %>
	</td>
</tr>
</table>
</form>

<br><br>

<% If vIdx <> "" Then %>
<table border="0" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<tr height="30">
	<td align="left" bgcolor="#FFFFFF">
		1�� ��Ʈ�� ���� :
		<% If session("ssBctId") = "tozzinet" Then %>
			<% If vSign1 = "" Then %>
				<input type="button" value="�����ϱ�" onClick="goSign()">
			<% Else %>
				���� �Ϸ�. <%=vSign1Date%>
			<% End If %>
		<% Else %>
			<% If vSign1 = "" Then %>
				���� ��.
			<% Else %>
				���� �Ϸ�. <%=vSign1Date%>
			<% End If %>
		<% End If %>
	</td>
</tr>
<tr height="30">
	<td align="left" bgcolor="#FFFFFF">
		2�� ���� ���� :
		<% If session("ssBctId") = "kobula" Then %>
			<% If vSign2 = "" Then %>
				<input type="button" value="�����ϱ�" onClick="goSign()">
			<% Else %>
				���� �Ϸ�. <%=vSign2Date%>
			<% End If %>
		<% Else %>
			<% If vSign2 = "" Then %>
				���� ��.
			<% Else %>
				���� �Ϸ�. <%=vSign2Date%>
			<% End If %>
		<% End If %>
	</td>
</tr>
</table>
<% End If %>

<% If vIdx <> "" Then %>
<!-- ####### �亯���� ####### //-->
<br>
<iframe src="iframe_program_ans.asp?pidx=<%=vIdx%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### �亯���� ####### //-->
<% End If %>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

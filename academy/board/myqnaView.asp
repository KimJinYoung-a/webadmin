<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/myqnaCls.asp"-->
<%
'####################################################
' Description :  1:1 ������ ��
' History : 2016.07.28 ������ ����
' History : 2016-11-08 ����ȭ �߰�
'####################################################
%>
<%
Dim oMyqna, i, idx, gridx, reidx
Dim masterQState, masterQGubun, masterQRegID, masterQRegName, masterQEmail, masterQRegdate, masterQLastRegdate, masterQPhoneChk, masterQPhoneNumber, masterQTitle, masterQSmsOK
Dim masterQorderserial , masterQq_itemid , masterQq_itemoption , masterQItemNames , masterQitemcount ,masterQitemoptionname , masterQtotalsum, masterQitemcost

Dim regIDnName
idx		= getNumeric(requestCheckVar(request("idx"),9))
gridx	= getNumeric(requestCheckVar(request("gridx"),9))

SET oMyqna = new CQna
	oMyqna.FRectIdx = idx
	oMyqna.FRectGroupIdx = gridx
	oMyqna.getOnemyqna

	If oMyqna.FResultCount < 1 Then
		response.write "<script>alert('������ �߻��߽��ϴ�.');location.replace('/academy/board/myqnaList.asp?menupos="&menupos&"');</script>"
		response.end
	End If

	masterQState		= oMyqna.FOneItem.getAnswerName
	masterQGubun		= oMyqna.FOneItem.FLecture_gubun
	masterQRegID		= oMyqna.FOneItem.FUserid
	masterQRegdate		= oMyqna.FOneItem.FRegdate
	masterQLastRegdate	= oMyqna.FOneItem.FLastRegdate
	masterQPhoneNumber	= oMyqna.FOneItem.FSmsnum & " (�亯����)"
	masterQSmsOK		= oMyqna.FOneItem.FSmsok
	masterQTitle		= oMyqna.FOneItem.FTitle

	masterQorderserial		=	oMyqna.FOneItem.Forderserial
	masterQq_itemid			=	oMyqna.FOneItem.Fq_itemid	
	masterQq_itemoption		=	oMyqna.FOneItem.Fq_itemoption
	masterQItemNames		=	oMyqna.FOneItem.FItemNames
	masterQitemcount		=	oMyqna.FOneItem.Fitemcount
	masterQitemoptionname	=	oMyqna.FOneItem.Fitemoptionname
	masterQtotalsum			=	oMyqna.FOneItem.Ftotalsum
	masterQitemcost			=	oMyqna.FOneItem.Fitemcost

	Dim GetItemNames
	if (masterQitemcount>1) then
		GetItemNames = masterQItemNames + " �� " + CStr(masterQitemcount-1) + "��"
	else
		GetItemNames = masterQItemNames
	end If
	
	Call getMyinfo(masterQRegID, masterQRegName, masterQEmail)
SET oMyqna = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function chggubun(){
	var frm = document.frm;
	if(confirm("���Ǻо߸� �����Ͻðڽ��ϱ�?")){
		frm.mode.value = "C";
		frm.submit();
	}
}
function fnqnaDel(){
	var frm = document.frm;
	if(confirm("���Ǳ��� �����Ͻðڽ��ϱ�?")){
		frm.mode.value = "D";
		frm.submit();
	}
}
function goView(vidx, vgridx){
	location.href='/academy/board/myqnaView.asp?menupos=<%=menupos%>&idx='+vidx+'&gridx='+vgridx;	
}
// �亯 �Ӹ��� �ֱ�
function chgCont(qcd, ccd, regid){
	var reStr;
	var rstStr = $.ajax({
		type: "POST",
		url: "ajax_myqnaTextarea.asp",
		data: "groupcd="+qcd+"&commcd="+ccd+"&regid="+regid,
		dataType: "text",
		async: false
	}).responseText;
	reStr = rstStr.split("|");
	if(reStr[0]=="OK"){
		$("#ansContents").val(reStr[1]);
	}
}
// �亯 ������ �Ӹ��� �ֱ�
function chgContEdit(qcd, ccd, regid){
	var reStr;
	var rstStr = $.ajax({
		type: "POST",
		url: "ajax_myqnaTextarea.asp",
		data: "groupcd="+qcd+"&commcd="+ccd+"&regid="+regid,
		dataType: "text",
		async: false
	}).responseText;
	reStr = rstStr.split("|");
	if(reStr[0]=="OK"){
		$("#ansContentsEdit").val(reStr[1]);
	}
}
// �亯 ���
function fnQnareplyAdd(){
	var frm = document.replyfrm;
	if(confirm("�亯���� ����Ͻðڽ��ϱ�?")){
		frm.mode.value = "addreply";
		frm.submit();
	}
}
// �亯 ����
function fnQnareplyDel(vidx){
	var frm = document.frm;
	if(confirm("�亯���� �����Ͻðڽ��ϱ�?")){
		frm.mode.value = "adel";
		frm.reidx.value = vidx;
		frm.submit();
	}
}
// �亯�� ������ ������
function fnQnareplyEditForm(vidx, commid){
	var editTrid = "QnAList"+vidx;
	var commVal = $("#"+commid+"").html();
	var repComm;
	repComm = commVal.replace(/<BR>/gi, "\n")
	$("#"+editTrid+"").hide();
	$("#replyEditTBL").show();
	$("#editidx").val(vidx);
	$("#ansContentsEdit").val(repComm);
}
// �亯�� ����
function fnQnareplyEdit(){
	var frm = document.replyEditfrm;
	if(confirm("�亯���� �����Ͻðڽ��ϱ�?")){
		frm.mode.value = "edit";
		frm.submit();
	}
}
</script>
<!-- ########################################### ������ ���� ���� ########################################### -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="/academy/board/doMyQnaProc.asp">
<input type="hidden" name="mode" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="reidx" value="<%= reidx %>" >
<input type="hidden" name="gridx" value="<%= gridx %>" >
<input type="hidden" name="menupos" value="<%= menupos %>" >
<col width="15%" />
<col width="35%" />
<col width="15%" />
<col width="35%" />
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left"><%= masterQState %></td>
	<td align="left" bgcolor="<%= adminColor("gray") %>">���Ǻо�</td>
	<td align="left">
		<select class="select" name="gubunVal">
			<option value="1" <%= Chkiif(masterQGubun = "1", "selected", "") %> >��ǰ(��ǰ) �ֹ�/����</option>
			<option value="2" <%= Chkiif(masterQGubun = "2", "selected", "") %>>�ֹ� ���/��ǰ/��ȯ</option>
			<option value="3" <%= Chkiif(masterQGubun = "3", "selected", "") %>>��ǰ ��� ���� ����</option>
			<option value="4" <%= Chkiif(masterQGubun = "4", "selected", "") %>>������û/���� ����</option>
			<option value="5" <%= Chkiif(masterQGubun = "5", "selected", "") %>>���� ���</option>
			<option value="6" <%= Chkiif(masterQGubun = "6", "selected", "") %>>�������� ���� ����</option>
			<option value="7" <%= Chkiif(masterQGubun = "7", "selected", "") %>>�̺�Ʈ/����/���ϸ���</option>
			<option value="8" <%= Chkiif(masterQGubun = "8", "selected", "") %>>ȸ��Ż��/�簡��</option>
			<option value="9" <%= Chkiif(masterQGubun = "9", "selected", "") %>>��Ÿ ����</option>
		</select>
		<input type="button" class="button" value="���к���" onclick="chggubun();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">�ۼ���</td>
	<td align="left"><%= masterQRegName %>(<%= masterQRegID %>)</td>
	<td align="left" bgcolor="<%= adminColor("gray") %>">�ۼ���(��������)</td>
	<td align="left"><%= masterQRegdate %>&nbsp;(<%= masterQLastRegdate %>) </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">�̸���</td>
	<td align="left"><%= Chkiif(masterQEmail <> "", masterQEmail & " (�亯����)", "") %></td>
	<td align="left" bgcolor="<%= adminColor("gray") %>">�޴���</td>
	<td align="left"><%= Chkiif(masterQSmsOK = "Y", masterQPhoneNumber, "") %></td>
</tr>
<% If masterQorderserial <> "" Then %>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">�ֹ� �� ������ȣ</td>
	<td align="left" colspan="3"><%=masterQorderserial %> / <%=GetItemNames%> / <%= masterQtotalsum %>��</td>
</tr>
<% End If %>
<% If masterQq_itemid <> "" Then %>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">��ǰ �ڵ�</td>
	<td align="left" colspan="3"><a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=masterQq_itemid %>" target="_blank"><%=masterQq_itemid %> / <%=masterQItemNames%>[<%=masterQitemoptionname%>] / <%= masterQitemcost %>��</a></td>
</tr>
<% End If %>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left"><%= masterQTitle %></td>
	<td align="center" colspan="2">
		<input type="button" class="button" value="����" onclick="fnqnaDel();" style=color:red;font-weight:bold>
	</td>
</tr>
</form>
</table>
<br>
<!-- ############################################ ������ ���� �� ############################################ -->
<!-- ########################################### ������ ���� ���� ########################################### -->
<%
Dim lastqna, qstContents, lastRegdate, lastSMSok, lastSmsNum
Dim QnaColor
SET oMyqna = new CQna
	oMyqna.FCurrPage = 1
	oMyqna.FPageSize = 500
	oMyqna.FRectGroupIdx = gridx
	oMyqna.getqnaDetailList
%>
<% If oMyqna.FResultCount > 0 Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
	For i = 0 to oMyqna.FResultCount - 1 
		If oMyqna.FItemList(i).FQna = "Q" Then
			QnaColor = "<font size='4' color='RED'><strong>"&oMyqna.FItemList(i).FQna&".</strong></font>"
		Else
			QnaColor = "<font size='4' color='BLUE'><strong>"&oMyqna.FItemList(i).FQna&".</strong></font>"
		End IF
%>
<tr align="LEFT" bgcolor="#FFFFFF" height="35" id="QnAList<%= oMyqna.FItemList(i).Fidx %>">
	<td><%= QnaColor %><br>
		<span id="QnAComm<%= oMyqna.FItemList(i).Fidx %>"><%= nl2br(oMyqna.FItemList(i).Fcomment) %></span>
	<% If oMyqna.FItemList(i).FanswerYN ="Y" and oMyqna.FItemList(i).Freply_num+1 >= oMyqna.FTotalCount AND oMyqna.FItemList(i).FQna = "A" Then %>
		<br><button type="button" onclick="fnQnareplyEditForm('<%= oMyqna.FItemList(i).Fidx %>', 'QnAComm<%= oMyqna.FItemList(i).Fidx %>');" class="button">����</button>
		&nbsp;<button type="button" onclick="fnQnareplyDel('<%= oMyqna.FItemList(i).Fidx %>');" class="button">����</button>
	<% End If %>
	</td>
</tr>
<% 
		lastqna			= oMyqna.FItemList(i).FQna 
		If lastqna = "Q" Then
			qstContents		= oMyqna.FItemList(i).Fcomment
			lastRegdate		= oMyqna.FItemList(i).FRegdate
			lastSMSok		= oMyqna.FItemList(i).FSmsok
			lastSmsNum		= oMyqna.FItemList(i).FSmsnum
		End If
%>
<%	Next %>
</table>
<br>
<% End If %>
<!-- ########################################### ������ ���� �� ########################################### -->
<!-- ################################### ������ �� �� �亯 ��� �� ���� ################################### -->
<% If lastqna = "Q" Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="replyfrm" method="POST" action="/academy/board/doMyQnaProc.asp">
<input type="hidden" name="mode" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="gridx" value="<%= gridx %>" >
<input type="hidden" name="menupos" value="<%= menupos %>" >
<!-- ���Ͽ� �ʿ��� ���� hidden ó�� -->
<input type="hidden" name="usermail" value="<%= masterQEmail %>" >
<input type="hidden" name="qstContents" value="<%= qstContents %>" >
<input type="hidden" name="lastRegdate" value="<%= lastRegdate %>" >
<input type="hidden" name="masterQRegName" value="<%= masterQRegName %>" >
<input type="hidden" name="masterQTitle" value="<%= masterQTitle %>" >
<!-- ################################-->
<!-- SMS���ۿ� �ʿ��� ���� hidden ó�� -->
<input type="hidden" name="lastSMSok" value="<%= lastSMSok %>" >
<input type="hidden" name="lastSmsNum" value="<%= lastSmsNum %>" >
<!-- ################################-->
<tr align="LEFT" bgcolor="#FFFFFF" height="35">
	<td>
		<font size='4' color='BLUE'><strong>A.</strong></font><br />
		�Ӹ���
		<select name="preface" id="preface" class="select" onchange="chgCont(this.value, compliment.value, '<%=masterQRegID%>')">
			<%= oMyqna.optPrfCd("'A000'", "H999")%>
		</select>
		/ �λ縻
		<select name="compliment" id="compliment" class="select" onchange="chgCont(preface.value, this.value, '<%=masterQRegID%>')">
			<option value="">����</option>
			<%= oMyqna.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" class="textarea" id="ansContents" rows="20" cols="100"></textarea>
		&nbsp;<input type="button" value="�亯�ϱ�" class="button" onclick="fnQnareplyAdd();">
	</td>
</tr>
</form>
</table>
<% End If %>
<!-- #################################### ������ �� �� �亯 ��� �� �� ##################################### -->
<!-- ################################### �亯 ���� Ŭ���� ������ �� ���� ################################### -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="replyEditTBL" style="display:none;">
<form name="replyEditfrm" method="POST" action="/academy/board/doMyQnaProc.asp">
<input type="hidden" name="mode" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="reidx" id="editidx" >
<input type="hidden" name="gridx" value="<%= gridx %>" >
<input type="hidden" name="menupos" value="<%= menupos %>" >
<!-- ���Ͽ� �ʿ��� ���� hidden ó�� -->
<input type="hidden" name="usermail" value="<%= masterQEmail %>" >
<input type="hidden" name="qstContents" value="<%= qstContents %>" >
<input type="hidden" name="lastRegdate" value="<%= lastRegdate %>" >
<input type="hidden" name="masterQRegName" value="<%= masterQRegName %>" >
<input type="hidden" name="masterQTitle" value="<%= masterQTitle %>" >
<!-- ################################-->
<!-- SMS���ۿ� �ʿ��� ���� hidden ó�� -->
<input type="hidden" name="lastSMSok" value="<%= lastSMSok %>" >
<input type="hidden" name="lastSmsNum" value="<%= lastSmsNum %>" >
<!-- ################################-->
<tr align="LEFT" bgcolor="#FFFFFF" height="35">
	<td>
		<font style=font-weight:bold>A.</font><br />
		�Ӹ���
		<select name="preface" id="preface" class="select" onchange="chgContEdit(this.value, compliment.value, '<%=masterQRegID%>')">
			<%= oMyqna.optPrfCd("'A000'", "H999")%>
		</select>
		/ �λ縻
		<select name="compliment" id="compliment" class="select" onchange="chgContEdit(preface.value, this.value, '<%=masterQRegID%>')">
			<option value="">����</option>
			<%= oMyqna.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContentsEdit" class="textarea" id="ansContentsEdit" rows="20" cols="100"></textarea><input type="button" value="�亯�ϱ�" class="button" onclick="fnQnareplyEdit();"></td>
</tr>
</form>
</table>
<% SET oMyqna = nothing %>
<!-- ################################### �亯 ���� Ŭ���� ������ �� �� ################################### -->
<!-- ######################################## ���� ���� ��� ���� ######################################## -->
<%
SET oMyqna = new CQna
	oMyqna.FCurrPage = 1
	oMyqna.FPageSize = 200
	oMyqna.FRectUserid = masterQRegID
	oMyqna.getUserQnAList

If oMyqna.FResultCount > 0 Then
%>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="35" align="left" bgcolor="BLACK">
	<td colspan="6"><font color="WHITE"><%=masterQRegID%> ȸ���� ���� ���� ���</font></td>
</tr>
<tr height="35" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">��ȣ</td>
	<td width="250">���Ǻо�</td>
	<td width="80">����</td>
	<td>����</td>
	<td width="140">�����</td>
	<td width="140">�����</td>
</tr>
<% For i=0 to oMyqna.FResultCount - 1 %>
<tr height="30" style="cursor:pointer;" align="center" bgcolor='#FFFFFF'" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff'; onclick="goView('<%= oMyqna.FItemList(i).FIdx %>','<%= oMyqna.FItemList(i).FReply_group_idx %>')">
	<td align="center"><%= oMyqna.FItemList(i).FIdx %></td>
	<td align="center"><%= oMyqna.FItemList(i).getQnaGubunName %></td>
	<td align="center"><%= oMyqna.FItemList(i).getAnswerName %></td>
	<td align="left"><%= oMyqna.FItemList(i).FTitle %></td>
	<td align="center"><%= oMyqna.FItemList(i).FUserid %></td>
	<td align="center"><%= FormatDate(oMyqna.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<% Next %>
</table>
<% End If %>
<% SET oMyqna = nothing %>
<!-- ################################### ���� ���� ��� �� ################################### -->
<script>
$(function(){
	chgCont("H999", "<%=masterQRegID%>")
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
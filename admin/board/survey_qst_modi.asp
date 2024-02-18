<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : ������ ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim lp, srv_sn, qst_sn
	srv_sn = requestCheckVar(getNumeric(request("ssn")),10)
	qst_sn = Request("qsn")

	'// �������� ���
	dim oSurveyQuestion
	Set oSurveyQuestion = new CSurvey
	oSurveyQuestion.FRectSn = qst_sn

	oSurveyQuestion.GetSurveyQuestCont

	if oSurveyQuestion.FResultCount<=0 then
		Call Alert_Close("�߸��� ��ȣ�̰ų� ������ �����Դϴ�.")
		dbget.Close: Response.End
	end if

	'//�������϶� ���� ����
	dim oSurveyPoll
	Set oSurveyPoll = new CSurvey
	oSurveyPoll.FRectSn = srv_sn
	oSurveyPoll.FRectqstSn = qst_sn

	if oSurveyQuestion.FItemList(1).Fqst_type="1" then	
		oSurveyPoll.GetSurveyPollList
	end if
%>
<script type='text/javascript'>
<!--
	function chgQstType(tp) {
		if(tp=="1") {
			document.getElementById("trQstPoll").style.display="";
			document.getElementById("trQstDiv").style.display="none";
		} else if(tp=="9") {
			document.getElementById("trQstPoll").style.display="none";
			document.getElementById("trQstDiv").style.display="";
		} else {
			document.getElementById("trQstPoll").style.display="none";
			document.getElementById("trQstDiv").style.display="none";
		}
	}

	var total_link = <%=oSurveyPoll.FResultCount+1%>;
	function fnAddPoll() {
		var oRow1 = tbl_poll.insertRow();
		var oRow2 = tbl_poll.insertRow();
		oRow1.style.backgroundColor="#FFFFFF";
		oRow1.style.textAlign="center";
		oRow2.style.backgroundColor="#FFFFFF";
		
		var oCell1 = oRow1.insertCell();
			oCell1.rowSpan = 2;
		var oCell2 = oRow1.insertCell();
			oCell2.colSpan = 2;
		var oCell3 = oRow2.insertCell();
		var oCell4 = oRow2.insertCell();
		
		oCell1.innerHTML = '���� #'+total_link + '<input type="hidden" name="poll_sn" value="" />';
		oCell2.innerHTML = '<textarea name="poll_content" class="textarea" style="width:100%; height:32px;"></textarea>';
		oCell3.innerHTML = '�߰��ǰ� <select name="poll_isAddAnswer" class="select"><option value="N" selected >����</option><option value="Y">����</option></select>';
		oCell4.innerHTML = '���ù��� ��ȣ : <input type="text" name="link_qst_sn" size="4" class="text">';

		total_link++;
	}

	//�� ����
	function fnQstSubmit() {
		var frm = document.frm_Qst;
		if(!frm.qst_type.value) {
			alert("���� ���¸� �������ּ���.");
			frm.qst_type.focus();
			return;
		}

		if(frm.qst_content.value.length<2) {
			alert("���� ������ �ۼ����ּ���.");
			frm.qst_content.focus();
			return;
		}

		// ������ ���� Ȯ��
		if(frm.qst_type.value=="1") {
			var chkPollCnt=0;
			for(var i=0;i<frm.poll_content.length;i++) {
				if(frm.poll_content[i].value) chkPollCnt++;
			}
			if(chkPollCnt<2) {
				alert("������ ������ �Է����ּ���.\n�������� �ּ� 2���̻� ����ؾߵ˴ϴ�.");
				return;
			}
		}

		if(confirm("�Է��� �������� ������ �����Ͻðڽ��ϱ�?")) {
			frm.submit();
		} else {
			return;
		}
	}
//-->
</script>
<!-- �Է����̺� ���� -->
<form name="frm_Qst" method="POST" action="survey_qst_process.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="qModi" />
<input type="hidden" name="srv_sn" value="<%=srv_sn%>" />
<input type="hidden" name="qst_sn" value="<%=qst_sn%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="#DDDDFF" align="left"><img src="/images/icon_star.gif" align="absmiddle"><b>���� ���� ����</b></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="20%" bgcolor="#EEEEEE">���� ��ȣ</td>
	<td width="80%" align="left"><b><%=srv_sn%></b></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="20%" bgcolor="#EEEEEE">���� ��ȣ</td>
	<td width="80%" align="left"><b><%=qst_sn%></b></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="#EEEEEE">��������</td>
	<td align="left">
		<select name="qst_type" class="select" onchange="chgQstType(this.value)">
			<option value="">::���¼���::</option>
			<option value="1">������</option>
			<option value="2">�ְ���</option>
			<option value="3">�ܴ���</option>
			<option value="9">������</option>
		</select>
		<script type='text/javascript'>
		frm_Qst.qst_type.value="<%=oSurveyQuestion.FItemList(1).Fqst_type%>";
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="#EEEEEE">���� ����</td>
	<td align="left">
		<textarea name="qst_content" class="textarea" style="width:100%; height:50px;"><%= ReplaceBracket(oSurveyQuestion.FItemList(1).Fqst_content) %></textarea>
	</td>
</tr>
<tr id="trQstPoll" align="center" bgcolor="#FFFFFF" style="display:<%=chkIIF(oSurveyQuestion.FItemList(1).Fqst_type="1","","none")%>;">
	<td bgcolor="#EEEEEE">
		����<br>
		<span style="cursor:pointer" onclick="fnAddPoll()">[�����߰�]</span>
	</td>
	<td align="left">
		<table width="100%" id="tbl_poll" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<%	
			if oSurveyPoll.FResultCount>0 then
				for lp=0 to (oSurveyPoll.FResultCount-1)
		%>
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="50" rowspan="2">
				���� #<%=lp+1%>
				<input type="hidden" name="poll_sn" value="<%=oSurveyPoll.FItemList(lp).Fpoll_sn%>" />
			</td>
			<td colspan="2">
				<textarea name="poll_content" class="textarea" style="width:100%; height:32px;"><%= ReplaceBracket(oSurveyPoll.FItemList(lp).Fpoll_content) %></textarea>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td>
				�߰��ǰ�
				<select name="poll_isAddAnswer" class="select">
					<option value="N" <%=chkIIF(oSurveyPoll.FItemList(lp).Fpoll_isAddAnswer="N","selected","")%>>����</option>
					<option value="Y" <%=chkIIF(oSurveyPoll.FItemList(lp).Fpoll_isAddAnswer="Y","selected","")%>>����</option>
				</select>
			</td>
			<td>���ù��� ��ȣ : <input type="text" name="link_qst_sn" size="4" class="text" value="<%=oSurveyPoll.FItemList(lp).Flink_qst_sn%>"></td>
		</tr>
		<%
				next
			end if
		%>
		</table>
	</td>
</tr>
<tr id="trQstDiv" align="center" bgcolor="#FFFFFF" style="display:none;">
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td align="left">�� �����ڴ� ������ �ƴմϴ�. ���׵��� �׷��� ��� ������ �ִ� �׸��Դϴ�.</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="#EEEEEE">�ʼ�����</td>
	<td align="left">
		<label><input type="radio" name="qst_isNull" value="N" <%=chkIIF(oSurveyQuestion.FItemList(1).Fqst_isNull="N","checked","")%> /> �亯�ʼ�</label>
		<label><input type="radio" name="qst_isNull" value="Y" <%=chkIIF(oSurveyQuestion.FItemList(1).Fqst_isNull="Y","checked","")%> /> �������</label>
	</td>
</tr>
</table>
</form>
<!-- �Է����̺� �� -->
<!-- ���׾׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="���׼���" onClick="fnQstSubmit()"></td>
</tr>
</table>
<!-- ���׾׼� �� -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
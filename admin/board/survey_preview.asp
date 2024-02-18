<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim page, lp, qstNo, using, strType, strDel, srv_sn

	srv_sn = Request("sn")

	'�⺻�� ����
	page=1
	using="Y"

	'// �������� ����
	dim oSurveyMaster
	Set oSurveyMaster = new CSurvey

	oSurveyMaster.FRectSn = srv_sn
	
	oSurveyMaster.GetSurveyCont

	'// �������� ���
	dim oSurveyQuestion
	Set oSurveyQuestion = new CSurvey

	oSurveyQuestion.FRectSn = srv_sn
	oSurveyQuestion.FPagesize = 100
	oSurveyQuestion.FCurrPage = page
	oSurveyQuestion.FRectUsing = using
	oSurveyQuestion.FRectOrder = "asc"

	oSurveyQuestion.GetSurveyQstList
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<title><%=oSurveyMaster.FitemList(1).Fsrv_subject%></title>
<style type="text/css">
<!--
body,table,tr,td {font-family: ����,AppleGothic,sans-serif; color:#888888; font-size:11px; word-spacing: -2px;scrollbar-face-color: F2F2F2; scrollbar-shadow-color:#bbbbbb; scrollbar-highlight-color: #bbbbbb; scrollbar-3dlight-color: #FFFFFF; scrollbar-darkshadow-color: #FFFFFF; scrollbar-track-color: #F2F2F2; scrollbar-arrow-color: #bbbbbb; scrollbar-base-color:#E9E8E8;}
td {word-break:break-all;}
img,table {border:0px;}
b {letter-spacing:-1px;}
input {padding-top:3px; height:21px;}
textarea {line-height:18px; padding:3px;}

.redtitle{font-family: ����; FONT-SIZE: 12px; COLOR: #c3080a; font-weight:bold;}
.graytext{font-family: ����; FONT-SIZE: 12px; COLOR: #888888; font-weight:bold;}
.grayNomal{font-family: ����; FONT-SIZE: 12px; COLOR: #888888;}
.input_text {border:1px #cccccc solid; FONT-FAMILY: "����"; font-size: 12px; color="#888888"; padding:1px;}
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>
<script language="javascript">
<!--
	function chkPollAdd(pollSn,dsp)
	{
		// �߰��亯 ���� �˻�
		if(document.all["addPoll"+pollSn])
		{
			if(dsp=='Y')
				document.all["addPoll"+pollSn].style.display="";
			else
				document.all["addPoll"+pollSn].style.display="none";
		}
	}
//-->
</script>
</head>
<body>
<!-- // �ٵ� ���� // -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<!-- // �Ӹ��� // -->
	<td><%=db2html(oSurveyMaster.FitemList(1).Fsrv_head)%></td>
</tr>
<%
	if oSurveyQuestion.FResultCount>0 then
%>
<form name="frmSurvey">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td style="padding:0 20 20 20;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			'## ���� ������
			qstNo = 1
			for lp=0 to oSurveyQuestion.FResultCount - 1
			
				'//���� ���к� ���
				Select Case oSurveyQuestion.FitemList(lp).Fqst_type
					Case "1"	'������
		%>
				<tr>
					<td style="padding:0 20px 0 20px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td style="padding:20px 0 20px 0">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td class="graytext" style="padding-bottom:8px;"><%=qstNo & ". " & oSurveyQuestion.FitemList(lp).Fqst_content %></td>
								</tr>
								<tr>
									<td><%=oSurveyQuestion.PrintSurveyPollList(oSurveyQuestion.FitemList(lp).Fqst_sn)%></td>
								</tr>
								</table>
							</td>
						</tr>
						<tr height="1"><td height="1" bgcolor="#dddddd"></td></tr>
						</table>
					</td>
				</tr>
		<%
					Case "2"	'�ְ���
		%>
				<tr>
					<td style="padding:0 20px 0 20px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td style="padding:20px 0 20px 0">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td class="graytext" style="padding-bottom:8px;"><%=qstNo & ". " & oSurveyQuestion.FitemList(lp).Fqst_content %></td>
								</tr>
								<tr>
									<td class="graytext"><textarea name="textfield5" class="input_text" style="width:100%;height:120px;"></textarea></td>
								</tr>
								</table>
							</td>
						</tr>
						<tr height="1"><td height="1" bgcolor="#dddddd"></td></tr>
						</table>
					</td>
				</tr>
		<%
					Case "3"	'�ܴ���
		%>
				<tr>
					<td style="padding:0 20px 0 20px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td style="padding:20px 0 20px 0">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td class="graytext" style="padding-bottom:8px;"><%=qstNo & ". " & oSurveyQuestion.FitemList(lp).Fqst_content %></td>
								</tr>
								<tr>
									<td class="graytext" style="padding-left:15px;"><input type="text" name="textfield5" class="input_text" style='width:450px;height:16px;'></td>
								</tr>
								</table>
							</td>
						</tr>
						<tr height="1"><td height="1" bgcolor="#dddddd"></td></tr>
						</table>
					</td>
				</tr>
		<%
					Case "9"	'���ױ���
						'���׹�ȣ���� ����
						qstNo = qstNo - 1
		%>
				<tr>
					<td class="redtitle" style="padding:20 0 0 0">[<%=oSurveyQuestion.FitemList(lp).Fqst_content%>]</td>
				</tr>
		<%
				End Select
				qstNo = qstNo + 1
			Next
		%>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" style="padding-bottom:20px;"><img src="http://fiximage.10x10.co.kr/web2008/etc/2008poll/poll_btn2.jpg" width="189" height="47" border="0"/></td>
</tr>
</form>
<% end if %>
<tr>
	<!-- // ������ // -->
	<td><%=db2html(oSurveyMaster.FitemList(1).Fsrv_tail)%></td>
</tr>
</table>
<!-- // �ٵ� �� // -->
</body>
</html>
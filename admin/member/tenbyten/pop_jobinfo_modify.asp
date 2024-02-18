<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenJobInfoCls.asp" -->
<%
	dim job_sn, strTitle
	dim job_name, job_isDel
	job_sn = Request("job_sn")



	dim oJob
	Set oJob = new CTenByTenJob

	oJob.FRectjob_sn = job_sn



	if job_sn<>"" then
		oJob.GetInfo
		strTitle = "직책정보 수정"
		if oJob.FResultCount>0 then
			job_name = oJob.FitemList(1).Fjob_name
			job_isDel = oJob.FitemList(1).Fjob_isDel
		end if
	else
		strTitle = "직책정보 등록"
	end if
%>
<html>
<head>
<title><%=strTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">
<!--
	function chk_form(form)
	{
		if(form.job_name.value.length<1)
		{
			alert("직책명을 입력해주십시요.");
			form.job_name.focus();
			return false;
		}
		else
		{
			form.action="tenbyten_job_process.asp";
			<% if job_sn<>"" then %>
			form.mode.value = "modi";
			<% else %>
			form.mode.value = "add";
			<% end if %>
			return true;
		}
	}

	function chk_modi(md)
	{
		var ms, form = document.frm_jobInfo;

		if(md=='Y') ms="삭제";
		else ms="복구";

		if(confirm("[<%=job_name%>]직책을 " + ms + "하시겠습니까?"))
		{
			form.action="tenbyten_job_process.asp";
			form.mode.value = "del";
			form.job_isDel.value = md;
			form.submit();
		}
	}
//-->
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm_jobInfo" method="POST" action="" onsubmit="return chk_form(this)">
<input type="hidden" name="mode" value="">
<input type="hidden" name="job_isDel" value="<%=job_isDel%>">
<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="bottom" bgcolor="F4F4F4">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" bgcolor="F4F4F4"><b><%=strTitle%></b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="350" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<% if job_sn<>"" then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">직책번호</td>
	<td>
		<b><%=job_sn%></b>
		<input type="hidden" name="job_sn" value="<%=job_sn%>">
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">직책명</td>
	<td>
		<input type="text" name="job_name" size="20" maxlength="60" value="<%=job_name%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2">
		<input type="submit" value="확인">
		<%
			if job_sn<>"" then
				if job_isDel="N" then
		%>
		<input type="button" value="삭제" onClick="chk_modi('Y')">
		<%		else %>
		<input type="button" value="복구" onClick="chk_modi('N')">
		<%
				end if
			end if
		%>
		<input type="button" value="취소" onClick="self.close()">
	</td>
</tr>
</table><br>
</form>
</body>
</html>
<% Set oJob = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
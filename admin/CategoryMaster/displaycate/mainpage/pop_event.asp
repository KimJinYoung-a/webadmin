<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear, vEventID, vCateCode, vType, vPage, vStartDate
vType = Request.Querystring("type")
vEventID = Request.Querystring("eventid")
If vEventID = "0" Then
	vEventID = ""
End IF
vCateCode = Request.Querystring("catecode")
vPage = Request.Querystring("page")
vStartDate = Request.Querystring("startdate")


%>
<script language="javascript">
<!--
document.domain = "10x10.co.kr";

	function jsUpload(){
		if(!document.frmImg.eventid.value){
			alert("�̺�Ʈ�ڵ�� �� �־��ּ���.");
			document.frmImg.eventid.focus();
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> <b><%=vType%></b> �̺�Ʈ���</div>
<table width="360" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="pop_event_proc.asp" onSubmit="return jsUpload();">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="type" value="<%=vType%>">
<input type="hidden" name="page" value="<%=vPage%>">
<input type="hidden" name="startdate" value="<%=vStartDate%>">
	<tr>
		<td bgcolor="#FFFFFF0" colspan="2">
			* <b><font color="red">[�ʵ�]</font></b><b>��� ��</b> �ش� �̺�Ʈ�� <b><font color="blue-green">����, �̹���, ī��, Ÿ��, ��ũ�� ������ �Ǹ�</font></b> �� �˾�â���� <b><font color="blue">�ٽ� Ȯ�� ��ư�� ����</font></b>�ּž� <b><font color="green">����� �������� ����</font></b>�˴ϴ�.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="eventid" value="<%=vEventID%>">
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
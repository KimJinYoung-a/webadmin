<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/etc/between/project/pop_project_uploadimg.asp?yr=&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsPjtSubmit(frm){
	if(frm.pjt_name.value==""){
		alert('��ȹ�� ���� �Է��ϼ���');
		frm.pjt_name.focus();
		return false;
	}
	if(frm.pjt_kind.value==""){
		alert('��ȹ�� ������ �����ϼ���');
		frm.pjt_gender.focus();
		return false;
	}
	if(frm.pjt_gender.value==""){
		alert('������ �����ϼ���');
		frm.pjt_gender.focus();
		return false;
	}
}
</script>
<script type="text/javascript">
$(function(){
	//�÷���Ŀ
	$("input[name='pjt_BGColor']").colorpicker();
});
</script>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<form name="frmPjt" method="post"  action="project_process.asp" onSubmit="return jsPjtSubmit(this);" style="margin:0px;">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr>
	<td> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��ȹ�� ���� ��� </td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<col width="150" />
		<col  />
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ȹ����</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<input type="text" name="pjt_name" size="60" maxlength="60" value="">
	   		</td>
	   	</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ȹ�� ����</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<% sbGetOptProjectCodeValue "pjt_kind","","" %>
	   		</td>
	   	</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
			<td bgcolor="#FFFFFF">
	   			<select name="pjt_gender" class="select">
	   				<option>- Choice -</option>
	   				<option value="A">��ü</option>
	   				<option value="M">����</option>
	   				<option value="F">����</option>
	   			</select>
			</td>
		</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<select class="select" name="pjt_state">
	   				<option value="0">��ϴ��</option>
	   				<option value="7">����</option>
	   			</select>
	   		</td>
	   	</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ǰ���Ĺ��</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<select name="pjt_sortType" class="select">
	   				<option value="1">�Ż�ǰ��</option>
	   				<option value="2">�����ݼ�</option>
	   				<option value="3">������ȣ��</option>
	   				<option value="4">����Ʈ������</option>
	   				<option value="5">���ݼ�</option>
	   			</select>
	   		</td>
	   	</tr>
	   	<tr>
			<td colspan="2" height="40" align="right"  bgcolor="#FFFFFF">
				<input type="image" src="/images/icon_save.gif">
				<a href="project_list.asp?menupos=<%=menupos%>"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

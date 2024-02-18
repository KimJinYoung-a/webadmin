<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim mode, pjt_code
Dim opjt
mode		= request("mode")
pjt_code	= request("pjt_code")

SET opjt = new cProject
	opjt.FRectPjt_code = pjt_code
	opjt.getProjectCont()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javaScript" src="/js/jquery.iframe-auto-height.js"></script>
<script type="text/javascript">
$(function(){
	//�÷���Ŀ
	$("input[name='pjt_BGColor']").colorpicker();
});
</script>
<script language="javascript">
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/etc/between/project/pop_project_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
function jsPjtSubmit(frm){
	if(frm.pjt_name.value==""){
		alert('��ȹ�� ���� �Է��ϼ���');
		frm.pjt_name.focus();
		return false;
	}

	if(frm.pjt_gender.value==""){
		alert('������ �����ϼ���');
		frm.pjt_gender.focus();
		return false;
	}
}
</script>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<form name="frmPjt" method="post"  action="project_process.asp" onSubmit="return jsPjtSubmit(this);" style="margin:0px;">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="ban" value="<%=opjt.FItemList(0).FPjt_topImgUrl%>">
<input type="hidden" name="pjt_kind" value="<%=opjt.FItemList(0).FPjt_kind%>">
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
	   			<input type="text" name="pjt_name" size="60" maxlength="60" value="<%= opjt.FItemList(0).FPjt_name %>">
	   		</td>
	   	</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ȹ�� ����</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<%= getDBcodeByName(opjt.FItemList(0).FPjt_kind) %>
	   		</td>
	   	</tr>
		<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ȹ�� ��� �̹���</B></td>
	   		<td bgcolor="#FFFFFF">
	   		<input type="button" name="btnBan" value="�̹��� ���" onClick="jsSetImg('<%=pjt_code%>','<%=opjt.FItemList(0).FPjt_topImgUrl%>','ban','spanban')" class="button">
	   			<div id="spanban" style="padding: 5 5 5 5">
   				<% If opjt.FItemList(0).FPjt_topImgUrl <> "" Then %>
	   				<img src="<%=opjt.FItemList(0).FPjt_topImgUrl%>" border="0">
	   				<a href="javascript:jsDelImg('ban','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
   				<% End If %>
	   			</div>
	   		</td>
	   	</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
			<td bgcolor="#FFFFFF">
	   			<select name="pjt_gender" class="select">
	   				<option>- Choice -</option>
	   				<option value="A" <%= Chkiif(opjt.FItemList(0).FPjt_gender = "A", "selected", "") %> >��ü</option>
	   				<option value="M" <%= Chkiif(opjt.FItemList(0).FPjt_gender = "M", "selected", "") %> >����</option>
	   				<option value="F" <%= Chkiif(opjt.FItemList(0).FPjt_gender = "F", "selected", "") %> >����</option>
	   			</select>
			</td>
		</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>����</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<select class="select" name="pjt_state">
	   				<option value="0" <%= Chkiif(opjt.FItemList(0).FPjt_state = "0", "selected", "") %> >��ϴ��</option>
	   				<option value="7" <%= Chkiif(opjt.FItemList(0).FPjt_state = "7", "selected", "") %> >����</option>
	   				<option value="9" <%= Chkiif(opjt.FItemList(0).FPjt_state = "9", "selected", "") %> >����</option>
	   			</select>
	   		</td>
	   	</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>��ǰ���Ĺ��</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<select name="pjt_sortType" class="select">
	   				<option value="1" <%= Chkiif(opjt.FItemList(0).FPjt_sortType = "1", "selected", "") %> >�Ż�ǰ��</option>
	   				<option value="2" <%= Chkiif(opjt.FItemList(0).FPjt_sortType = "2", "selected", "") %> >�����ݼ�</option>
	   				<option value="3" <%= Chkiif(opjt.FItemList(0).FPjt_sortType = "3", "selected", "") %> >������ȣ��</option>
	   				<option value="4" <%= Chkiif(opjt.FItemList(0).FPjt_sortType = "4", "selected", "") %> >����Ʈ������</option>
	   				<option value="5" <%= Chkiif(opjt.FItemList(0).FPjt_sortType = "5", "selected", "") %> >���ݼ�</option>
	   			</select>
	   		</td>
	   	</tr>
	   	<tr>
	   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>�������</B></td>
	   		<td bgcolor="#FFFFFF">
	   			<input type="radio" name= "pjt_using" value= "Y" <%= Chkiif(opjt.FItemList(0).Fpjt_using = "Y", "checked", "") %> >Y
	   			<input type="radio" name= "pjt_using" value= "N" <%= Chkiif(opjt.FItemList(0).Fpjt_using = "N", "checked", "") %> >N
	   		</td>
	   	</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��ȹ�� �׷� ��� </td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	   	<tr>
	   		<td width="30" align="center" rowspan="2"  bgcolor="<%= adminColor("tabletop") %>"> ��<br>ȹ<br>��<br><br>��<br>��<br>��<br>��<br>��<br> </td>
	   		<td width="65" align="center"  bgcolor="<%= adminColor("tabletop") %>">ȭ�����ø�</td>
	   		<td bgcolor="#FFFFFF">�׷���</td>
	   	</tr>
	   	<tr>
	   		<td width="65" align="center"  bgcolor="<%= adminColor("tabletop") %>">�׷�<br>����Ʈ</td>
	   		<td bgcolor="#FFFFFF">
	   			<div id="divFrm3" style="display:;">
	   				<iframe id="iframG" src="iframe_projectitem_group.asp?pjt_code=<%= pjt_code %>" frameborder="0" width="100%" class="autoheight"></iframe>
	   			</div>
	   		</td>
		</tr>
	   	<tr>
			<td colspan="3" height="40" align="right"  bgcolor="#FFFFFF">
				<input type="image" src="/images/icon_save.gif">
				<a href="project_list.asp?menupos=<%=menupos%>"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<% SET opjt = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
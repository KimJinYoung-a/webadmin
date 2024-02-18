<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

Dim i
Dim iMaxLength : iMaxLength = 20
Dim isall : isall="A"

Dim sqlStr, ArrRows

sqlStr = "select 'MAP_CURR' as tbnm, count_BIG(*) CNT, MAX(lastupdate) as lastupdt from [db_analyze_etc].[dbo].[tbl_nvshop_mapItem] WITH(NOLOCK)"
sqlStr = sqlStr + " union ALL"
sqlStr = sqlStr + " select 'MAP_PRE', count_BIG(*), MAX(lastupdate) as lastupdt from [db_analyze_etc].[dbo].[tbl_nvshop_mapItem_Pre] WITH(NOLOCK)"
sqlStr = sqlStr + " union ALL"
sqlStr = sqlStr + " select 'MIMAP_ALL', count_BIG(*), MAX(lastupdate) as lastupdt from [db_analyze_etc].[dbo].[tbl_nvshop_mapItem_ALL] WITH(NOLOCK);"
rsAnalget.CursorLocation = adUseClient
rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly,adLockReadOnly
If Not (rsAnalget.Eof or rsAnalget.bof ) Then			 
	ArrRows = rsAnalget.getRows()
end if
rsAnalget.close

 
%>
<script language="javascript">
function fnChkFile(sFile, sMaxSize, arrExt){
	//���� ���ε� ����Ȯ��
	if (!sFile){
		return true;
	}
	var blnResult = false;
	var maxsize = sMaxSize * 1024 * 1024;
	
	//���� Ȯ���� Ȯ��
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();
	for (var i = 0; i < arrExt.length; i++){
		if (arrExt[i].toLowerCase() == fExet){
			blnResult =  true;
		}
	}
	return blnResult;
}

function jsChkNull(type,obj,msg){
	switch (type) {
		// text, password, textarea, hidden
		case "text" :
		case "password" :
		case "textarea" :
		case "hidden" :
		if (jsChkBlank(obj.value)) {
			alert(msg);
			//obj.focus();
			return false;
		}else {
			return true;
		}
		break;
		// checkbox
		case "checkbox" :
		if (!obj.checked) {
			alert(msg);
			return false;
		}else {
			return true;
		}
		break;
		// radiobutton
		case "radio" :
		var objlen = obj.length;
		for (i=0; i < objlen; i++) {
			if (obj[i].checked == true)
			return true;
		}
		if (i == objlen) {
			alert(msg);
			return false;
		}else{
			return true;
		}
		break;
		// ���ڰ˻�
		case "numeric" :
		if (!jsChkNumber(obj.value)||jsChkBlank(obj.value)) {
			alert(msg);
			return false;
		}else {
			return true;
		}
		break;
	}
	
	// select list
	if (obj.type.indexOf("select") != -1) {
		if (obj.options[obj.selectedIndex].value == 0 || obj.options[obj.selectedIndex].value == ""){
			alert(msg);
			return false;
		}else{
			return true;
		}
	}
	return true;
}

function XLSumbit(){
	var frm = document.frmFile;
	var uploadFile = document.getElementById("ifrmtg")

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length]  = "zip";

	//���� Ȯ��
	if(!jsChkNull("text",frm.sFile,"������ �Է��Ͻʽÿ�.")){
		frm.sFile.focus();
		return;
	}

	//������ȿ�� üũ
	if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
		alert("������ <%=iMaxLength%>MB������ xls���ϸ� ���ε� �����մϴ�.");
		return;
	}

	frm.target = "ifrmtg";

	if ( confirm('����Ͻðڽ��ϱ�?') ){
		frm.submit();
	}
//	frm.reset();
}

function execFileArr(subfldr,monthFolder,ifilelist){
	var ifiles = ifilelist.split("|");
	var frm = document.frmAct;

	for (var i=0;i<ifiles.length-1;i++){
		if (ifiles[i].length>0){
			setTimeout(function(isubfldr,imonthFolder,ifilename){ 
				frm.target = "ifrmtgtodb";
				frm.subfldr.value = isubfldr;
				frm.monthFolder.value = imonthFolder;
				frm.filename.value = ifilename;
				frm.submit();
			}, 5000*i,subfldr,monthFolder,ifiles[i]);
		}
	}


}


</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmFile" method="post" action="/admin/etc/naverEp/procFileUploadNvMapItem.asp"  enctype="MULTIPART/FORM-DATA"  >
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ϸ�:</td>
	<td align="left" colspan="2">
		<input type="file" name="sFile" id="sFile" class="file" >
		&nbsp;&nbsp;&nbsp;
		1.������ <font color="red">xlsx zip ����</font>���� ��.<br>
	</td>
</tr> 
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ο���</td>
	<td align="left" colspan="2">
		<input type="radio" name="isall" value="M" <%=CHKIIF(isall="M","checked","")%> >���οϷ��ǰ
		<input type="radio" name="isall" value="A" <%=CHKIIF(isall="A","checked","")%> >��ü
	</td>
	
</tr>
</form>
<% if isArray(ArrRows) then %>
<% For i =0 To UBound(ArrRows,2) %>
<tr bgcolor="#FFFFFF">
	<td><%=ArrRows(0,i)%></td>
	<td><%=FormatNumber(ArrRows(1,i),0)%></td>
	<td><%=ArrRows(2,i)%></td>
</tr>
<% Next %>
<% end if %>

<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="3">
	    <input type="button" class="button" value="���" onClick="XLSumbit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</table>

<!--
<form name="frmAct" method="post" action="/admin/etc/naverEp/procFileUploadNvMapFileToDB.asp" >
<input type="hidden" name="monthFolder" value="">
<input type="hidden" name="subfldr" value="">
<input type="hidden" name="filename" value="">
</form>

<iframe id="ifrmtgtodb" name="ifrmtgtodb" frameborder="1" width="1100" height="280"></iframe>
<br>
-->
<iframe id="ifrmtg" name="ifrmtg" frameborder="1" width="1100" height="1100"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

Dim  iMaxLength
IF iMaxLength = "" THEN iMaxLength = 10

Dim partnercompanyid  : partnercompanyid = requestCheckVar("partnercompanyid",32)


%>
<script language="javascript">
function XLSumbit(){
	var frm = document.frmFile;

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length]  = "xls";

	if (frm.tplcompanyid.value == '') {
		alert('���縦 �����ϼ���.');
		return;
	}

	if (frm.xltype.value == '') {
		alert('���������� �����ϼ���.');
		return;
	}

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

	frm.submit();
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
            }
            else {
                return true;
            }
            break;
        // checkbox
        case "checkbox" :
            if (!obj.checked) {
				alert(msg);
                return false;
            }
            else {
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
            }
            else {
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

function fnChkFile(sFile, sMaxSize, arrExt){
    //���� ���ε� ����Ȯ��
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //���� �뷮 Ȯ��
    var maxsize = sMaxSize * 1024 * 1024;

     //	var img = new Image();
    //	img.dynsrc = sFile;
    //var fSize = img.fileSize ;
    	//if (fSize > maxsize){
    		//alert("����ũ��� "+sMaxSize+"MB���ϸ� �����մϴ�.");
    		//return false;
    	//}

    //���� Ȯ���� Ȯ��
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();

	for (var i = 0; i < arrExt.length; i++)
	   	{
	    	if (arrExt[i].toLowerCase() == fExet)
	    	{
	   			blnResult =  true;
	   		}
		}

	return blnResult;
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left" bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="2">
			<b>1. �ֹ�����Ÿ �������</b>
		</td>
	</tr>
	<form name="frmFile" method="post" action="procFileUpload.asp"  enctype="MULTIPART/FORM-DATA">
	<input type="hidden" name="iML" value="<%=iMaxLength%>">
	<tr bgcolor="#FFFFFF">
		<td width="80" align="right" bgcolor="<%= adminColor("tabletop") %>"> ���� :</td>
		<td>
			<% CALL drawPartner3plCompany("tplcompanyid","","") %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" align="right" bgcolor="<%= adminColor("tabletop") %>"> �������� :</td>
		<td>
			<select class="select" name="xltype">
				<option></option>
				<option value="sabangnet">���� �ֹ�</option>
				<option value="default">�⺻ ����</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" align="right" bgcolor="<%= adminColor("tabletop") %>">���ϸ�:</td>
		<td><input type="file" name="sFile" ></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
		    <input type="button" class="button" value="���" onClick="XLSumbit();">
		    <input type="button" class="button" value="���" onClick="self.close();">
		</td>
	</tr>

	</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹ� ��� ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim iMaxLength	,sellsite
	sellsite = requestCheckVar("sellsite",32)

IF iMaxLength = "" THEN iMaxLength = 10
%>

<script language="javascript">

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

function XLSumbit(){
	var frm = document.frmFile;
    if (frm.sellsite.value.length<1){
        alert('���޸��� �����ϼ���.');
        frm.sellsite.focus();
        return;
    }
    
    /*
    if (frm.sellsite.value=="interpark"){
        alert('������ũ ������');
        return;
    }
    */

	if (frm.sellsite.value=="gmarket1010"){
		if (!confirm('Gmarket XL ��Ͻ� ���Ǹűݾ� ��꿡 ������ �ֽ��ϴ�. �׷��� ����Ͻðڽ��ϱ�? ��� �� �����ڿ��� �˷��ּ���.')){
			return;
		}
	}

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length]  = "xls";

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

</script>

<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>1. �ֹ�����Ÿ �������</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmFile" method="post" action="procFileUpload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���θ� ���� :</td>
	<td align="left">
		<% call drawSelectBoxXSiteOrderInputPartner("sellsite", sellsite) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ϸ�:</td>
	<td align="left">
		<input type="file" name="sFile" class="file">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2">
	    <input type="button" class="button" value="���" onClick="XLSumbit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</form>
</table>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=70>
		���ε����
	</td>
	<td align="left">
		* ������ũ ����: �߼۰���>�ֹ��������ٿ�ε�(������������)
		<Br>
		<% get_xsite_excel_order_sample() %>
	</td>
</tr>
</table>
<br/>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=70>
		���� ����
	</td>
	<td align="left">
		*���� ������ ���� �������� �ؽ�Ʈ�� �Ǿ��ֽ��ϴ�.<br />
		���� �������� ���ð� ���õ����ʹ�� �Է��Ͻø� �˴ϴ�.<br />
		���� �۾��� �ʼ��Է� ���̰�, ���� �۾��� �ɼ��Դϴ�.<br/> 
	<%
		response.write "<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/Common_Sample.xls' onfocus='this.blur()'>"
		response.write "<font color='red'>*���� ���</font></a>"
	%>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'###########################################################
' Description : ��ǰ ����ɼ� ���
' History : 2013.12.16 ������ �ɼǰ��� ����   
'###########################################################
%>
<%
dim i,iRowMax
iRowMax = 19 '�ɼ� �ִ밹��
%>

<script language="javascript">
<!--
function AddOption()
{
	var frm = document.itemopt;
    var addedCnt = 0;
    
	if(!frm.optTypeNm.value){
		alert("�߰��� �ɼ� ���� ���� �Է����ֽʽÿ�.");
		frm.optTypeNm.focus();
		return false;
	}

    for (var i=0;i<frm.optNm.length;i++){
        if (frm.optNm[i].value.length>0){
            opener.InsertOptionWithGubun(frm.optTypeNm.value, frm.optNm[i].value, "0000");
            addedCnt++;
        }
    }

    if (addedCnt>0){
	    self.close();
	}else{
	    alert('�߰��� �ɼ��� �Է��� �ּ���.');
	}
}
//-->
</script>
<body onload="window.resizeTo(550,890);document.itemopt.optTypeNm.focus();">
<table width="500" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
<form name="itemopt" >
    <tr height="30" bgcolor="#DDDDFF">
		<td width="120" align="center">�ɼ� ���� ��</td>
		<td bgcolor="#FFFFFF" align="left"><input type="text" name="optTypeNm" size="20" maxlength="20"> ����</td>
	</tr>
	<% for i=0 to iRowMax %>
	<tr height="30" bgcolor="#DDDDFF">
		<td width="120" align="center">�ɼ� �� <%= i+1 %></td>
		<td bgcolor="#FFFFFF" align="left"><input type="text" name="optNm" size="32" maxlength="20"> <%= chkIIF(i=0,"����","") %><%= chkIIF(i=1,"�Ķ�","") %><%= chkIIF(i=2,"���","") %></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" value="�ɼ��߰�" class="button" onClick="AddOption();">
			<input type="button" value=" �� �� "  class="button" onclick="self.close()">
		</td>
	</tr>
</form> 
</table>
<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2013.12.16 </div>
</body>
<!-- #include virtual="/admin/lib/poptail.asp"-->
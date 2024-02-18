<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������ǰ�����
' History : 2018.01.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/safetycert/safetycert_cls.asp"-->
<%
dim i, identikey

dim osafety
set osafety = new Csafetycert
	osafety.FPageSize = 1
	osafety.FCurrPage = 300
	osafety.fsafetycert
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function chkAllItem() {
	if($("input[name='infoDiv']:first").attr("checked")=="checked") {
		$("input[name='infoDiv']").attr("checked",false);
	} else {
		$("input[name='infoDiv']").attr("checked","checked");
	}
}

//����Ʈ ��ü ����
function savestandingList() {
	var chk=0;
	$("form[name='frmsafety']").find("input[name='infoDiv']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� �׸��� �������ּ���.");
		return;
	}

	var identikey;
	for (i=0; i< frmsafety.infoDiv.length; i++){
		if (frmsafety.infoDiv[i].checked == true){
			identikey = frmsafety.infoDiv[i].value;

			if (!eval("frmsafety.SafetyTargetY_" + identikey).checked && !eval("frmsafety.SafetyTargetN_" + identikey).checked){
				alert('����������󿩺θ� �����ϼ���');
				eval("frmsafety.SafetyTargetY_" + identikey).focus();
				return false;
			}else{
				if (eval("frmsafety.SafetyTargetY_" + identikey).checked == true && eval("frmsafety.SafetyTargetN_" + identikey).checked == true){
					eval("frmsafety.SafetyTargetYN_" + identikey).value = 'S'
				}else if (eval("frmsafety.SafetyTargetY_" + identikey).checked == true){
					eval("frmsafety.SafetyTargetYN_" + identikey).value = 'Y'
				}else{
					eval("frmsafety.SafetyTargetYN_" + identikey).value = 'N'
				}
			}

			if (eval("frmsafety.SafetyCertYN_" + identikey).value==""){
				alert('�����������θ� �����ϼ���');
				eval("frmsafety.SafetyCertYN_" + identikey).focus();
				return false;
			}

			if (eval("frmsafety.SafetyConfirmYN_" + identikey).value==""){
				alert('����Ȯ�ο��θ� �����ϼ���');
				eval("frmsafety.SafetyConfirmYN_" + identikey).focus();
				return false;
			}

			if (eval("frmsafety.SafetySupplyYN_" + identikey).value==""){
				alert('���������ռ����θ� �����ϼ���');
				eval("frmsafety.SafetySupplyYN_" + identikey).focus();
				return false;
			}

			if (eval("frmsafety.SafetyComply_" + identikey).value==""){
				alert('���������ؼ����θ� �����ϼ���');
				eval("frmsafety.SafetyComply_" + identikey).focus();
				return false;
			}
	    }
	}

	if(confirm("�����Ͻ� ����Ʈ ������ ���� �Ͻðڽ��ϱ�?")) {
		frmsafety.mode.value="safetylistedit";
		frmsafety.action="/admin/itemmaster/safetycert/safecert_process.asp";
		frmsafety.submit();
	}
}

function CheckClick(identikey){
	var f = document.frmsafety;
	var objStr = "infoDiv";
	var chk_flag = true;

	for(var i=0; i<f.infoDiv.length; i++){
		if(f.infoDiv[i].value==identikey){
			f.infoDiv[i].checked=true;
			break;
		}
	}
}

</script>

<form name="frmsafetyedit" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<form name="frmsafety" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type="button" onClick="savestandingList();" value="��������" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= osafety.FtotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=30><input type="button" value="��ü" class="button" onClick="chkAllItem();"></td>
    <td width=50>ǰ���ȣ</td>
    <td>ǰ���</td>
    <td width=120>�����������<Br>����</td>
    <td width=80>��������<Br>����</td>
    <td width=80>����Ȯ��<Br>����</td>
    <td width=80>���������ռ�<Br>����</td>
    <td width=80>���������ؼ�<Br>����</td>
    <td width=150>��������</td>
    <td width=80>���</td>
</tr>
<% if osafety.FtotalCount>0 then %>
<%
for i=0 to osafety.FResultCount - 1

identikey = osafety.FItemList(i).finfoDiv
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
    <td align="center"><input type="checkbox" name="infoDiv" value="<%= osafety.FItemList(i).finfoDiv %>" /></td>
    <td align="center">
    	<%= osafety.FItemList(i).finfoDiv %>
    </td>
    <td>
    	<%= osafety.FItemList(i).finfoDivName %>
    </td>
    <td align="center">
    	<input type="hidden" name="SafetyTargetYN_<%= identikey %>" value="<%= osafety.FItemList(i).fSafetyTargetYN %>">
    	<input type="checkbox" name="SafetyTargetY_<%= identikey %>" <% if osafety.FItemList(i).fSafetyTargetYN="Y" or osafety.FItemList(i).fSafetyTargetYN="S" then response.write " checked" %> onclick="CheckClick('<%= identikey %>');" >���
    	&nbsp;&nbsp;
    	<input type="checkbox" name="SafetyTargetN_<%= identikey %>" <% if osafety.FItemList(i).fSafetyTargetYN="N" or osafety.FItemList(i).fSafetyTargetYN="S" then response.write " checked" %> onclick="CheckClick('<%= identikey %>');" >����
    </td>
    <td align="center">
    	<% drawSelectBoxisusingYN "SafetyCertYN_"&identikey, osafety.FItemList(i).fSafetyCertYN, " onchange='CheckClick("""& identikey &""");'" %>
    </td>
    <td align="center">
    	<% drawSelectBoxisusingYN "SafetyConfirmYN_"&identikey, osafety.FItemList(i).fSafetyConfirmYN, " onchange='CheckClick("""& identikey &""");'" %>
    </td>
    <td align="center">
    	<% drawSelectBoxisusingYN "SafetySupplyYN_"&identikey, osafety.FItemList(i).fSafetySupplyYN, " onchange='CheckClick("""& identikey &""");'" %>
    </td>
    <td align="center">
    	<% drawSelectBoxisusingYN "SafetyComply_"&identikey, osafety.FItemList(i).fSafetyComply, " onchange='CheckClick("""& identikey &""");'" %>
    </td>
    <td align="center">
    	<%= osafety.FItemList(i).flastupdate %>
    	<% if osafety.FItemList(i).flastadminid <> "" then %>
    		<br><%= osafety.FItemList(i).flastadminid %>
    	<% end if %>
    </td>
    <td align="center">
    </td>
</tr>
<%
Next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">�˻������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>
</form>

<%
set osafety=nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
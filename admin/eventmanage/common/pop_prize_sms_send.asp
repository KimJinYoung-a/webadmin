<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event  
' History : 2008.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%

 Dim cEvtPrize
 Dim arrPrize, intLoop
 Dim eCode, egKindCode
  
 IF egKindCode = "" THEN egKindCode = 0	
	
	eCode 	= replace(Request("cksel")," ","")		 '���� ������ ��ȣ
if eCode <> "" then
	if checkNotValidHTML(eCode) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
'	Response.write eCode
'	Response.end
	set cEvtPrize = new ClsEventPrize
	cEvtPrize.FECode	  	= eCode			'�̺�Ʈ �ڵ�
	cEvtPrize.FEGKindCode 	= egKindCode	'�׷��ڵ�(�ΰŽ�,��ȭ�̺�Ʈ ȸ��)
	arrPrize = cEvtPrize.fnGetPrizeListUserInfo		'��÷����
	set cEvtPrize = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">
<!--
$(function(){
	$("#addcell").on("click",function(){
		if($("#addusercell").val()==""){
			alert("�߰��� ��ȭ��ȣ�� �Է��� �ּ���.");
		}else{
			var oRow;
			oRow += "<tr id='s" + (Number($("#cnt").val())+1) + "'>"
			oRow += "	<td bgcolor='#FFFFFF' align='center'>" + (Number($("#cnt").val())+1) + "</td>"
			oRow += "	<td bgcolor='#FFFFFF' align='center'><input type='text' name='usercell' value='" + $("#addusercell").val() + "'></td>"
			oRow += "	<td bgcolor='#FFFFFF' align='center'>�����Է�</td>"
			oRow += "	<td bgcolor='#FFFFFF' align='center'><a href=javascript:DelList('s" + (Number($("#cnt").val())+1) + "')>����</a></td>"
			oRow += "</tr>"
			//alert(oRow);
			$("#ucell").append(oRow);
			$("#cnt").val((Number($("#cnt").val())+1));
		}
	});
});
function SendSMS(){
	if($("#smstext").val()==""){
		alert("���� ������ �Է��ϼ���.");
	}else{
		$("#smstxt").val($("#smstext").val());
		document.frmPrize.submit();
	}
}
function updateChar() {
	var length = calculate_msglen(document.getElementById("smstext").value);

	if (length <= 80) {
		document.getElementById("charlen").innerHTML = "(" + length + "/80)<br><br>SMS";
	} else {
		document.getElementById("charlen").innerHTML = "(" + length + "/2000)<br><br><font color='red'>LMS</font>";
	}
}
function calculate_msglen(message) {
	var nbytes = 0;

	for (i=0; i<message.length; i++) {
		var ch = message.charAt(i);

		if(escape(ch).length > 4) {
			nbytes += 2;  // �ѱ��϶� 2�� ����
		} else if(ch == '\n') {
			if (message.charAt(i-1) != '\r') {
				nbytes += 1;  // Enter�϶� 1�� ����
			}
		} else {
			nbytes += 1;  // ��Ÿ ���ڵ��϶� 1�� ����
		}
	}

	return nbytes;
}

function DelList(obj){
	$("#ucell #"+ obj).remove();
}
//-->
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>SMS�߼�</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="post">
    <input type="hidden" name="mode" value="send">
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>" align="center">
    		���ڳ���<br>
    		<div id="charlen"></div>
    	</td>
    	<td bgcolor="#FFFFFF" colspan="3"><textarea name="smstext" id="smstext" class="textarea" cols="52" rows="10" onKeyUp="updateChar()"></textarea><div id="charlen"></div></td>
    </tr>
    </form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <input type="button" class="button" value="����" onclick="SendSMS(frm);">
			<input type="button" class="button" value="���" onclick="self.close();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="text" name="addusercell" id="addusercell">&nbsp;<input type="button" id="addcell" class="button" value="�߰�">
		�� ȸ�� ���� ���ŵ��ǿ��ο� ���þ��� ��� �߼� �˴ϴ�.
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frmPrize" method="post" action="/admin/eventmanage/common/prize_sms_send.asp">
<table width="100%" border="0" align="left" id="ucell" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">

<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="cnt" id="cnt" value="<%=UBound(arrPrize,2)+1%>">
<input type="hidden" name="smstxt" id="smstxt">
<tbody>
<tr>
	<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td align="center"  width="200" bgcolor="<%= adminColor("tabletop") %>">��ȭ��ȣ</td>							
	<td align="center"  width="100" bgcolor="<%= adminColor("tabletop") %>">�Է±���</td>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
</tr>

<%IF isArray(arrPrize) THEN%>
	<%For intLoop = 0 To UBound(arrPrize,2)	%>
<tr id="s<%=intLoop+1%>">
	<td bgcolor="#FFFFFF" align="center"><%=intLoop+1%></td>
	<td bgcolor="#FFFFFF" align="center"><input type='text' name='usercell' value='<%=arrPrize(0,intLoop)%>'></td>
	<td bgcolor="#FFFFFF" align="center"><%=arrPrize(1,intLoop)%></td>
	<td bgcolor="#FFFFFF" align="center"><a href="javascript:DelList('s<%=intLoop+1%>')">����</a></td>
</tr>
	<%Next%>				
<%else%>	
	<tr>
		<td bgcolor="#FFFFFF" colspan="10" align="center">������ �����ϴ�.</td>
	</tr>
<%END IF%>

</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

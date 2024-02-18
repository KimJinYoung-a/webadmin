<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ SMS �߼� ������
' History : �̻� ����
'           2020.12.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim reqhp, smstext, orderserial, userid, defaultMsg, makerid, itemid, orderdetailidx, title
dim sqlstr, smstext1, smstext2, smstext3
	reqhp 			= requestCheckvar(request("reqhp"),16)
	smstext 		= request("smstext")
	orderserial 	= requestCheckvar(request("orderserial"),11)
	userid 			= requestCheckvar(request("userid"),32)
	defaultMsg 		= request("defaultMsg")
	title 			= request("title")
	makerid 		= requestCheckvar(request("makerid"),32)
	itemid 			= requestCheckvar(request("itemid"),10)
	orderdetailidx 	= requestCheckvar(request("orderdetailidx"),10)

if (defaultMsg="") then defaultMsg="[�ٹ�����]"

if (reqhp<>"" and smstext<>"") then

	if LenB(smstext) > 80 then
    	Call SendNormalLMS(reqhp, title, "", smstext)
    else
    	Call SendNormalSMS_LINK(reqhp, "", smstext)
    end if

    response.write "<script>alert('���۵Ǿ����ϴ�.');</script>"

	''// ��ȭ��ȣ/�ֹ���ȣ ������ CS�޸� ����, skyer9, 2015-05-06
    ''if (orderserial<>"") or (userid<>"") then
        call AddCsMemo(orderserial,"1",userid,session("ssBctId"),"[SMS "+ reqhp + "]" + smstext)
        response.write "<script>alert('�߼۳��뿡 MEMO�� ����Ǿ����ϴ�.')</script>"
    ''end if

    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
 }

function SendSMS(frm){

	if (frm.reqhp.value.length<9){		// - ���� üũ�ϰ� ��������. �����Ϳ��� �ʹ� �����ϴٰ���. 2020.12.10 �ѿ��
		alert('�ڵ�����ȣ�� ��Ȯ�� �Է��ϼ���.');
		return;
    }
	if ((frm.smstext.value.length<1)||(frm.smstext.value=="[�ٹ�����]")){
		alert('���ڳ����� �Է��ϼ���.');
		return;
	}

    if (getByteLength(frm.smstext.value) > 2000) {
        alert('1000õ ���� �̻��� �Է��� �� �����ϴ�.\n\n���ڼ��� �ٿ��ּ���');
        return;
    }

    if (getByteLength(frm.smstext.value) > 80) {
        var varconfirmMsg = 'LMS �� �̿��� ���ڸ� �����մϴ�. ���� �Ͻðڽ��ϱ�?';
    }else{
        var varconfirmMsg = '���� �Ͻðڽ��ϱ�?';
    }

	var ret= confirm(varconfirmMsg);
	if(ret){
		frm.submit();
	}
}

function updateChar() {
	var length = calculate_msglen(document.getElementById("smstext").value);

	if (length <= 80) {
		document.getElementById("charlen").innerHTML = "(" + length + "/80)<br><br>SMS";

		document.getElementById("title").className = "text_ro";
		document.getElementById("title").readOnly = true;
		document.getElementById("title").value = "";
	} else {
		document.getElementById("charlen").innerHTML = "(" + length + "/2000)<br><br><font color='red'>LMS</font>";

		document.getElementById("title").className = "text";
		document.getElementById("title").readOnly = false;
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

function getOnload(){
	updateChar();
}

function TnCSTemplateGubunChanged(gubun) {
	var frm = document.frm;

	var orderserial = frm.orderserial.value;
	var userid = frm.userid.value;

	var makerid = frm.makerid.value;
	var itemid = frm.itemid.value;
	var orderdetailidx = frm.orderdetailidx.value;

	CSTemplateFrame.location.href="/cscenter/board/cs_template_select_process.asp?mastergubun=10&gubun=" + gubun + "&orderserial=" + orderserial + "&userid=" + userid + "&makerid=" + makerid + "&itemid=" + itemid + "&orderdetailidx=" + orderdetailidx;
}

 function TnCSTemplateGubunProcess(v, errMSG) {

	if (errMSG != "") {
		alert(errMSG);
		document.frm.smstext.value = "";
		return;
	}

	if(v == ''){
	}
	else{
		document.frm.smstext.value = v;
		// alert(v);
	}
 }

function GetOrderInfo() {
	var orderserial = document.getElementById("orderserial").value;
	if (orderserial.length == 11) {
		var userid = document.getElementById("userid").value;
		var reqhp = document.getElementById("reqhp").value;
		if ((userid.length == 0) && (reqhp.length == 0)) {
			$.ajax({
				type: 'GET',
				url: '/cscenter/action/ajax_get_order_info.asp?orderserial=' + orderserial,
				data: { get_param: 'value' },
				success: function (data) {
					var obj = jQuery.parseJSON(data);
					if (obj.code == "00") {
						var userid = document.getElementById("userid");
						var reqhp = document.getElementById("reqhp");

						userid.value = obj.userid;
						reqhp.value = obj.reqhp;
					}
				},
			});
		}
	}
}

window.onload = getOnload;

<% if (orderserial="") and (reqhp = "") and (userid = "") then %>
$( document ).ready(function() {
    $('input[name=orderserial]').keyup(function() {
		GetOrderInfo();
	});
    $('input[name=orderserial]').change(function() {
		GetOrderInfo();
	});
	$('input[name=orderserial]').bind('input propertychange', function() {
		GetOrderInfo();
	});
});
<% end if %>
// window.resizeTo('280','320');

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		���ڹ߼�.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" method=post action="pop_cs_sms_send.asp" style="margin:0px;" >
<input type="hidden" name="mode" value="send">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="orderserial" id="orderserial" value="<%= orderserial %>" size="15" maxlength="16" <% if (orderserial <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
	</td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" id="userid" name="userid" value="<%= userid %>" size="15" maxlength="32" <% if (userid <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
	</td>
</tr>
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
	<td bgcolor="#FFFFFF"><input type="text" id="reqhp" name="reqhp" class="text" value="<%= reqhp %>" size="15" maxlength="16"></td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="makerid" value="<%= makerid %>" size="15" maxlength="32" <% if (makerid <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
	</td>
</tr>
<input type="hidden" name="orderdetailidx" value="<%= orderdetailidx %>">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemid" value="<%= itemid %>" size="15" maxlength="32" <% if (itemid <> "") or True then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
	</td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF"></td>
</tr>
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input id="title" type="text" name="title" class="text" value="" size="25" maxlength="30" readonly>

		<% SelectBoxCSTemplateGubun "10", "" %>
		<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">
		���ڳ���<br>
		<div id="charlen"></div>
	</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea id="smstext" name="smstext" class="textarea" cols="52" rows="10" onKeyUp="updateChar()"><%= defaultMsg %></textarea></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="4" align="center">
		<input type="button" class="button" value="SMS�߼�" onclick="SendSMS(frm);">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/board/cs_templatecls.asp"-->
<%

'// SMS �߼� ������

dim reqhp, smstext
dim orderserial, userid, defaultMsg, makerid, itemid, orderdetailidx
dim title

reqhp 			= RequestCheckvar(request("reqhp"),16)
smstext 		= request("smstext")
orderserial 	= RequestCheckvar(request("orderserial"),16)
userid 			= RequestCheckvar(request("userid"),32)
defaultMsg 		= request("defaultMsg")
title 			= request("title")

makerid 		= RequestCheckvar(request("makerid"),32)
itemid 			= RequestCheckvar(request("itemid"),10)
orderdetailidx 	= RequestCheckvar(request("orderdetailidx"),10)
if smstext <> "" then
	if checkNotValidHTML(smstext) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if defaultMsg <> "" then
	if checkNotValidHTML(defaultMsg) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if title <> "" then
	if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if (defaultMsg="") then defaultMsg="[�ΰŽ�]"

dim sqlstr
dim smstext1, smstext2, smstext3

if (reqhp<>"" and smstext<>"") then

	if LenB(smstext) > 80 then
    	Call SendNormalLMS(reqhp, title, CS_MAIN_PHONENO, smstext)
    else
    	Call SendNormalSMS_LINK(reqhp, CS_MAIN_PHONENO, smstext)
    end if


    response.write "<script>alert('���۵Ǿ����ϴ�.');</script>"

	call AddCsMemo(orderserial,"1",userid,session("ssBctId"),"[SMS "+ reqhp + "]" + smstext)
	response.write "<script>alert('�߼۳��뿡 MEMO�� ����Ǿ����ϴ�.')</script>"

    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

%>

<script language='javascript'>
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

	if (frm.reqhp.value.length<12){
		alert('�ڵ�����ȣ�� ��Ȯ�� �Է��ϼ���.');
		return;
    }
	if ((frm.smstext.value.length<1)||(frm.smstext.value=="[�ΰŽ�]")){
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

	CSTemplateFrame.location.href="/cscenterv2/board/cs_template_select_process.asp?mastergubun=10&gubun=" + gubun + "&orderserial=" + orderserial + "&userid=" + userid + "&makerid=" + makerid + "&itemid=" + itemid + "&orderdetailidx=" + orderdetailidx;
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

window.onload = getOnload;

// window.resizeTo('280','320');

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
    <form name="frm" method=post action="pop_cs_sms_send.asp">
    <input type="hidden" name="mode" value="send">
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="orderserial" value="<%= orderserial %>" size="15" maxlength="16" <% if (orderserial <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    	</td>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="userid" value="<%= userid %>" size="15" maxlength="32" <% if (userid <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    	</td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="reqhp" class="text" value="<%= reqhp %>" size="15" maxlength="16"></td>
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
    	<td bgcolor="#FFFFFF" colspan="3"><textarea name="smstext" class="textarea" cols="52" rows="10" onKeyUp="updateChar()"><%= defaultMsg %></textarea></td>
    </tr>
    </form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <input type="button" class="button" value="SMS�߼�" onclick="SendSMS(frm);">
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
<br>* 2,000 ����Ʈ ���� ���۰���

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

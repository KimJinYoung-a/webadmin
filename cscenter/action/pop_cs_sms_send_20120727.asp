<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
dim reqhp, smstext
dim orderserial, userid, defaultMsg
dim title

reqhp 			= request("reqhp")
smstext 		= request("smstext")
orderserial 	= request("orderserial")
userid 			= request("userid")
defaultMsg 		= request("defaultMsg")
title 			= request("title")

if (defaultMsg="") then defaultMsg="[�ٹ�����]"

dim sqlstr
dim smstext1, smstext2, smstext3

if (reqhp<>"" and smstext<>"") then

	if LenB(smstext) > 80 then
    	Call SendNormalLMS(reqhp, title, "", smstext)
    else
    	Call SendNormalSMS(reqhp, "", smstext)
    end if





''����
''    if LenB(smstext)>160 then
''        smstext1 = LeftB(smstext,80)
''        smstext2 = MidB(smstext,81,80)
''        smstext3 = MidB(smstext,161,80)
''    elseif LenB(smstext)>80 then
''        smstext1 = LeftB(smstext,80)
''        smstext2 = MidB(smstext,81,80)
''    else
''        smstext1 = smstext
''    end if
''
''
''    if (Trim(smstext1)<>"") then
''    	sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
''    	sqlStr = sqlStr + " values('" + reqhp + "',"
''    	sqlStr = sqlStr + " '1644-6030',"
''    	sqlStr = sqlStr + " '1',"
''    	sqlStr = sqlStr + " getdate(),"
''    	sqlStr = sqlStr + " '" + html2db(smstext1) + "')"
''
''    	dbget.Execute sqlStr
''    end if
''
''    if (Trim(smstext2)<>"") then
''    	sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
''    	sqlStr = sqlStr + " values('" + reqhp + "',"
''    	sqlStr = sqlStr + " '1644-6030',"
''    	sqlStr = sqlStr + " '1',"
''    	sqlStr = sqlStr + " getdate(),"
''    	sqlStr = sqlStr + " '" + html2db(smstext2) + "')"
''
''    	dbget.Execute sqlStr
''    end if
''
''    if (Trim(smstext3)<>"") then
''    	sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
''    	sqlStr = sqlStr + " values('" + reqhp + "',"
''    	sqlStr = sqlStr + " '1644-6030',"
''    	sqlStr = sqlStr + " '1',"
''    	sqlStr = sqlStr + " getdate(),"
''    	sqlStr = sqlStr + " '" + html2db(smstext3) + "')"
''
''    	dbget.Execute sqlStr
''    end if

    response.write "<script>alert('���۵Ǿ����ϴ�.');</script>"

    if (orderserial<>"") or (userid<>"") then
        call AddCsMemo(orderserial,"1",userid,session("ssBctId"),"[SMS "+ reqhp + "]" + smstext)
        response.write "<script>alert('�߼۳��뿡 MEMO�� ����Ǿ����ϴ�.')</script>"
    end if

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
    		<input type="text" name="orderserial" value="<%= orderserial %>" size="13" maxlength="16" <% if (orderserial <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    		* �ֹ���ȣ �Է½� CS�޸� ��ϵ˴ϴ�.
    	</td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="userid" value="<%= userid %>" size="15" maxlength="32" <% if (userid <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    		* ���̵� �Է½� CS�޸� ��ϵ˴ϴ�.
    	</td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ڵ�����ȣ</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="reqhp" class="text" value="<%= reqhp %>" size="15" maxlength="16"></td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF"><input id="title" type="text" name="title" class="text" value="" size="30" maxlength="30" readonly></td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>" align="center">
    		���ڳ���<br>
    		<div id="charlen"></div>
    	</td>
    	<td bgcolor="#FFFFFF"><textarea name="smstext" class="textarea" cols="52" rows="10" onKeyUp="updateChar()"><%= defaultMsg %></textarea></td>
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

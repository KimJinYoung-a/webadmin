<html>
<head>
<title></title>
<meta http-equiv="cache-control" content="no-cache" />
<link rel="stylesheet" type="text/css" href="default.css">
<meta http-equiv="content-type" content="text/html; charset=euc-kr">
<script src="/cscenter/js/sha512.js"></script>
</head>
<!--��Ʈ�ѷ�-->
<script language=javascript src="EventClientCtrlObj1012.js"></script>

<script>

var ISCALL=0;
var STAT=0;
var timerID=null;
var isExtened=0;
var PhoneNum="";
var PhonePeer="";
var PhoneCaller="";
var RestStatus="0";
var READY=0;
var chk = "checked";
var MemberStatus = "";
var PhoneStatus = "";
var FORWARD_WHEN = "";
var FORWARD_NUM = "";

//------------���󼭹�IP---------
var strServerIP = "203.84.251.211";
//������Ʈ
var strServerPort = "8083"; //8083

//------------��Ʈ�ѷ� ����---------
if(document.all.EventClientCtrl.readyState == 4 )
{
        READY=1;
}

function decodeURL(str){
    var s0, i, j, s, ss, u, n, f;
    s0 = "";                // decoded str
    for (i = 0; i < str.length; i++){   // scan the source str
        s = str.charAt(i);
        if (s == "+"){s0 += " ";}       // "+" should be changed to SP
        else {
            if (s != "%"){s0 += s;}     // add an unescaped char
            else{               // escape sequence decoding
                u = 0;          // unicode of the character
                f = 1;          // escape flag, zero means end of this sequence
                while (true) {
                    ss = "";        // local str to parse as int
                        for (j = 0; j < 2; j++ ) {  // get two maximum hex characters for parse
                            sss = str.charAt(++i);
                            if (((sss >= "0") && (sss <= "9")) || ((sss >= "a") && (sss <= "f"))  || ((sss >= "A") && (sss <= "F"))) {
                                ss += sss;      // if hex, add the hex character
                            } else {--i; break;}    // not a hex char., exit the loop
                        }
                    n = parseInt(ss, 16);           // parse the hex str as byte
                    if (n <= 0x7f){u = n; f = 1;}   // single byte format
                    if ((n >= 0xc0) && (n <= 0xdf)){u = n & 0x1f; f = 2;}   // double byte format
                    if ((n >= 0xe0) && (n <= 0xef)){u = n & 0x0f; f = 3;}   // triple byte format
                    if ((n >= 0xf0) && (n <= 0xf7)){u = n & 0x07; f = 4;}   // quaternary byte format (extended)
                    if ((n >= 0x80) && (n <= 0xbf)){u = (u << 6) + (n & 0x3f); --f;}         // not a first, shift and add 6 lower bits
                    if (f <= 1){break;}         // end of the utf byte sequence
                    if (str.charAt(i + 1) == "%"){ i++ ;}                   // test for the next shift byte
                    else {break;}                   // abnormal, format error
                }
            s0 += String.fromCharCode(u);           // add the escaped character
            }
        }
    }
    return s0;
}

//------------������ �α� ------------
function loginIppbx(id, pass, exten, comp)
{
        if(READY ==0){
                alert("���ͷ��� ���۵��� ���߽��ϴ�.");
                return false;
        }
        if(document.all.EventClientCtrl.IsConnected() == true){
                document.all.EventClientCtrl.DisconnectServer();
                setTimeout("chgButton()",1000);
        } else {
			var strLoginID = id+"_"+exten+"_0_"+comp;
			var strLoginPwd = hex_sha512(pass);
			// var strLoginPwd = pass;

			remove_box();
			displayText("S", "ConnectServer('"+strServerIP+"', '"+strServerPort+"', '"+strLoginID+"', '"+strLoginPwd+"')");
			document.all.EventClientCtrl.SetEncryption(0);
			document.all.EventClientCtrl.ConnectServer(strServerIP, strServerPort, strLoginID, strLoginPwd);
        }
        return false;
}

//------------���� �α׾ƿ� ------------
function ippbxLogout()
{
        var bConnect = document.all.EventClientCtrl.IsConnected();
        if(bConnect == true)
        {
		displayText("S", "DisconnectServer()");
                document.all.EventClientCtrl.DisconnectServer();
                setTimeout("chgButton()",1000);
        }
}

function OnPageLoad()
{
        if(document.all.EventClientCtrl != null )
        {
                document.all.EventClientCtrl.SetLogMode(true);
        }
}

function SetEncrypt(v)
{
        if(document.all.EventClientCtrl != null )
        {
                document.all.EventClientCtrl.SetEncryption(v);
        }
}
//-------------Ŭ������----------
function click2call(num, cid)
{
	displayText("S", "Click2Call('"+cid+"', '"+num+"', 'outbound')");
        var calll=document.all.EventClientCtrl.Click2Call(cid,num,"outbound");
        return false;
}
function click2dial(id,num,servive)
{
                var calll=document.all.EventClientCtrl.Click2Call(id,num,servive);
                return false;
}
//----------��ư ��Ʈ��------------
function chgButton()
{
        var bConnect = document.all.EventClientCtrl.IsConnected();
        ISCON=(bConnect)?1:0;
        if(ISCON == 1)
        {
                document.all.LogTR.innerHTML= "<img src='/image/left_dot01.gif'> "
			+"<b>[����:"+PhoneCaller+"/"+PhoneNum+"/"+PhonePeer+"]</b>"
			+"&nbsp;<input name='mbtn' type='button' style='width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A' value='�α׾ƿ�'  onClick='javascript:logout()'>"
			+"&nbsp;/&nbsp;<input type=button name=phonestatus value='��ȭ�����' style='width:80px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000'>"
			+"&nbsp;<input type=button name=forwardstatus value='������ȯ����:' style='width:200px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000'>";


		parseForwarding(FORWARD_NUM,FORWARD_WHEN);
		parsePhoneStatus(PhoneStatus);
		parseMemberStatus(MemberStatus);
		parsePDSMemberStatus('3');

        } else {
                //alert("�������");
                //��α�Ǿ����� ����
                location.reload();
        }
}

//------------������ ��ɾ���� ------------
function SIPCommand(strCommand)
{
        if(document.all.EventClientCtrl.IsConnected() == false)
        {
                displayText("S", "������ ������� ���� ����");
                return;
        }
        if(strCommand != "" && strCommand != "undefined" )
        {
		if(strCommand != "PONG")
		{
			displayText("S", "SendSIPCommand('CMD|"+strCommand+"')");
		}
                document.all.EventClientCtrl.SendSIPCommand("CMD|"+strCommand);
        }
        return false;
}


<!------------MESSAGE PARSE START------------>

//�α�
function parseLogin(kind,data1,data2,data3,data4, data5,data6,data7)
{
        //LOGIN|KIND:LOGIN_OK|DATA1:300|DATA2:����1|DATA3:0|DATA4:OK|DATA5:11110002
        if(kind == "LOGIN_OK")
        {
                PhoneNum=data1;
                PhonePeer=data5;
                PhoneCaller=decodeURL(data2);
                MemberStatus = data3;
                PhoneStatus = data4;
                FORWARD_WHEN = data6;
                FORWARD_NUM = data7;

                setTimeout("chgButton()",1000);
		SIPCommand("LOGIN_ACK");

        } else {
                document.all.EventClientCtrl.DisconnectServer();
                alert("�α� ����");
        }
        return;
}
function parseCallStatus(kind,data1,data2)
{
        if(kind == "REDIRECT")
        {
                if(data2 == "NOCHAN")
                {
                        //alert("�����ֱ��� ä���� ����");
                        return;
                } else if(data2 == "BUSY") {
                        //alert(data1+"�� ��ȭ��");
                        return;
                }
        }
}
function parseCallEvent(kind,data1,data2,data3,data4,data5,data6)
{
/*
        if(kind == "IR")
        {
                alert("**"+PhoneNum+" �ιٿ�� ��ȭ�� ["+data1+"]���� ����");
        } else if(kind == "ID") {
                alert("**"+PhoneNum+" �ιٿ�� ��ȭ ["+data1+"]�� ��ȭ��");
        } else if(kind == "OR") {
                alert("**"+PhoneNum+" �ƿ��ٿ�� ��ȭ ["+data1+"]�� �õ���");
        } else if(kind == "OD") {
                alert("**"+PhoneNum+" �ƿ��ٿ�� ��ȭ ["+data1+"]�� ��ȭ��");
        } else if(kind == "PICKUP") {
                alert("**"+PhoneNum+" ��ܹޱ� ��ȭ ["+data1+"]�� ��ȭ��");
        }
*/
}

function parseHangupEvent(kind,data1,data2,data3,data4)
{
        //alert("**"+PhoneNum+" ��ȭ���� ["+data1+","+data2+"]");
        STAT=0;
        ISCALL=0;
}
function toUTF8(szInput)
{
 var wch,x,uch="",szRet="";
 for (x=0; x<szInput.length; x++)
  {
  wch=szInput.charCodeAt(x);
  if (!(wch & 0xFF80)) {
   szRet += "%" + wch.toString(16);
  }
  else if (!(wch & 0xF000)) {
   uch = "%" + (wch>>6 | 0xC0).toString(16) +
      "%" + (wch & 0x3F | 0x80).toString(16);
   szRet += uch;
  }
  else {
   uch = "%" + (wch >> 12 | 0xE0).toString(16) +
      "%" + (((wch >> 6) & 0x3F) | 0x80).toString(16) +
      "%" + (wch & 0x3F | 0x80).toString(16);
   szRet += uch;
  }
 }
 return(szRet);
}

function parseEtc(msg)
{
	displayText("R", msg);

        var msgs=msg.split("|");
        if(msgs == null || msgs.length < 2)
        {
                return;
        }
        var event = msgs[0];

        var Insp=new Object();
        for(i=1;i<msgs.length;i++)
        {
                keyval=msgs[i].split(":");
                Insp[keyval[0]]=keyval[1];
        }
        var kind = Insp["KIND"];
        var data1 = Insp["DATA1"];
        var data2 = Insp["DATA2"];
        var data3 = Insp["DATA3"];
        var data4 = Insp["DATA4"];
        var data5 = Insp["DATA5"];
        var data6 = Insp["DATA6"];
        var data7 = Insp["DATA7"];
        var data8 = Insp["DATA8"];
        var data9 = Insp["DATA9"];


        if(event == "LOGIN")
        {
                parseLogin(kind,data1,data2,data3,data4,data5,data6,data7);
                return;
        }
        else if(event == "PEER")
        {
		parsePhoneStatus(data2);
                return;
        }
        else if(event == "CALLEVENT")
        {
                parseCallEvent(kind,data1,data2,data3,data4,data5,data6);
                return;
        }
        else if(event == "FORWARDING")
        {
		if(kind == 'OK')
		{
                	parseForwarding(data1,data2);
		}
                return;
        }
        else if(event == "CALLSTATUS")
        {
                parseCallStatus(kind,data1,data2);
                return;
        }
        else if(event == "DTMFCARDEVENT")
        {
                document.all.card_num.value = document.all.card_num.value+kind;
        }
        else if(event == "DTMFCVCEVENT")
        {
                document.all.cvc_num.value = document.all.cvc_num.value+kind;
        }
        else if(event == "HANGUPEVENT")
	{

                if(data8 == "")
                {
                        data8 = "NORMAL";
                }
		SIPCommand("HANGUP_ACK|"+data5+","+data8);
                return;
        } else if(event == "MEMBERSTATUS") {
		parseMemberStatus(kind);
        }
        else if(event == "PDSMEMBERSTATUS")
        {
                parsePDSMemberStatus(kind);
        }
        else if(event == "BYE")
        {
                //alert(kind);
                //document.all.EventClientCtrl.DisconnectServer();
        }
        else
        {
                //alert("ELSE:"+msg);
        }
        return;
}
function parsePDSMemberStatus(kind)
{
        if(kind =='0')
        {
                document.all.pdsstatus0.style.background = "red";
                document.all.pdsstatus1.style.background = "white";
                document.all.pdsstatus2.style.background = "white";
                document.all.pdsstatus3.style.background = "white";
        } else if(kind =='1') {
                document.all.pdsstatus0.style.background = "white";
                document.all.pdsstatus1.style.background = "red";
                document.all.pdsstatus2.style.background = "white";
                document.all.pdsstatus3.style.background = "white";
        } else if(kind =='2') {
                document.all.pdsstatus0.style.background = "white";
                document.all.pdsstatus1.style.background = "white";
                document.all.pdsstatus2.style.background = "red";
                document.all.pdsstatus3.style.background = "white";
        } else if(kind =='3') {
                document.all.pdsstatus0.style.background = "white";
                document.all.pdsstatus1.style.background = "white";
                document.all.pdsstatus2.style.background = "white";
                document.all.pdsstatus3.style.background = "red";
        } else {
        }
}
function parseMemberStatus(kind)
{
        if(kind =='0')
        {
                document.all.memberstatus0.style.background = "red";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
                document.all.memberstatus8.style.background = "white";
        } else if(kind =='1') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "red";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
                document.all.memberstatus8.style.background = "white";
        } else if(kind =='2') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "red";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
                document.all.memberstatus8.style.background = "white";
        } else if(kind =='3') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "red";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
                document.all.memberstatus8.style.background = "white";
        } else if(kind =='4') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "red";
                document.all.memberstatus5.style.background = "white";
                document.all.memberstatus8.style.background = "white";
        } else if(kind =='5') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "red";
                document.all.memberstatus8.style.background = "white";
        } else if(kind =='8') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
                document.all.memberstatus8.style.background = "red";
        } else {
        }
}
function parseForwarding(num, when)
{
		var label = "������ȯ����:";
		if(when == 'A')
		{
			document.all.forward_when[1].checked = true;
			label = label+"�׻�["+num+"]";
                	document.all.forwardstatus.style.background = "yellow";
		} else if(when == 'B') {
			document.all.forward_when[2].checked = true;
			label = label+"��ȭ��["+num+"]";
                	document.all.forwardstatus.style.background = "yellow";
		} else if(when == 'C') {
			document.all.forward_when[3].checked = true;
			label = label+"������["+num+"]";
                	document.all.forwardstatus.style.background = "yellow";
		} else {
			document.all.forward_when[0].checked = true;
			label = label+"����";
                	document.all.forwardstatus.style.background = "white";
		}
		document.all.forwarding.value=num;
		document.all.forwardstatus.value=label;
}
function parsePhoneStatus(kind)
{
        if(kind =='OK' || kind =='REGISTERED' ||kind =='REACHABLE' )
        {
                document.all.phonestatus.style.background = "lightgreen";
        } else if(kind =='NOK' || kind=='UNREACHABLE' || kind=='UNREGISTERED') {
                document.all.phonestatus.style.background = "gray";
        } else {
                document.all.phonestatus.style.background = "white";
        }
}
</script>
<!--Activex�� ��ũ��Ʈ �����κ� START-->

<script id="OnSendEtcEvent" for="EventClientCtrl" event="SendEtcEvent(strEventName,strEventValue)">
        if(strEventValue != 'PING')
        {
                if(strEventValue == "Bye")
                {
                        alert("������ ������ ������ϴ�. \n�ٸ������� ���� ���̵�� �α�Ǿ���");
                        document.all.EventClientCtrl.DisconnectServer();
                } else {
                        parseEtc(strEventValue);
                }
        } else {
                SIPCommand("PONG");
        }
        return false;
</script>
<script id="OnSendNetworkErrorEvent" for="EventClientCtrl" event="SendNetworkErrorEvent()">
        document.all.EventClientCtrl.DisconnectServer();
        setTimeout("chgButton()",1000);
        alert("������ ���� ����!");
</script>


<!--��ȭ�ɱ� ���� ��Ÿ��ɾ� �׽�Ʈ START-->
<script>
function logout()
{
        alert("�α׾ƿ�");
        ippbxLogout();
        return false;
}
function login()
{
        if(document.input_form.userid.value == null || document.input_form.userid.value=="")
        {
                alert("������ȣ�� �Է��ϼ���");
                return;
        }
        loginIppbx(document.input_form.userid.value, document.input_form.pass.value,  document.input_form.exten.value,document.input_form.companyid.value);
        return false;
}

function click2call_test()
{
        alert(document.input_form.number.value+"�� ��ȭ�ɱ�");
        click2call(document.input_form.number.value, document.input_form.cid.value);
        return false;
}
function sipcommand_attended()
{
        if(document.all.attended.value == null || document.all.attended.value=="")
        {
                alert("������ ������ ������ �Է��ϼ���");
                return;
        } else {
                var rtn = confirm("["+document.all.attended.value+"] �� ��ȭ�� �����ðڽ��ϱ�?");

                if(rtn == false)
                {
                        return;
                }
        }
        SIPCommand("ATTENDED|"+document.all.attended.value);
        return false;
}
function sipcommand_redirect()
{
        if(document.all.redirect.value == null || document.all.redirect.value=="")
        {
                alert("������ ������ ������ �Է��ϼ���");
                return;
        } else {
                var rtn = confirm("["+document.all.redirect.value+"] �� ��ȭ�� �����ðڽ��ϱ�?");

                if(rtn == false)
                {
                        return;
                }
        }
        SIPCommand("REDIRECT|"+document.all.redirect.value);
        return false;
}
function sipcommand_msg()
{
        if(document.all.msg_exten.value == null || document.all.msg_exten.value=="")
        {
                alert("�޼������� ������ ������ �Է��ϼ���");
                return;
        }
        if(document.all.msg.value == null || document.all.msg.value=="")
        {
                alert("�޼����� �Է��ϼ���");
                return;
        }
        SIPCommand("MSG|"+document.all.msg_exten.value+"|"+document.all.msg.value);
        return false;
}
function sipcommand_spy()
{
        if(document.all.spy_exten.value == null || document.all.spy_exten.value=="")
        {
                alert("������� ������ ������ �Է��ϼ���");
                return;
        }
        SIPCommand("SPY|"+PhonePeer+","+document.all.spy_exten.value);
        return false;
}
function sipcommand_rec(mode)
{
	if(mode == 'start')
	{
        	SIPCommand("REC_START|"+PhonePeer);
	} else {
        	SIPCommand("REC_STOP|"+PhonePeer);
	}
        return false;
}
function sipcommand_card(mode)
{
        SIPCommand("DTMFCARD|"+mode);
        return false;
}
function sipcommand_cvc(mode)
{
        SIPCommand("DTMFCVC|"+mode);
        return false;
}
function sipcommand_memberstatus(s)
{
        SIPCommand("MEMBERSTATUS|"+s+","+PhoneNum);
}
function pds_memberstatus(s)
{
        SIPCommand("PDSMEMBERSTATUS|"+s+","+PhoneNum);
}
function sipcommand_hangup()
{
        SIPCommand("HANGUP|"+PhonePeer);
}
function sipcommand_receive()
{
        SIPCommand("RECEIVE|"+PhonePeer);
}
function sipcommand_reject()
{
        SIPCommand("REJECT|"+PhonePeer);
}
function sipcommand_pickup()
{
        SIPCommand("PICKUP|"+PhonePeer);
}
function selectForward(value)
{
	FORWARD_WHEN = value;
}
function sipcommand_forwarding()
{
	if(FORWARD_WHEN != 'N' && document.all.forwarding.value=='')
	{
		alert("������ȯ�� ��ȣ�� �Է����ּ���");
		return;
	}
        SIPCommand("FORWARDING|"+PhoneNum+","+document.all.forwarding.value+","+FORWARD_WHEN);
}
function displayText(fsend, text)
{
        event_num = event_num+1;
        if(fsend == "S")
        {
                document.all.snd_text.value = document.all.snd_text.value+"\nC->S["+event_num+"] "+text;
        } else {
                document.all.snd_text.value = document.all.snd_text.value+"\nS->C["+event_num+"] "+text;
        }
}
function remove_box()
{
        event_num=0;
        document.all.snd_text.value = "";
}
</script>
<body leftmargin="0" topmargin="0">
<br>
&nbsp;<b>�׽�ƮŬ���̾�Ʈ(Activex)</b><br>
<br>
<form name=input_form method="post">
<input type='hidden' name='selradio' value=''>

<table width="100%" border="0" cellpadding="5" cellspacing="5">
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
	<td>
		<DIV id=LogTR style='display:;border:solid 0'>
		<img src='/image/left_dot01.gif'>
		<b>COMP_ID : </b><input type=text size=10 name=companyid value='10x10'>&nbsp;
		<b>ID : </b><input type=text size=10 name=userid value='tozzinet'>&nbsp;
		<b>PW : </b><input type=text size=10 name=pass value='cube1010'>
		<b>���� : </b><input type=text size=10 name=exten value='808'>
		<input name="mbtn" type="button" style="width:60px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" value='�α���'  onClick="javascript:login()">
		</DIV>
	</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
	<td>
<img src='/image/left_dot01.gif'><b>��������:</b>
		<input type=button name=memberstatus0 value='���(0)' onClick=javascript:sipcommand_memberstatus('0') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus1 value='�����(1)' onClick=javascript:sipcommand_memberstatus('1') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus2 value='��ó��(2)' onClick=javascript:sipcommand_memberstatus('2') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus3 value='�޽�(3)' onClick=javascript:sipcommand_memberstatus('3') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus4 value='�Ļ�(4)' onClick=javascript:sipcommand_memberstatus('4') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus5 value='���Űź�(5)' onClick=javascript:sipcommand_memberstatus('5') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus8 value='PDS(8)' onClick=javascript:sipcommand_memberstatus('8') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
	</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td>
<img src='/image/left_dot01.gif'><b>PDS����:</b>
                <input type=button name=pdsstatus0 value='���(0)' onClick=javascript:pds_memberstatus('0') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
                <input type=button name=pdsstatus1 value='�����(1)' onClick=javascript:pds_memberstatus('1') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
                <input type=button name=pdsstatus2 value='��ó��(2)' onClick=javascript:pds_memberstatus('2') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
                <input type=button name=pdsstatus3 value='Ÿ����(3)' onClick=javascript:pds_memberstatus('3') style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
        </td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
	<td>
		<img src='/image/left_dot01.gif'>
		<b>��ȭ�ɱ� : </b>
		����ȣ: <input type=text size=15 name=number value=''>&nbsp;
		RID:<input type=text size=15 name=cid value=''>
		<input name="mbtn" type="button" value='��ȭ�ɱ�' style="width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:click2call_test()"><br>
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
	<td>
		<img src='/image/left_dot01.gif'>
		<b>�ޱ�,����,��ܹޱ� : </b>
                <input name="mbtn" type="button" value='��ȭ�ޱ�' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_receive()">
                <input name="mbtn" type="button" value='��ȭ����' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_hangup()">
                <input name="mbtn" type="button" value='��ܹޱ�' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_pickup()">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
<img src='/image/left_dot01.gif'>
<b>�����ֱ�(����)-���ٵ�� : </b><input type=text size=15 name='attended' value=''>
<input name="mbtn" type="button" value='�����ֱ�' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_attended()">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
		<img src='/image/left_dot01.gif'>
		<b>�����ֱ�(����, ��Ʈ��ȣ, ��ǥ��ȣ)-����ε� : </b>
		<input type=text size=15 name='redirect' value=''>
		<input name="mbtn" type="button" value='�����ֱ�' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_redirect()">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
		<img src='/image/left_dot01.gif'>
		<b>������ȯ : </b>
		<input type=text size=15 name='forwarding' value=''>
                    <input type="radio" name=forward_when value='N' onClick=selectForward('N')>������ȭ����
                    <input type="radio" name=forward_when value='A' onClick=selectForward('A')>�׻�
                    <input type="radio" name=forward_when value='B' onClick=selectForward('B')>��ȭ��
                    <input type="radio" name=forward_when value='C' onClick=selectForward('C')>������
		<input name="mbtn" type="button" value='������ȯ' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_forwarding()">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
		<img src='/image/left_dot01.gif'>
		<b>�ܹ��޼��� : </b>
		��������:<input type=text size=10 name='msg_exten' value=''>
		�޼���: <input type=text size=50 name='msg' value=''>
		<input name="mbtn" type="button" value='������' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_msg()">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
		<img src='/image/left_dot01.gif'>
		<b>����� : </b>
		����:<input type=text size=10 name='spy_exten' value=''>
		<input name="mbtn" type="button" value='����' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_spy()">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
		<img src='/image/left_dot01.gif'>
		<b>�κг��� : </b>
		<input name="mbtn" type="button" value='����' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_rec('start')">
		<input name="mbtn" type="button" value='����' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_rec('stop')">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr>
        <td align=left >
                <img src='/image/left_dot01.gif'>
                <b>ī���ȣ �� CVS ����: </b>
</td>
</tr>
<tr>
        <td align=left> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <b>ī���ȣ: </b>
                <input type=text size=20 name='card_num' value=''>
                <input name="mbtn" type="button" value='����' style="width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_card('Y')">
                <input name="mbtn" type="button" value='����' style="width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_card('N')">
</td>
</tr>
<tr>
        <td align=left> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <b>CVC: </b>
                <input type=text size=10 name='cvc_num' value=''>
                <input name="mbtn" type="button" value='����' style="width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_cvc('Y')">
                <input name="mbtn" type="button" value='����' style="width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:sipcommand_cvc('N')">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=10><td></td></tr>
<tr>
<td>
                <br><b>��ɾ�/�̺�Ʈ:</b>&nbsp;&nbsp;<input name="mbtn" type="button" value='����' style="width:60px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" onClick="javascript:remove_box()"><br>
                <textarea name="snd_text" cols=160 rows=28 >
                </textarea>
</td>
</tr>
</table>

</form>
</body>
</html>

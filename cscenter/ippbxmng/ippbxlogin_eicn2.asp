<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

Dim ippbxLocalUser, localCallNo, autoLogin

autoLogin = request("al")
if (autoLogin = "") then
	autoLogin = "Y"
end if

dim sqlStr

sqlStr = " select top 1 localcallno "
sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_ippbx_user "
sqlStr = sqlStr + " where userid = '" & session("ssBctId") & "' "
sqlStr = sqlStr + " and useyn = 'Y' "
'response.write sqlStr

ippbxLocalUser = ""
if (session("ssBctId") <> "") then
	rsget.Open sqlStr, dbget, 1
	if  not rsget.EOF  then
		ippbxLocalUser = session("ssBctId")
		''if (ippbxLocalUser = "tozzinet") then
		''	ippbxLocalUser = "hasora"
		''end if
		localCallNo = rsget("localcallno")
	end if
	rsget.close
end if

dim i

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<script src="/cscenter/js/jquery-1.9.1.js"></script>
<script src="/cscenter/js/sha512.js"></script>

</head>

<script language=javascript src="EventClientCtrlObj1012.js"></script>
<script language=javascript >

var strServerIP = "110.93.128.96";
var strServerPort = "8083";

var ISCALL = 0;
var STAT = 0;
var timerID = null;
var isExtened = 0;
var PhoneNum = "";
var PhoneCaller = "";
var RestStatus = "0";
var READY = 0;

var event_num = 0;
var PhonePeer = "";
var PhoneStatus = "";
var FORWARD_WHEN = "";
var FORWARD_NUM = "";
var MemberStatus = "";

var objCtrl = document.all.EventClientCtrl;
if (objCtrl != undefined) {
	if(objCtrl.readyState == 4) {
		READY = 1;
	}
}

function click2call(num)
{
	if ($("#userid").val()=="") {
		alert("로그인 후 이용가능합니다.");
		return false;
	}

	// var calll = objCtrl.Click2Call(PhoneCaller,num,"outbound");
	// cid 빈값이면 대표번호
	var calll = objCtrl.Click2Call("", num, "outbound");
	return false;
}

function click2dial(id,num,context)
{
		var calll = objCtrl.Click2Call(id,num,context);
		return false;
}

function LoginIppbx()
{
    //alert(1);
    if($("#company_id").val() == "" || $("#userid").val()=="" || $("#exten").val()=="" || $("#passwd").val() == "") {
        alert("로긴 정보를 입력하세요");
        return;
    }
    //alert(2);
	ConnectServer();

}

function ConnectServer()
{
	if(READY == 0) {
		alert("컨터롤이 시작되지 못했습니다.");
		return false;
	}

	if(objCtrl.IsConnected() == true) {
		alert("이미 로그인 되어 있습니다.");
		return false;
	}
    //alert(3);
	var strLoginID = $("#userid").val() + "_" + $("#exten").val() + "_0_" + $("#company_id").val();
	var strLoginPwd = hex_sha512($("#passwd").val());

	displayText("S", "ConnectServer('"+strServerIP+"', '"+strServerPort+"', '"+strLoginID+"', '"+strLoginPwd+"')");

	objCtrl.SetEncryption(0);
	//alert(4);
    objCtrl.ConnectServer(strServerIP, strServerPort, strLoginID, strLoginPwd);
  //  alert(5);
	return false;
}

function DisconnectServer()
{
	if(objCtrl.IsConnected() == true) {
		objCtrl.DisconnectServer();
	}
}

function OnPageLoad()
{
    if(objCtrl != null) {
        objCtrl.SetLogMode(true);
    }
}

function SetEncrypt(v)
{
    if(objCtrl != null) {
        objCtrl.SetEncryption(v);
    }
}

function parseMessage(msg)
{
	displayText("R", msg);

    var msgs=msg.split("|");
    if(msgs == null || msgs.length < 2) {
        return;
    }
    var event = msgs[0];

    var Insp=new Object();
    for(i=1;i<msgs.length;i++) {
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
    var data10 = Insp["DATA10"];
    var data11 = Insp["DATA11"];

    if(event == "LOGIN") {
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
        //$("#card_num").val($("#card_num").val()+kind);
    }
    else if(event == "DTMFCVCEVENT")
    {
        //$("#cvc_num").val($("#cvc_num").val()+kind);
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
        return;
    }
    else if(event == "PDSMEMBERSTATUS")
    {
        //parsePDSMemberStatus(kind);
    }
    else if(event == "SERVER_STATUS")
    {
		//ClientServer 데몬 상태
		parseLogout(kind);
	}
    else if(event == "NODESVC_STATUS")
    {
		//nodejs 데몬 상태
		parseNodeSvc(kind);
	}
    else if(event == "BYE")
    {
		parseBye(kind,data1,data3);
	}
    else
    {
        // alert("ELSE:"+msg);
    }
    return;
}

function SIPCommand(strCommand)
{
    if(objCtrl.IsConnected() == false) {
        alert("서버와 연결되지 않은 상태");
        return;
    }
    if(strCommand != "" && strCommand != "undefined") {
		if(strCommand != "PONG") {
			// displayText("S", "SendSIPCommand('CMD|"+strCommand+"')");
		}
        objCtrl.SendSIPCommand("CMD|"+strCommand);
    }

    return false;
}

function parseLogin(kind,data1,data2,data3,data4,data5,data6,data7)
{
    //LOGIN|KIND:LOGIN_OK|DATA1:300|DATA2:상담원1|DATA3:0|DATA4:OK|DATA5:11110002
    if(kind == "LOGIN_OK") {
		PhoneNum=data1;
		PhonePeer=data5;
		PhoneCaller=decodeURL(data2);
		MemberStatus = data3;
		PhoneStatus = data4;
		FORWARD_WHEN = data6;
		FORWARD_NUM = data7;

		SIPCommand("LOGIN_ACK");

        doLogin();
    } else {
        objCtrl.DisconnectServer();
        alert("로긴 실패");
    }

    return;
}

function doLogin()
{
	$("#LoginButton").hide();
	$("#LogoutButton").show();

	$("#RestSetButton").show();
	$("#RestResetButton").hide();

	// parseForwarding(FORWARD_NUM,FORWARD_WHEN);
	parsePhoneStatus(PhoneStatus);
}

function parseForwarding(num, when)
{
	var label = "착신전환 ";
	if(when == '')
	{
		when="N";
	}
	var forwarding = $("#forwarding");
	if(forwarding.length>0)
	{
		forwarding.val(num);
	}
	var forward_when = $('#forward_when');
	if(forward_when.length>0)
	{
		$("input[name=forward_when]").each(function(){
			if($(this).val() == when)
			{
				$(this).attr("checked", true);
			}
		});
	}
	if(when == 'A')
	{
		label = label+"항상["+num+"]";
        $("#forwardstatus").css("background","yellow");
	} else if(when == 'B') {
		label = label+"통화중["+num+"]";
        $("#forwardstatus").css("background","yellow");
	} else if(when == 'C') {
		label = label+"부재중["+num+"]";
        $("#forwardstatus").css("background","yellow");
	} else if(when == 'T') {
		label = label+"부재중+통화중["+num+"]";
        $("#forwardstatus").css("background","yellow");
	} else {
		label = label+"안함";
        $("#forwardstatus").css("background","#EBF3FC");
	}
	var forwardstatus = $("#forwardstatus");
	if(forwardstatus.length>0)
	{
		forwardstatus.html(label);
	}
}

function parsePhoneStatus(kind)
{
	var phonestatus = $("#phonestatus");
	if(phonestatus.length ==0)
	{
		return;
	}
    if(kind =='OK' || kind =='REGISTERED' ||kind =='REACHABLE' )
    {
        phonestatus.css("background","#EBF3FC");
		phonestatus.html("전화 정상");
    } else if(kind =='NOK' || kind=='UNREACHABLE' || kind=='UNREGISTERED') {
        phonestatus.css("background","gray");
		phonestatus.html("전화 없음");
    } else {
        phonestatus.css("background","#EBF3FC");
		phonestatus.html("ERR");
    }
}






function LogoutIppbx()
{
	DisconnectServer();
	location.href = "ippbxlogin_eicn2.asp?al=N";
}

//후처리 시작
function rest_set()
{
	$("#RestSetButton").hide();
	$("#RestResetButton").show();

    command_memberstatus('2');
}

//후처리 끝
function rest_reset()
{
	$("#RestSetButton").show();
	$("#RestResetButton").hide();

    command_memberstatus('0');
}

function redirect(num)
{
    SIPCommand("REDIRECT|" + num);
    return false;
}

function parseLogout(kind)
{
	doLogout();
}

function parseBye(kind, uid, name)
{
	alert("["+kind+"]"+name+"("+uid+")");
	// alert("로그아웃됨");
	// location.reload();
}

function parseCallStatus(kind,data1,data2)
{
    if(kind == "REDIRECT") {
        if(data2 == "NOCHAN") {
            //alert("돌려주기할 채널이 없음");
            return;
        } else if(data2 == "BUSY") {
            //alert(data1+"이 통화중");
            return;
        }
    }
}

function parseNodeSvc(kind)
{
	alert("전화기 서버 장애["+kind+"]");
	LogoutIppbx();
}



function doLogout()
{
	$("#LoginButton").show();
	$("#LogoutButton").hide();

	$("#RestSetButton").hide();
	$("#RestResetButton").hide();

	$("#memberstatus").html("Stop");
	$("#phonestatus").html("");

	$("#memberstatus").css("background","#EBF3FC");
	$("#phonestatus").css("background","#EBF3FC");

	DisconnectServer();
}





function parseMemberStatus(kind)
{
    MemberStatus = kind;

	$("#RestSetButton").attr("disabled", "disabled");
	$("#RestResetButton").attr("disabled", "disabled");

    if(kind =='0')
    {
		$("#RestSetButton").removeAttr("disabled");
		$("#memberstatus").html(PhoneCaller+" " + "<font color='blue'>대기</font>");
    } else if(kind =='1') {
        $("#memberstatus").html(PhoneCaller+" " + "<font color='green'>상담중</font>");
    } else if(kind =='2') {
		$("#RestResetButton").removeAttr("disabled");
        $("#memberstatus").html(PhoneCaller+" " + "<font color='green'>후처리</font>");
    } else if(kind =='3') {
        $("#memberstatus").html(PhoneCaller+" " + "<font color='green'>휴식</font>");
    } else if(kind =='4') {
        $("#memberstatus").html(PhoneCaller+" " + "<font color='green'>식사</font>");
    } else if(kind =='5') {
        $("#memberstatus").html(PhoneCaller+" " + "<font color='green'>수신거부</font>");
    } else {
		$("#memberstatus").html("<font color='red'>ERR</font>");
    }
}

function parseRecordType(type)
{
	var label = "";
	if(type == '')
	{
		return;
	}
	var rec = $("#record_type");
	if(rec.length>0)
	{
		if(type == 'M') {
			rec.html(label+"전수녹취");
		} else if(type == 'P') {
			rec.html(label+"부분녹취");
		}
	}
}

function parseCallEvent(kind,data1,data2,data3,data4,data5,data6)
{
	if(kind == "IR")
	{
		popCallRing($("#userid").val(), PhoneNum, data1,'','','');
	}
}

function parseHangupEvent(kind,data1,data2,data3,data4)
{
    //alert("**"+PhoneNum+" 전화끊음 ["+data1+","+data2+"]");
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

function command_memberstatus(s)
{
	SIPCommand("MEMBERSTATUS|"+s+","+PhoneNum+","+MemberStatus);
}

function displayText(fsend, text)
{
	var val = $("#logMsg").val();
	$("#logMsg").val(val + fsend + " " + text + "\n");
}

var iWinArray = new Array();
function popCallRing(ippbxuser,intel,caller,memoid,iorderserial,iuserid){
    //권한 문제로.. 계속 새창으로 띠울지여부..
    var popwinName = "popCallRing<%= Replace(CStr(FormatDateTime(now(),4)),":","") %><%= Right(CStr(FormatDateTime(now(),3)),2) %>_";
    var arrIdx = 0;
    var isFound = false;

    if (iWinArray.length>0){
        if (!isFound){
            arrIdx = iWinArray.length;
            popwinName = popwinName + arrIdx;
        }
    }else{
        popwinName = popwinName + arrIdx;
    }

	// SSL 적용(2014-03-11 skyer9)
	var popwin = window.open('https://webadmin.10x10.co.kr/cscenter/ordermaster/ordermasterWithCallRing.asp?ippbxuser=' + ippbxuser + '&intel=' + intel + '&phoneNumber=' + caller + '&id=' + memoid + '&orderserial=' + iorderserial + '&userid=' + iuserid,popwinName,'width=1680,height=1000,scrollbars=yes,resizable=yes');

    popwin.focus();
    iWinArray[arrIdx] = popwin;
}

function Click2CallBox(comp){
    if (comp.value.length<7){
        alert('전화번호를 입력하세요.');
        comp.focus();
        return;
    }

    click2call(comp.value);
}

</script>

<script language='javascript'>

function jsPopLog() {
	alert($("#logMsg").val());
}

function getOnLoad(){
    <% if (ippbxLocalUser<>"") and autoLogin = "Y" then %>
    LoginIppbx();
    <% end if %>
}

window.onload = getOnLoad;
</script>

<script id="OnSendEtcEvent" for="EventClientCtrl" event="SendEtcEvent(strEventName,strEventValue)">
	<% if session("ssBctId")="skyer9" then %>
	// alert(strEventValue);
	<% end if %>

	if(strEventValue != 'PING') {
		if(strEventValue == "Bye") {
			alert("서버와 연결이 끊겼습니다. \n다른곳에서 같은 아이디로 로긴되었음");
			objCtrl.DisconnectServer();
		} else {
			parseMessage(strEventValue);
		}
	} else {
		SIPCommand("PONG");
	}

	return false;
</script>
<script id="OnSendNetworkErrorEvent" for="EventClientCtrl" event="SendNetworkErrorEvent()">
	objCtrl.DisconnectServer();
	// setTimeout("chgButton()",1000);
	// alert("서버와 연결 끊음!");
</script>




<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#CCCCCC">

<TABLE cellpadding=0 cellspacing=0 border=0 width=480 bgcolor="#AED8EE">
<tr>
    <td>
        <TABLE cellpadding=0 cellspacing=0 border=0 width=100% bgcolor="#EBF3FC">
        <TR>
            <td class="a" width="180">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td id='memberstatus' width="50%" align=center style='color:#000000;background-color:#EBF3FC;border:0 solid #000000; font-size:9pt;height:20px'>
							<%= ChkIIF(ippbxLocalUser="","아이디 미설정.","Stop") %>
						</td>
						<td id='phonestatus' width="50%" align=center style='color:#000000;background-color:#EBF3FC;border:0 solid #000000; font-size:9pt;height:20px'>
						</td>
					</tr>
				</table>
            </td>
            <TD align=right class="t_txt3" id=UserTR width="10">
				<form name="ippbxlogin_eicn">
					<input type='hidden' id='company_id' value='10x10'>
					<input type='hidden' id='userid' value='<%= ippbxLocalUser %>'>
					<input type='hidden' id='exten' value='<%= localCallNo %>'>
					<input type='hidden' id='passwd' value='cube1010??'>
					<input type='hidden' id='option' value='0'>						<!-- 로긴후 상태(0:대기,2,후처리,3:휴식,4:식사 등) -->
					<input type='hidden' id='usertype' value='M'>					<!-- M:일반상담원 -->
					<input type='hidden' id='serverip' value='110.93.128.96'>
					<input type='hidden' id='serverport' value='8083'>
				</form>
            </TD>
            <td align="right">
                <a href="javascript:popCallRing('','','','','','');">pop</a>
				<!--
				<input type="button" class="button" value="[TEST]로그" onClick="jsPopLog()">
				-->
                &nbsp;
                <input type=button id=LoginButton class="button" onclick="LoginIppbx()" value=" 로 긴 ">
				<input type=button id=LogoutButton class="button" onclick="LogoutIppbx()" value=" 종 료 " style="display:none;">
                &nbsp;
				<input type=button id=RestSetButton class=button value='전화허용중임' onClick='javascript:rest_set()' style="display:none;">
				<input type=button id=RestResetButton class=button value='전화비허용중임' onClick='javascript:rest_reset()' style="display:none;">
            </td>
        </tr>
        </TABLE>
     </td>
</tr>
</table>

<input type="hidden" id="logMsg" value="">

</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

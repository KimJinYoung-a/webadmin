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
<script src="/cscenter/js/socket.io.js"></script>
<script src="/cscenter/js/sha512.js"></script>

<script language=javascript >

var event_num = 0;
var PhoneNum = "";
var PhonePeer = "";
var UserName = "";
var PhoneStatus = "";
var FORWARD_WHEN = "";
var FORWARD_NUM = "";
var MemberStatus = "";
var eicn_mid_connector = "https://ippbx.10x10.co.kr:8089";
var socket = null;

$(window).bind('beforeunload', function() {
	if(socket != null) {
		SendCommand("Bye.");
		// alert("상담원 어플리케이션이 정상적으로 로그아웃되었습니다.");
	}
});

function LoginIppbx()
{
    if($("#company_id").val() == "" || $("#userid").val()=="" || $("#exten").val()=="" || $("#passwd").val() == "") {
        alert("로긴 정보를 입력하세요");
        return;
    }

	ConnectServer();
}

function LogoutIppbx()
{
	location.href = "ippbxlogin_eicn.asp?al=N";
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

function click2call(num)
{
	if ($("#userid").val()=="") {
		alert("로그인 후 이용가능합니다.");
		return false;
	}

	// 빈값이면 대표번호
	SendCommand("CLICKDIAL|," + num + ",oubbound");

	return false;
}

function redirect(num)
{
    SendCommand("REDIRECT|" + num);
    return false;
}

function ConnectServer()
{
	try{
        socket = io.connect(eicn_mid_connector, {
			'secure' : true,
            'reconnect' : true,
            'resource' : 'socket.io'
        });

		var passwd = $('#passwd').val();
		var passwd_sha512 = hex_sha512(passwd);
		$('#passwd').val(passwd_sha512);

        socket.emit('climsg_login', {
            company_id : $('#company_id').val(),
            userid : $('#userid').val(),
            exten : $('#exten').val(),
            passwd : $('#passwd').val(),
            serverip : $('#serverip').val(),
            serverport : $('#serverport').val(),
            usertype : $('#usertype').val(),
            option : $('#option').val()
        });

        socket.on('connect', function() {
            parseMessage("NODEJS|KIND:CONNECT");
        });
        socket.on('svcmsg', function(data) {
            parseMessage(data);
        });
        socket.on('svcmsg_ping', function() {
            socket.emit('climsg_pong');
        });
        socket.on('disconnect', function() {
            parseMessage("NODESVC_STATUS|KIND:DISCONNECT");
        });
        socket.on('error', function() {
            parseMessage("NODESVC_STATUS|KIND:ERROR");
        });
        socket.on('end', function() {
            parseMessage("NODESVC_STATUS|KIND:END");
        });
        socket.on('close', function() {
            parseMessage("NODESVC_STATUS|KIND:CLOSE");
        });
	}catch(error){
        alert("서버가 정상인지 확인후 사용해주세요[1]\n" + error.message);
        LogoutIppbx();
	}
}

function DisconnectServer()
{
	try {
		// socket.disconnect();
		socket = null;
	} catch(error) {
        alert("서버가 정상인지 확인후 사용해주세요[0]");
        LogoutIppbx();
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
        parseLogin(kind,data1,data2,data3,data4,data5,data6,data7,data8);
		SendCommand("LOGIN_ACK");
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
		SendCommand("HANGUP_ACK|"+data5+","+data8);
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

function parseLogin(kind,data1,data2,data3,data4, data5,data6,data7,data8)
{
    //LOGIN|KIND:LOGIN_OK|DATA1:300|DATA2:상담원1|DATA3:0|DATA4:OK|DATA5:11110002
    if(kind == "LOGIN_OK") {
        PhoneNum=data1;
        PhonePeer=data5;
        UserName=data2;
        MemberStatus = data3;
        PhoneStatus = data4;
        FORWARD_WHEN = data6;
        FORWARD_NUM = data7;
        RECORD_TYPE = data8;

        doLogin();
    } else if(kind == "LOGOUT"){
        alert("로그아웃");
    } else {
        alert("로긴 실패");
    }

    return;
}

function parseLogout(kind)
{
	/*
	PhoneNum = "";
    PhonePeer = "";
    UserName = "";
    MemberStatus = "";
    PhoneStatus = "";
    FORWARD_WHEN = "";
    FORWARD_NUM = "";
   RECORD_TYPE = "";
   */

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

function doLogin()
{
	$("#LoginButton").hide();
	$("#LogoutButton").show();

	$("#RestSetButton").show();
	$("#RestResetButton").hide();

	parseForwarding(FORWARD_NUM,FORWARD_WHEN);
	parsePhoneStatus(PhoneStatus);
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

function parseMemberStatus(kind)
{
    MemberStatus = kind;

	$("#RestSetButton").attr("disabled", "disabled");
	$("#RestResetButton").attr("disabled", "disabled");

    if(kind =='0')
    {
		$("#RestSetButton").removeAttr("disabled");
		$("#memberstatus").html(UserName+" " + "<font color='blue'>대기</font>");
    } else if(kind =='1') {
        $("#memberstatus").html(UserName+" " + "<font color='green'>상담중</font>");
    } else if(kind =='2') {
		$("#RestResetButton").removeAttr("disabled");
        $("#memberstatus").html(UserName+" " + "<font color='green'>후처리</font>");
    } else if(kind =='3') {
        $("#memberstatus").html(UserName+" " + "<font color='green'>휴식</font>");
    } else if(kind =='4') {
        $("#memberstatus").html(UserName+" " + "<font color='green'>식사</font>");
    } else if(kind =='5') {
        $("#memberstatus").html(UserName+" " + "<font color='green'>수신거부</font>");
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

function SendCommand(strCommand)
{
	if(PhoneNum == null || PhoneNum == "") {
		alert("로긴후 사용해주세요");
		return false;
	}

	var cmd = "";
	if(strCommand == 'Bye.' || strCommand == 'BYE' )
	{
		displayText("S", "SendCommand("+strCommand+")");
		cmd = strCommand;
	} else {
		displayText("S", "SendCommand('CMD|"+strCommand+"')");
		cmd = "CMD|"+strCommand;
	}
    if(socket != null)
    {
        socket.emit('climsg_command',cmd);
    } else {
        parseMessage("NODESVC_STATUS|KIND:RELOADED");
    }
    return false;
}

function command_memberstatus(s)
{
	SendCommand("MEMBERSTATUS|"+s+","+PhoneNum+","+MemberStatus);
}

function displayText(fsend, text)
{
	return;

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

</head>

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
				<input type="button" class="button" value="pop" onClick="popCallRing('','','','','','');">
				<!-- -->
				<input type="button" class="button" value="로그" onClick="jsPopLog()">
				<!-- -->11
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

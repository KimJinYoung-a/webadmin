<!DOCTYPE html>
<html>
<head>
<title>테스트클라이언트(nodejs)</title>
<meta http-equiv="cache-control" content="no-cache" />
<link rel="stylesheet" type="text/css" href="default.css">
<meta http-equiv="content-type" content="text/html; charset=utf-8">

<script src="/cscenter/js/jquery-1.9.1.js"></script>
<script src="/cscenter/js/socket.io.js"></script>
<script src="/cscenter/js/sha512.js"></script>

<script>

/*********************************************************************
*   화일명:main.js
*   EICN 에서 제공하는 전화 이벤트 연동 javascript
*********************************************************************/
var event_num=0;
var PhoneNum="";
var PhonePeer="";
var UserName="";
var PhoneStatus = "";
var FORWARD_WHEN = "";
var FORWARD_NUM = "";
var MemberStatus = "";
//var eicn_mid_connector = "https://203.84.251.211:8083"; 
var eicn_mid_connector = "https://ippbx.10x10.co.kr:8089";
var socket = null;

$(window).bind('beforeunload', function(){
	if(socket != null)
	{
        	SendCommand("Bye.");
        	alert("상담원 어플리케이션이 정상적으로 로그아웃되었습니다.");
	}
});

//UI 연동
$(document).ready(function() {

	//처음엔 로그아웃 버튼을 숨김
	$("#logout_btn").hide();

        //로긴버튼
        $("#login_btn").click(function(){
		login();
        });
        //로그아웃버튼
        $("#logout_btn").click(function(){
		logout();
        });

	//상태변경버튼 시작
        $("#memberstatus0").click(function(){
		command_memberstatus('0');
        });
        $("#memberstatus1").click(function(){
		command_memberstatus('1');
        });
        $("#memberstatus2").click(function(){
		command_memberstatus('2');
        });
        $("#memberstatus3").click(function(){
		command_memberstatus('3');
        });
        $("#memberstatus4").click(function(){
		command_memberstatus('4');
        });
        $("#memberstatus5").click(function(){
		command_memberstatus('5');
        });
	//상태변경버튼 끝

	//전화걸기버튼
        $("#dial_btn").click(function(){
		click2call();
        });
	//전화받기버튼
        $("#receive_btn").click(function(){
		command_receive();
        });
	//전화끊기버튼
        $("#hangup_btn").click(function(){
		command_hangup();
        });
	//당겨받기버튼
        $("#pickup_btn").click(function(){
		command_pickup();
        });
	//당겨받기버튼1-번호보이게 전화기에서
        $("#pickup_btn1").click(function(){
		command_pickup1();
        });
	//돌려주기버튼-어텐디드
        $("#attended_btn").click(function(){
		command_attended();
        });
	//돌려주기버튼-블라인드
        $("#redirect_btn").click(function(){
		command_redirect();
        });
	//돌려주기 전화끊기버튼
        $("#attended_hangup_btn").click(function(){
		command_attended_hangup();
        });
	//돌려주기버튼-외부어텐디드
        $("#attendedout_btn").click(function(){
		command_attended_out();
        });
	//돌려주기버튼-외부어텐디드
        $("#redirectout_btn").click(function(){
		command_redirect_out();
        });
	//돌려주기버튼-블라인드
        $("#redirecthunt_btn").click(function(){
		command_redirecthunt();
        });
	//착신전환버튼
        $("#forward_btn").click(function(){
		command_forwarding();
        });
	//비우기버튼
        $("#remove_btn").click(function(){
		remove_box();
        });
});
function login()
{
        if($("#company_id").val() == "" || $("#userid").val()==""
		|| $("#exten").val()=="" || $("#passwd").val() == "")
        {
                alert("로긴 정보를 입력하세요");
                return;
        }
	ConnectServer();
}
function remove_box()
{
        event_num=0;
        $("#snd_text").val("");
}
//------------서버연동 ------------
function ConnectServer()
{
try{
        socket = io.connect(eicn_mid_connector, {
                'secure' : true,  
                'reconnect' : true,
				'resource' : 'socket.io',
				secure : true
        });
	//socket = io.connect(eicn_mid_connector);
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
        alert("서버가 정상인지 확인후 사용해주세요");
        logout();
}
}
//------------서버로 명령어보내기 ------------
function SendCommand(strCommand)
{
	if(PhoneNum == null || PhoneNum == "")
	{
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

//-------------클릭투콜----------
function click2call()
{
	var number = $("#number");
	var cid_num = $("#cid");
	if(number.length == 0 || number.val() == "")
	{
		alert("전화번호를 입력하세요");
		return;
	}
        alert(number.val()+"로 전화걸기");
	num = number.val();
	cid = cid_num.val();
	SendCommand("CLICKDIAL|"+cid+","+num+",oubbound");
	displayText("S", "Click2Call('"+cid+"', '"+num+"', 'outbound')");
        return false;
}
//----------버튼 콘트롤------------
function changeLogout()
{
	var logtr = $("#LogTR");
	if(logtr.length==0)
	{
		return;
	}
	//로긴버튼은 감추고 로그아웃버튼을 보여줌
	$("#login_btn").hide();
	$("#logout_btn").show();
	logtr.html( "<img src='left_dot01.gif'> <b>[상담원:"+UserName+"/"+PhoneNum+"/"+PhonePeer+"]</b>");

	parseForwarding(FORWARD_NUM,FORWARD_WHEN);
	parsePhoneStatus(PhoneStatus);
	parseMemberStatus(MemberStatus);
	parseRecordType(RECORD_TYPE);
}

<!------------MESSAGE PARSE START------------>

//로긴
function parseLogin(kind,data1,data2,data3,data4, data5,data6,data7,data8)
{
        //LOGIN|KIND:LOGIN_OK|DATA1:300|DATA2:상담원1|DATA3:0|DATA4:OK|DATA5:11110002
        if(kind == "LOGIN_OK")
        {
                PhoneNum=data1;
                PhonePeer=data5;
                UserName=data2;
                MemberStatus = data3;
                PhoneStatus = data4;
                FORWARD_WHEN = data6;
                FORWARD_NUM = data7;
                RECORD_TYPE = data8;

                //setTimeout("changeLogout()",1000);
                changeLogout();

        } else if(kind == "LOGOUT"){
                alert("로그아웃");
        } else {
                alert("로긴 실패");
        }
        return;
}
function parseCallStatus(kind,data1,data2)
{
        if(kind == "REDIRECT")
        {
                if(data2 == "NOCHAN")
                {
                        //alert("돌려주기할 채널이 없음");
                        return;
                } else if(data2 == "BUSY") {
                        //alert(data1+"이 통화중");
                        return;
                }
        }
}
function parseCallEvent(kind,data1,data2,data3,data4,data5,data6)
{
/*
        if(kind == "IR")
        {
                alert("**"+PhoneNum+" 인바운드 전화가 ["+data1+"]에서 왔음");
        } else if(kind == "ID") {
                alert("**"+PhoneNum+" 인바운드 전화 ["+data1+"]와 통화중");
        } else if(kind == "OR") {
                alert("**"+PhoneNum+" 아웃바운드 전화 ["+data1+"]와 시도중");
        } else if(kind == "OD") {
                alert("**"+PhoneNum+" 아웃바운드 전화 ["+data1+"]와 통화중");
        } else if(kind == "PICKUP") {
                alert("**"+PhoneNum+" 당겨받기 전화 ["+data1+"]와 통화중");
        }
*/
}

function parseHangupEvent(kind,data1,data2,data3,data4)
{
        //alert("**"+PhoneNum+" 전화끊음 ["+data1+","+data2+"]");
}

function parseNodeSvc(kind)
{
	alert("Nodejs 서버 장애["+kind+"]");
	//alert("로그아웃됨");
//	location.reload();
}
function parseLogout(kind)
{
	//alert("로그아웃됨");
//	location.reload();
}
function parseBye(kind, uid, name)
{
	alert("["+kind+"]"+name+"("+uid+")");
	//alert("로그아웃됨");
//	location.reload();
}
function parseMemberStatus(kind)
{
        MemberStatus = kind;
        if(kind =='0')
        {
                document.all.memberstatus0.style.background = "red";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
        } else if(kind =='1') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "red";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
        } else if(kind =='2') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "red";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
        } else if(kind =='3') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "red";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "white";
        } else if(kind =='4') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "red";
                document.all.memberstatus5.style.background = "white";
        } else if(kind =='5') {
                document.all.memberstatus0.style.background = "white";
                document.all.memberstatus1.style.background = "white";
                document.all.memberstatus2.style.background = "white";
                document.all.memberstatus3.style.background = "white";
                document.all.memberstatus4.style.background = "white";
                document.all.memberstatus5.style.background = "red";
        } else {
        }
}
function parseRecordType(type)
{
	var label = "녹취형태:";
	if(type == '')
	{
		return;
	}
	var rec = $("#record_type");
	if(rec.length>0)
	{
		if(type == 'M')
		{
			rec.html(label+"전수녹취");
		} else if(type == 'P') {
			rec.html(label+"부분녹취");
		}
	}
}
function parseForwarding(num, when)
{
	var label = "착신전환상태:";
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
               	$("#forwardstatus").css("background","white");
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
                phonestatus.css("background","lightgreen");
        } else if(kind =='NOK' || kind=='UNREACHABLE' || kind=='UNREGISTERED') {
                phonestatus.css("background","gray");
        } else {
                phonestatus.css("background","white");
        }
}
function parseMessage(msg)
{
//alert(msg);
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
        var data10 = Insp["DATA10"];
        var data11 = Insp["DATA11"];


        if(event == "LOGIN")
        {
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
                //alert("ELSE:"+msg);
        }
        return;
}


//UI연동////////////////////////////////////////////////////////////////////////////

function logout()
{
	var rtn = confirm("로그아웃하시겠습니까?");

	if(rtn == true)
	{
		SendCommand("BYE");
	}
        return false;
}
//내선-어텐디드 
function command_attended()
{
        if($("#transfer_num").val() == "")
        {
                alert("돌려줄 상담원의 내선을 입력하세요");
                return;
        } else {
                var rtn = confirm("["+$("#transfer_num").val()+"] 로 전화를 돌리시겠습니까?");

                if(rtn == false)
                {
                        return;
                }
        }
        SendCommand("ATTENDED|"+$("#transfer_num").val());
        return false;
}
//내선-블라인드 
function command_redirect()
{
        if($("#transfer_num").val()=="")
        {
                alert("돌려줄 상담원의 내선을 입력하세요");
                return;
        } else {
                var rtn = confirm("["+$("#transfer_num").val()+"] 로 전화를 돌리시겠습니까?");

                if(rtn == false)
                {
                        return;
                }
        }
        SendCommand("REDIRECT|"+$("#transfer_num").val());
        return false;
}
//내선-블라인드 
function command_redirecthunt()
{
        if($("#redirecthunt_num").val()=="")
        {
                alert("돌려줄 번호(헌트,대표)를 입력하세요");
                return;
        } else {
                var rtn = confirm("["+$("#redirecthunt_num").val()+"] 로 전화를 돌리시겠습니까?");

                if(rtn == false)
                {
                        return;
                }
        }
        SendCommand("REDIRECT_HUNT|"+$("#redirecthunt_num").val());
        return false;
}
//외부-어텐디드 
function command_attended_out()
{
        if($("#transferout_num").val() == "")
        {
                alert("돌려줄 번호를 입력하세요");
                return;
        } else {
                var rtn = confirm("["+$("#transferout_num").val()+"] 로 전화를 돌리시겠습니까?");

                if(rtn == false)
                {
                        return;
                }
        }
        SendCommand("ATTENDED_OUT|"+$("#transferout_num").val());
        return false;
}
//외부-블라인드 
function command_redirect_out()
{
        if($("#transferout_num").val() == "")
        {
                alert("돌려줄 번호를 입력하세요");
                return;
        } else {
                var rtn = confirm("["+$("#transferout_num").val()+"] 로 전화를 돌리시겠습니까?");

                if(rtn == false)
                {
                        return;
                }
        }
        SendCommand("REDIRECT_OUT|"+$("#transferout_num").val());
        return false;
}
function command_lastevent()
{
        SendCommand("GET_LASTEVENT|callevent");
        return false;
}
function command_rec(mode)
{
	if(mode == 'start')
	{
        	SendCommand("REC_START|"+PhonePeer);
	} else {
        	SendCommand("REC_STOP|"+PhonePeer);
	}
        return false;
}
function command_memberstatus(s)
{
/*
        if(MemberStatus == '5')
        {
                alert("전화기가 수신거부상태입니다.전화기에서 풀어주세요");
                return;
        }
*/
        SendCommand("MEMBERSTATUS|"+s+","+PhoneNum+","+MemberStatus);
}
function command_hangup()
{
        SendCommand("HANGUP|"+PhonePeer);
}
function command_attended_hangup(){
	SendCommand( "ATTENDEDHANGUP|"+PhonePeer );
}
function command_receive()
{
        SendCommand("RECEIVE|"+PhonePeer);
}
function command_reject()
{
        SendCommand("REJECT|"+PhonePeer);
}
function command_pickup()
{
        SendCommand("PICKUP|"+PhonePeer);
}
function command_pickup1()
{
        SendCommand("PICKUP1|"+PhonePeer);
}
function selectForward(value)
{
	FORWARD_WHEN = value;
}
function command_forwarding()
{
	if(FORWARD_WHEN != 'N' && $("#forwarding").val()=='')
	{
		alert("착신전환할 번호를 입력해주세요");
		return;
	}
        SendCommand("FORWARDING|"+PhoneNum+","+$("#forwarding").val()+","+FORWARD_WHEN);
}
function displayText(fsend, text)
{
        event_num = event_num+1;
        if(fsend == "S")
        {
                $("#snd_text").val($("#snd_text").val()+"\nC->S["+event_num+"] "+text);
        } else {
                $("#snd_text").val($("#snd_text").val()+"\nS->C["+event_num+"] "+text);
        }
}



</script>

</head>
<body leftmargin="0" topmargin="0">
<br>
&nbsp;<b>테스트클라이언트(socket.io) 1111</b><br>
<br>
<form name=input_form id=input_form method="post">
<input type='hidden' name='option' id='option' value='0'>
<input type='hidden' name='usertype' id='usertype' value='M'>
<input type='hidden' name='serverip' id='serverip' value='203.84.251.211'>
<input type='hidden' name='serverport' id='serverport' value='8083'>

<table width="100%" border="0" cellpadding="5" cellspacing="5">
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
	<td>
		<table border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<DIV id=LogTR style='display:;border:solid 0'>
			<img src='left_dot01.gif'>
			<b>COMP_ID : </b><input type=text size=10 name='company_id' id='company_id' value='10x10'>&nbsp;
			<b>ID : </b><input type=text size=10 name='userid' id='userid' value=''>&nbsp;
			<b>PW : </b><input type=text size=10 name='passwd' id='passwd' value=''>
			<b>내선 : </b><input type=text size=10 name='exten' id='exten' value=''>
			</DIV>
		</td>
		<td width=5></td>
		<td>
			<input name="login_btn" id="login_btn" type="button" style="width:60px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" value='로그인'>
			<input name="logout_btn" id="logout_btn" type="button" style="width:60px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" value='로그아웃'>
		</td>
		<td width=5></td>
		<td>
			<table border="1" cellpadding="0" cellspacing="0">
			<tr>
			<td id='phonestatus' align=center style='width:80px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000;'>
				전화기상태
			</td>
			<td id='forwardstatus' align=center style='width:300px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000;'>
				착신전환상태
			</td>
			<td id='record_type' align=center style='width:120px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000;'>
				녹취형태
			</td>
			</tr>
			</table>
		</td>
		</tr>
		</table>
	</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
	<td>
<img src='left_dot01.gif'><b>상담원상태:</b>
		<input type=button name=memberstatus0 id='memberstatus0' value='대기(0)' style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus1 id='memberstatus1' value='상담중(1)' style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus2 id='memberstatus2' value='후처리(2)' style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus3 id='memberstatus3' value='휴식(3)' style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus4 id='memberstatus4' value='식사(4)' style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
		<input type=button name=memberstatus5 id='memberstatus5' value='수신거부(5)' style="width:70px; height:20px;color:#000000;background-color:#FFFFFF;border:1 solid #000000">
	</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
	<td>
		<img src='left_dot01.gif'>
		<b>전화걸기 : </b>
		고객번호: <input type=text size=15 name=number id=number value=''>&nbsp;
		RID:<input type=text size=15 name=cid id=cid value=''>
		<input name="dial_btn" type="button" id='dial_btn' value='전화걸기' style="width:70px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A"><br>
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
	<td>
		<img src='left_dot01.gif'>
		<b>받기,끊기,당겨받기 : </b>
                <input name="receive_btn" id='receive_btn' type="button" value='전화받기' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
                <input name="hangup_btn" id='hangup_btn' type="button" value='전화끊기' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
                <input name="pickup_btn" id='pickup_btn' type="button" value='당겨받기' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
                <input name="pickup_btn1" id='pickup_btn1' type="button" value='당겨받기1' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
        <td align=left >
<img src='left_dot01.gif'>
<b>돌려주기(내선) : </b><input type=text size=15 name='transfer_num' id='transfer_num' value=''>
<input name="redirect_btn" type="button" id="redirect_btn" value='돌려주기(BLIND)' style="width:140px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
<input name="attended_btn" type="button" id='attended_btn' value='돌려주기(ATTENDED)' style="width:140px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
<input name="attended_hangup_btn" id='attended_hangup_btn' type="button" value='돌려준전화끊기' style="width:120px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
        <td align=left >
<img src='left_dot01.gif'>
<b>돌려주기(외부로) : </b><input type=text size=15 name='transferout_num' id='transferout_num' value=''>
<input name="redirectout_btn" type="button" id='redirectout_btn' value='돌려주기(BLIND)' style="width:140px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
<input name="attendedout_btn" type="button" id='attendedout_btn' value='돌려주기(ATTENDED)' style="width:140px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
        <td align=left >
		<img src='left_dot01.gif'>
		<b>돌려주기(헌트번호, 대표번호) : </b>
		<input type=text size=15 name='redirecthunt_num' id='redirecthunt_num' value=''>
		<input name="redirecthunt_btn" type="button" id="redirecthunt_btn" value='돌려주기(BLIND)' style="width:140px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
        <td align=left >
		<img src='left_dot01.gif'>
		<b>착신전환 : </b>
		<input type=text name='forwarding' id='forwarding' value='' size=15>
                    <input type="radio" size=15 name='forward_when' id='forward_when' value='N' onClick="selectForward('N')">착신전화안함
                    <input type="radio" name='forward_when' id='forward_when' value='A' onClick="selectForward('A')">항상
                    <input type="radio" name='forward_when' id='forward_when' value='B' onClick="selectForward('B')">통화중
                    <input type="radio" name='forward_when' id='forward_when' value='C' onClick="selectForward('C')">부재중
                    <input type="radio" name='forward_when' id='forward_when' value='T' onClick="selectForward('T')">부재중+통화중
		<input name="forward_btn" type="button" id='forward_btn' value='착신전환' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=25>
        <td align=left >
		<img src='left_dot01.gif'>
		<b>마지막콜이벤트다시받기 : </b>
		<input name="lastevent_btn" type="button" id='lastevent_btn' value='다시받기' style="width:80px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A">
</td>
</tr>
<tr height=1><td bgcolor=lightgray></td></tr>
<tr height=10><td></td></tr>
<tr height=25>
<td>
                <br><b>명령어/이벤트:</b>&nbsp;&nbsp;<input name="remove_btn" type="button" id="remove_btn" value='비우기' style="width:60px; height:20px;color:#FFFFFF;background-color:#51881A;border:0 solid #51881A" ><br>
                <textarea name="snd_text" id='snd_text' cols=160 rows=28 >
                </textarea>
</td>
</tr>
</table>

</form>
</body>
</html>

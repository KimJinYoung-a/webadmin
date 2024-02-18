<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
Dim ippbxLocalCallArr, ippbxMatchUserArr
Dim ippbxLocalUser

ippbxLocalCallArr = Array("901","902","903","904","905","906","907","908","909","910","911","913")
ippbxMatchUserArr = Array("   ","bseo","limpid727","durida22","ilovecozie","greenmon","hasora","908","porco0805","zerogirl0730","wowwooy","icommang")

''ippbxLocalCallArr = Array("901","902","903","904","905","906","907","908","909","801")
''ippbxMatchUserArr = Array("   ","bseo","limpid727","durida22","ilovecozie","greenmon","hasora","908","909","icommang")
''''''''''''''''''''''''''성민정,이수정,임희훈,홍예린,기성숙,,,''coolhas

dim i
for i=LBound(ippbxMatchUserArr) to UBound(ippbxMatchUserArr)
    if (session("ssBctId")<>"") and (session("ssBctId")=ippbxMatchUserArr(i)) then
        ippbxLocalUser = "user" & ippbxLocalCallArr(i)
        Exit For
    end if
next
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>



<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#CCCCCC">
<TABLE cellpadding=0 cellspacing=0 border=0 width=480 bgcolor="#AED8EE">
<tr>
    <td>
        <TABLE cellpadding=0 cellspacing=0 border=0 width=100% bgcolor="#EBF3FC">
        <TR>
            <td class="a" width="120">
                <DIV id=Logininfo style='border:solid 0 red;weight:900;font-size:9pt;height:20px;text-align:right'><%= ChkIIF(ippbxLocalUser="","아이디 미설정.","Stop") %></div>
            </td>
            <td class="a" width="260">
                <DIV id=dispinfo style='display:none;border:solid 0 gray;weight:900;font-size:9pt;height:20px;text-align:center'></div>
            </td>
            <TD align=right class="t_txt3" id=UserTR width="10">
                <INPUT TYPE=hidden size=13 ID=LOGIN_ID NAME=LOGIN_ID VALUE="<%= ippbxLocalUser %>">
            	<INPUT TYPE=hidden size=13 ID=LOGIN_PWD NAME=LOGIN_PWD VALUE="1234">
            </TD>
            <td align="right">
                <a href="javascript:popCallRing('','','','','','');">pop</a>
                &nbsp;
                <input type=button id=ConnButton name=ConnButton class="button" onclick="return loginIppbx()" value=" 로 긴 ">
            </td>
        </tr>
        </TABLE>
     </td>
</tr>
</table>

<script language=javascript src=EventClientCtrlObj.js></script>

<script language=javascript >
var Hndlw = null;
var iWinNo = 0;
var iWinArray = new Array();

function popCallRing(ippbxuser,intel,caller,memoid,iorderserial,iuserid){
    //권한 문제로.. 계속 새창으로 띠울지여부..
    var popwinName = "popCallRing<%= Replace(CStr(FormatDateTime(now(),4)),":","") %><%= Right(CStr(FormatDateTime(now(),3)),2) %>_";
    var arrIdx = 0;
    var isFound = false;
    
    if (iWinArray.length>0){
        /* 무조건 새창으로
        for (var i=0;i<iWinArray.length;i++){
            if (iWinArray[i]){
                try{
                    if (!iWinArray[i].callring.NowDoing){
                        arrIdx = i;
                        popwinName = popwinName + arrIdx;
                        isFound = true;
                        break;
                    }
                }catch(e){
                    //창을 닫은 경우
                    arrIdx = i;
                    popwinName = popwinName + arrIdx;
                    isFound = true;
                    break;
                }
            }
        }
        */
        if (!isFound){
            arrIdx = iWinArray.length;
            popwinName = popwinName + arrIdx;
        }
    }else{
        popwinName = popwinName + arrIdx;
    }
    
    ////var popwin = window.open('/cscenter/history/history_memo_write.asp?ippbxuser=' + ippbxuser + '&intel=' + intel + '&caller=' + caller + '&id=' + memoid + '&orderserial=' + iorderserial + '&userid=' + iuserid,'popCallRing','width=500,height=800,scrollbars=yes,resizable=yes');
    ////var popwin = window.open('/cscenter/ippbxmng/popCallRing.asp?ippbxuser=' + ippbxuser + '&intel=' + intel + '&phoneNumber=' + caller + '&id=' + memoid + '&orderserial=' + iorderserial + '&userid=' + iuserid,popwinName,'width=500,height=800,scrollbars=yes,resizable=yes');
    //주문내역 프레임 않에 있을경우
    var popwin = window.open('/cscenter/ordermaster/ordermasterWithCallRing.asp?ippbxuser=' + ippbxuser + '&intel=' + intel + '&phoneNumber=' + caller + '&id=' + memoid + '&orderserial=' + iorderserial + '&userid=' + iuserid,popwinName,'width=1680,height=1000,scrollbars=yes,resizable=yes');
    popwin.focus();
    iWinArray[arrIdx] = popwin;
    
}


// 여기정보를 수정해 주세요 
//var strServerIP = "203.84.251.210"; //"192.168.1.254"; //사설IP 가능. 구버전
//var strServerIP = "RGE4dUdpYi9IZkAoRE4veEhTQTk="; //203.84.251.211 : 테스트
var strServerIP = "RGE4dUdpYi9IZkAoRE4veEhTPTk="; //203.84.251.210 : 레알


//데몬포트
//var strServerPort = "8083";       //구버전
var strServerPort = "Rjs4L0h2ODw="; //8083

var ISCALL=0;
var STAT=0;
var timerID=null;
var isExtened=0;
var PhoneNum="";
var PhoneCaller="";
var RestStatus="0";
var READY=0;

//------------콘트롤러 상태---------
if(document.all.EventClientCtrl.readyState == 4 ){
	READY=1;
}

//var POPUPURL="http://203.84.251.210/ippbxmng/user/mini_custom.jsp?category=2";
//구버전
function ViewCallerInfo(caller, kind, intel){
	//url=POPUPURL+"&userid="+document.all.LOGIN_ID.value+"&callerCID="+caller;
	//parent.custom_info.location.href=url;
        if(kind != "1")
        {
                //url=POPUPURL+"&userid="+document.all.LOGIN_ID.value+"&callerCID="+caller;
                //parent.custom_info.popupwin(document.all.LOGIN_ID.value,intel, caller);
                //alert(intel + ',' + caller);
                popCallRing(document.all.LOGIN_ID.value,intel, caller,'','','');
        }
}


//window.focus();
function Focus(){
//	if(document.all.chkTop.checked)
//		window.focus();
}



////parent.custom_info.location.href="index.html";
////parent.menu_frame.location.href="index.html";
function addrbook()
{
	isExtened=(isExtened)?0:1;
	document.all.ADDRTABLE.style.display=(isExtened)?"":"none";
	if(isExtened){
		window.resizeTo(WinWidth,WinHeight+document.all.ADDRTABLE.offsetHeight+1);
	}else{
		window.resizeTo(WinWidth,WinHeight);
	}	
}

//-------------클릭투콜----------
function click2call(num)
{
	var calll=document.all.EventClientCtrl.Click2Call(PhoneCaller,num,"outbound");
	return false;
}

function click2dial(id,num,context)
{
		var calll=document.all.EventClientCtrl.Click2Call(id,num,context);
		return false;
}

function chgButton()
{

	var bConnect = document.all.EventClientCtrl.IsConnected();
	document.getElementById("ConnButton").value=(bConnect)?" 종 료 ":" 로 긴 ";
	ISCON=(bConnect)?1:0;		
	//document.all.Logininfo.innerHTML=(bConnect)?document.all.LOGIN_ID.value +" : 내선("+PhoneNum+"/"+PhoneCaller+")":"Stop";
	if(ISCON == 1)
	{
		////parent.custom_info.location.href="inbound.jsp?userid="+document.all.LOGIN_ID.value+"&userInTel="+PhoneNum;
		////parent.menu_frame.location.href="menu.jsp?userid="+document.all.LOGIN_ID.value+"&userInTel="+PhoneNum;
	} else {
		////parent.custom_info.location.href="";
		////parent.menu_frame.location.href="";
	}
	document.all.Logininfo.innerHTML=(bConnect)?document.all.LOGIN_ID.value +" : 내선("+PhoneNum+")":"Stop";
	document.all.UserTR.style.display=(bConnect)?"none":"";
	document.all.dispinfo.style.display=(bConnect)?"":"none";
}

//------------서버로 로긴 ------------
function loginIppbx()
{
	if(READY ==0){
		alert("컨터롤이 시작되지 못했습니다.");
		return false;
	}
	
	if (document.all.LOGIN_ID.value.length<1){
	    alert("ippbx 아이디가 설정 되지 않았습니다. 관리자 문의 요망");
		return false;
	}
	
	document.all.dispinfo.innerHTML="";	
	if(document.all.EventClientCtrl.IsConnected() == true){
		document.all.EventClientCtrl.DisconnectServer();
		setTimeout("chgButton()",1000);
	}else{
        var strLoginID = document.all.LOGIN_ID.value+"_IR_OR";
		var strLoginPwd = document.all.LOGIN_PWD.value;
        document.all.EventClientCtrl.EncConnectServer(strServerIP, strServerPort, strLoginID, strLoginPwd);

	}
	return false;
	
}

//------------서버 로그아웃 ------------
function ippbxLogout()
{
    var bConnect = document.all.EventClientCtrl.IsConnected();
    if(bConnect == true){
            document.all.EventClientCtrl.DisconnectServer();
    }
}

function CheckConnect()
{
	if(document.all.EventClientCtrl != null ) 
	{
		var bConnect = document.all.EventClientCtrl.IsConnected();
		alert( "연결상태 : " + bConnect );
	}
}

function OnPageLoad()
{
	if(document.all.EventClientCtrl != null ) 
	{
		document.all.EventClientCtrl.SetLogMode(true);
		document.all.EventClientCtrl.SetEncryption(false);
	}
}

function SetEncrypt(v)
{
	if(document.all.EventClientCtrl != null ) 
	{
		document.all.EventClientCtrl.SetEncryption(v);
	}
}

//로긴
function parseLogin(kind,data1,data2,data3,data4)
{
	//LOGIN|KIND:LOGIN_OK|DATA1:108|DATA2:이만우|DATA3:1|DATA4:OK (DATA1:번호 DATA2:사용자이름, DATA3:후처리상태(X:헌트멤버아님), DATA4:폰상태(OK,NOK))
	if(kind == "LOGIN_OK")
	{
		PhoneNum=data1;
		PhoneCaller=data2;
		setTimeout("chgButton()",1000);
		RestStatus = data3;
		
	} else {
		document.all.EventClientCtrl.DisconnectServer();
		alert("로긴 실패:");
	}
        return;
}
function parseCallEvent(kind,data1,data2,data3,data4)
{
	if(kind == "IR")
	{
		//alert("인바운드 전화가 ["+data1+"]--->["+PhoneNum+"]에서 왔음");
		//팝업띄우기
		//ViewCallerInfo(PhoneNum,"1",data1);//기존
		
		popCallRing(document.all.LOGIN_ID.value,PhoneNum, data1,'','','');
		
	} else if(kind == "ID") {
		////alert("인바운드 전화 ["+data1+"]와 통화중");
		//disp="통화중";
	} else if(kind == "OR") {
		//disp="거는중";
	} else if(kind == "OD") {
		//disp="통화중";
	}
}
function parseHangupEvent(kind,data1,data2,data3,data4)
{
	////alert("전화끊음");
	STAT=0;
	ISCALL=0;
}

function parseEtc(msg)
{
	var msgs=msg.split("|");
	if(msgs == null || msgs.length < 2)
	{
		return;
	}
	var Insp=new Object();
	Insp["EVENT"]=msgs[0];
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

    if(Insp["EVENT"] == "LOGIN")
    {
		parseLogin(kind,data1,data2,data3,data4);
		return;
	}
	else if(Insp["EVENT"] == "CALLEVENT")
	{
		parseCallEvent(kind,data1,data2,data3,data4);
		return;
        } 
	else if(Insp["EVENT"] == "HANGUPEVENT") {

		parseHangupEvent(kind,data1,data2,data3,data4);
		return;

        } else if(Insp["EVENT"] == "RESETREST") {

		RestStatus = "0";
		document.all.restinfo.innerHTML="<input type=button value='전화허용중임' STYLE='width:100; height:24;' onClick='javascript:rest_set()'>";

        } else if(Insp["EVENT"] == "SETREST") {
		RestStatus = "1";
		document.all.restinfo.innerHTML="<input type=button value='전화비허용중임' STYLE='width:100; height:24; color:red;' onClick='javascript:rest_reset()'>";
	} else {
		//alert("ELSE:"+msg);
	}
        return;
}


//구버전
function parseMsg(msg){
//window.focus();
    //alert(msg);
	var msgs=msg.split("|");
	var Insp=new Object();
	Insp["EVENT"]=msgs[0];
	var disp="";
	for(i=1;i<msgs.length;i++){
		keyval=msgs[i].split(":");
		Insp[keyval[0]]=keyval[1];	
	}	
	
	var caller=Insp["CALLERID"];
	var caller1=Insp["CALLER1ID"];
	var caller2=Insp["CALLER2ID"];
	var channel1=Insp["CHANNEL1"];
	var channel2=Insp["CHANNEL2"];
	var msg=Insp["MSG"];
	var status=Insp["STATUS"];
	clearTimeout(timerID);
	if(Insp["EVENT"] == "RINGEVENT"){
		disp=Insp["CALLERID"];
		disp+=(Insp["ISDIAL"] == "1")?"로 전화를 걸고있습니다.":"에서 전화가 오고 있습니다.";
		ViewCallerInfo(Insp["CALLERID"],Insp["ISDIAL"],PhoneNum);
		ISCALL=Insp["ISDIAL"];
		STAT=1;
	}else if(Insp["EVENT"] == "LOGINRESULT"){
		if(status == "1"){
			linfos=msg.split(",");
			PhoneNum=linfos[0];
			PhoneCaller=linfos[1];
			setTimeout("chgButton()",1000);
		}else{
			 document.all.EventClientCtrl.DisconnectServer();

			alert("로긴 실패:"+msg);

		}


	}else if(Insp["EVENT"] == "CHANNELLIST"){
		var CALL=(ISCALL == 0)?caller1:caller2;
		disp=CALL+"와 통화 중입니다.";
		STAT=2;
	}else if(Insp["EVENT"] == "CHANNELOUT"){
		disp="통화종료되었습니다.";
		STAT=0;
		ISCALL=0;
		timerID=setTimeout("clearInfo()",4000);
	}
	document.all.dispinfo.innerHTML=disp;	
	return false;
}
function clearInfo(){
	document.all.dispinfo.innerHTML="";	
}


function chkResult(){
	var res=document.all.resulttext.value;
	document.all.resulttext.value="";
	alert(res);

}

//------------서버로 명령어보내기 ------------
function SIPCommand(strCommand)
{
	 if(strCommand != "" && strCommand != "undefined" )
        {
                document.all.EventClientCtrl.SendSIPCommand("CMD|"+strCommand);
        }
        return false;
}


//돌려주기
function redirect(num)
{
        SIPCommand("REDIRECT|"+PhoneNum+",outbound,"+num);
        return false;
}
//후처리시작
function rest_set()
{
        SIPCommand("SETREST|"+PhoneNum);
        return false;
}
//후처리끝
function rest_reset()
{
        SIPCommand("RESETREST|"+PhoneNum);
        return false;
}


//function CommandResultEvent(bstrCommandResult){
//	document.all.resulttext.value+=bstrCommandResult;

//}
//function EtcEvent(strEventName,strEventValue){
//	alert(strEventName+","+strEventValue);
//}

function Click2CallBox(comp){
    if (comp.value.length<7){
        alert('전화번호를 입력하세요.');
        comp.focus();
        return;
    }
    
    click2call(comp.value);
}
</script>

<script id="OnSendEtcEvent" for="EventClientCtrl" event="SendEtcEvent(strEventName,strEventValue)">
        <% if session("ssBctId")="icommang" then %>
        alert(strEventValue);
        <% end if %>
                
        if(strEventValue != 'aaaa')
        {
                
                parseEtc(strEventValue);
                
        }
        return false;
</script>
<script id="OnSendNetworkErrorEvent" for="EventClientCtrl" event="SendNetworkErrorEvent()">
        document.all.EventClientCtrl.DisconnectServer();
	setTimeout("chgButton()",1000);
	alert("서버와 연결 끊음!");
</script>

<script language='javascript'>

function getOnLoad(){
    <% if (ippbxLocalUser<>"") then %>
    loginIppbx();
    
    //js 권한문제로 새로고침되면 무조건 띠움.. => 주문내역창과 통합.
    //popCallRing('','','','','','');
    <% end if %>
}

window.onload = getOnLoad;
</script>
</body>
</html>
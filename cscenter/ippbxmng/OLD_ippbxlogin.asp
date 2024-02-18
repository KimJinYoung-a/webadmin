<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
Dim ippbxLocalCallArr, ippbxMatchUserArr
Dim ippbxLocalUser

ippbxLocalCallArr = Array("901","902","903","904","905","906","907","908","909","910","911")
ippbxMatchUserArr = Array("   ","bseo","limpid727","durida22","ilovecozie","greenmon","hasora","908","porco0805","zerogirl0730","wowwooy")
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
                <input type=button id=ConnButton name=ConnButton class="button" onclick="return sConnect()" value=" 로 긴 ">
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
var strServerIP = "203.84.251.210"; //"192.168.1.254"; //사설IP 가능.
var strServerPort = "8083";
var POPUPURL="http://203.84.251.210/ippbxmng/user/mini_custom.jsp?category=2";
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
timer=setInterval("Focus()",1000);
var ISCON=0;
var ISCALL=0;
var STAT=0;
var READY=0;
var timerID=null;
var isExtened=0;
var PhoneNum="";
var PhoneCaller="";
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

function click2call(num)
{
    
    //// var calll=document.all.EventClientCtrl.Click2Call(PhoneCaller,num,"outbound");
	var calll=document.all.EventClientCtrl.Click2Call("0216446030",num,"outbound");
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
function sConnect()
{
	if(!READY ){
		alert("컨터롤이 시작되지 못했습니다.");
		return false;
	}
	
	if (document.all.LOGIN_ID.value.length<1){
	    alert("ippbx 아이디가 설정 되지 않았습니다. 관리자 문의 요망");
		return false;
	}
	
	document.all.dispinfo.innerHTML="";	
	if(ISCON){
		document.all.EventClientCtrl.DisconnectServer();
		setTimeout("chgButton()",1000);
	}else{
		var strLoginID = document.all.LOGIN_ID.value;
		var strLoginPwd = document.all.LOGIN_PWD.value;
		document.all.EventClientCtrl.ConnectServer(strServerIP, strServerPort, strLoginID, strLoginPwd);

	//	var strMessage = document.all.EventClientCtrl.GetLogMessage();
	}
	//setTimeout("chgButton()",1000);
	return false;	
}

function CheckConnect()
{
	if(document.all.EventClientCtrl != null ) 
	{
		var bConnect = document.all.EventClientCtrl.IsConnected();
		alert( "연결상태 : " + bConnect );
	}
}

function OnRingEvent( bstrRingEvent )
{
	//window.open("NewCall.html?CallDAta="+bstrRingEvent);
}
if(document.all.EventClientCtrl.readyState == 4 ){
 READY=1;
}
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

function chkResult(){
	var res=document.all.resulttext.value;
	document.all.resulttext.value="";
	alert(res);

}
function SIPCommand(strCommand)
{
	if(strCommand != "" && strCommand != "undefined" && READY == 1 && document.all.EventClientCtrl != null ) 
	{
		document.all.EventClientCtrl.SendSIPCommand("CMD|"+strCommand);
	}
	setTimeout("chkResult()",4000);
	return false;
}
function CommandResultEvent(bstrCommandResult){
	document.all.resulttext.value+=bstrCommandResult;

}
function EtcEvent(strEventName,strEventValue){
//	alert(strEventName+","+strEventValue);

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


<script id="OnSendRingEvent" for="EventClientCtrl" event="SendRingEvent(bstrRingEvent)">
	parseMsg(bstrRingEvent);
</script>

<script id="OnSendChannelListEvent" for="EventClientCtrl" event="SendChannelListEvent(bstrChannelList)">
	parseMsg(bstrChannelList);
</script>
<script id="OnSendChannelOutEvent" for="EventClientCtrl" event="SendChannelOutEvent(bstrChannelOut)">
	parseMsg(bstrChannelOut);
</script>
<script id="OnSendLoginResultEvent" for="EventClientCtrl" event="SendLoginResultEvent(bstrLoginResult)">
	parseMsg(bstrLoginResult);
</script>

<script id="OnSendCommandResultEvent" for="EventClientCtrl" event="SendCommandResultEvent(bstrCommandResult)">
	CommandResultEvent(bstrCommandResult);
</script>

<script id="OnSendEtcEvent" for="EventClientCtrl" event="SendEtcEvent(strEventName,strEventValue)">
	EtcEvent(strEventName,strEventValue);
</script>
<script id="OnSendNetworkErrorEvent" for="EventClientCtrl" event="SendNetworkErrorEvent()">
               document.all.EventClientCtrl.DisconnectServer();
               setTimeout("chgButton()",1000);
        alert("서버와 연결 끊음!");
</script>


<script language='javascript'>

function getOnLoad(){
    <% if (ippbxLocalUser<>"") then %>
    sConnect();
    
    //js 권한문제로 새로고침되면 무조건 띠움.. => 주문내역창과 통합.
    //popCallRing('','','','','','');
    <% end if %>
}

window.onload = getOnLoad;
</script>
</body>
</html>
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
''''''''''''''''''''''''''������,�̼���,������,ȫ����,�⼺��,,,''coolhas

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
                <DIV id=Logininfo style='border:solid 0 red;weight:900;font-size:9pt;height:20px;text-align:right'><%= ChkIIF(ippbxLocalUser="","���̵� �̼���.","Stop") %></div>
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
                <input type=button id=ConnButton name=ConnButton class="button" onclick="return loginIppbx()" value=" �� �� ">
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
    //���� ������.. ��� ��â���� ���������..
    var popwinName = "popCallRing<%= Replace(CStr(FormatDateTime(now(),4)),":","") %><%= Right(CStr(FormatDateTime(now(),3)),2) %>_";
    var arrIdx = 0;
    var isFound = false;
    
    if (iWinArray.length>0){
        /* ������ ��â����
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
                    //â�� ���� ���
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
    //�ֹ����� ������ �ʿ� �������
    var popwin = window.open('/cscenter/ordermaster/ordermasterWithCallRing.asp?ippbxuser=' + ippbxuser + '&intel=' + intel + '&phoneNumber=' + caller + '&id=' + memoid + '&orderserial=' + iorderserial + '&userid=' + iuserid,popwinName,'width=1680,height=1000,scrollbars=yes,resizable=yes');
    popwin.focus();
    iWinArray[arrIdx] = popwin;
    
}


// ���������� ������ �ּ��� 
//var strServerIP = "203.84.251.210"; //"192.168.1.254"; //�缳IP ����. ������
//var strServerIP = "RGE4dUdpYi9IZkAoRE4veEhTQTk="; //203.84.251.211 : �׽�Ʈ
var strServerIP = "RGE4dUdpYi9IZkAoRE4veEhTPTk="; //203.84.251.210 : ����


//������Ʈ
//var strServerPort = "8083";       //������
var strServerPort = "Rjs4L0h2ODw="; //8083

var ISCALL=0;
var STAT=0;
var timerID=null;
var isExtened=0;
var PhoneNum="";
var PhoneCaller="";
var RestStatus="0";
var READY=0;

//------------��Ʈ�ѷ� ����---------
if(document.all.EventClientCtrl.readyState == 4 ){
	READY=1;
}

//var POPUPURL="http://203.84.251.210/ippbxmng/user/mini_custom.jsp?category=2";
//������
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

//-------------Ŭ������----------
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
	document.getElementById("ConnButton").value=(bConnect)?" �� �� ":" �� �� ";
	ISCON=(bConnect)?1:0;		
	//document.all.Logininfo.innerHTML=(bConnect)?document.all.LOGIN_ID.value +" : ����("+PhoneNum+"/"+PhoneCaller+")":"Stop";
	if(ISCON == 1)
	{
		////parent.custom_info.location.href="inbound.jsp?userid="+document.all.LOGIN_ID.value+"&userInTel="+PhoneNum;
		////parent.menu_frame.location.href="menu.jsp?userid="+document.all.LOGIN_ID.value+"&userInTel="+PhoneNum;
	} else {
		////parent.custom_info.location.href="";
		////parent.menu_frame.location.href="";
	}
	document.all.Logininfo.innerHTML=(bConnect)?document.all.LOGIN_ID.value +" : ����("+PhoneNum+")":"Stop";
	document.all.UserTR.style.display=(bConnect)?"none":"";
	document.all.dispinfo.style.display=(bConnect)?"":"none";
}

//------------������ �α� ------------
function loginIppbx()
{
	if(READY ==0){
		alert("���ͷ��� ���۵��� ���߽��ϴ�.");
		return false;
	}
	
	if (document.all.LOGIN_ID.value.length<1){
	    alert("ippbx ���̵� ���� ���� �ʾҽ��ϴ�. ������ ���� ���");
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

//------------���� �α׾ƿ� ------------
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
		alert( "������� : " + bConnect );
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

//�α�
function parseLogin(kind,data1,data2,data3,data4)
{
	//LOGIN|KIND:LOGIN_OK|DATA1:108|DATA2:�̸���|DATA3:1|DATA4:OK (DATA1:��ȣ DATA2:������̸�, DATA3:��ó������(X:��Ʈ����ƴ�), DATA4:������(OK,NOK))
	if(kind == "LOGIN_OK")
	{
		PhoneNum=data1;
		PhoneCaller=data2;
		setTimeout("chgButton()",1000);
		RestStatus = data3;
		
	} else {
		document.all.EventClientCtrl.DisconnectServer();
		alert("�α� ����:");
	}
        return;
}
function parseCallEvent(kind,data1,data2,data3,data4)
{
	if(kind == "IR")
	{
		//alert("�ιٿ�� ��ȭ�� ["+data1+"]--->["+PhoneNum+"]���� ����");
		//�˾�����
		//ViewCallerInfo(PhoneNum,"1",data1);//����
		
		popCallRing(document.all.LOGIN_ID.value,PhoneNum, data1,'','','');
		
	} else if(kind == "ID") {
		////alert("�ιٿ�� ��ȭ ["+data1+"]�� ��ȭ��");
		//disp="��ȭ��";
	} else if(kind == "OR") {
		//disp="�Ŵ���";
	} else if(kind == "OD") {
		//disp="��ȭ��";
	}
}
function parseHangupEvent(kind,data1,data2,data3,data4)
{
	////alert("��ȭ����");
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
		document.all.restinfo.innerHTML="<input type=button value='��ȭ�������' STYLE='width:100; height:24;' onClick='javascript:rest_set()'>";

        } else if(Insp["EVENT"] == "SETREST") {
		RestStatus = "1";
		document.all.restinfo.innerHTML="<input type=button value='��ȭ���������' STYLE='width:100; height:24; color:red;' onClick='javascript:rest_reset()'>";
	} else {
		//alert("ELSE:"+msg);
	}
        return;
}


//������
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
		disp+=(Insp["ISDIAL"] == "1")?"�� ��ȭ�� �ɰ��ֽ��ϴ�.":"���� ��ȭ�� ���� �ֽ��ϴ�.";
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

			alert("�α� ����:"+msg);

		}


	}else if(Insp["EVENT"] == "CHANNELLIST"){
		var CALL=(ISCALL == 0)?caller1:caller2;
		disp=CALL+"�� ��ȭ ���Դϴ�.";
		STAT=2;
	}else if(Insp["EVENT"] == "CHANNELOUT"){
		disp="��ȭ����Ǿ����ϴ�.";
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

//------------������ ��ɾ���� ------------
function SIPCommand(strCommand)
{
	 if(strCommand != "" && strCommand != "undefined" )
        {
                document.all.EventClientCtrl.SendSIPCommand("CMD|"+strCommand);
        }
        return false;
}


//�����ֱ�
function redirect(num)
{
        SIPCommand("REDIRECT|"+PhoneNum+",outbound,"+num);
        return false;
}
//��ó������
function rest_set()
{
        SIPCommand("SETREST|"+PhoneNum);
        return false;
}
//��ó����
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
        alert('��ȭ��ȣ�� �Է��ϼ���.');
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
	alert("������ ���� ����!");
</script>

<script language='javascript'>

function getOnLoad(){
    <% if (ippbxLocalUser<>"") then %>
    loginIppbx();
    
    //js ���ѹ����� ���ΰ�ħ�Ǹ� ������ ���.. => �ֹ�����â�� ����.
    //popCallRing('','','','','','');
    <% end if %>
}

window.onload = getOnLoad;
</script>
</body>
</html>
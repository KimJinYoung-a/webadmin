<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/NoUSBAllowIpList.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/admin/scmBGCls.asp" -->
<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dim UserOsInfo
dim vSavedID,vSavedEno
dim sBGImg, ClsscmBG
dim lgnMethod
lgnMethod = requestCheckVar(trim(request("lgnMethod")),1)
if lgnMethod ="" then lgnMethod ="U"
UserOsInfo = Request.ServerVariables("HTTP_USER_AGENT")
vSavedID = tenDec(request.cookies("SCMSave")("SAVED_ID"))
vSavedEno = tenDec(request.cookies("SCMSave")("SAVED_Eno"))
  
  
''USB �������� �α��� üũ
'Dim NoUsbValidIP : NoUsbValidIP = fnIsNoUsbAllowIp
Dim NoUsbValidIP
NoUsbValidIP = fncheckAllowIPWithByDB("Y", "", "")

if Application("scmBGdiv") = 0 then
	set ClsscmBG = new CscmBG
	ClsscmBG.fnGetBGUrl
	sBGImg = ClsscmBG.FBGImg 
	set ClsscmBG = nothing
	Application.lock
	Application("scmBG") = sBGImg
	Application("scmBGdiv") = 1
	Application.unlock
end if 


Function fnExistFile(filePath) 
  Dim fso, result 
  Set fso = CreateObject("Scripting.FileSystemObject") 
  If fso.FileExists(filePath) Then  
    result = 1 
  Else 
    result = 0 
  End If 
  fnExistFile = result 
End Function 

'if fnExistFile(Application("scmBG")) = 0 Then Application("scmBG")=""
 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<title>10x10 SCM LOGIN</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="Robots" content="noindex,nofollow">  
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" /> 
<link REL="SHORTCUT ICON" href="http://fiximage.10x10.co.kr/icons/10x10SCM.ico">
<link REL="apple-touch-icon" href="/images/iphone_icon_SCM.png"/>
<style>
html {overflow:auto;}
</style>
<!--[if lt IE 9]>
	<script src="/js/respond.min.js"></script>
	<link rel="stylesheet" type="text/css" href="/css/adminIe.css" />
<![endif]-->
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script> 
<script type="text/javascript" src="/js/jquery.tabslet.min.js"></script>
<script type="text/javascript" src="/js/jquery.tabslet.js"></script>
<!-- ActiveX ��ġ�κ� ���� -->
<OBJECT ID='MaGerAuth' WIDTH='0' HEIGHT='0'	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'>
</OBJECT>
<!-- ActiveX ��ġ�κ� �� -->
<SCRIPT  type="text/javascript">  
// �α��� ���˻�/����
function chkForm() { 
	if(document.frmLogin.lgnMethod.value=="S") {

			 if(!document.frmLogin.usid.value) {
					alert('���̵� �Է����ּ���.'); 
					document.frmLogin.usid.focus();
					return  ;
				}
				
				if(!document.frmLogin.uspwd.value) {
				alert('��й�ȣ�� �Է����ּ���.'); 
				document.frmLogin.uspwd.focus();
				return ;
				}
				
				if(document.frmLogin.sAuthNo.value.length<6) {
				alert('�޴������� ������ ������ȣ�� �Է����ּ���.');
				document.frmLogin.sAuthNo.focus();
				return  ;
				}
							
				document.frmLogin.usn.value=""; 
				document.frmLogin.action="<%=getSCMSSLURL%>/login/dologin.asp";
				 
	} else if(document.frmLogin.lgnMethod.value=="N") {
		
				if(!document.frmLogin.usn.value) {
					alert('����� �Է����ּ���.');
					document.frmLogin.usn.focus();
					return  ;
				}
			 

			if(!document.frmLogin.unpwd.value) {
				alert('��й�ȣ�� �Է����ּ���.');
				document.frmLogin.unpwd.focus();
				return  ;
			}
			
			document.frmLogin.uid.value="";
			document.frmLogin.usid.value="";
			document.frmLogin.action="<%=getSCMSSLURL%>/login/dologinbyempno.asp";	
	} else{

			var sn='';
			try { sn = MaGerAuth.GetSN();   }
			catch(e) {
			    <% if (Not NoUsbValidIP) then %>
				alert('�ٹ����� ������ ActiveX ��Ʈ���� ��ġ���ּ���');
				self.location='/login/USB_Token_Info.htm';
				return  ;
				<% end if %>
			}

			if (sn != '') {
				document.frmLogin.tokenSn.value=sn 
			} else {
			    <% if (Not NoUsbValidIP) then %>
				alert('�ٹ����� ������ ������ USBŰ�� PC�� �Ű� �ٽ� �α������ּ���.');
				return  ; 
				<% end if %>
			}

		 if(!document.frmLogin.uid.value) {
				alert('���̵� �Է����ּ���.');
				document.frmLogin.uid.focus();
				return  ;
			}
			
			if(!document.frmLogin.upwd.value) {
			alert('��й�ȣ�� �Է����ּ���.');
			document.frmLogin.upwd.focus();
			return ;
		 }
		 
		 document.frmLogin.usn.value="";
		 document.frmLogin.action="<%=getSCMSSLURL%>/login/dologin.asp";
}  
  
	document.frmLogin.submit();
 }


// SMS�α��� ������ȣ �߼�
function popSMSAuthNo() {
	if(!document.frmLogin.usid.value) {
		alert('���̵� �Է����ּ���.');
		document.frmLogin.usid.focus();
		return;
	}
	
	if(!document.frmLogin.uspwd.value) {
		alert('��й�ȣ�� �Է����ּ���.'); 
		document.frmLogin.uspwd.focus();
		return ;
	}
	 
	hidFrm.location.href="/admin/member/tenbyten/iframe_adminLogin_SendSMS.asp?uid="+document.frmLogin.usid.value;  
	
}

//sms �����ܰ�
function jsSetStep(iValue){
 if(iValue==2){
 	 document.all.dvid.style.display = 'none'; 
 	 document.all.dvAuth.style.display = ''
 }else{
 	document.all.dvid.style.display = ''; 
 	 document.all.dvAuth.style.display = 'none'
 }
}

// SMS�Է� ī���� �۵�(3�а�:180��)
var iSecond=180;
var timerchecker = null;

function startLimitCounter(cflg) { 
	if(cflg=="new") {
		if(timerchecker != null) {
			alert("�̹� ������ȣ�� �߼��Ͽ����ϴ�.\n�޴����� SMS�� Ȯ�����ּ���.");
			return ;
		}
		iSecond=180;
	}
    rMinute = parseInt(iSecond / 60);
    rSecond = iSecond % 60;
    if(rSecond<10) {rSecond="0"+rSecond};
 
    if(iSecond > 0)
    {
        document.frmLogin.sLimitTime.value = rMinute+":"+rSecond; 
        iSecond--;
        timerchecker = setTimeout("startLimitCounter()", 1000); // 1�� �������� üũ
    }
    else
    {
        clearTimeout(timerchecker);
        document.frmLogin.sLimitTime.value = "0:00";
        timerchecker = null;
        alert("������ȣ �Է� �ð��� ����Ǿ����ϴ�.\n\nSMS�� ���� ���ߴٸ� �ٽ� ��ȣ�� �޾��ּ���.");
    }
}

// �޴�����ȣ ����/����Ȯ�� �˾�
function PopChgHPNum() {
	alert("IP(or USB) �α��� �� ������������ �޴��� ����Ȯ�� �� �̿밡���մϴ�.");
	return;
//	if(confirm("����Ȯ���� ���� ���� ���� ���̵��Դϴ�.\n���� Ȯ���� �����ðڽ��ϱ�?")) {
//		if(!document.frmLogin.usid.value) {
//			alert('���̵� �Է����ּ���.');
//			document.frmLogin.usid.focus();
//			return;
//		} else {
//			var popwin = window.open("pop_ChangeHPIdentify.asp?uid="+document.frmLogin.usid.value,"PopChgHPNum","width=400 height=270 scrollbars=yes");
//			popwin.focus();
//		}
//	}
}

// �����ȳ� �˾�
function popSecLgnInfo(flg) {
	if(flg=="U") {
		var InfoPop = window.open("/login/USB_Token_Info.htm","LoginInfoPop","width=820,height=600,scrollbars=yes");
		InfoPop.focus();
	} else if(flg=="S") {
		var InfoPop = window.open("/login/SMS_Auth_Info.htm","LoginInfoPop","width=690,height=600,scrollbars=yes");
		InfoPop.focus();
	}
}



 $(function(){  
	/* tab */
	$(".tabCont").hide();
	$(".tabNav").find("li:first").addClass("current");	
	$(".tabContainer").find(".tabCont:first").show();
	$(".tabNav li").click(function(){
		$(this).siblings("li").removeClass("current");
		$(this).addClass("current"); 
		$(this).closest(".tabNav").nextAll(".tabContainer:first").find(".tabCont").hide();
		var activeTab = $(this).find("a").attr("href"); 
		$(activeTab).show();
		var tidx =	$(this).index(); 
		if (tidx==1){
			document.frmLogin.lgnMethod.value="S"; 
			document.frmLogin.usid.focus()
		}else if(tidx==2){ 
			document.frmLogin.lgnMethod.value="N"; 
			document.frmLogin.usn.focus()
		}else{
			document.frmLogin.lgnMethod.value="U"; 
			document.frmLogin.uid.focus()
		}
		return false;
	});

	// input action
	$(".inpForm input").focus(function(){
		$(this).addClass('onInput');
		$(this).siblings("label").hide();
	});
	$(".inpForm input").focusout(function(){
		$(this).removeClass('onInput');
		if($(this).val() == ""){
			$(this).siblings("label").show();
		}
	});

	// family site
	$(".tenFamily dt").click(function(){
		if($(".tenFamily dd").is(":hidden")){
			$(this).parent().children('dd').show();
		}else{
			$(this).parent().children('dd').hide();
		};
	});
	$(".tenFamily dd li").click(function(){
		var evtName = $(this).text();
		$(this).parent().parent().parent().children('dt').text(evtName);
		$(this).parent().parent().hide();
		 document.getElementById("hidL").value = evtName; 
	});
	$(".tenFamily dl").mouseleave(function(){
		$(this).children("dd").hide();
	});
});


function jsGoUrl(){
	var strUrl;
	if( document.getElementById("hidL").value=="ONLINE"){
		strUrl = "http://www.10x10.co.kr/"
	}else if(document.getElementById("hidL").value=="OFFLINE"){
		strUrl = "http://www.10x10.co.kr/offshop/index.asp"
		
	}else if(document.getElementById("hidL").value=="THE FINGERS"){
		strUrl = "http://www.thefingers.co.kr/"
	}
	
	var winOp = window.open("about:blank");
	winOp.location.href = strUrl;
} 


$(document).ready(function(){
 document.frmLogin.uid.focus() 
  var iniactiveTab ;
 if(document.getElementById("lgnMethod").value=="N"){  
 		$(".tab3").siblings("li").removeClass("current");
		$(".tab3").addClass("current"); 
		$(".tab3").closest(".tabNav").nextAll(".tabContainer:first").find(".tabCont").hide();
 		 iniactiveTab = "#tab3";
 		$(iniactiveTab).show(); 
 		document.frmLogin.usn.focus()
 }else if (document.getElementById("lgnMethod").value=="S"){ 
 		$(".tab2").siblings("li").removeClass("current");
		$(".tab2").addClass("current"); 
		$(".tab2").closest(".tabNav").nextAll(".tabContainer:first").find(".tabCont").hide();
 		 iniactiveTab = "#tab2";
 		$(iniactiveTab).show(); 
 		document.frmLogin.usid.focus()
 }
 }); 
 
</SCRIPT>
</head> 
<body class="scmLogin <%if Application("scmBG") = "" then%>noImage<%end if%>" <%if Application("scmBG") <> "" then%>style="background-image:url(<%=Application("scmBG")%>);"<%end if%>>
	<h1><span></span><% if (application("Svr_Info")="Dev") then %>[Dev] <% end if %>10X10 SCM LOGIN</h1>
	<form name="frmLogin" method="post" action="<%=getSCMSSLURL%>/login/dologin.asp"  >
    <input type="hidden" name="backpath" value="<%= request("backpath") %>">
    <input type="hidden" name="tokenSn" value="">
    <input type="hidden" name="lgnMethod" id="lgnMethod" value="<%=lgnMethod%>"> 
	<div class="loginBox">
		<div class="tabNav">
			<ul>
				<li class="tab1"><a href="#tab1"><%=CHKIIF(NoUsbValidIP,"IP","USBŰ")%></a></li>
				<li class="tab2"><a href="#tab2">SMS</a></li>
				<li class="tab3"><a href="#tab3">���</a></li>
			</ul>
		</div>
		<div class="tabContainer">
			<!-- IP -->	
			<div id="tab1" class="tabCont"> 
					<p class="inpForm">  
						<label for="memId1">���̵�</label> 
						<input type="text" id="uid" name="uid"  value="<%=vSavedID%>"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwd.focus();"/>
					</p>
					<p class="inpForm">
						<label for="memPw1">��й�ȣ</label>
						<input type="password" id="upwd" name="upwd" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();"/>
					</p>
					<p class="rt">
						<input type="checkbox" id="saved_id" name="saved_id" value="o" <%=chkIIF(vSavedID<>"","checked","")%>/> 
						<label for="saveId1">���̵�����</label>
					</p>
					<div class="btnArea"><button class="btn" type="button" onClick="chkForm()" >�α���</button></div> 
			</div>
			<!-- SMS -->
			<div 	id="tab2" class="tabCont">
				 <div> 
					<p class="inpForm">
						<label for="memId1">���̵�</label>  
						<input type="text" id="usid" name="usid"  value="<%=vSavedID%>"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.uspwd.focus();"/>
					</p>
					<p class="inpForm">
						<label for="memPw1">��й�ȣ</label>
						<input type="password" id="uspwd" name="uspwd" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) popSMSAuthNo();"/>
					</p>
					<p class="rt" >
						<input type="checkbox" id="saved_sid" name="saveds_sid" value="o" <%=chkIIF(vSavedID<>"","checked","")%>/> 
						<label for="saveId1">���̵�����</label>
					</p>
					<div class="btnArea" id="dvid"><button class="btn" type="button" onClick="popSMSAuthNo()" >������ȣ �ޱ�</button> </div>
				</div>
				 <!-- ������ȣ �Է� -->
				<div  id="dvAuth" style="display:none;">
					<p class="timeLimit">�Է���ȿ�ð� <strong><input type="text" name="sLimitTime" id="sLimitTime" value="-:--" readonly  style="width:100px;display:inline-block; margin-top:-4px; padding-left:0.5rem; font-size:2.5rem; font-family:arial; vertical-align:middle;border:0;"></strong></p> 
					<!--<p><button class="btn btnReapply" type="button" onClick="document.frmLogin.lgnStep.value=1;chkForm();">������ȣ ��߼�</button></p> -->
					<p class="inpForm tMar20">
						<label for="smsNum">SMS ������ȣ �Է�</label>
						<input type="text" id="sAuthNo" name="sAuthNo"   value="" AUTOCOMPLETE="off"/>
					</p>
					<div class="btnArea" style="margin-top:0;"><button class="btn" type="button" onClick="chkForm()" >�α���</button></div>
				</div> 
			</div>  
			 
			<!-- ��� -->
			<div id="tab3" class="tabCont">
				<p class="inpForm">
					<label for="staffNum">���</label>
					<input type="text" id="usn" name="usn" value="<%=vSavedEno%>" onKeyPress="if (event.keyCode == 13) frmLogin.unpwd.focus();"/>
				</p>
				<p class="inpForm">
					<label for="memPw3">��й�ȣ</label>
					<input type="password" id="unpwd" name="unpwd" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();" />
				</p>
				<p class="rt">
					<input type="checkbox" id="saved_eno" name="saved_eno"  value="o" <%=chkIIF(vSavedEno<>"","checked","")%>/>
					<label for="saveId3">�������</label>
				</p>
				<div class="btnArea"><button class="btn" type="button" onClick="chkForm()">�α���</button></div>
			</div>
		</div>
	</div>
</form>
	<ul class="help">
		<li><a href="javascript:popSecLgnInfo('U')">USBŰ �����ȳ�</a></li>
		<li><a href="javascript:popSecLgnInfo('S')">SMS �����ȳ�</a></li>
	</ul>
	<p class="slogan">YOU ARE ALREADY DIFFERENT <a href="http://www.10x10.co.kr/" target="_blank">10X10.CO.KR</a></p>
	<div class="tenFamily">		
		<input type="hidden" name="hidL" id="hidL" value="">
		<dl>
			<dt>���� �ٷΰ���</dt>
			<dd>
				<ul>
					<li value="http://www.10x10.co.kr/">ONLINE</li>
					<li>OFFLINE</li>
					<li>THE FINGERS</li>
				</ul>
			</dd>
		</dl>
		<button class="btnGo" type="button" onClick="jsGoUrl();">�̵�</button>
	</div> 
<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</body>
</html>

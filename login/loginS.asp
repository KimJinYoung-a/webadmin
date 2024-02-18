<%@ language="vbscript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%
'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Call serverupdate_underconstruction()

dim manageUrl
IF application("Svr_Info")="Dev" THEN
	manageUrl 	 = "http://testwebadmin.10x10.co.kr"
ELSE
	manageUrl 	 = "http://webadmin.10x10.co.kr"
END IF
dim UserOsInfo
dim vSavedID,saved_id
dim vUserid, vPassWd, vChkAuth, vIsSec
  
UserOsInfo = Request.ServerVariables("HTTP_USER_AGENT") 
vUserid = session("tmpUID")
vPassWd = session("tmpUPWD")
vChkAuth = requestCheckVar(trim(request.Form("chkAuth")),1)
vIsSec  = requestCheckVar(trim(request.Form("hidSec")),1)
saved_id= requestCheckVar(trim(request.Form("saved_id")),1) 
 
if vChkAuth <> "Y" or vUserid="" or vPassWd="" then
	response.write("<script>window.alert('���̵�/1�� ��й�ȣ Ȯ�� �� 2�� ������ �����մϴ�.');</script>")
  response.write("<script>self.location='"&manageUrl&"'</script>")
end if

if inStr(UserOsInfo,"Windows CE")>0 then
	response.redirect manageUrl&"/PDAadmin/indexPDA.asp"
	dbget.close()	:	response.End
end if
%> 
<!-- #include virtual="/partner/lib/adminHead_NoJs.asp" -->
<link REL="SHORTCUT ICON" href="http://fiximage.10x10.co.kr/icons/10x10SCM.ico">
<script type="text/javascript"> 
	// �н����� ���⵵ �˻�
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

    if((numAlpha>0&&numNums>0)||(numAlpha>0&&numSpecials>0)||(numNums>0&&numSpecials>0)) {
    	sRst=true;
    } else {
    	sRst=false;
    }
    return sRst;
}

	function validate(){  
		var frm = document.frmLogin;		
		<%if vIsSec ="N" then '2�� ��� �̼����� ���������ϵ���%>
			
			if(!frm.upwdS1.value){
				 alert("��й�ȣ�� �Էµ��� �ʾҽ��ϴ�.");
				 frm.upwdS1.focus();
				 return;
			}  
			
			if (frm.upwdS1.value.length < 8 || frm.upwdS1.value.length > 16){
			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
			frm.upwdS1.focus();
			return ;
		 }
	
	
			if (!fnChkComplexPassword(frm.upwdS1.value)) {
				alert('�н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
				frm.upwdS1.focus();
				return;
			}
	
		 	if(!frm.upwdS2.value){
					 alert("��й�ȣ�� Ȯ�����ּ���");
					 frm.upwdS2.focus();
					 return;
				}  
				
			if (frm.upwdS1.value!=frm.upwdS2.value){
				alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.");
				frm.upwdS1.focus();
				return ;
			} 
	
	<%else%>
	
		if(!document.frmLogin.upwdS.value){
			 alert("��й�ȣ�� �Էµ��� �ʾҽ��ϴ�.");
			 document.frmLogin.upwdS.focus();
			 return;
		}  
		
	<%end if%> 
	 
		document.frmLogin.submit();
	}
 
 
 function jsSearchPWD(){
 	location.href = "/login/searchPwd.asp";
}

$(function(){
	var contH = $('.loginBoxV16').outerHeight();	
	$('.loginBoxV16').css('margin-top',-contH/2+70+'px');
});

</script>
 
</head> 
<body <%if vIsSec <>"N" then%>onLoad="document.frmLogin.upwdS.focus()"<%end if%>>  
<div   id="login">
	<div class="container scrl">
		<% if (application("Svr_Info")="Dev") then %>
		<h1>This is 2009  Test Server...</h1> 
		<% end if %> 
		<div class="loginBoxV16">
			<h1><img src="/images/partner/admin_login_logo_2016.png" alt="Partner Login - �����ΰ���ä�� �ٹ������� ���»� �������Դϴ�." /></h1>
			<div class="loginCont">
				<form method="post" name="frmLogin" action="<%=getSCMSSLURL%>/login/dologinByPartner.asp">
    			<input type="hidden" name="backpath" value="<%= request("backpath") %>">
    			<input type="hidden" name="loginNo" value="2">
    			<input type="hidden" name="hidSec" value="<%=vIsSec%>"> 
				<div class="loginInput">
					<fieldset>
						<p class="inputArea"><label for="id">���̵�</label><input type="text" id="id"  class="formTxt" value="<%=vUserid%>" disabled="disabled" maxlength="32" /></p>
						<p class="inputArea tPad10"><label for="pwr">1�� ��й�ȣ</label><input type="password" id="pwr"  class="formTxt" value="********" disabled="disabled" /></p>
						<%if vIsSec ="N" then
							dim islongtimeNotUsingID
							islongtimeNotUsingID = IsLongTimeNotLoginUserid(vUserid)
							
							    if (islongtimeNotUsingID) then
							%>
							<div class="cautionMsg">
								<p class="cRd3">��Ⱓ �α��� ������ �����ϴ�.  </p>
								<p class="cRd3"> 2�� ��й�ȣ ���� ������ ����<br/> �����ͷ� �����ּ���</p>
								 <div class="cBl3"><br/>������: 070-4868-1799</div></div>
							<%			
							    else
							%>
						
										<div class="cautionMsg">������ ���� 2�ܰ� ������ �����մϴ�.<br />�α��ο� ����Ͻ� <span class="cRd3">2�� ��й�ȣ�� ����</span>���ּ���.</div>
										<p class="inputArea tPad10"><label for="pwr2">2�� ��й�ȣ</label>
											<input type="password" id="upwdS1" name="upwdS1" class="formTxt" placeholder="2�� ��й�ȣ" maxlength="32"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS2.focus();"/></p>  
									 	<p class="inputArea tPad10"><label for="pwr2">2�� ��й�ȣ Ȯ��</label>
									 	<input type="password" id="upwdS2" name="upwdS2" class="formTxt" placeholder="2�� ��й�ȣ Ȯ��" maxlength="32"  AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) validate();"/></p>
									 	<div class="alertMsg">������� 8~16���� ����/���� ����<br /> ��ҹ��ڱ���</div>
									 	<!-- for dev msg : �α������� �߸��Է��� ��� ���� //-->
										<div id="divMsg" class="alertMsg" style="display:none;">���̵� ��й�ȣ�� �ùٸ��� �ʽ��ϴ�.</div>
										<button type="button" class="loginBtnV16" onClick="validate();">Login</button>
										<p class="tPad10 fs11 cGy3"><input type="checkbox" id="saved_id" class="formCheck" name="saved_id" value="o" <%=chkIIF(saved_id="o","checked","")%>/> <label for="idSave">���̵�����</label></p>
										<span class="helpTxt" onclick="jsSearchPWD();">1 / 2�� ��й�ȣ ã��</span>
					 	    <%	end if %>
						<%else%>
						
							<div class="cautionMsg">2�� ��й�ȣ�� �Է����ּ���</div>
						<p class="inputArea tPad10"><label for="pwr2">2�� ��й�ȣ</label><input type="password" id="upwdS" name="upwdS" class="formTxt"    AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) validate();"/></p>
						<!-- for dev msg : �α������� �߸��Է��� ��� ���� //-->
						<div id="divMsg" class="alertMsg" style="display:none;">���̵� ��й�ȣ�� �ùٸ��� �ʽ��ϴ�.</div>
						<button type="button" class="loginBtnV16" onClick="validate();">Login</button>
						<p class="tPad10 fs11 cGy3"><input type="checkbox" id="saved_id" class="formCheck" name="saved_id" value="o" <%=chkIIF(saved_id="o","checked","")%>/> <label for="idSave">���̵�����</label></p>
						<span class="helpTxt" onclick="jsSearchPWD();">1 / 2�� ��й�ȣ ã��</span>
						<%end if%> 						
					</fieldset>
				</div>
				</form>
				<div class="helpTxtBox" style="display:;">
					<dl>
						<dt>2�ܰ� ������ �� �ϳ���?</dt>
						<dd>�α��� ���Ȱ�ȭ�� ���� 2�ܰ� ������ ����˴ϴ�.<br />������ ���̵�� ��й�ȣ �ܿ� 2�� ��й�ȣ�� �Է��ϴ� ���ߺ��� �����Դϴ�.</dd>
					</dl>
				</div>
			</div>
			<div class="linkWrapV16">
				<ul class="goLink">
					<li class="link01"><a href="http://company.10x10.co.kr/inquiry_write.asp" target="_blank">�ű�����</a></li>
					<li class="link02"><a href="http://www.10x10.co.kr" target="_blank">�¶��μ�</a></li>
					<li class="link03"><a href="http://www.10x10.co.kr/offshop/index.asp" target="_blank">�������μ�</a></li>
					<li class="link04"><a href="http://company.10x10.co.kr/company_04.htm" target="_blank">���ô±�</a></li>
				</ul>
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div>
		</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

 
	  
 
 

 
  
              	  
         
<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%
dim userid, searchID
dim manager_name, manager_hp, manager_shp, manager_email, manager_smail
dim jungsan_name, jungsan_hp, jungsan_shp, jungsan_email, jungsan_smail
dim deliver_name, deliver_hp, deliver_shp, deliver_email, deliver_smail

dim sql,reFAddr
dim recentqcount

 userid  = session("reauthUID")
 reFAddr = request.ServerVariables("REMOTE_ADDR")
 
 
if userid <> "" then
	'�ʱ�ȭ
		manager_name	= ""
	 	manager_hp		= ""
		manager_email	= ""
	 	jungsan_name	= ""
	 	jungsan_hp		= ""
		jungsan_email	= ""
	 	deliver_name	= ""
	 	deliver_hp		= ""
		deliver_email	= ""
	
	'���̵���ȸ �α� ���
	sql = "exec db_partner.dbo.sp_Ten_partner_searchPWD_log '"&userid&"','"&Left(reFAddr,16)&"'"
    dbget.Execute sql
	 	 
	'10�� ���� 10ȸ �̻� �˻��� ���� 	 
	recentqcount = 0 	 
	sql = "select count(idx) as cnt "
	sql = sql & " from db_partner.dbo.tbl_partner_searchPWD_log  "
	sql = sql & " where refip='" + Left(reFAddr,16) + "' "
	sql = sql & " and datediff(n,regdate,getdate())<=10" 
	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
	if not rsget.eof then
		recentqcount = rsget("cnt")
	end if
	rsget.close

	if recentqcount>=10 then
		response.write "<script type='text/javascript'>alert('�ܽð� ���� �������� ������ �����Ͽ����ϴ�.\n��� �� �ٽ� �õ����ּ���.');</script>"
	  
	else

	sql =" select id, manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp " &_
		" ,email as manager_email ,jungsan_email, deliver_email " &_
		" from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
  	    searchID			=rsget("id")
   		manager_name 	=rsget("manager_name")
   		if manager_name <> "" then manager_name= left(manager_name,1)
   		manager_hp =rsget("manager_hp")
   		if manager_hp <> "" then manager_shp= left(manager_hp,4)&"****"&right(manager_hp,5)
		manager_email =rsget("manager_email")
   		if inStr(manager_email,"@")>0 then manager_smail= left(split(manager_email,"@")(0),4)&"****@"&split(manager_email,"@")(1)

   		jungsan_name 	=rsget("jungsan_name")
   		if jungsan_name <> "" then jungsan_name=left(jungsan_name,1)
   		jungsan_hp =rsget("jungsan_hp")
   		if jungsan_hp <> "" then jungsan_shp= left(jungsan_hp,4)&"****"&right(jungsan_hp,5)
		jungsan_email =rsget("jungsan_email")
   		if inStr(jungsan_email,"@")>0 then jungsan_smail= left(split(jungsan_email,"@")(0),4)&"****@"&split(jungsan_email,"@")(1)

   		deliver_name 	=rsget("deliver_name")
   		if deliver_name <> "" then deliver_name=left(deliver_name,1)
   		deliver_hp =rsget("deliver_hp") 
   		if deliver_hp <> "" then deliver_shp= left(deliver_hp,4)&"****"&right(deliver_hp,5)  
		deliver_email =rsget("deliver_email")
   		if inStr(deliver_email,"@")>0 then deliver_smail= left(split(deliver_email,"@")(0),4)&"****@"&split(deliver_email,"@")(1)
    end if
    rsget.close
	end if
end if
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<style type="text/css">
.lineBtnV16 {width:97px;}
.rowUnderline {border-bottom:solid 1px #E0E0E0;}
</style>
<script type="text/javascript">
// SMS�Է� ī���� �۵�(2��30��:150��)
var iSecond=0;
var timerchecker = null;

	
function jsSMSSend(shp){  
	document.frmSMS.sHp.value = shp;
	document.frmSMS.action = "/login/reconfirmip_sendSMS.asp";
	document.frmSMS.submit();
}

function jsEmailSend(shp){  
	document.frmSMS.sHp.value = shp;
	document.frmSMS.action = "/login/reconfirmip_sendEmail.asp";
	document.frmSMS.submit();
}
	


function startLimitCounter(cflg) {  
	
	switch (cflg) {
		case "new":
			iSecond=180;	//SMS 3��
			break;
		case "newMail":
			iSecond=600;	//�̸��� 10��
			break;
	}
	 
    rMinute = parseInt(iSecond / 60);
    rSecond = iSecond % 60;
    if(rSecond<10) {rSecond="0"+rSecond};

    if(iSecond > 0)
    {
        document.frmAuth.sLimitTime.value  = rMinute+":"+rSecond; 
        iSecond--;
        timerchecker = setTimeout("startLimitCounter()", 1000); // 1�� �������� üũ
    }
    else
    {
        clearTimeout(timerchecker);
        document.frmAuth.sLimitTime.value = "0:00";
        timerchecker = null;
        alert("������ȣ �Է� �ð��� ����Ǿ����ϴ�.\n\nSMS�� ���� ���ߴٸ� �ٽ� ��ȣ�� �޾��ּ���.");
    }
}

function jsChkAuthno(){ 
		if(document.frmAuth.sAuthNo.value.length<6) {
			alert('�޴������� ������ ������ȣ�� �Է����ּ���.');
			document.frmAuth.sAuthNo.focus();
			return;
		}
		
		 if(document.frmAuth.sLimitTime.value == "0:00"){
		 	alert("������ȣ �Է� �ð��� ����Ǿ����ϴ�.\n\nSMS�� ���� ���ߴٸ� �ٽ� ��ȣ�� �޾��ּ���.");
		 	return;
		}
		 
		document.frmAuth.submit();
}


</script>
</head>
<body>
<div  id="login">
	<div class="container scrl"><!--class="container scrl"-->
		<div class="pwrBoxV16">
			<div class="titWrap">
				<h1>���� ȯ�� IP ����</h1>
			</div>
			<div class="pwrContWrap">
			    <% if (userid="") then %>
			    <p class="cBk1">��ϵ��� ���� ���̵��̰ų� �߸��� �����Դϴ�. �ٽ� �ѹ� Ȯ�����ּ���</p>
			    
			    <% else %>
				<p class="cBk1">��ϵ� ������� �޴��� �������� ����ȯ�� IP ������ �� �� �ֽ��ϴ�.</p>
					<form name="frmID" method="post">
				
				<%if userid <> "" and searchID ="" then%><p class="tPad10 cRd3">��ϵ��� ���� ���̵��̰ų� �߸��� �����Դϴ�. �ٽ� �ѹ� Ȯ�����ּ���.</p><%end if%>
				</form>
				<div class="sectionWrap" id="idinfo" style="display:<%if searchID ="" then %>none<%end if%>;">
					<h2>���̵� ��ȸ ���</h2>
					<p class="tPad10">�Ʒ� ����� ������ Ȯ���Ͻð� '������ȣ �ޱ�' ��ư�� Ŭ���� �ּ���.</p>
					<form name="frmSMS" method="post" target="hidFrm"  action="/login/reconfirmip_sendSMS.asp">
						<input type="hidden" name="uid" value="<%=userid%>">
						<input type="hidden" name="sHp" value="">
						<input type="hidden" name="sKey" value="<%=md5(userid&"TPUSMS")%>">
					<table class="resultList">
						<colgroup>
							<col width="*" /><col width="60px" /><col width="160px" /><col width="100px" />
						</colgroup>
						<tr>
							<td rowspan="2" class="rowUnderline">���������</td>
							<td rowspan="2" class="rowUnderline"><%=manager_name%>**</td>
							<td><%=manager_shp%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsSMSSend('M');">�������� �ޱ�</button></td>
						</tr>
						<tr>
							<td class="rowUnderline"><%=manager_smail%></td>
							<td class="rt rowUnderline"><button type="button" class="lineBtnV16" onClick="jsEmailSend('M');">�������� �ޱ�</button></td>
						</tr>
						<tr>
							<td rowspan="2" class="rowUnderline">��������</td>
							<td rowspan="2" class="rowUnderline"><%=jungsan_name%>**</td>
							<td><%=jungsan_shp%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsSMSSend('J');">�������� �ޱ�</button></td>
						</tr>
						<tr>
							<td class="rowUnderline"><%=jungsan_smail%></td>
							<td class="rt rowUnderline"><button type="button" class="lineBtnV16" onClick="jsEmailSend('J');">�������� �ޱ�</button></td>
						</tr>
						<tr>
							<td rowspan="2">��۴����</td>
							<td rowspan="2"><%=deliver_name%>**</td>
							<td><%=deliver_shp%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsSMSSend('D');">�������� �ޱ�</button></td>
						</tr>
						<tr>
							<td><%=deliver_smail%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsEmailSend('D');">�������� �ޱ�</button></td>
						</tr>
					</table>
					</form>
				</div>
				<div id="dvAuth" class="sectionWrap" style="display:none;">
					<form name="frmAuth" method="post" target="hidFrm" action="/login/reconfirmipProc.asp">
						<input type="hidden" name="hidM" value="A">
						<input type="hidden" name="uid" value="<%=userid%>">
						<input type="hidden" name="sKey" value="<%=md5(userid&"TPUAUTH")%>">
					<h2>������ȣ �Է�</h2>
					<p class="tPad10">�޴���ȭ�� ������ ������ȣ�� �Է����ּ���. <strong>[ ������ȣ ��ȿ�ð� <span id="spTime" ><input type=text class="cRd3" name="sLimitTime" value="-:--" readolny style="width:30px; border:0px dotted #E0E0E0; text-align:center;background-color:#F8F8F8;font-weight:bold;"></span> ]</strong></p>
					<div class="inputBox"> 
						<p class="inputArea"><label for="code">������ȣ</label><input type="text" id="sAuthNo" name="sAuthNo" class="formTxt ftLt"   style="width:180px;" maxlength="6" onKeyPress="if (event.keyCode == 13) jsChkAuthno();"/></p>
						<button type="button" class="viewBtnV16" style="width:120px;" onClick="jsChkAuthno();">�Է�</button>
					</div>
					</form> 
				</div>		
				<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
			    <% end if %>
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div> 
		</div>
	</div>
</div>

</body>
</html>

<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim userid, searchID
dim manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp, manager_shp, jungsan_shp,deliver_shp
dim sql,reFAddr
dim   recentqcount
dim searchType
dim company_no, ceoname ,Bno1,Bno2, Bno3

 userid  = requestCheckVar(trim(request("uid")),32)
 reFAddr = request.ServerVariables("REMOTE_ADDR")
 searchType= requestCheckVar(request("rdoAType"),1)
 Bno1 = requestCheckVar(trim(request("Bno1")),3)
 Bno2 = requestCheckVar(trim(request("Bno2")),2)
 Bno3 = requestCheckVar(trim(request("Bno3")),5)
 company_no = Bno1&"-"&Bno2&"-"&Bno3
 ceoname = requestCheckVar(trim(request("Cnm")),32)
 	 	  
if searchType ="" then searchType = "2" '�⺻��:����ڵ�Ϲ�ȣ �˻�
	
if userid <> "" then
	'�ʱ�ȭ
		manager_name= ""
	 	manager_hp	=""
	 	jungsan_name=""
	 	jungsan_hp 	=""
	 	deliver_name=""
	 	deliver_hp 	=""
	
	'���̵���ȸ �α� ���
	sql = "exec db_partner.dbo.sp_Ten_partner_searchPWD_log '"&userid&"','"&Left(reFAddr,16)&"'"
  dbget.Execute sql
	 	 
	'10�� ���� 10ȸ �̻� �˻��� ���� 	 
	recentqcount = 0 	 
	sql = "select count(idx) as cnt "
	sql = sql & " from db_partner.dbo.tbl_partner_searchPWD_log  "
	sql = sql & " where refip='" + Left(reFAddr,16) + "' "
	sql = sql & " and datediff(n,regdate,getdate())<=10" 
	rsget.Open sql, dbget, 1
	if not rsget.eof then
		recentqcount = rsget("cnt")
	end if
	rsget.close

	if recentqcount>=10 then
		response.write "<script type='text/javascript'>alert('�ܽð� ���� �������� ������ �����Ͽ����ϴ�.\n��� �� �ٽ� �õ����ּ���.');</script>"
	  
	else

	sql =" select id, manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	if searchType ="1" then
		sql = sql & " and left(replace(ceoname,' ',''),3) =left('"&ceoname&"',3) "
	else
		sql = sql & " and company_no ='"&company_no&"' "
	end if
	 
	rsget.Open sql,dbget,1
  if  not rsget.EOF  then
  	  searchID			=rsget("id")
   		manager_name 	=rsget("manager_name")
   		if manager_name <> "" then manager_name= left(manager_name,1)
   		manager_hp =rsget("manager_hp") 
   		if manager_hp <> "" then manager_shp= left(manager_hp,4)&"****"&right(manager_hp,5)
   		jungsan_name 	=rsget("jungsan_name")
   		if jungsan_name <> "" then jungsan_name=left(jungsan_name,1)
   		jungsan_hp =rsget("jungsan_hp")
   		if jungsan_hp <> "" then jungsan_shp= left(jungsan_hp,4)&"****"&right(jungsan_hp,5)
   		deliver_name 	=rsget("deliver_name")
   		if deliver_name <> "" then deliver_name=left(deliver_name,1)
   		deliver_hp =rsget("deliver_hp") 
   		if deliver_hp <> "" then deliver_shp= left(deliver_hp,4)&"****"&right(deliver_hp,5)  
  end if
  rsget.close
	end if
end if
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/partner/lib/adminHead_NoJs.asp" -->
<script type="text/javascript">
		// SMS�Է� ī���� �۵�(2��30��:150��)
var iSecond=150;
var timerchecker = null;

	function jsSearchID(){
		if(document.frmID.rdoAType[0].checked){
			if(!document.frmID.Cnm.value){
			alert("��ǥ�ڸ��� �Է����ּ���");
			document.frmID.Cnm.focus();
			return;
			}
		}else{
			if(!document.frmID.BNo1.value){
			alert("����ڵ�Ϲ�ȣ�� �Է����ּ���");
			document.frmID.BNo1.focus();
			return;
			}
				if(!document.frmID.BNo2.value){
			alert("����ڵ�Ϲ�ȣ�� �Է����ּ���");
			document.frmID.BNo2.focus();
			return;
			}
				if(!document.frmID.BNo3.value){
			alert("����ڵ�Ϲ�ȣ�� �Է����ּ���");
			document.frmID.BNo3.focus();
			return;
			}
	  }
		 
		if(!document.frmID.uid.value){
			alert("���̵� �Է����ּ���");
			document.frmID.uid.focus();
			return;
		}
		document.frmID.submit();
	}
	
	function jsSMSSend(shp){  
	  document.frmSMS.sHp.value = shp;
	  document.frmSMS.submit();
	}
	


function startLimitCounter(cflg) {  
	
	if(cflg=="new") {
//		if(timerchecker != null) {
//			alert("�̹� ������ȣ�� �߼��Ͽ����ϴ�.\n�޴����� SMS�� Ȯ�����ּ���.");
//			return;
//		}
		iSecond=150;	 
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

function jsChangeField(iValue){ 
  location.href = "/login/searchPwd.asp?rdoAType="+iValue;
}
$(function(){
	var contH = $('.pwrBoxV16').outerHeight();
	var winH = $(window).height();
	if(winH < contH){
		$('.pwrBoxV16').css('top',0);
	} else {
		$('.pwrBoxV16').css('margin-top',-contH/2+'px');
	}
});
</script>
<style>
.pwrBoxV16 {width:336px; top:50%; margin-left:-168px; padding-top:15px;}
.pwrBoxV16 .titWrap {background:url(/images/partner/admin_login_box2_2016.png) 0 0 no-repeat;}
.pwrBoxV16 .titWrap h1 {padding:10px 17px 17px 17px;}
.pwrContWrap {padding:25px 30px 35px 30px; background:url(/images/partner/admin_login_box_2016.png) 0 100% no-repeat;}
.pwrBoxV16 .inputBox {margin-top:10px;}
.sectionWrap {margin-top:25px; padding-top:25px;}
.lineBtnV16 {width:90px; padding:0;}
.copy {padding-bottom:15px;}
</style>
</head>
<body>
<div id="login">
	<div class="container scrl"><!--class="container scrl"-->
		<div class="pwrBoxV16">
			<div class="titWrap">
				<h1>��й�ȣ ã��</h1>
			</div>
			<div class="pwrContWrap">
				<p class="cBk1">��ϵ� ������� �޴��� �������� ��й�ȣ�� ã�� �� �ֽ��ϴ�.</p>
					<form name="frmID" method="post">
				<div class="tMar10">
					<span><label><input type="radio" name="rdoAType" value="1" <%if searchType="1" then%>checked<%end if%> onClick="jsChangeField(1);"> ��������: ����</label></span>
					<span class="lPad20"><label><input type="radio" name="rdoAType" value="2" <%if searchType="2" then%>checked<%end if%> onClick="jsChangeField(2);"> ����ڵ�Ϲ�ȣ ����: �����</label></span>
				</div> 
				<div class="inputBox" id="dvB" style="display:<%if searchType<>"2" then%>none<%end if%>;">
					<p class="inputArea">
						<label>����ڵ�Ϲ�ȣ</label>
						<span class="ftRt">
							<input type="text" id="BNo1" name="BNo1" value="<%=Bno1%>" class="formTxt ftNone" maxlength="3"  style="width:27%;" onKeyPress="if (event.keyCode == 13) document.frmID.BNo2.focus();" />
							-
							<input type="text" id="BNo2" name="BNo2" value="<%=Bno2%>" class="formTxt ftNone" maxlength="2"  style="width:25%;" onKeyPress="if (event.keyCode == 13) document.frmID.BNo3.focus();" />
							-
							<input type="text" id="BNo3" name="BNo3" value="<%=Bno3%>" class="formTxt ftNone" maxlength="5" style="width:39%;" onKeyPress="if (event.keyCode == 13) document.frmID.uid.focus();" />
						</span>
					</p>
				</div>
				<div class="inputBox" id="dvC"  style="display:<%if searchType<>"1" then%>none<%end if%>;">
					<p class="inputArea">
						<label>��ǥ�ڸ�</label>
						<input type="text" id="Cnm" name="Cnm" value="<%=ceoname%>" class="formTxt" style="width:100%;" onKeyPress="if (event.keyCode == 13) document.frmID.uid.focus();" />
					</p>
				</div>
				<div class="inputBox">
					<p class="inputArea">
						<label for="id">ID</label>
						<input type="text" id="uid" name="uid" value="<%=userid%>" class="formTxt" style="width:100%;" onKeyPress="if (event.keyCode == 13) jsSearchID();" />
					</p>
				</div>
				<button type="button" class="viewBtnV16 tMar20" style="width:100%;" onClick="jsSearchID();">��ȸ</button>
				<!-- for dev msg : ���̵����� �߸��Է��� ��� ���� //-->
				<%if userid <> "" and searchID ="" then%><p class="tPad10 cRd3">��ϵ��� ���� ���̵��̰ų� �߸��� �����Դϴ�. �ٽ� �ѹ� Ȯ�����ּ���.</p><%end if%>
				</form>
				<div class="sectionWrap" id="idinfo" style="display:<%if searchID ="" then %>none<%end if%>;">
					<h2>��ȸ ���</h2>
					<p class="tPad10">�Ʒ� ����� ������ Ȯ���Ͻð� '������ȣ �ޱ�' ��ư�� Ŭ���� �ּ���.</p>
					<form name="frmSMS" method="post" target="hidFrm"  action="/login/searchPwd_sendSMS.asp">
						<input type="hidden" name="uid" value="<%=userid%>">
						<input type="hidden" name="sHp" value="">
						<input type="hidden" name="sKey" value="<%=md5(userid&"TPUSMS")%>">
					<table class="resultList">
						<colgroup>
							<col width="*" /><col width="*" /><col width="*" /><col width="100px" />
						</colgroup>
						<tr>
							<td>���������</td>
							<td class="ct" style="padding-left:5px; padding-right:5px;"><%=manager_name%>**</td>
							<td class="ct"><%=manager_shp%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsSMSSend('M');">������ȣ �ޱ�</button></td>
						</tr>
						<tr>
							<td>��������</td>
							<td class="ct" style="padding-left:5px; padding-right:5px;"><%=jungsan_name%>**</td>
							<td class="ct"><%=jungsan_shp%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsSMSSend('J');">������ȣ �ޱ�</button></td>
						</tr>
						<tr>
							<td>��۴����</td>
							<td class="ct" style="padding-left:5px; padding-right:5px;"><%=deliver_name%>**</td>
							<td class="ct"><%=deliver_shp%></td>
							<td class="rt"><button type="button" class="lineBtnV16" onClick="jsSMSSend('D');">������ȣ �ޱ�</button></td>
						</tr>
					</table>
					</form>
				</div>
				<div id="dvAuth" class="sectionWrap" style="display:none;">
					<form name="frmAuth" method="post" target="hidFrm" action="/login/searchPwdProc.asp">
						<input type="hidden" name="hidM" value="A">
						<input type="hidden" name="uid" value="<%=userid%>">
						<input type="hidden" name="sKey" value="<%=md5(userid&"TPUAUTH")%>">
						<h2>������ȣ �Է�</h2>
						<p class="tPad10">�޴���ȭ�� ������ ������ȣ�� �Է����ּ���. <strong>[������ȣ ��ȿ�ð� <span id="spTime" ><input type=text class="cRd3" name="sLimitTime" value="-:--" readolny style="width:33px; border:0px dotted #E0E0E0; text-align:center;background-color:#F8F8F8;font-weight:bold;"></span>]</strong></p>
						<div class="inputBox"> 
							<p class="inputArea"><label for="code">������ȣ</label><input type="text" id="sAuthNo" name="sAuthNo" class="formTxt ftLt" style="width:100px;" maxlength="6" onKeyPress="if (event.keyCode == 13) jsChkAuthno();"/></p>
							<button type="button" class="viewBtnV16" style="width:85px;" onClick="jsChkAuthno();">��ȸ</button>
						</div>
					</form>
				</div>
				<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>		
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div> 
		</div>
	</div>
</div>

</body>
</html>

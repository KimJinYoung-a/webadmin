<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_event_winner.asp
' Description :  �̺�Ʈ ��÷���
' History : 2007.02.22 ������ ����
'           2009.08.06 ������ SMS/�̸��� �߼� �߰�
'			2020.04.09 �ѿ�� ����(����ǰ���� üũ �߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode, cEvtCont
Dim egKindCode, ename, ebrand
	eCode		= requestCheckVar(Request("eC"),10)
 	egKindCode 	= requestCheckVar(Request("egKC"),10) 	

if eCode<>"" and eCode<>"0" then
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	
	cEvtCont.fnGetEventCont	 '�̺�Ʈ ���� ��������
	ename 		= db2html(cEvtCont.FEName)
	ebrand 		= cEvtCont.FEBrand
	set cEvtCont = nothing
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function jsChType(iVal){	
		$("#span1").hide();

		if(iVal == "2"){
			$("#div1").hide();
			$("#div2").show();
			$("#div3").hide();
		}else if	(iVal == "3"){
			$("#div1").show();
			$("#div2").hide();
			$("#div3").hide();
			$("#div3_1").show();
		}else if	(iVal == "5"){
			$("#div1").show();
			$("#div2").hide();
			$("#div3").show();
			$("#div3_1").hide();
		}else if	(iVal == "1"){
			$("#div1").hide();
			$("#div2").hide();
			$("#div3").hide();
			$("#span1").show();
		}else{
			$("#div1").hide();
			$("#div2").hide();
			$("#div3").hide();
		}	
	}
	
	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	
	function jsWinnerSubmit(){
		var frm=document.frmWin;
		if(!frm.sR.value){
			alert("����� �Է����ּ���");
			frm.sR.focus();
			return false;
		}
		
		if(!IsDigit(frm.sR.value)){
			alert("����� ���ڸ� �Է°����մϴ�.");
			frm.sR.focus();
			return false;
		}

		if(frm.uploadtype[0].checked){
			if(!frm.sW.value){
				alert("��÷�ڸ� �Է����ּ���");
				frm.sW.focus();
				return false;
			}
		}
		
		if(frm.evtprizetype.value == "3"){
			if(!frm.sGKN.value){
				alert("����ǰ����  �Է��� �ּ���");
				frm.sGKN.focus();
				return false;
			}

			if(!frm.iGK.value){
				alert("����ǰ���� Ȯ�� ��ư�� ������ Ȯ���� �ּ���");
				return false;
			}

			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('��� ��û���� �����ϼ���.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    // Ư�� ��ǰ��� �̺�Ʈ�� ��÷Ȯ�� �Ⱓ ������ ����û�� ���� ����(2015.04.28; ������)
			    if(!confirm('��� ��û���� ��÷Ȯ�αⰣ �������� �����Ǿ��ֽ��ϴ�.\n\nȮ���ϼ̽��ϱ�?')){
				    frm.reqdeliverdate.focus();
				    return false;
			    }
			    //alert('��� ��û���� ��÷Ȯ�αⰣ ���ķ� �Է��ϼž� �մϴ�.');
			    //frm.reqdeliverdate.focus();
			    //return false;
			}

			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('��� ������ �����ϼ���.');
        		return false;
        	}
			if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&(frm.jungsanValue.value=="")){
			    alert('�����(���԰�)�� �Է��ϼ���');
			    frm.jungsanValue.focus();
			    return false;
			}
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('��ü ���̵� �����ϼ���.');
        		return false;
            }
            
            
		}
		
		if(frm.evtprizetype.value == "2"){
			if(!frm.couponvalue.value){
				alert("�����ݾ� �Ǵ� �������� �Է����ּ���!");
				frm.couponvalue.focus();
				return false;
			}
			
			if(!frm.minbuyprice.value){
				alert("�ּұݾ��� �Է����ּ���!");
				frm.minbuyprice.focus();
				return false;
			}
			
			 if(!frm.sDate.value || !frm.eDate.value ){
			  	alert("�Ⱓ�� �Է����ּ���");
			  	frm.sDate.focus();
			  	return false;
			  }
		
			  if(frm.sDate.value > frm.eDate.value){
			  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.sDate.focus();
			  	return false;
			  }	  		
		}
		
		if(frm.evtprizetype.value == "5"){
			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('��� ��û���� �����ϼ���.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    alert('��� ��û���� ��÷Ȯ�αⰣ ���ķ� �Է��ϼž� �մϴ�.');
			    frm.reqdeliverdate.focus();
			    return false;
			}
			
			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('��� ������ �����ϼ���.');
        		return false;
        	}
            
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('��ü ���̵� �����ϼ���.');
        		return false;
            }
            
			 if(frm.itemuse_itemid.value == ""){
			  	alert("�׽��ͻ�ǰ�� �������ּ���");
			  	return false;
			  }
			  
			 if(frm.itemuse.value == ""){
			  	alert("���Բ� �������� �׽��ͻ�ǰ���� �Է����ּ���");
			  	frm.itemuse.focus();
			  	return false;
			  }
            
            if(GetByteLength(frm.itemuse.value) > 100)
            {
			  	alert("�׽��ͻ�ǰ���� 100 Byte �̳��� �Է����ּ���");
			  	frm.itemuse.focus();
			  	return false;
            }
            
			 if(!frm.itemuse_sDate.value || !frm.itemuse_eDate.value ){
			  	alert("�׽��ͻ�ǰ���Ⱓ�� �Է����ּ���");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
		
			  if(frm.itemuse_sDate.value > frm.itemuse_eDate.value){
			  	alert("�׽��ͻ�ǰ���Ⱓ �������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
			  
			 if(!frm.usewrite_sDate.value || !frm.usewrite_eDate.value ){
			  	alert("�׽����ı��ϱⰣ�� �Է����ּ���");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }
		
			  if(frm.usewrite_sDate.value > frm.usewrite_eDate.value){
			  	alert("�׽����ı��ϱⰣ �������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }	 
		}
		

		if(confirm("����Ͻ� ������ ���� �Ǵ� ������ �Ұ����ϸ� ������ �ٷ� ����˴ϴ�.\n\n��� �Ͻðڽ��ϱ�? ")){
			if(frm.uploadtype[1].checked){
				frm.target = "excelframe";
				$("#normalSubmit").hide();
				$("#excelprocing").show();
			} else {
				frm.target = "";
			}
			frm.action="eventprize_process.asp";
			frm.submit();
			return true;
		}else{
		    return false;
		}
	}

	function jsAutoWinnerSubmit(){
		var frm=document.frmWin;
		if(!frm.sR.value){
			alert("����� �Է����ּ���");
			frm.sR.focus();
			return false;
		}
		
		if(!IsDigit(frm.sR.value)){
			alert("����� ���ڸ� �Է°����մϴ�.");
			frm.sR.focus();
			return false;
		}

		if(frm.evtprizetype.value == "3"){
			if(!frm.sGKN.value){
				alert("����ǰ����  �Է��� �ּ���");
				frm.sGKN.focus();
				return false;
			}

			if(!frm.iGK.value){
				alert("����ǰ���� Ȯ�� ��ư�� ������ Ȯ���� �ּ���");
				return false;
			}

			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('��� ��û���� �����ϼ���.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    // Ư�� ��ǰ��� �̺�Ʈ�� ��÷Ȯ�� �Ⱓ ������ ����û�� ���� ����(2015.04.28; ������)
			    if(!confirm('��� ��û���� ��÷Ȯ�αⰣ �������� �����Ǿ��ֽ��ϴ�.\n\nȮ���ϼ̽��ϱ�?')){
				    frm.reqdeliverdate.focus();
				    return false;
			    }
			    //alert('��� ��û���� ��÷Ȯ�αⰣ ���ķ� �Է��ϼž� �մϴ�.');
			    //frm.reqdeliverdate.focus();
			    //return false;
			}

			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('��� ������ �����ϼ���.');
        		return false;
        	}
			if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&(frm.jungsanValue.value=="")){
			    alert('�����(���԰�)�� �Է��ϼ���');
			    frm.jungsanValue.focus();
			    return false;
			}
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('��ü ���̵� �����ϼ���.');
        		return false;
            }
            
            
		}
		
		if(frm.evtprizetype.value == "2"){
			if(!frm.couponvalue.value){
				alert("�����ݾ� �Ǵ� �������� �Է����ּ���!");
				frm.couponvalue.focus();
				return false;
			}
			
			if(!frm.minbuyprice.value){
				alert("�ּұݾ��� �Է����ּ���!");
				frm.minbuyprice.focus();
				return false;
			}
			
			 if(!frm.sDate.value || !frm.eDate.value ){
			  	alert("�Ⱓ�� �Է����ּ���");
			  	frm.sDate.focus();
			  	return false;
			  }
		
			  if(frm.sDate.value > frm.eDate.value){
			  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.sDate.focus();
			  	return false;
			  }	  		
		}
		
		if(frm.evtprizetype.value == "5"){
			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('��� ��û���� �����ϼ���.');
			    return false;
			}
			
			if (frm.reqdeliverdate.value<=frm.dAEDate.value){
			    alert('��� ��û���� ��÷Ȯ�αⰣ ���ķ� �Է��ϼž� �մϴ�.');
			    frm.reqdeliverdate.focus();
			    return false;
			}
			
			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('��� ������ �����ϼ���.');
        		return false;
        	}
            
            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('��ü ���̵� �����ϼ���.');
        		return false;
            }
            
			 if(frm.itemuse_itemid.value == ""){
			  	alert("�׽��ͻ�ǰ�� �������ּ���");
			  	return false;
			  }
			  
			 if(frm.itemuse.value == ""){
			  	alert("���Բ� �������� �׽��ͻ�ǰ���� �Է����ּ���");
			  	frm.itemuse.focus();
			  	return false;
			  }
            
            if(GetByteLength(frm.itemuse.value) > 100)
            {
			  	alert("�׽��ͻ�ǰ���� 100 Byte �̳��� �Է����ּ���");
			  	frm.itemuse.focus();
			  	return false;
            }
            
			 if(!frm.itemuse_sDate.value || !frm.itemuse_eDate.value ){
			  	alert("�׽��ͻ�ǰ���Ⱓ�� �Է����ּ���");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
		
			  if(frm.itemuse_sDate.value > frm.itemuse_eDate.value){
			  	alert("�׽��ͻ�ǰ���Ⱓ �������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.itemuse_sDate.focus();
			  	return false;
			  }
			  
			 if(!frm.usewrite_sDate.value || !frm.usewrite_eDate.value ){
			  	alert("�׽����ı��ϱⰣ�� �Է����ּ���");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }
		
			  if(frm.usewrite_sDate.value > frm.usewrite_eDate.value){
			  	alert("�׽����ı��ϱⰣ �������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.usewrite_sDate.focus();
			  	return false;
			  }	 
		}

		if(confirm("����Ͻ� ������ ���� �Ǵ� ������ �Ұ����ϸ� ������ �ٷ� ����˴ϴ�.\n\n��� �Ͻðڽ��ϱ�? ")){
			if(frm.uploadtype[1].checked){
				frm.target = "excelframe";
				$("#normalSubmit").hide();
				$("#excelprocing").show();
			} else {
				frm.target = "";
			}
			frm.action="eventprize_auto_process.asp";
			frm.submit();
			return true;
		}else{
		    return false;
		}
	}

	function disabledBox(comp){
        var frm = comp.form;
        if (comp.value=="Y"){
            frm.makerid.disabled = false;
            frm.jungsan.disabled = false;

			frm.jungsanValue.disabled = false;
	        //frm.jungsan.checked = true;
        }else{
            frm.makerid.selectedIndex = 0;
            frm.makerid.value = '';
            frm.makerid.disabled = true;
            frm.jungsan.disabled = true;

	        frm.jungsanValue.value = '';
	        frm.jungsanValue.disabled = true;
	        frm.jungsan.checked = false;
        }
    }

	//����ǰ ���� ���
	function jsSetGiftKind(){
		var gift_delivery, isupchebeasong;
		var sGKN;
		var makerid;

		for (var i=0;i < frmWin.isupchebeasong.length; i++){
			if (frmWin.isupchebeasong[i].checked){
				isupchebeasong=frmWin.isupchebeasong[i].value;
			}
		}

		sGKN=frmWin.evt_name.value
		makerid=frmWin.ebrand.value

		if (isupchebeasong==""){
			alert("��۱����� ������ �ּ���.");
			return;
		}
		gift_delivery=isupchebeasong

		var winkind;
		winkind = window.open('/admin/shopmaster/gift/popgiftKindReg.asp?gift_delivery='+gift_delivery+'&makerid='+makerid+'&sGKN='+sGKN,'popkind','width=1280px, height=960px, scrollbars=yes');
		winkind.focus();
	}

	//SMS���� Ȯ��
	function chkSMSTextLength(cont) {
		if(GetByteLength(cont)>80) {
			alert("SMS�� 80 Byte������ �߼� �����մϴ�.");
		}
		$("#smsCnt").html(GetByteLength(cont));
	}

	function swSMS() {
		if(frmWin.chkSMS.checked) {
			frmWin.smsCont.className="textarea";
			frmWin.smsCont.disabled=false;
		} else {
			frmWin.smsCont.className="textarea_ro";
			frmWin.smsCont.disabled=true;
		}
	}
	function swEmail() {
		if(frmWin.chkEmail.checked) {
			frmWin.emailCont.className="textarea";
			frmWin.emailCont.disabled=false;
		} else {
			frmWin.emailCont.className="textarea_ro";
			frmWin.emailCont.disabled=true;
		}
	}
	
	function GetByteLength(val){
	 	var real_byte = val.length;
	 	for (var ii=0; ii<val.length; ii++) {
	  		var temp = val.substr(ii,1).charCodeAt(0);
	  		if (temp > 127) { real_byte++; }
	 	}
	
	   return real_byte;
	}
	
	function ViewByteLength()
	{
		frmWin.bytecheck.value = GetByteLength(frmWin.itemuse.value);
	}

	function jungsanYN(){
		var frm = document.frmWin;
		if(frm.jungsan.checked==true){
			frm.jungsanValue.disabled = false;
		}else{
			frm.jungsanValue.value = '';
			frm.jungsanValue.disabled = true;
		}
	}
	function checkover1(obj) {
		var val = obj.value;
		if (val) {
			if (val.match(/^\d+$/gi) == null) {
				alert("���ڸ� ��������!");
				document.frmWin.jungsanValue.value = '';
				obj.select();
				return;
			}
		}
	}
	
	function jsUploadType(a){
		if(a == "direct"){
			$("#spandirect").show();
			$("#spanexcel").hide();
		} else {
			$("#spanexcel").show();
			$("#spandirect").hide();
		}
	}
	
	function jsGoExcelUp(){
		var winexcel;
		winexcel = window.open('/admin/eventmanage/event/pop_event_winner_excelupload.asp?eventid=<%=eCode%>','winexcel','width=400px, height=150px');
		winexcel.focus();
	}
	
	function jsPageReload(){
		opener.location.reload();
	}
</script>

<script type="text/javascript">
var speed = 350 //�����̴� �ӵ� - 1000�� 1��

function doBlink(){
var blink = $("blink");
for (var i=0; i < blink.length; i++)
blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : ""
}

function startBlink() {
setInterval("doBlink()",speed)
}
window.onload = startBlink;
</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷�� ���&nbsp;&nbsp;&nbsp;<font color="red"><b>�� <blink><u>�׽��� �̺�Ʈ ��÷ ��Ͻ�</u></blink> �ݵ�� <blink><u>������ �׽��� �̺�Ʈ��</u></blink> �ϼ���.</b></font></div>
<table border="0" width="100%" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmWin" method="post">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="egKC" value="<%=egKindCode%>">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="evt_name" value="<%= ename %>">
<input type="hidden" name="ebrand" value="<%= ebrand %>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="130" align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF">
					<%sbGetOptCommonCodeArr "evtprizetype", "", False,True,"onChange=jsChType(this.value);"%>
				</td>
			</tr>

			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">���</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sR"></td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����Ī</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="sRN" size="20" maxlength="32">
					<span id="span1" style="display:none;color:darkred;">(��÷Ȯ���������� �̺�Ʈ�� �߰��Ǿ� ǥ�õ�)</span>
				</td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷Ȯ�αⰣ</td>
				<td bgcolor="#FFFFFF"><input type="text" name="dASDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('dASDate');" style="cursor:hand;">
					~<input type="text" name="dAEDate" size="10"  maxlength="10" value="<%=dateadd("d",14,date())%>" onClick="jsPopCal('dAEDate');" style="cursor:hand;"></td>
			</tr>
			<% If eCode="4" Then %>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷�ο�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="prizecnt" value=""  size="2" maxlength="2" style="cursor:hand;"></td>
			</tr>
			<% End If %>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷��</td>
				<td bgcolor="#FFFFFF">
					<label id="labeldirect" style="cursor:pointer;"><input type="radio" id="labeldirect" name="uploadtype" value="direct" onClick="jsUploadType('direct');" checked>ID �����Է�(100�� ����)</label>
					<label id="labelexcel" style="cursor:pointer;"><input type="radio" id="labelexcel" name="uploadtype" value="excel" onClick="jsUploadType('excel');">Excel�� ���(100�� �̻�)</label>
					<strong><a href="" onClick="$('#excelexplain').show();return false;"><font color="red">[�ʵ�!!]</font></a></strong>
					<div style="display:none;padding:5 0 5 0;" id="excelexplain">
					<strong><a href="" onClick="$('#excelexplain').hide();return false;"><font size="3" color="blue">[ �� �� ]</font></a></strong><br>
					* �ٿ��� �ȵɶ��� �Ʒ� �ּҸ� �ܾ� ������ �� ���ͳ� â �ּҿ� �ٿ� �־� ����.<br>
					* ���� ���� â�� <strong>�ٿ���� ���</strong> ����� ����Ʈ�� ���� Ȯ���ϰ�, ����Ȱ� <strong>�ϳ��� ������ ������ �ٽ� �÷���</strong> ���� �ϸ�ǰ�, ����Ȱ� <strong>�ִ� ��� ���â�� ���� "Excel�� ���" �� �����ϰ� ���ε� ���� �ʰ� ������ ������ ��� �� ����</strong> �ϸ� �˴ϴ�.
					</div>
					<div style="padding:5 0 5 0;">
						<span id="spandirect" style="display:block;">
							�޸ӷ� ����, ������� (��: aaa,bbb,ccc)<br>
							<textarea name="sW" rows="5" style="width:100%"></textarea>
						</span>
						<span id="spanexcel" style="display:none;">
							<input type="button" value="Excel ���" onClick="jsGoExcelUp();"> ������ ���� �ٿ�޾� ��� [<a href="/admin/eventmanage/event/event_winner_userlist.xls" target="_blank"><u><strong>Download �� </strong></u></a>] <%=manageUrl%>/admin/eventmanage/event/event_winner_userlist.xls<br>
							
						</span>
					</div>
				</td>
			</tr>	
		</table>
	</td>
		
</tr>
<tr>
	<td>
		<div id="div1" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">							
			<tr>
				<td align="center" width="130"  bgcolor="<%= adminColor("tabletop") %>">����� ��ϱ���</td>
				<td bgcolor="#FFFFFF">
					<input type=radio name=rdgubun value="U">User�� ����� �Է�
					<input type=radio name=rdgubun value="F" checked>User �⺻ �ּ� ��� <font color="blue">[������ �⺻ �ּ��� ���]</font>
				</td>
			</tr>				
			<!-- ��� ���� �߰� : ������ -->
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����û��</td>
            	<td bgcolor="#FFFFFF">
            		<input type="text" name="reqdeliverdate" size="10" maxlength="10"  value="" >
		            <a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
            	</td>
            </tr>
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��۱���</td>
            	<td bgcolor="#FFFFFF">
            		<input type=radio name="isupchebeasong" value="N" checked onClick="disabledBox(this);">�ٹ����ٹ��
            		<input type=radio name="isupchebeasong" value="Y" onClick="disabledBox(this);">��ü�������
            	</td>
            </tr>
			<tr id="div3_1" style="display:block;">
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ��</td>
				<td bgcolor="#FFFFFF"><input type="hidden" name="iGK" >
					<input type="text" name="sGKN" size="10" onkeyup="document.frmWin.iGK.value='';"> 
					<input type="button" class="button" value="Ȯ��" onClick="jsSetGiftKind();">				
					<div id="spanImg"></div>	
				</td>
			</tr>				
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���꿩��</td>
            	<td bgcolor="#FFFFFF">
					<input type="checkbox" class="checkbox" name="jungsan" id="jungsan" onclick="javascript:jungsanYN();">������&nbsp;&nbsp;
					�����(���԰�) : <input type="text" class="text" id="jungsanValue" name="jungsanValue" onkeyup="checkover1(this)">
            	</td>
            </tr>
            <tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ü��۽�<br>��üID</td>
            	<td bgcolor="#FFFFFF">
            	    <% drawSelectBoxDesignerwithName "makerid","" %>
            	    <script language='javascript'>
            	    document.frmWin.makerid.disabled=true;
            	    document.frmWin.jungsan.disabled=true;
            	    </script>
            	</td>
            </tr>
		</table>	
		</div>	
		<div id="div2" style="display:none;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">							
			<tr>
				<td align="center" width="130" bgcolor="<%= adminColor("tabletop") %>">����Ÿ��</td>
				<td bgcolor="#FFFFFF">
					<input type=text name=couponvalue maxlength=7 size=10>
					<input type=radio name=coupontype value="1" onclick="alert('% ���� �����Դϴ�.');">%����
					<input type=radio name=coupontype value="2" checked >������
					(�ݾ� �Ǵ� % ����)
				</td>
			</tr>						
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�ּұ��űݾ�</td>
				<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>�� �̻� ���Ž� ��밡��(����)</td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ȿ�Ⱓ</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('sDate');" style="cursor:hand;">
					~<input type="text" name="eDate" size="10"  maxlength="10" onClick="jsPopCal('eDate');" style="cursor:hand;">
				</td>
			</tr>	
		</table>	
		</div>
	</td>
</tr>
<tr id="div3" style="display:none;">
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" width="130" bgcolor="<%= adminColor("tabletop") %>">�׽��ͻ�ǰ</td>
				<td bgcolor="#FFFFFF">
					�� �ɼ��� ������ ��ǰ�� �ڿ� �Է��ϸ� �˴ϴ�.<br>&nbsp;&nbsp;&nbsp;&nbsp;�̰� �Է¶��� ���� ���Բ� �������� �׽�Ʈ ��ǰ���Դϴ�.<br>
					<input type="button" value="��ǰ" onClick="window.open('/admin/eventmanage/event/pop_CateItemList.asp','popWinn','width=800, height=500, scrollbars=yes');">
					<input type="text" name="itemuse" value="" size="50" onkeyup="ViewByteLength()">
					<input type="text" name="bytecheck" value="" size="2">
					<input type="hidden" name="itemuse_itemid" value="">
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�׽��ͻ�ǰ���Ⱓ</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="itemuse_sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('itemuse_sDate');" style="cursor:hand;">
					~<input type="text" name="itemuse_eDate" size="10"  maxlength="10" onClick="jsPopCal('itemuse_eDate');" style="cursor:hand;">
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�׽����ı��ϱⰣ</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="usewrite_sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('usewrite_sDate');" style="cursor:hand;">
					~<input type="text" name="usewrite_eDate" size="10"  maxlength="10" onClick="jsPopCal('usewrite_eDate');" style="cursor:hand;">
				</td>
			</tr>
		</table>
		
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" width="130"  bgcolor="<%= adminColor("tabletop") %>">��÷�� SMS<br>������</td>
			<td bgcolor="#FFFFFF">
				<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td><textarea name="smsCont" rows="2" style="width:100%" class="textarea" onkeyup="chkSMSTextLength(this.value)">[�ٹ�����] �̺�Ʈ��÷�� �����մϴ�. �������� �� �����ٹ������� Ȯ�����ּ���.</textarea></td>
					<td width="110" valign="bottom"><input type=checkbox name=chkSMS value="Y" checked onClick="swSMS()">SMS���ù߼�</td>
				</tr>
				<tr>
					<td align="right">�� 80Byte���� �Է°���(���� <span id="smsCnt">76</span>Byte)</td>
					<td></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" width="130"  bgcolor="<%= adminColor("tabletop") %>">��÷�� �̸���<br>������</td>
			<td bgcolor="#FFFFFF">
				<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td><textarea name="emailCont" rows="5" style="width:100%" class="textarea_ro" disabled></textarea></td>
					<td width="110" valign="bottom"><input type=checkbox name=chkEmail value="Y" onClick="swEmail()">�̸��� ���ù߼�</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<span id="normalSubmit">
		<% If eCode="4" Then %>
		<input type="button" class="button" value="�ڵ���÷" onclick="jsAutoWinnerSubmit();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<% End If %>

		<a href="" onclick="jsWinnerSubmit();return false;"><img src="/images/icon_confirm.gif" border="0"></a>
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</span>
		&nbsp;<br>
		<span id="excelprocing" style="display:none;"><blink><strong>* ó�� �Ǵ� ���Դϴ�. â�� �������� ���� �Ϸ���� ó�����ּ���.</strong></blink></span>
		<span id="excelprocdetail"></span>
		<span id="excelSubmit" style="display:none;"><input type="submit" value="���� 100�� ����" style="height:30px;"></span>
	</td>
</tr>	
</form>	
</table>
<iframe id="excelframe" src="about:blank" name="excelframe" width="0" height="0"></iframe>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	'�α��� Ȯ��
	''if session("ssBctId")="" or isNull(session("ssBctId")) then
	if session("ssnTmpUIDPartner")="" or isNull(session("ssnTmpUIDPartner")) then
		Call Alert_Return("�߸��� �����Դϴ�.")
		response.End
	end if
%>
 
<!-- #include virtual="/partner/lib/adminHead.asp" -->
 
<script language='JavaScript'>
<!--
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

	function chkForm() {
		var frm = document.frmLogin;
		
		if(!frm.upwd.value) {
			alert("��й�ȣ�� �Է����ּ���.");
			frm.upwd.focus();
			return  ;
		}
		
	
		if (frm.upwd.value.length < 8 || frm.upwd.value.length > 16){
			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
			frm.upwd.focus();
			return ;
		 }
		 
		 	if(frm.upwd.value==frm.uid.value) {
			alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
			frm.upwd.focus();
			return  ;
		}
		
		 if (!fnChkComplexPassword(frm.upwd.value)) {
				alert('�н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
				frm.upwd.focus();
				return;
			}
		 
		 	if(!frm.upwd2.value) {
			alert("��й�ȣ Ȯ���� �Է����ּ���.");
			frm.upwd2.focus();
			return  ;
		}
		
			if(frm.upwd.value!=frm.upwd2.value) {
			alert("��й�ȣ Ȯ���� Ʋ���ϴ�.\n��Ȯ�� ��й�ȣ�� �Է����ּ���.");
			frm.upwd.focus();
			return  ;
			} 
			
		
			if(!frm.upwdS1.value) {
			alert("2�� ��й�ȣ�� �Է����ּ���.");
			frm.upwdS1.focus();
			return  ;
		}
		
	
		if (frm.upwdS1.value.length < 8 || frm.upwdS1.value.length > 16){
			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
			frm.upwdS1.focus();
			return ;
		 }
		 
		 	if(frm.upwdS1.value==frm.uid.value) {
			alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
			frm.upwdS1.focus();
			return  ;
		}
		
			if(frm.upwdS1.value==frm.upwd.value) {
			alert("1�� ��й�ȣ��  �ٸ� ��й�ȣ�� ������ּ���.");
			frm.upwdS1.focus();
			return  ;
		}

		if (!fnChkComplexPassword(frm.upwd.value)) {
			alert('1�� ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
			frm.upwd.focus();
			return;
		}
		if (!fnChkComplexPassword(frm.upwdS1.value)) {
			alert('2�� ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
			frm.upwdS1.focus();
			return;
		}

		 	if(!frm.upwdS2.value) {
			alert("��й�ȣ Ȯ���� �Է����ּ���.");
			frm.upwdS2.focus();
			return  ;
		}
		
			if(frm.upwdS1.value!=frm.upwdS2.value) {
			alert("��й�ȣ Ȯ���� Ʋ���ϴ�.\n��Ȯ�� ��й�ȣ�� �Է����ּ���.");
			frm.upwdS1.focus();
			return  ;
			}  

		 frm.submit(); 
	}
//-->
</script>
</head>
<body onLoad="document.frmLogin.upwd.focus()">  
<div class="wrap" id="login">
	<div class="container scrl">
		<div class="pwrBoxV16">
			<div class="titWrap">
				<h1>��й�ȣ ����</h1>
			</div>
			<div class="pwrContWrap">
				<p> 2008�� 12�� 15�Ϻ��� <span class="cRd3">��й�ȣ ��ȭ ��å</span>���� <br />���ȿ� ����� �н������ �����ϼž� �ٹ����� ������ ����Ͻ� �� �ֽ��ϴ�. 
			    ���� ��й�ȣ�� �ּ� 3������ �ѹ� �̻� ������ �ֽñ� �ٶ��ϴ�.<br><br>
			    ��ȭ�� ��й�ȣ ��å�� �Ʒ��� �����ϴ�.<br />
						<span class="cBd3"> &nbsp; 1. �ּ� 8�ڸ� �̻��� ��й�ȣ ���<br />
			    &nbsp; 2. ���̵�� �����ϰų� ���̵� ������ �н����� ����<br />
			    &nbsp; 3. ���� ���ڸ� �������� 3�� �̻� ����<br />
			    &nbsp; 4. ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)<br /><br /></span>
			</p>
				<form name="frmLogin" method="post" action="/login/doPasswordModi_partner.asp" target="FrameCKP"  >
    		<input type="hidden" name="backpath" value="<%= request("backpath") %>">				 
						<strong class="fs14 cBk1">ID:<%=session("ssnTmpUIDPartner")%><input type=hidden name=uid value='<%=session("ssnTmpUIDPartner")%>'></strong> 
						<div class="sectionWrap">
							<div class="partitionZone">
								<h2>1�� ��й�ȣ</h2>
								<div class="ftRt" style="width:265px;">
									<p class="inputArea"><label for="id">1�� ��й�ȣ</label><input type="password" id="upwd" name="upwd" class="formTxt" placeholder="1�� ��й�ȣ" style="width:100%;" maxlength="32" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwd2.focus();"/></p>
									<p class="inputArea tPad10"><label for="pwr">1�� ��й�ȣ Ȯ��</label><input type="password" id="upwd2" name="upwd2" class="formTxt" placeholder="1�� ��й�ȣ Ȯ��" style="width:100%;" maxlength="32" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS1.focus();"/></p> 
									<p class="tPad10 cRd3">���� ���� 8~16�� ����/���ڷ� �Է����ּ���. <!-- ��й�ȣ�� ��ġ���� �ʽ��ϴ�. �ٽ� Ȯ�����ּ���. --></p>
								</div>
						</div>
						<div class="partitionZone tMar20">
							<h2>2�� ��й�ȣ</h2>
							<div class="ftRt" style="width:265px;">
								<p class="inputArea"><label for="id2">2�� ��й�ȣ</label><input type="password" id="upwdS1" name="upwdS1" class="formTxt" placeholder="2�� ��й�ȣ" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS2.focus();"/></p>
								<p class="inputArea tPad10"><label for="pwr2">2�� ��й�ȣ Ȯ��</label><input type="password" id="upwdS2" name="upwdS2" class="formTxt" placeholder="2�� ��й�ȣ Ȯ��" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();"/></p> 
							</div>
						</div>
					</div>
				<button type="button" class="viewBtnV16 tMar20" style="width:100%;" onClick="chkForm();">��й�ȣ ����</button>
			</form>
			<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0"></iframe>
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div>
		</div>
	</div>
</div>

</body>
</html>


 
<%@ language="vbscript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim userid, searchID
dim manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp, manager_shp, jungsan_shp,deliver_shp
dim sql
userid  = session("AuthUID")
 
 if userid ="" or session("AuthChk") <>"Y" then
  response.write("<script>alert('���������� �ùٸ��� �ʽ��ϴ�.Ȯ�� �� �ٽ� �õ����ּ���') ;location.href='/login/searchPwd.asp';</script>") 
     response.End
 end if%>
<!-- #include virtual="/partner/lib/adminHead.asp" -->
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

	function chkForm() {
		var frm = document.frmPWD;
		
		if(!frm.upwd.value) {
			alert("��й�ȣ�� �Է����ּ���.");
			frm.upwd.focus();
			return  ;
		}
		
	
//		if (frm.upwd.value.length < 8 || frm.upwd.value.length > 16){
//			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
//			frm.upwd.focus();
//			return ;
//		 }
		 
		 	if(frm.upwd.value==frm.uid.value) {
			alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
			frm.upwd.focus();
			return  ;
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
		
	
//		if (frm.upwdS1.value.length < 8 || frm.upwdS1.value.length > 16){
//			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
//			frm.upwdS1.focus();
//			return ;
//		 }
		 
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

		 frm.submit();
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
.sectionWrap {margin-top:0; padding-top:15px; background:none;}
.copy {padding-bottom:15px;}
</style>
</head>
<body>
<div id="login">
	<div class="container scrl">
		<div class="pwrBoxV16">
			<div class="titWrap">
				<h1>��й�ȣ �缳��</h1>
			</div>
			<div class="pwrContWrap">
				<div style="background-color:#eee; padding:10px 7px;">
					<strong class="fs14 cBk1">ID : <%=userid%></strong>
					<p class="tPad05">������ �Ϸ�Ǿ����ϴ�. ��й�ȣ�� �缳�����ּ���.<br />(���� ���� 8~16���� ����/����)</p>
				</div>
				<form name="frmPWD" method="post" action="/login/searchPwdProc.asp">
					<input type="hidden" name="hidM" value="C">
					<input type="hidden" name="uid" value="<%=userid%>">
					<div class="sectionWrap">
						<div class="partitionZone">
							<div>
								<p class="inputArea"><label for="id">1�� ��й�ȣ</label><input type="password" id="upwd" name="upwd" class="formTxt" placeholder="1�� ��й�ȣ" style="width:100%;" maxlength="16" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwd2.focus();"/></p>
								<p class="inputArea tPad10"><label for="pwr">1�� ��й�ȣ Ȯ��</label><input type="password" id="upwd2" name="upwd2" class="formTxt" placeholder="1�� ��й�ȣ Ȯ��" style="width:100%;" maxlength="16" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS1.focus();"/></p> 
								<p class="tPad10 cRd3">1�� ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)</p>
							</div>
						</div>
						<div class="partitionZone tMar20 tPad15" style="background:url(/images/partner/admin_login_dot.png) 0 0 repeat-x;">
							<div>
								<p class="inputArea"><label for="id2">2�� ��й�ȣ</label><input type="password" id="upwdS1" name="upwdS1" class="formTxt" placeholder="2�� ��й�ȣ" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) frmLogin.upwdS2.focus();"/></p>
								<p class="inputArea tPad10"><label for="pwr2">2�� ��й�ȣ Ȯ��</label><input type="password" id="upwdS2" name="upwdS2" class="formTxt" placeholder="2�� ��й�ȣ Ȯ��" style="width:100%;" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13) chkForm();"/></p> 
								<p class="tPad10 cRd3">2�� ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)</p>
							</div>
						</div>
					</div>
					<button type="button" class="viewBtnV16 tMar20" style="width:100%;" onClick="chkForm();">��й�ȣ ����</button>
				</form>
			</div>
			<div class="copy">COPYRIGHT&copy; 10x10.co.kr ALL RIGHTS RESERVED.</div>
		</div>
	</div>
</div>
</body>
</html>
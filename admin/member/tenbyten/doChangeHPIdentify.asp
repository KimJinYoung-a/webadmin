<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ȯ���� ����� �޴�����ȣ ���� ó��
' History : 2011.05.30 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/common/member/oneself/nice.nuguya.oivs.asp"-->
<%
dim cMember
dim C_dumiKey 
C_dumiKey = session.sessionid 

dim mode, userid, username, juminno, dumiKey, strSql
dim MobileNo0, MobileNo1, MobileNo2, MobileNo3, MobileNo
dim manageUrl
dim NiceId, KeyString, ReturnURL, ConfirmMsg, strProcessType, strSendInfo, strOrderNo, SIKey

'// ���� �Ҵ�
mode = requestCheckVar(request.form("mode"),5)
userid = requestCheckVar(request.form("userid"),32)
dumiKey = request.form("dumiKey")
'----------
MobileNo0 = requestCheckVar(Request.form("hpNum0"),3)	' �޴�����ȣ0
MobileNo1 = requestCheckVar(Request.form("hpNum1"),3)	' �޴�����ȣ1
MobileNo2 = requestCheckVar(Request.form("hpNum2"),4)	' �޴�����ȣ2
MobileNo3 = requestCheckVar(Request.form("hpNum3"),4)	' �޴�����ȣ3

'// ���ǰ� Ȯ��
if (dumiKey<>C_dumiKey) then 
    Call Alert_close("���������� �ùٸ��� �ʽ��ϴ�.")
    response.end
end if

'// ���� �⺻���� ����
Set cMember = new CTenByTenMember
	cMember.Fuserid = userid
	cMember.fnGetScmMyInfo

	username      	= cMember.Fusername
	juminno			= Replace(Trim(cMember.FJuminno),"-","")

Set cMember = Nothing

if username="" or isNull(username) then
    Call Alert_close("���������� �������� �ʽ��ϴ�.")
    response.end
end if

IF application("Svr_Info")="Dev" THEN
 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 ELSE
 	manageUrl 	    = "http://webadmin.10x10.co.kr"
 END IF	

'// ��庰 �б�
Select Case mode
	Case "chgHP"
		MobileNo = MobileNo1 + "-" + MobileNo2 + "-" + MobileNo3

		strSql = "Update db_partner.dbo.tbl_user_tenbyten " &_
				" Set usercell='" & MobileNo & "'" &_
				"	, isIdentify='Y' " &_
				" Where userid='" & userid & "'"
		dbget.Execute(strSql)
	%>
		<script language="javascript">
		alert('����Ȯ�� �� �Է��Ͻ� �޴�����ȣ�� ����Ǿ����ϴ�.');
		parent.opener.history.go(0);
		parent.close();
		</script>
	<%

	Case "ActH"
		MobileNo = MobileNo0 + "-" + MobileNo1 + "-" + MobileNo2 + "-" + MobileNo3
		'=======================================================================================================
		'=====	�� ���ÿ� �߱� ���� ������ ���� : ���ÿ� �߱޵� ȸ���� ID �� KeyString���� �����Ͻʽÿ�. ��
		'=======================================================================================================
		NiceId= "Ntenxten1"	' �ѱ��ſ������� ���� ���� ���� ȸ���� ID ("Nxxx~")
		KeyString = "r6cA3YS9s8WTktrzgfNSOqQXsKf6GnNNpVEnn4DeDuwzgXhICcDpFhTefoTvFUbsux9EvPsbadplISwb" ' Ű��Ʈ��(80�ڸ�)�� �־��ּ���.
	
		'========================================================================================
		'=====	�� �����̿�� �ʿ��� ������ ���� ��
		'========================================================================================
		' �������� �޾Ƽ� ó���� URL�� �������ּ���.
		ReturnURL = manageUrl & "/admin/member/tenbyten/actChangeHPIdentify.asp" ' �������� ����� ���� ���� POPUP URL
		' �޴������� �� ������ȣ�� ���� �����ϰ� ���� �� ������ �� �ֽ��ϴ�.
		' Ư���� ��쿡 ���Ǵ� �����̴� �������� �ʰ�, ����Ͻø� �ڵ����� ���۵˴ϴ�.
		ConfirmMsg = ""	' ������ ������ȣ (6�ڸ� ���ڷ� �Է����ּ���.)
		'========================================================================================
	
		oivsObject.AthKeyStr = KeyString
	
		strProcessType = "5" '//�����ڵ� �������� ������.
		strSendInfo = makeSendInfo(NiceId, juminno, SIKey, ReturnURL, MobileNo, ConfirmMsg) '//������û�� �ʿ��� ��ȣȭ ������ ��������
	
		randomize(time())     
		strOrderNo = Replace(date, "-", "")  & round(rnd*(999999999999-100000000000)+100000000000) '//�ֹ���ȣ.. �� ��û���� �ߺ����� �ʵ��� ����
		
		'// ��ŷ������ ���� ��û������ ���ǿ� ����
		session("niceRsdNo") = juminno
		session("niceOrderNo") = strOrderNo
	%>
	<form name="resFrom" method="post" action="https://secure.nuguya.com/nuguya/NiceCert.do">
		<input type="text" name="SendInfo" value="<%=strSendInfo%>">
		<input type="hidden" name="ProcessType" value="<%=strProcessType%>">
		<input type="hidden" name="OrderNo" value="<%=strOrderNo%>">
		<input type="hidden" name="CertMethod" value="CM">	
		</form>		
		<script language="javascript">		
		<!--		
			var w="433";
			var h="540";
		    var x=window.screenLeft;
		    var y=window.screenTop;
		    var l=x+((document.body.offsetWidth-w)/2);
		    var t=y+((document.body.offsetHeight-h)/2);
	
			var frm = document.resFrom;
			var certWin = window.open("","niceCert","toolbars=0,resizable=0,scrolling=0,width="+w+",height="+h+",statusbar=1,top="+t+",left="+l);
			frm.target = "niceCert";
			frm.submit();
			certWin.focus();
			self.location.href ="about:blank";		
		//-->
		</script>
	<%
End Select
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
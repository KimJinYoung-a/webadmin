<%@ language=vbscript %>
<%
	Option Explicit
	Response.Expires = -1440
	
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  ����Ȯ�� ����� ó��
' History : 2011.05.31 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/common/member/oneself/nice.nuguya.oivs.asp"-->
<%
	dim KeyString, strRecvData, ssResidNo, ssOrderNo
	dim sConnectIP,sUserName,sUserPW
	Dim clsConnDB, clsUILog,clsSearchPW	
	dim iUserSeq, sEmail,dRegdate
		
	 sConnectIP = Left(request.ServerVariables("REMOTE_ADDR"),32)
	 sUserName 	=  requestCheckVar(Request.Form("sHUN"),20)	

	'=================================================================================================
	'=====	�� ���ÿ� �߱� ���� ������ ���� : ���ÿ� �߱޵� KeyString���� �����Ͻʽÿ�. ��
	'=================================================================================================
	KeyString = "r6cA3YS9s8WTktrzgfNSOqQXsKf6GnNNpVEnn4DeDuwzgXhICcDpFhTefoTvFUbsux9EvPsbadplISwb"  '//Ű��Ʈ��(80�ڸ�)�� �־��ּ���.
	
	oivsObject.AthKeyStr = KeyString
	
	strRecvData = Request.Form( "SendInfo" )
	oivsObject.resolveDatas(strRecvData)
	
	'// ��ŷ������ ���� ���ǿ� ����� ���� �� .. 
	ssResidNo = session("niceRsdNo")
	ssOrderNo = session("niceOrderNo")
	
	If  ssResidNo <> oivsObject.residNo or ssOrderNo <> oivsObject.ordNo then
		response.write("���������� �������� �ʽ��ϴ�.")
	End if
	
'	response.write("<BR>�ֹ���ȣ : " + oivsObject.ordNo)
'	response.write("<BR>�������� �������� : " + oivsObject.retCd + "(1:���� / 0:����)")
'	response.write("<BR>�����ڵ� : " + oivsObject.resCd)
'	response.write("<BR>���� �޽��� : " + oivsObject.message)
'	response.write("<BR>ȸ���� ID : " + oivsObject.niceId)
'	response.write("<BR>�ֹι�ȣ : " + oivsObject.residNo)
'	response.write("<BR>�޴�����ȣ : " + oivsObject.phoneNo)
	
	
	IF  oivsObject.retCd = "1" THEN
		'// ����Ȯ�� ����
	%>
		<script language="javascript">
		opener.parent.actChgHP();
		self.close();
		</script>
	<%
	ELSE
	    Call Alert_close("���������� �����Ͽ����ϴ�.\n�Է��� ������ Ȯ�����ּ���.")
	    response.end
	END IF	
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->
<!-- #include virtual="/lib/classes/cscenter/sp_tenCashCls.asp" -->

<%
	Dim vUserID, vUserName, oTenCash, vReturnCash, vCurrentDeposit, vDivCD, vReturnMethod, vBankName, vBankAccount, vBankOwnerName, vTitle, vRegUserID, vCSID, vOrderSerial
	Dim vQuery, vOpenerRe, vIsOK, vErr
	vUserID 		= Request("userid")
	vUserName		= Request("username")
	vReturnCash		= Request("returncash")
	vDivCD			= "A003"	'### CS ���� - ȯ������
	vReturnMethod	= "R007"	'### ������
	'vTitle			= GetCSRefundTitle("", vDivCD, "", vReturnMethod, "ȯ��(������)")
	vTitle			= "��ġ���� ���������� ȯ��"
	vRegUserID		= session("ssbctid")
	vBankName		= Request("rebankname")
	vBankAccount	= Request("rebankaccount")
	vBankOwnerName	= Request("rebankownername")
	vOrderSerial	= Request("orderserial")
	vOpenerRe		= "opener.document.location.reload();window.close();"

	If vUserID = "" Then
		Response.Write "<script>alert('���̵� �����ϴ�.');"&vOpenerRe&"</script>"
		dbget.close()
		Response.End
	End IF

	If vReturnCash = "" OR vReturnCash = "0" Then
		Response.Write "<script>alert('ȯ�Ҿ��� ���ų� 0�� �Դϴ�.');"&vOpenerRe&"</script>"
		dbget.close()
		Response.End
	End IF

	If IsNumeric(vReturnCash) = false Then
		Response.Write "<script>alert('�߸��� ȯ�ұݾ� �Դϴ�.');"&vOpenerRe&"</script>"
		dbget.close()
		Response.End
	End IF


	'####### �Էµ� �ֹ���ȣ�� �ִ� ��� �ֹ���ȣ�� ���̵� �´��� ��.
	Dim vOrderCheck
	If vOrderSerial <> "" Then
		vOrderCheck = "x"
		sqlStr = "SELECT count(orderserial) FROM [db_order].[dbo].[tbl_order_master] WHERE userid = '" & vUserID & "' AND orderserial = '" & vOrderSerial & "'"
		rsget.Open sqlStr,dbget,1
		If rsget(0) < 1 Then
			vOrderCheck = "x"
		Else
			vOrderCheck = "o"
		End IF
		rsget.close()
		If vOrderCheck = "x" Then
			sqlStr = "SELECT count(orderserial) FROM [db_log].[dbo].[tbl_old_order_master_2003] WHERE userid = '" & vUserID & "' AND orderserial = '" & vOrderSerial & "'"
			rsget.Open sqlStr,dbget,1
			If rsget(0) < 1 Then
				Response.Write "<script>alert('" & vOrderSerial & " �ֹ���ȣ�� " & vUserID & " ���� �ֹ���ȣ�� �ƴմϴ�.');"&vOpenerRe&"</script>"
				dbget.close()
				Response.End
			End IF
			rsget.close()
		End IF
	End IF


	'####### ������ ��ġ�� ��������.
	Set oTenCash = New CTenCash
	oTenCash.FRectUserID = vUserID
	oTenCash.getUserCurrentTenCash
	vCurrentDeposit = oTenCash.Fcurrentdeposit
	Set oTenCash = Nothing

	If vCurrentDeposit = "0" Then
		Response.Write "<script>alert('��ġ���� 0�� �Դϴ�.');"&vOpenerRe&"</script>"
		dbget.close()
		Response.End
	End IF

	If (CDbl(vCurrentDeposit) - CDbl(vReturnCash)) < 0 Then
		Response.Write "<script>alert('ȯ�Ҿ��� ��ġ�ݺ��� Ů�ϴ�.');"&vOpenerRe&"</script>"
		dbget.close()
		Response.End
	End IF


    On Error Resume Next
        dbget.beginTrans


		If (Err.Number = 0) Then
			vErr = "1"
		End IF

		'### cs��� �� idx ��������.
		'vCSID = RegCSMaster(vDivCD, orderserial, vRegUserID, vTitle, "��ġ�� ������ ȯ��", "C004", "CD99")
	    '' CS Master ����
	    'userid �� �ٸ� ���뵵 ���� ������ �Ϸ��� RegCSMaster �� �Լ��� Ǯ� �����Ͽ� ����.
	    Dim sqlStr
	    sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	    rsget.Open sqlStr,dbget,1,3
	    rsget.AddNew
	        rsget("divcd")          = vDivCD
	    	rsget("customername")   = vUserName
	    	rsget("userid")         = vUserID
	    	rsget("writeuser")      = vRegUserID
	    	rsget("title")          = vTitle
	    	rsget("contents_jupsu") = "��ġ�� ������ ȯ��"
	    	rsget("gubun01")        = "C004"
	    	rsget("gubun02")        = "CD99"
	    	rsget("currstate")      = "B001"
	    	rsget("deleteyn")       = "N"
	    	rsget("opentitle")      = "ȯ��"
	    	rsget("opencontents")   = ""
	    	rsget("orderserial")   	= CHKIIF(vOrderSerial<>"",vOrderSerial,"")
	    rsget.update
		    vCSID = rsget("id")
		rsget.close



		'### ������ ȯ�� insert.
		If (Err.Number = 0) Then
			vErr = "2"
		End IF
        Call RegCSMasterRefundInfo(vCSID, vReturnMethod, vReturnCash, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vBankName, vBankAccount, vBankOwnerName, "0")

		If (Err.Number = 0) Then
			vErr = "3"
		End IF
        Call AddCSMasterRefundInfo(vCSID, 0, 0, 0, 0)



        '### ���� ��ȣȭ �߰�.
		If (Err.Number = 0) Then
			vErr = "4"
		End IF
        Call EditCSMasterRefundEncInfo(vCSID, "AE2", vBankAccount)



		'### �α� ���� �� ���� ��ġ�ݾ� ����.
		If (Err.Number = 0) Then
			vErr = "5"
		End IF
		vQuery = "INSERT INTO [db_user].[dbo].[tbl_depositlog](userid,deposit,jukyocd,jukyo,orderserial,deleteyn,asid) " & _
				 "	VALUES('" & vUserID & "',-" & vReturnCash & ",'300','��ġ�� ������ ȯ��'," & CHKIIF(vOrderSerial<>"","'"&vOrderSerial&"'","null") & ",'N'," & vCSID & ") "
		dbget.Execute vQuery



		If (Err.Number = 0) Then
			vErr = "6"
		End IF
		vQuery = "UPDATE [db_user].[dbo].[tbl_user_current_deposit] SET " & VBCRLF
		vQuery = vQuery & "		currentdeposit = currentdeposit - " & vReturnCash & " " & VBCRLF
		vQuery = vQuery & "		,spenddeposit = spenddeposit + " & vReturnCash & " " & VBCRLF
		vQuery = vQuery & "	WHERE userid = '" & vUserID & "' "

		dbget.Execute vQuery


		If (Err.Number = 0) Then
			dbget.CommitTrans
			vIsOK = "o"
		Else
			dbget.RollBackTrans
			vIsOK = "x"
		End If
	On Error Goto 0
%>


<script language="javascript">
<% If vIsOK = "o" Then %>
	alert("ó���Ǿ����ϴ�.");
<% Else %>
	alert("����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. �����ڿ��� �����ڵ� <%=vErr%> �� �Բ� ���� ���.");
<% End IF %>
	<%=vOpenerRe%>
</script>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

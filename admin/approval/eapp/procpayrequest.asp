<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����  ���
' History : 2011.03.16 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<% 
Dim objCmd, returnValue, sMode
Dim reportIdx
Dim payrequestidx,payrequestdate,payrequestPrice,InBank,accountNo,accountHolder,payrequestState,adminID,authId1,authId2,authId,authposition,AuthState,Comment
Dim paydate, outBank, yyyymm, yyyy,mm,isTakeDoc,paycomment,payrealdate
Dim eapppartIdx, partMoney
dim arreapppartIdx,lasteapppartIdx,returnUrl
Dim fileName,referLink, filelink, filetype
Dim i ,blnusing, commentidx, isLast
Dim payRequestTitle, payrequesttype, divMoney,arap_cd,cust_cd
Dim iDockind, sVatKind, dIssuedate, sItemName, mTotPrice, mSupplyPrice, mVatPrice, setaxkey, sDocbigo, sfile2,paydocidx
Dim igbn ,iRectMenu
Dim ipayType, sCurrencyType, mCurrencyPrice
Dim sPayrequesttitle, taxkey, MatchTaxkey ''2013/09/16 �߰�

iRectMenu =	requestCheckvar(Request("iRM"),10)
sMode		= requestCheckvar(Request("hidM"),2)
reportIdx	= requestCheckvar(Request("irIdx"),10)
payrequestidx	= requestCheckvar(Request("iprIdx"),10)
payrequestdate	= requestCheckvar(Request("dprd"),10)
payrequestPrice	= requestCheckvar(Request("mprp"),20)
cust_cd			= requestCheckvar(Request("hidcustcd"),13)
InBank			= requestCheckvar(Request("selIB"),50)
accountNo		= requestCheckvar(Request("san"),50)
accountHolder	= requestCheckvar(Request("sah"),16)
payrequestState	= requestCheckvar(Request("hidPRS"),4)
AuthState		= requestCheckvar(Request("hidAS"),4)
igbn				=requestCheckvar(Request("igbn"),1)

paydate 	= requestCheckvar(Request("dPD"),10)
outBank 	= requestCheckvar(Request("selOB"),10)
yyyy  		= requestCheckvar(Request("selY"),4)
mm			= requestCheckvar(Request("selM"),4)
if yyyy <> "" or mm <> "" then
	yyyymm = yyyy&"-"&mm
end if
payrealdate =requestCheckvar(Request("dprld"),10)
isTakeDoc 	= requestCheckvar(Request("rdoTD"),1)
paycomment 	= ReplaceRequestSpecialChar(Request("tPCmt"))

adminID		= session("ssBctId")
authId		= requestCheckvar(Request("hidAI"),32)
authId1		= requestCheckvar(Request("hidAI1"),32)
authId2		= requestCheckvar(Request("hidAI2"),32)
authposition= requestCheckvar(Request("iAP"),10)
Comment		= ReplaceRequestSpecialChar(Request("tCmt"))

fileName 	= ReplaceRequestSpecialChar(Request("sFile"))

referLink	= ReplaceRequestSpecialChar(Request("sL"))
blnusing	= requestCheckvar(Request("blnU"),1)
commentidx = requestCheckvar(Request("iCidx"),10)
isLast		= requestCheckvar(Request("blnL"),1)
 IF isLast = "" THEN isLast = 1
eapppartIdx= ReplaceRequestSpecialChar(Request("ip"))
partMoney	= ReplaceRequestSpecialChar(Request("mP"))
IF  partMoney = "" THEN partMoney = 0

returnUrl   =  requestCheckvar(Request("hidRU"),100)

arap_cd		=  requestCheckvar(Request("iAIdx"),10)
payRequestTitle =requestCheckvar(Request("sprt"),200)

iDockind =requestCheckvar(Request("rdoDK"),1)
sVatKind =requestCheckvar(Request("rdoVK"),1)

dIssuedate =requestCheckvar(Request("dID"),10)
sItemName =requestCheckvar(Request("sINm"),50)
mTotPrice =requestCheckvar(Request("mTP"),20)
mSupplyPrice =requestCheckvar(Request("mSP"),20)
mVatPrice =requestCheckvar(Request("mVP"),20)
setaxkey =requestCheckvar(Request("sEK"),32)
sDocbigo =requestCheckvar(Request("sDB"),50)
sfile2 =requestCheckvar(Request("sfile2"),120)
paydocidx =requestCheckvar(Request("hidPDidx"),10)
ipayType	= requestCheckvar(Request("selPT"),4)
sCurrencyType	= requestCheckvar(Request("selCT"),3)
mCurrencyPrice	= requestCheckvar(Request("sCP"),20)
sPayrequesttitle= requestCheckvar(Request("sPayrequesttitle"),100)

'���ڰ���  type 1/�繫ȸ�� �űԵ�� type 2
payrequesttype =requestCheckvar(Request("iptt"),10)
divMoney = 2000000

Dim sqlStr,AssignedRow

SELECT CASE sMode
Case "I"
	IF Cdbl(payrequestPrice) >= Cdbl(divMoney) or payrequestState = 0 THEN
		authstate = 0
	ELSE
		authstate = 1
	END IF

    ''//2013/11/04 �߰�
	if (authstate=0) then
	    '' authId1 �� ���� ���� ������ �ִ°�� 1
	    sqlStr = " select count(*) as CNT from db_partner.dbo.tbl_eAppAuthLine"
        sqlStr = sqlStr&" where authID='"&authId1&"'"
        sqlStr = sqlStr&" and reportIdx="&reportidx
        sqlStr = sqlStr&" and isUsing=1"
        sqlStr = sqlStr&" and authState=1"
        sqlStr = sqlStr&" and authPosition=999" ''����

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            if rsget("CNT")>0 then
                authstate=1
            end if
        end if
        rsget.Close
	end if

  dbget.beginTrans
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_Insert]( "&payrequesttype&","&reportidx&" ,'"&payRequestTitle&"',"&arap_cd&",'"&payrequestdate&"', '"&payrequestPrice&"'"&_
						+",'"&cust_cd&"','"&InBank&"','"&accountNo&"','"&accountHolder&"',"&payrequestState&",'"&adminId&"','"&authId1&"','"&authId2&"',"&authstate&", '"&divMoney&"',"&ipayType&",'"&sCurrencyType&"','"&mCurrencyPrice&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = 0 THEN
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
		dbget.RollBackTrans
	response.end
	END IF

	'����÷��
		payrequestidx = returnValue
		fileName = split(fileName,",")
		For i = 0 To UBound(fileName)
		if(trim(fileName(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&","&payrequestidx&" ,'"&trim(fileName(i))&"', 1)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		  returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
					Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
					dbget.RollBackTrans
				response.end
				END IF
		end if
		Next

		referLink = split(referLink,",")
		For i = 0 To UBound(referLink)
		if(trim(referLink(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&","&payrequestidx&" ,'"&trim(referLink(i))&"', 0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
			IF returnValue = 0 THEN
					Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
					dbget.RollBackTrans
				response.end
				END IF
		end if
		Next

		'�������� ���
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Insert]("&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF

		'�μ��� �ڱݱ��� ���
		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,"&payrequestidx&",'"&eapppartIdx(i)&"','"&partMoney(i)&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
					Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
					dbget.RollBackTrans
				response.end
				END IF
		Next
		END IF

	dbget.CommitTrans
		IF payrequesttype = "2" THEN
	%>
		<script language="javascript">
			<!--
				alert("��ϵǾ����ϴ�");
				opener.location.reload();
				self.close();
			//-->
			</script>
<%
		ELSE
%>
		<script language="javascript">
			<!--
				alert("��ϵǾ����ϴ�");
				opener.top.location.href = "<%=returnUrl%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestidx%>&iRM=<%=iRectMenu%>";
				self.close();
			//-->
			</script>
<%	END IF

	response.end
Case "U"
dbget.beginTrans
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_Update]( "&reportidx&","&payrequestIdx&" ,'"&payrequesttitle&"',"&arap_cd&",'"&payrequestdate&"', '"&payrequestPrice&"'"&_
						+",'"&cust_cd&"','"&InBank&"','"&accountNo&"','"&accountHolder&"',"&payrequestState&",'"&adminId&"',"&authposition&_
						+",'"&divMoney&"',"&ipayType&",'"&sCurrencyType&"','"&mCurrencyPrice&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue = 0 THEN
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
		dbget.RollBackTrans
		response.end
	END IF

		 Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_Delete]("&reportIdx&","&payrequestidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
				dbget.RollBackTrans
			response.end
		END IF

		fileName = split(fileName,",")
		For i = 0 To UBound(fileName)
		if(trim(fileName(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&","&payrequestidx&" ,'"&trim(fileName(i))&"', 1)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
			IF returnValue = 0 THEN
				Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
				dbget.RollBackTrans
				response.end
			END IF
		end if
		Next

		referLink = split(referLink,",")
		For i = 0 To UBound(referLink)
		if(trim(referLink(i)) <> "") then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppAttachFile_insert]( "&reportIdx&","&payrequestidx&" ,'"&trim(referLink(i))&"', 0)}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
		end if
		Next

	IF paydocidx <> "" THEN
	'�������� ����
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Update]("&paydocidx&","&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
	ELSE
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Insert]("&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
	END IF

		'�μ��� �ڱݱ��� ���
	 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_Delete]( "&reportIdx&" ,"&payrequestidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF

			IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,"&payrequestidx&",'"&eapppartIdx(i)&"','"&partMoney(i)&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF

		Next
		END IF
	dbget.CommitTrans
		IF payrequesttype = "2" THEN
	%>
		<script language="javascript">
			<!--
				alert("ó���Ǿ����ϴ�");
				opener.location.reload();
				self.close();
			//-->
			</script>
<%
		ELSE
%>
		<script language="javascript">
			<!--
				alert("ó���Ǿ����ϴ�");
				top.location.href  = "<%=returnUrl%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestidx%>&iRM=<%=iRectMenu%>";
				self.close();
			//-->
			</script>
<%	END IF

	response.end
Case "C"
dbget.beginTrans
	IF outBank = "" THEN outBank = 0
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_Confirm]("&reportIdx&","&payrequestIdx&" ,'"&paydate&"',"&outBank&",'"&yyyymm&"','"&payrealdate&"',"&isTakeDoc&_
						+",'"&paycomment&"',"&payrequestState&","&authposition&",'"&authId&"', "&authstate&","&arap_cd&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF returnValue =2 THEN
		  Call Alert_return ("����ó�� ������ �����ϴ�.")
			dbget.RollBackTrans
		response.end
	ELSEIF 	returnValue =0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
	END IF

		'�μ��� �ڱݱ��� ���
	 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_Delete]( "&reportIdx&" ,"&payrequestIdx&" )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF

		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,"&payrequestIdx&" ,'"&eapppartIdx(i)&"','"&partMoney(i)&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
	 	IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
		Next
		END IF

		dbget.CommitTrans

 		'''2012-12-14 ������ �߰�. //�������ν� ����.
 		if (payrequestState="7") then
        	IF (setaxkey<>"")  then ''((iDockind="1") or (iDockind="2")) and
        		sqlStr = "exec db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne_etcBuy] "&payrequestidx&",'"&setaxkey&"'"
        		dbget.Execute sqlStr
            end if
        end if

		IF igbn = "1" THEN
			%>
				<script language="javascript">
			<!--
				alert("ó���Ǿ����ϴ�");
				parent.location.href = "<%=returnUrl%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestidx%>&iRM=<%=iRectMenu%>";
			//-->
			</script>
			<%
		ELSE
		%>
		<script language="javascript">
			<!--
				alert("ó���Ǿ����ϴ�");
				top.location.href = "<%=returnUrl%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestidx%>&iRM=<%=iRectMenu%>";
			//-->
			</script>
<%	 END IF
	response.end
CASE "T"
Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_ModTakeDoc]("&payrequestIdx&" ,"&isTakeDoc&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	IF returnValue ="1" THEN
%>
	<script language="javascript">
	<!--
	alert("��ϵǾ����ϴ�.");
	opener.location.reload();
	self.close();
	//-->
	</script>
<%

	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
	END IF
response.end
CASE "S"	 '����ó�� �� �������
dbget.beginTrans
IF outBank = "" THEN outBank = 0
Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_ConfirmUpdate]("&reportIdx&","&payrequestIdx&" ,'"&paydate&"',"&outBank&",'"&yyyymm&"',"&isTakeDoc&_
						+",'"&paycomment&"',"&payrequestState&","&authposition&",'"&authId&"',"&arap_cd&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue =2 THEN
		  Call Alert_return ("����ó�� ������ �����ϴ�.")
		dbget.RollBackTrans
		response.end
	ELSEIF 	returnValue =0 THEN
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
		dbget.RollBackTrans
		response.end
	END IF

	IF paydocidx <> "" THEN
	'�������� ����
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Update]("&paydocidx&","&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
	ELSE
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Insert]("&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
	END IF

		'�μ��� �ڱݱ��� ���
	 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_Delete]( "&reportIdx&"  ,"&payrequestidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF


		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)

		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,"&payrequestidx&",'"&eapppartIdx(i)&"','"&partMoney(i)&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			dbget.RollBackTrans
			response.end
		END IF
		Next
		END IF

	 	dbget.CommitTrans
%>
	<script language="javascript">
	<!--
		alert("��ϵǾ����ϴ�");
		self.location.href = "confirmpayrequest.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestidx%>&iRM=<%=iRectMenu%>&ias=<%=authstate%>";
	//-->
	</script>
<%
response.end
CASE "D"
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_Delete]("&reportIdx&","&payrequestIdx&" ,'"&adminId&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing

 	IF	returnValue =1 THEN
 		%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�");
				top.location.href  = "<%=returnUrl%>?iRM=<%=iRectMenu%>";
				self.close();
			//-->
			</script>
<%
response.end

	ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
		response.end
	END IF
CASE "AU"	 '�����׸� ����
	dbget.beginTrans

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayRequest_ARAPUpdate]("&reportIdx&","&payrequestIdx&","&arap_cd&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
		End With
		returnValue = objCmd(0).Value
	Set objCmd = nothing

	IF returnValue =0 THEN
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
		dbget.RollBackTrans
		response.end
	END IF

	dbget.CommitTrans
%>
	<script type="text/javascript">
		alert("����Ǿ����ϴ�");
		location.href = '/admin/approval/eapp/modeappPayDoc.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
	</script>
<%
response.end
CASE "DU" '�������� ����
	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eAppPayDoc_Update]("&paydocidx&","&payrequestidx&",'"&iDockind&"','"&sVatKind&"','"&dIssuedate&"','"&sItemName&"','"&mTotPrice&"','"&mSupplyPrice&"','"&mVatPrice&"','"&setaxkey&"','"&sDocbigo&"','"&sfile2&"','"&adminid&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 1 THEN
		    '''2012-02 ������ �߰�.
        	IF ((iDockind="1") or (iDockind="2")) and (setaxkey<>"") then
        		sqlStr = "exec  db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne_etcBuy] "&payrequestidx&",'"&setaxkey&"'"
        		dbget.Execute sqlStr
            end if
	%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�");

			<% If Request("returnurl") <> "" Then %>
				location.href = '<%=Request("returnurl")%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			<% Else %>
				<%IF returnUrl = "top" THEN %>
					top.location.href  = "popindex.asp?iRM=M028";
				<%ELSE%>
				opener.location.reload();
				<%END IF%>
				self.close();
			 <% End If %>
			//-->
			</script>
<%
response.end
		ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
CASE "PM"  '�μ��� �ڱݱ��� ����.
   ' response.write "������ ���� ���"
   ' response.end
    '�μ��� �ڱݱ��� ���
	 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_Delete]( "&reportIdx&"  ,"&payrequestidx&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF

''rw eapppartIdx
		IF eapppartIdx <> "" THEN
		eapppartIdx = split(eapppartIdx,",")
		partMoney = split(partMoney,",")

		For i = 0 To UBound(eapppartIdx)
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPartMoney_insert]( "&reportIdx&" ,"&payrequestidx&",'"&eapppartIdx(i)&"','"&partMoney(i)&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing
		IF returnValue = 0 THEN
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
		Next
		END IF
''rw UBound(eapppartIdx)
		if (UBound(eapppartIdx)>=0) then
		    ''����Ⱥ��� ��� ERP �Ⱥ� ����. // ���۵� ��츸.
		    ''and (convert(varchar(7),D.issueDate,21)=@YYYYMM or convert(varchar(7),P.payREALDATE,21)=@YYYYMM )
		    ''������ ���� �ȵǾ����� �ǹ̾���.
		    sqlStr = "select P.payrequestIdx,convert(varchar(7),D.issueDate,21) as YYYYMM"
            sqlStr = sqlStr&" from db_partner.dbo.tbl_eappPayRequest p"
            sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_eAppPayDoc D"
            sqlStr = sqlStr&" 	on p.payrequestIdx=D.payrequestIdx"
            sqlStr = sqlStr&" where P.payrequestIdx="&payrequestidx
            sqlStr = sqlStr&" and D.erpDocLinkType is Not NULL"
            sqlStr = sqlStr&" and D.erpDocLinkKey is Not NULL"
            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
                YYYYMM = rsget("YYYYMM")
            else
                YYYYMM = ""
            end if
            rsget.Close

            if (YYYYMM<>"") then
                IF application("Svr_Info")="Dev" THEN
                    sqlStr = " exec db_SCM_LINK.[dbo].[sp_SCM2ERP_payreqDIV_MAKE_TEST] '"&YYYYMM&"',"&payrequestidx
                ELSE
        	        sqlStr = " exec db_SCM_LINK.[dbo].[sp_SCM2ERP_payreqDIV_MAKE] '"&YYYYMM&"',"&payrequestidx
                END IF

                rw sqlStr
                server.Execute("/lib/db/dbiTmsOpen.asp")
                dbiTms_dbget.Execute sqlStr, AssignedRow
                server.Execute("/lib/db/dbiTmsClose.asp")
                rw "�Ⱥ����� ���� ["&AssignedRow&"]��"
            end if
		end if
%>
		<script language="javascript">
		<!--
			alert("�����Ǿ����ϴ�");
		 location.href = '/admin/approval/eapp/modeappPayDoc.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
		//-->
		</script>
<%
    response.end
CASE "FS" '�濵���������׸�  ����

	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_eappPayRequest_FSUpdate]("&payrequestIdx&" ,'"&paydate&"',"&outBank&",'"&payrealdate&"','"&yyyymm&"',"&isTakeDoc&",'"&paycomment&"',"&CHKIIF(request("frcfin")="on",1,0)&")}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
		Set objCmd = nothing

		IF returnValue = 1 THEN
	%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�");
			 location.href = '/admin/approval/eapp/modeappPayDoc.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			//-->
			</script>
<%
response.end
		ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
CASE "DP" '��������׸�  ����
    sqlStr = "update db_partner.dbo.tbl_eapppayrequest " &VbCRLF
    sqlStr = sqlStr&" SET payType="&ipayType&VbCRLF
    sqlStr = sqlStr&" , currencyType='"&sCurrencyType&"'"&VbCRLF
    IF (mCurrencyPrice<>"") then
        sqlStr = sqlStr&" , currencyPrice="&mCurrencyPrice&VbCRLF
    end if
    sqlStr = sqlStr&" WHERE payrequestIdx="&payrequestIdx
'rw  sqlStr
    dbget.Execute sqlStr,returnValue
	IF returnValue = 1 THEN
	%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�...");
			<% If Request("returnurl") <> "" Then %>
			location.href = '<%=Request("returnurl")%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			<% Else %>
			 location.href = '/admin/approval/eapp/modeappPayDoc.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			 <% End If %>
			//-->
			</script>
<%
response.end
        ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
CASE "ET" '�ڱݿ뵵��Ī  ����
    sqlStr = "update db_partner.dbo.tbl_eapppayrequest " &VbCRLF
    sqlStr = sqlStr&" SET payrequestTitle='"&sPayrequesttitle&"'"&VbCRLF
    sqlStr = sqlStr&" WHERE payrequestIdx="&payrequestIdx
'rw  sqlStr
    dbget.Execute sqlStr,returnValue

    if (returnValue = 1) then
        sqlStr = "select top 1 T.taxkey,M.taxkey as MatchTaxkey"&VbCRLF
        sqlStr = sqlStr & " from db_partner.dbo.tbl_eAppPayDoc D"&VbCRLF
        sqlStr = sqlStr & " 	join db_partner.dbo.tbl_esero_tax T"&VbCRLF
        sqlStr = sqlStr & " 	on D.etaxkey=T.taxkey"&VbCRLF
        sqlStr = sqlStr & " 	left join db_partner.dbo.tbl_esero_taxMatch M"&VbCRLF
        sqlStr = sqlStr & " 	on T.taxkey=M.taxkey"&VbCRLF
        sqlStr = sqlStr & " where D.payrequestIdx="&payrequestIdx&VbCRLF

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            taxkey = rsget("taxkey")
            MatchTaxkey = rsget("MatchTaxkey")
        else
            taxkey = ""
            MatchTaxkey = ""
        end if
        rsget.Close

        if (taxkey<>"") and isNULL(MatchTaxkey) then
            sqlStr = "exec db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne_etcBuy] "&payrequestidx&",'"&taxKey&"'"
            dbget.Execute sqlStr
        end if

        if (taxkey<>"") then
            sqlStr = "update db_Partner.dbo.tbl_Esero_Tax"
            sqlStr = sqlStr & " set dtlnameorg=isNULL(dtlnameorg,dtlname)" & vbCRLF
            sqlStr = sqlStr & " ,dtlname='"&sPayrequesttitle&"'"& vbCRLF
            sqlStr = sqlStr & " where taxKey='"&taxKey&"'" & vbCRLF
            dbget.Execute sqlStr

            sqlStr = "update db_partner.dbo.tbl_eAppPayDoc"
            sqlStr = sqlStr & " set itemname='"&sPayrequesttitle&"'"& vbCRLF
            sqlStr = sqlStr & " where payrequestIdx="&payrequestIdx&VbCRLF
            dbget.Execute sqlStr
        end if
    end if

	IF returnValue = 1 THEN
	%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�...");
			<% If Request("returnurl") <> "" Then %>
			location.href = '<%=Request("returnurl")%>?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			<% Else %>
			 location.href = '/admin/approval/eapp/modeappPayDoc.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			 <% End If %>
			//-->
			</script>
<%
response.end
        ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
CASE "SC" '��������  ����
    sqlStr = "update db_partner.dbo.tbl_eapppayrequest " &VbCRLF
    sqlStr = sqlStr&" SET payrequestState="&payrequestState&VbCRLF
    sqlStr = sqlStr&" WHERE payrequestIdx="&payrequestIdx&VbCRLF
    sqlStr = sqlStr&" and IsNULL(payType,0)<>2"

    dbget.Execute sqlStr,returnValue
	IF returnValue = 1 THEN
	%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�...");
			 location.href = '/admin/approval/eapp/modeappPayDoc.asp?iridx=<%=reportIdx%>&ipridx=<%=payrequestIdx%>'
			//-->
			</script>
<%
response.end
        ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
CASE "SD" '������û�� ����
    sqlStr = "update db_partner.dbo.tbl_eapppayrequest " &VbCRLF
    sqlStr = sqlStr&" SET isusing=0"&VbCRLF
    sqlStr = sqlStr&" WHERE payrequestIdx="&payrequestIdx&VbCRLF
    sqlStr = sqlStr&" and payrequestState=5"

    dbget.Execute sqlStr,returnValue
	IF returnValue = 1 THEN
	%>
	<script language="javascript">
			<!--
				alert("�����Ǿ����ϴ�...");
				opener.location.reload();
			    window.close();
			//-->
			</script>
<%
response.end
        ELSE
			Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.1")
			response.end
		END IF
		
CASE "SR"		'����Ϸ� �� �濵������ �ݷ�ó��
 
    sqlStr = "update db_partner.dbo.tbl_eapppayrequest " &VbCRLF
    sqlStr = sqlStr&" SET payrequestState=5, erplinktype = null, erplinkkey = null " &VbCRLF 
    sqlStr = sqlStr&" WHERE payrequestIdx="&payrequestIdx 
    dbget.Execute sqlStr 
    
   
	sqlStr = "update  db_partner.dbo.tbl_eappauthline "&VbCRLF
	sqlStr = sqlStr&" set authid = '"&adminID&"', authstate = 5 , confirmdate = getdate() "&VbCRLF
	sqlStr = sqlStr&" where payrequestidx ="&payrequestIdx&" and reportidx ="&reportIdx&" and authposition = 2  "
	dbget.Execute sqlStr 
	 
	 
	%>
	<script language="javascript">
			<!--
				alert("�ݷ�ó�� �Ǿ����ϴ�..");
			 location.href = '/admin/approval/payreqlist/?selPRS=5 '
			//-->
			</script>
<%
CASE ELSE
	Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.0")
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

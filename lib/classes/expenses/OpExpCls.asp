<%
'###########################################################
' Description : ������ Ŭ����
' History : 2011.06.03 ������ ����
'###########################################################

Class OpExp
	public FYYYYMM
	public FPartTypeIdx
	public FOpExpidx
	public FOpExpPartIdx
	public FOpExpDailyIdx
	public FYYYYMMDD
	public Farap_cd
	public Farap_nm
	public FinExp
	public FOutExp
	public FOpExpObj
	public FDetailCOnts
	public Fbizsection_cd
	public Fbizsection_nm
	public FsupExp
	public FvatExp
	public FauthNo
  	public FLastMonthExp
	public FTotExp
	public FOpExpPartName

	public FSPageNo
	public FEPageNo
	public FPageSize
	public FCurrPage
	public FTotCnt

	public FadminID
	public FSYYYYMM
	public FEYYYYMM
	public FPart_sn
	public FRectDepartmentID
	public FMode
	public FRectUserid
	public FRectPartsn
	public FState

	public Faccountidx
	public Finouttype


	'��� ����Ʈ
	' /admin/expenses/opexp/index.asp
	public Function fnGetOpExpMonthlyList
		IF FPartTypeIdx ="" THEN FPartTypeIdx = 0
		IF FOpExpPartIdx ="" THEN FOpExpPartIdx = 0
		IF FRectPartsn = "" THEN FRectPartsn = 0
		IF FRectDepartmentID = "" THEN FRectDepartmentID = "-1"

		Dim strSql
		strSql = "[db_partner].[dbo].sp_Ten_OpExpMonthly_getList('"&FSYYYYMM&"','"&FEYYYYMM&"',"&FPartTypeIdx&","&FOpExpPartIdx&",'"&FRectUserid&"',"&FRectPartsn&",'"&FState&"','"&FRectDepartmentID&"')"

		'response.write strSql & "<Br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpMonthlyList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'��� �󼼸���Ʈ
	public Function fnGetOpExpDailyList
		IF FPartTypeIdx = "" THEN FPartTypeIdx = 0
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0
 		IF Farap_cd = "" THEN Farap_cd = 0
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDaily_getListCnt]('"&FYYYYMM&"',"&FPartTypeIdx&","&FOpExpPartIdx&" ,"&Farap_cd&",'"&Fbizsection_nm&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_OpExpDaily_getList('"&FYYYYMM&"',"&FPartTypeIdx&", "&FOpExpPartIdx&","&Farap_cd&",'"&Fbizsection_nm&"',"&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailyList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'��� �� �հ� ����Ʈ
	public Function fnGetOpExpDailySumList
		Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpDaily_getSumList('"&FYYYYMM&"',"&FpartTypeIdx&","&FOpExpPartIdx&","&Farap_cd&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailySumList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'��� ��������
	public Function fnGetOpExpDailyData
	Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDaily_getData]("&FOpExpDailyIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FYYYYMMDD 		= rsget("YYYYMMDD")
			FOpExpPartIdx 	= rsget("OpExpPartIdx")
			Farap_cd 		= rsget("arap_cd")
			FinExp 			= rsget("inExp")
			FOutExp 		= rsget("OutExp")
			FOpExpObj 		= rsget("OpExpObj")
			FDetailCOnts 	= rsget("DetailCOnts")
			Fbizsection_Cd= rsget("bizsection_cd")
			FsupExp 		= rsget("supExp")
			FvatExp 		= rsget("vatExp")
			FauthNo 		= rsget("authNo")
			Finouttype	= rsget("inouttype")
			Fbizsection_nm = rsget("bizsection_nm")
		END IF
		rsget.close
	End Function

	public Function fnGetOpExpMonthlyData
		Dim strSql

		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0
		IF FopExpIdx = "" THEN FopExpIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpMonthly_getData]('"&Fyyyymm&"',"&FOpExpPartIdx&","&FopExpIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FOpExpidx 	= rsget("OpExpidx")
			Fyyyymm		= rsget("yyyymm")
			FOpExpPartIdx= rsget("OpExpPartIdx")
			FLastMonthExp= rsget("LastMonthExp")
			FInExp		= rsget("InExp")
			FOutExp		= rsget("OutExp")
			FTotExp 	= rsget("TotExp")
			FOpExpPartName=rsget("OpExpPartName")
			FPartTypeIdx = rsget("PartTypeIdx")
			Fstate       = rsget("state")
		END IF
		rsget.close
	End Function

	'//���,����,���� ó���� ���� üũ
	public Function fnGetOpExpAuth
	Dim objCmd
 
	 Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExp_getAuth]('"&Fyyyymm&"',"&FOpExpPartIdx&",'"&FMode&"','"&FadminID&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    fnGetOpExpAuth = objCmd(0).Value
	Set objCmd = nothing
	End Function

	'//����� ���� üũ 
	public Function fnGetOpExpPartAuth
	Dim objCmd
	 Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpPart_getAuth]( "&FOpExpPartIdx&",'"&FadminID&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    fnGetOpExpPartAuth = objCmd(0).Value
	Set objCmd = nothing
	End Function
End Class

'//���°�
Function fnGetStateDesc(ByVal iState)
	Dim strState
	IF iState = "1" THEN
		strState = "�ۼ��Ϸ�"
	ELSEIF iState = "5" THEN
		strState = "����������"
	ELSEIF iState = "7" THEN
		strState = "<font color='#3333FF'>����Ϸ�</font>"
	ELSEIF iState = "9" THEN
		strState = "<font color='#11AA11'>Ȯ�οϷ�</font>"
	ELSEIF iState = "10" THEN
		strState = "<font color='#FF33FF'>���ۿϷ�</font>"
	ELSE
		strState = "<font color='red'>�ۼ���</font>"
	END IF
	fnGetStateDesc = strState
End Function

Sub SbOptState(ByVal iState)
	%>
	<option value="">--����--</option>
	<option value="0" <%IF iState ="0" THEN%>selected<%END IF%>>�ۼ���</option>
	<option value="1" <%IF iState ="1" THEN%>selected<%END IF%>>�ۼ��Ϸ�</option>
	<option value="5" <%IF iState ="5" THEN%>selected<%END IF%>>����������</option>
	<option value="7" <%IF iState ="7" THEN%>selected<%END IF%>>����Ϸ�</option>
	<option value="9" <%IF iState ="9" THEN%>selected<%END IF%>>Ȯ�οϷ�</option>
	<option value="10" <%IF iState ="10" THEN%>selected<%END IF%>>���ۿϷ�</option>
	<%
End Sub

'//���Ѱ��� - ������ ����
Function fnChkAdminAuth( ByVal  authLevel, ByVal Partsn)
	Dim strAuth
	strAuth = False
	IF (authLevel<=2  or partsn= 8) THEN
			strAuth = True
	END IF
	fnChkAdminAuth = strAuth
End Function

%>
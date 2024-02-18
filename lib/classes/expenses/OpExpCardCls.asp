<%
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
  public Fauthdate
  public FSAuthDate
  public FEAuthDate
  public FsevExp
  public Fdeducttype
	public FisYYYYMM
	public FPartTypeName

	'��� ����Ʈ
	public Function fnGetOpExpMonthlyList
		IF FPartTypeIdx ="" THEN FPartTypeIdx = 0
		IF FOpExpPartIdx ="" THEN FOpExpPartIdx = 0
		IF FRectPartsn = "" THEN FRectPartsn = 0
		IF FRectDepartmentID = "" THEN FRectDepartmentID = "-1"

		Dim strSql
		strSql = "[db_partner].[dbo].sp_Ten_OpExpMonthlyCard_getList('"&FSYYYYMM&"','"&FEYYYYMM&"',"&FPartTypeIdx&","&FOpExpPartIdx&",'"&FRectUserid&"',"&FRectPartsn&",'"&FState&"','"&FRectDepartmentID&"')"
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpMonthlyList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'��� �󼼸���Ʈ
	public Function fnGetOpExpDailyList
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0
 		IF Farap_cd = "" THEN Farap_cd = 0
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDailyCard_getListCnt]('"&FYYYYMM&"', "&FPartTypeIdx&" ,"&FOpExpPartIdx&" ,"&Farap_cd&",'"&Fbizsection_nm&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_OpExpDailyCard_getList('"&FYYYYMM&"', "&FPartTypeIdx&",  "&FOpExpPartIdx&","&Farap_cd&",'"&Fbizsection_nm&"',"&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailyList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function



	'��� ��û��  �󼼸���Ʈ
	public Function fnGetOpExpDailyNoSetList
	IF FPartTypeIdx = "" THEN FPartTypeIdx = 0
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0
 		IF Farap_cd = "" THEN Farap_cd = 0
 		IF FRectPartsn = "" THEN FRectPartsn = 0
		IF FRectDepartmentID = "" THEN FRectDepartmentID = "NULL"

		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDailyCard_getNoSetListCnt]('"&FSAuthDate&"','"&FEAuthDate&"',"&FPartTypeIdx&" ,"&FOpExpPartIdx&" ,"&Farap_cd&",'"&Fbizsection_nm&"','"&FisYYYYMM&"','"&FRectUserid&"',"&FRectPartsn&","&FRectDepartmentID&")"
		''response.write strSql & "<Br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_OpExpDailyCard_getNoSetList('"&FSAuthDate&"','"&FEAuthDate&"',"&FPartTypeIdx&" ,  "&FOpExpPartIdx&","&Farap_cd&",'"&Fbizsection_nm&"','"&FisYYYYMM&"','"&FRectUserid&"',"&FRectPartsn&","&FSPageNo&","&FEPageNo&","&FRectDepartmentID&")"

		if session("ssBctId")="tozzinet" then
			response.write strSql & "<Br>"
		else
		''response.write strSql & "<Br>"
		end if
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailyNoSetList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'��� �� �հ� ����Ʈ
	public Function fnGetOpExpDailySumList
		Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpDailyCard_getSumList('"&FYYYYMM&"',"&FPartTypeIdx&","&FOpExpPartIdx&","&Farap_cd&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDailySumList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'��� ��������
	public Function fnGetOpExpDailyData
	Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpDailyCard_getData]("&FOpExpDailyIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FYYYYMM 		= rsget("YYYYMM")
			Fauthdate		= rsget("authdate")
			FOpExpPartIdx 	= rsget("OpExpPartIdx")
			Farap_cd 		= rsget("arap_cd")
			FOutExp 		= rsget("OutExp")
			FOpExpObj 		= rsget("OpExpObj")
			FDetailCOnts 	= rsget("DetailCOnts")
			Fbizsection_Cd= rsget("bizsection_cd")
			FsupExp 		= rsget("supExp")
			FvatExp 		= rsget("vatExp")
			FsevExp 		= rsget("sevExp")
			FauthNo 		= rsget("authNo")
			Fdeducttype	= rsget("deducttype")
			Finouttype	= rsget("inouttype")
			Fbizsection_nm = rsget("bizsection_nm")
		END IF
		rsget.close
	End Function

	public Function fnGetOpExpMonthlyData
		Dim strSql
		IF FOpExpPartIdx = "" THEN FOpExpPartIdx = 0
		IF FopExpIdx = "" THEN FopExpIdx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpMonthlyCard_getData]('"&Fyyyymm&"',"&FOpExpPartIdx&","&FopExpIdx&")"

		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FOpExpidx 		= rsget("OpExpCardidx")
			Fyyyymm				= rsget("yyyymm")
			FOpExpPartIdx	= rsget("OpExpPartIdx")
			FOutExp				= rsget("OutExp")
			FOpExpPartName=rsget("OpExpPartName")
			FPartTypeName=rsget("PartTypeName")
			FPartTypeIdx = rsget("PartTypeIdx")
			Fstate       = rsget("state")
		END IF
		rsget.close
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


	'### ERP ������ �������⿡ ���� üũ
	public Sub fERPdataCount()
		Dim strSql, addSql, i
		strSql = "SELECT COUNT(A.Seq) " & vbCrLf
		strSql = strSql & " FROM [TMSDB].[db_scm_Link].[dbo].[vw_reged_Card_AppList] AS A " & vbCrLf
		strSql = strSql & " 	LEFT JOIN [db_partner].[dbo].[tbl_OpExpPart] AS B ON A.CardNo = B.CardNo AND B.isUsing = 1 " & vbCrLf
		strSql = strSql & " 	LEFT JOIN [db_partner].[dbo].[tbl_OpExpDailyCard] AS C ON A.ApprNo = C.authno AND Convert(datetime,(A.TransDate + ' ' + A.TransTime)) = C.authdate " & vbCrLf
		strSql = strSql & " WHERE C.authno is Null AND A.TransDate >= '2012-07-01' " & vbCrLf
'		rw strSql
'		response.write "���� ������"
'		response.end
		rsget.Open strSql,dbget,1
		If Not rsget.Eof Then
			FTotCnt = rsget(0)
		Else
			FTotCnt = 0
		End If
		rsget.close
	end sub

public FRectBankAccNo
public FRectBankNm

    '### ��������ݸ���Ʈ
    public Function fnGetIpkumList
        Dim strSql
        strSql = " SELECT TOP 1 L.inoutidx, L.acct_nm, L.acct_no, L.acct_txday, jeokyo, branch,inout_gubun, tx_amt,tx_cur_bal,erp_datetime "& vbCrLf
        strSql = strSql & " FROM [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT L "& vbCrLf
        strSql = strSql & "    left join [db_order].[dbo].tbl_bank_div A on L.acct_no=A.accountno "& vbCrLf
        strSql = strSql & "    left join db_order.dbo.tbl_ipkum_list i on i.inoutidx = L.inoutidx  "& vbCrLf
        strSql = strSql & " WHERE l.acct_no ='"&FRectBankAccNo&"' and acct_nm ='"&FRectBankNm&"' "& vbCrLf
        strSql = strSql & " ORDER BY  acct_txday desc "
        rsget.Open strSql,dbget,1
		If Not rsget.Eof Then
		    fnGetIpkumList = rsget.getRows()
        END IF
        rsget.close
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

function fnGetPartTypeDesc(ByVal iPartType)
    Dim iPartTypeName
	IF iPartType = "4" THEN
		iPartTypeName = "�ſ�"
	ELSEIF iPartType = "10" THEN
		iPartTypeName = "<font color='#FF33FF'>üũ</font>"
	ELSE
		iPartTypeName = ""
	END IF
	fnGetPartTypeDesc = iPartTypeName
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

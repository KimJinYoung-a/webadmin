<%
'####################################
'	수지항목, 문서 연동 관리
'####################################


Class CArapLinkEdms

public FEdmsName
public FARAPlinkedmsIdx
public FARAP_CD
public FARAPUse
public FARAPDel
public FARAP_NM
public FACC_CD
public FACC_USE_CD
public FACC_NM
public FACCUse
public FACCDel
public Fedmsidx
public FisUsing
public FedmsUsing
public FadminId

public FCateidx1
public FCateidx2

public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage
public FTotCnt

public FARAP_GB
public FCASH_FLOW
public FACC
public Fmatch

'리스트
	public Function fnGetArapLinkEdmsList
	Dim strSql

		strSql ="[db_partner].[dbo].[sp_Ten_ARAPLinkedms_getListCnt]('"&FARAP_GB&"','"&FCASH_FLOW&"','"&FARAP_NM&"','"&FACC&"','"&Fedmsname&"','"&Fmatch&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql = "db_partner.dbo.sp_Ten_ARAPLinkedms_getList('"&FARAP_GB&"','"&FCASH_FLOW&"','"&FARAP_NM&"','"&FACC&"','"&Fedmsname&"','"&Fmatch&"',"&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetArapLinkEdmsList = rsget.getRows()
			END IF
			rsget.close
		END IF
	End Function

'내용보기
	public Function fnGetArapLinkEdmsData
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_ARAPLinkedms_getData]( "&FARAP_CD&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FARAPlinkedmsIdx= rsget("ARAPlinkedmsIdx")
			FARAP_CD       	= rsget("ARAP_CD")
			FARAP_NM       	= rsget("ARAP_NM")
			Fedmsidx       	= rsget("edmsidx")
			Fedmsname       = rsget("edmsname")
			FACC_USE_CD			= rsget("ACC_USE_CD")
			FACC_NM					= rsget("ACC_NM")
		END IF
		rsget.close
	End Function

	'전자결재 문서 선택 리스트 - 수지항목 연계된 경우 수지항목 포함된 리스트
	public Function fnGetEappArapLinkEdmsList
	Dim strSql 
		strSql ="[db_partner].[dbo].[sp_Ten_ARAPLinkedms_getEappListCnt]("&FCateIdx1&","&FCateIdx2&",'"&Fedmsname&"','"&FARAP_NM&"')" 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		if session("ssBctId")="tozzinet" then
		response.write strSql & "<Br>"
		end if
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql = "db_partner.dbo.sp_Ten_ARAPLinkedms_getEappList("&FCateIdx1&","&FCateIdx2&",'"&Fedmsname&"','"&FARAP_NM&"',"&FSPageNo&","&FEPageNo&")"

		'response.write strSql & "<Br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetEappArapLinkEdmsList = rsget.getRows()
			END IF
			rsget.close
		END IF
	End Function

	'전자결재 문서 선택 리스트 - 시급계약직 리스트
	public Function fnGetPartTimeEappArapLinkEdmsList
	Dim strSql 
		strSql ="[db_partner].[dbo].[sp_Ten_ARAPLinkedms_getPartTimeEappListCnt]("&FCateIdx1&","&FCateIdx2&",'"&Fedmsname&"','"&FARAP_NM&"','Y')" 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql = "db_partner.dbo.sp_Ten_ARAPLinkedms_getPartTimeEappList("&FCateIdx1&","&FCateIdx2&",'"&Fedmsname&"','"&FARAP_NM&"',"&FSPageNo&","&FEPageNo&",'Y')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetPartTimeEappArapLinkEdmsList = rsget.getRows()
			END IF
			rsget.close
		END IF
	End Function

	'리스트
	public Function fnGetEappArapLinkNPayEdmsList
	Dim strSql
		IF Fedmsidx = "" THEN Fedmsidx = 0
		strSql ="[db_partner].[dbo].[sp_Ten_ARAPLinkNPayedms_getEappListCnt]('"&FARAP_NM&"','"&Fedmsname&"',"&Fedmsidx&",'"&FadminId&"')"

		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql = "db_partner.dbo.sp_Ten_ARAPLinkNPayedms_getEappList('"&FARAP_NM&"','"&FEdmsName&"',"&Fedmsidx&","&FSPageNo&","&FEPageNo&",'"&FadminId&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetEappArapLinkNPayEdmsList = rsget.getRows()
			END IF
			rsget.close
		END IF
	End Function
End Class
%>
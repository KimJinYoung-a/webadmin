<%
'######################################
' 수지항목 관리 Class
'######################################
Class CARAP
public FARAP_GB
public FCASH_FLOW
public FARAP_NM
public FACC
	public Function fnGetARAPCD 
	Dim strSql
	strSql = "db_partner.dbo.sp_Ten_TMS_BA_ARAP_CD_getList('"&FARAP_GB&"','"&FCASH_FLOW&"','"&FARAP_NM&"','"&FACC&"')" 
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetARAPCD = rsget.getRows()
		END IF
		rsget.close 
	End Function
End Class

'######################################
' 수지항목 관리 function
'######################################
function fnGetARAP_GB(ByVal iValue)
	Dim RValue
	RValue = ""
	IF iValue = "1" THEN
		RValue = "수입"
	ELSEIF iValue="2" THEN
		RValue = "지출"
	END IF	
	fnGetARAP_GB = RValue
End function

function fnGetARAP_Cash(ByVal iValue)
	Dim RValue
	RValue = ""
	IF iValue = "001" THEN
		RValue = "영업"
	ELSEIF iValue="002" THEN
		RValue = "투자"
	ELSEIF iValue="003" THEN
		RValue = "재무"	
	END IF	
	fnGetARAP_Cash = RValue
End function
%>

		
		 
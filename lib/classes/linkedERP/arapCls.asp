<%
'######################################
' �����׸� ���� Class
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
' �����׸� ���� function
'######################################
function fnGetARAP_GB(ByVal iValue)
	Dim RValue
	RValue = ""
	IF iValue = "1" THEN
		RValue = "����"
	ELSEIF iValue="2" THEN
		RValue = "����"
	END IF	
	fnGetARAP_GB = RValue
End function

function fnGetARAP_Cash(ByVal iValue)
	Dim RValue
	RValue = ""
	IF iValue = "001" THEN
		RValue = "����"
	ELSEIF iValue="002" THEN
		RValue = "����"
	ELSEIF iValue="003" THEN
		RValue = "�繫"	
	END IF	
	fnGetARAP_Cash = RValue
End function
%>

		
		 
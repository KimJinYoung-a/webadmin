<%
'###########################################################
' Description : 매출로그 함수관리
' Hieditor : 2013.12.27 정윤정 생성 
'########################################################### 

'//PC사 구분 리스트
Function fnGetCommCode(ByVal sType)
Dim strSql
	strSql = "db_order.dbo.sp_Ten_commcode_getTypeList('"&sType&"')"  
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCommCode = rsget.getRows()
		END IF
	rsget.close
End Function

'//PC사 구분 옵션값
Sub sbGetOptPGgubun(sValue)
	Dim arrList,intLoop
	arrList = fnGetCommCode("PGgubun")
	IF isArray(arrList) THEN
		FOR intLoop = 0 To UBound(arrList,2)
 %> 
<option value="<%=arrList(1,intLoop)%>" <%IF sValue=arrList(1,intLoop) THEN%>selected<%END IF%>><%=arrList(2,intLoop)%></option> 
<% NEXT
	END IF
End Sub

'//PC사ID 리스트
Sub sbGetOptPGID(sValue)
Dim arrList,intLoop
	arrList = fnGetCommCode("PGID")
	IF isArray(arrList) THEN
		FOR intLoop = 0 To UBound(arrList,2)
 %> 
<option value="<%=arrList(1,intLoop)%>" <%IF sValue=arrList(1,intLoop) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option> 
<% NEXT
	END IF
End Sub

%>
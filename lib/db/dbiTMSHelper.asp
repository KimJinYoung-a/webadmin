<%
'==========================================================================
'	Description: DB ������ ���� Ŭ���� ����
'	History: 2007.08.02
'==========================================================================


' ##########################################################################
'// ���ν��� ������ return�� ��ȯ(output ����)
'// fnExecSPReturnValue(������ sp, db��������)
' ##########################################################################
Function fnExecSPReturnValue(ByVal strSql)

    Dim objCmd
    Dim intResult
    Set objCmd = Server.CreateObject("ADODB.Command")
    With objCmd
    	.ActiveConnection = dbiTms_dbget
    	.CommandType = adCmdText
    	.CommandText = "{?=call "&strSql&" }"
    	objCmd(0).Direction =adParamReturnValue
    	.Execute, , adExecuteNoRecords
    End With
    	intResult = objCmd(0).Value
    Set objCmd = nothing
    
    fnExecSPReturnValue = intResult

End Function
	
	
' ##########################################################################
'// ���ν��� ������ ��� ���ڵ帮��Ʈ  ��ȯ 
'// fnExecSPReturnRS(������ sp, db��������)
' ##########################################################################
Function fnExecSPReturnRS(ByVal strSql)

    Dim  arrList	
    dbiTms_rsget.CursorLocation = adUseClient
    dbiTms_rsget.Open strSql,dbiTms_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc	
    If Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF ) Then
    	arrList = dbiTms_rsget.GetRows()
    Else
    	arrList = NULL
    End If
    dbiTms_rsget.Close
    
    fnExecSPReturnRS = arrList

End Function

' ##########################################################################
'// ���ν��� ������ ��� ���ڵ� �迭  ��ȯ 
'// fnExecSPReturnArr(������ sp, db��������, �迭������)
' ##########################################################################
Function fnExecSPReturnArr(ByVal strSql,ByVal iArrCount)	

    Dim  arrValue,intLoop	
    dbiTms_rsget.CursorLocation = adUseClient		
    dbiTms_rsget.Open strSql,dbiTms_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    If Not dbiTms_rsget.EOF Then
    	IF iArrCount = 1 THEN
    		arrValue =  rsget(0)
    	ELSE	
    		ReDim arrValue(iArrCount)
    		For intLoop = 0 To (iArrCount-1)
    		 arrValue(intLoop) = rsget(intLoop)
    		Next
    	END IF	
    End If
    dbiTms_rsget.Close
    
    fnExecSPReturnArr = arrValue

End Function

' ##########################################################################
'// ���ν��� ������ RS���� Output��ȯ
'// fnExecSPReturnRSOutput(������ sp, ��ȯ�迭)
' ##########################################################################
Function fnExecSPReturnRSOutput(ByVal strSql, ByRef params)

	Dim cmd, i
    Set cmd = CreateObject("ADODB.Command")

    cmd.ActiveConnection = dbiTms_dbget
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc
    Set cmd = collectParams(cmd, params)
    'cmd.Parameters.Refresh

    dbiTms_rsget.CursorLocation = adUseClient
    dbiTms_rsget.Open cmd, ,adOpenForwardOnly, adLockReadOnly

    For i = 0 To cmd.Parameters.Count - 1	  
      If cmd.Parameters(i).Direction = adParamOutput OR cmd.Parameters(i).Direction = adParamInputOutput OR cmd.Parameters(i).Direction = adParamReturnValue Then
        If IsObject(params) Then	    
          If params is Nothing Then
            Exit For	        
          End If	      
        Else
          params(i)(4) = cmd.Parameters(i).Value
        End If
      End If
    Next	

	Set cmd.ActiveConnection = Nothing
	Set cmd = Nothing
    Set dbiTms_rsget.ActiveConnection = Nothing

	'Set fnExecSPReturnRSOutput = rsget

End Function


' ##########################################################################
'// ���ν��� ������ RS���� Output��ȯ
'// fnExecSP(������ sp, ��ȯ�迭)
' ##########################################################################
Function fnExecSP(ByVal strSql, ByRef params)

	Dim cmd, i
    Set cmd = CreateObject("ADODB.Command")

    cmd.ActiveConnection = dbiTms_dbget
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc
    Set cmd = collectParams(cmd, params)

	cmd.Execute

    For i = 0 To cmd.Parameters.Count - 1	  
      If cmd.Parameters(i).Direction = adParamOutput OR cmd.Parameters(i).Direction = adParamInputOutput OR cmd.Parameters(i).Direction = adParamReturnValue Then
        If IsObject(params) Then	    
          If params is Nothing Then
            Exit For	        
          End If	      
        Else
          params(i)(4) = cmd.Parameters(i).Value
        End If
      End If
    Next	

	Set cmd.ActiveConnection = Nothing
	Set cmd = Nothing

	fnExecSP = True 

End Function

' ##########################################################################
'// ���ν��� ������ Output �� ��ȯ
'// fnExecSPOutput(������ sp, ��ȯ�迭 ũ��)
' ##########################################################################
Function fnExecSPOutput(ByVal strSql,ByVal arrParm)	
    
    
    Dim objCmd, i
    Set objCmd = Server.CreateObject("ADODB.Command")
    With objCmd
    
    	.ActiveConnection = dbiTms_dbget
    	.CommandType = adCmdStoredProc
    	.CommandText = strSql
    	.Prepared = true
    				
      Set objCmd = collectParams(objcmd, arrParm)
      .Execute 
      For i = 0 To objCmd.Parameters.Count - 1	  
          If objCmd.Parameters(i).Direction = adParamOutput OR objCmd.Parameters(i).Direction = adParamInputOutput OR objCmd.Parameters(i).Direction = adParamReturnValue Then
            If IsObject(arrParm) Then	    
              If arrParm is Nothing Then
                Exit For	        
              End If	      
            Else
              arrParm(i)(4) = objCmd.Parameters(i).Value
            End If
          End If
        Next			
    End With
    Set objCmd = nothing
    	fnExecSPOutput = arrParm
End Function

'---------------------------------------------------
' �迭�� �Ű������� �����.
'---------------------------------------------------
Function MakeParam(PName,PType,PDirection,PSize,PValue)
  MakeParam = Array(PName, PType, PDirection, PSize, PValue)
End Function

'---------------------------------------------------
' �Ű����� �迭 ������ ������ �̸��� �Ű����� ���� ��ȯ�Ѵ�.
'---------------------------------------------------		
Function GetValue(arrParm, paramName)
	Dim param
  For Each param in arrParm           
    If param(0) = paramName Then        	
      GetValue = param(4)
      Exit Function
    End If
  Next
End Function


'---------------------------------------------------------------------------
'Array�� �Ѱܿ��� �Ķ���͸� Parsing �Ͽ� Parameter ��ü��
'�����Ͽ� Command ��ü�� �߰��Ѵ�.
'---------------------------------------------------------------------------
Function collectParams(objCmd,arrParm)
	Dim i,l,u,v

    If VarType(arrParm) = 8192 or VarType(arrParm) = 8204 or VarType(arrParm) = 8209 then 		'�迭���� Ȯ��    
	    For i = LBound(arrParm) To UBound(arrParm)
		    l = LBound(arrParm(i))
		    u = UBound(arrParm(i))
		     
		    ' Check for nulls.
		    If u - l = 4 Then
				
			    If VarType(arrParm(i)(4)) = vbString Or VarType(arrParm(i)(4)) = 0 Then
				    If arrParm(i)(4) = "" Then
					    v = Null
				    Else
					    v = arrParm(i)(4)
				    End If
			    Else
				    v = arrParm(i)(4)
			    End If				   
'rw v
			    objCmd.Parameters.Append objCmd.CreateParameter(arrParm(i)(0), arrParm(i)(1), arrParm(i)(2), arrParm(i)(3), v)
		    End If			   
	    Next

	    Set collectParams = objCmd
	    Exit Function
    Else
	    Set collectParams = objCmd
    End If
End Function


' ##########################################################################
	'// Ʈ����� �̿� ��Ƽ ���ν��� ������ return�� ��ȯ(output ����)
'// fnMultiExecSPReturnValue(������ sp, db��������)
' ##########################################################################
Function fnMultiExecSPReturnValue(ByVal strSql)

    Dim objCmd
    Dim intResult
    
    Set objCmd = Server.CreateObject("ADODB.Command")
    With objCmd
    
    	.ActiveConnection = dbiTms_dbget
    	.CommandType = adCmdText
    	.CommandText = "{?=call "&strSql&" }"		
    	objCmd(0).Direction =adParamReturnValue
    	.Execute, , adExecuteNoRecords
    End With
    	intResult = objCmd(0).Value
    Set objCmd = nothing
    
    fnMultiExecSPReturnValue = intResult

End Function



'---------------------------------------------------
' SQL RecordSet ��ȯ
'---------------------------------------------------
Function RecordSQL(ByVal strSQL, ByVal params)
	Dim cmd	: Set cmd = CreateObject("ADODB.Command")
	cmd.ActiveConnection = dbiTms_dbget
	cmd.CommandText = strSQL
	cmd.CommandType = adCmdText							' SQL(Prepared Statement)

    Set cmd = collectParams(cmd, params)				' Append Parameters
	dbiTms_rsget.CursorLocation = adUseClient					' Ŭ���̾�Ʈ Ŀ��(Disconnected Recordset)
	'dbiTms_rsget.Open cmd, ,adOpenStatic, adLockReadOnly		' RecordCount, ����¡ ����
	dbiTms_rsget.Open cmd, ,adOpenForwardOnly, adLockReadOnly	' ���� ����

	Set cmd.ActiveConnection = Nothing
	Set cmd = Nothing
	Set dbiTms_rsget.ActiveConnection = Nothing

End Function

' prepared SQL�� �Ķ���� �迭 �߰� �Լ�
Sub redimParam(ByRef params, ByVal pName, ByVal pType, ByVal pDirection, ByVal pSize, ByVal pValue)
	Dim i
	If IsArray(params) Then
		i = UBound(params)+1
		ReDim Preserve params(i)
	Else
		i = 0
		ReDim params(0)
	End If 
	params(i) = Array(pName, pType, pDirection, pSize, pValue)
End Sub


' SQL RecordSet ���� �����
' Count  ���� : �÷� �Ķ���� ����
' Paging ���� : ����¡ �Ķ���� �ʿ�
' SELECT ���� : ����¡ �Ķ���� ����
Function makeQuery(ByVal sqlColumn, ByVal sqlTable, ByVal sqlWhere, ByVal sqlOrder, ByVal CurrPage, ByVal PageSize, ByVal sqlGroup)
	Dim sql 
	If PageSize <> "" Then	' CTE ����¡ SELECT ����
		sql = "	;WITH CTE_LIST AS (" & vbCrLf
		sql = sql & "		SELECT" & vbCrLf
		sql = sql & "			ROW_NUMBER() OVER ( " & sqlOrder & " ) AS RowNum,	" & vbCrLf
		sql = sql & "			" & sqlColumn & vbCrLf
		sql = sql & "		" & sqlTable & vbCrLf
		sql = sql & "		" & "WHERE 1=1" & vbCrLf
		sql = sql & "		" & sqlWhere & vbCrLf
		sql = sql & "		" & sqlGroup & vbCrLf
		sql = sql & "	) SELECT * FROM CTE_LIST " & vbCrLf
		sql = sql & "	WHERE RowNum BETWEEN " & (PageSize * (CurrPage - 1) + 1) & " AND " & (PageSize * CurrPage) & vbCrLf
		sql = sql & "	ORDER BY RowNum ASC " & vbCrLf
	Else					' ī��Ʈ ���� or �Ϲ� SELECT ����
		sql = "	SELECT" & vbCrLf
		If sqlColumn = "" Then	' ī��Ʈ ������ ��� ORDER BY �� ����
			sqlOrder = ""
			sql = sql & "		Count(*) cnt " & vbCrLf
		Else
			sql = sql & "		" & sqlColumn & vbCrLf
		End If 
		sql = sql & "	" & sqlTable & vbCrLf
		sql = sql & "	" & "WHERE 1=1" & vbCrLf
		sql = sql & "	" & sqlWhere & vbCrLf
		sql = sql & "	" & sqlGroup & vbCrLf
		sql = sql & "	" & sqlOrder & vbCrLf

		If sqlColumn = "" And sqlGroup <> "" Then	' ī��Ʈ������ GROUP BY ���� ������ ���������� �ٽ� ī��Ʈ
			sql = "SELECT Count(*) cnt FROM (" & vbCrLf & sql
			sql = sql & ") t" & vbCrLf
		End If 
	End If 
	makeQuery = sql
End Function 

%>

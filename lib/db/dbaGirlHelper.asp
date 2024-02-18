<%

' ##########################################################################
'// 프로시져 실행후 RS리턴 Output반환
'// fnExecSPReturnRSOutput(실행할 sp, 반환배열)
' ##########################################################################
Function dbaGirl_fnExecSPReturnRSOutput(ByVal strSql, ByRef params)

	Dim cmd, i
    Set cmd = CreateObject("ADODB.Command")

    cmd.ActiveConnection = dbagirl_dbget
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc
    Set cmd = dbagirl_collectParams(cmd, params)
    'cmd.Parameters.Refresh

    dbagirl_rsget.CursorLocation = adUseClient
    dbagirl_rsget.Open cmd, ,adOpenForwardOnly, adLockReadOnly

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
    Set dbagirl_rsget.ActiveConnection = Nothing

	'Set fnExecSPReturnRSOutput = rsget

End Function

' ##########################################################################
'// 프로시져 실행후 Output 값 반환
'// fnExecSPOutput(실행할 sp, 반환배열 크기)
' ##########################################################################
Function dbagirl_fnExecSPOutput(ByVal strSql,ByVal arrParm)	
    
    
    Dim objCmd, i
    Set objCmd = Server.CreateObject("ADODB.Command")
    With objCmd
    
    	.ActiveConnection = dbagirl_dbget
    	.CommandType = adCmdStoredProc
    	.CommandText = strSql
    	.Prepared = true
    				
      Set objCmd = dbagirl_collectParams(objcmd, arrParm)
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
    dbagirl_fnExecSPOutput = arrParm
End Function

'---------------------------------------------------------------------------
'Array로 넘겨오는 파라메터를 Parsing 하여 Parameter 객체를
'생성하여 Command 객체에 추가한다.
'---------------------------------------------------------------------------
Function dbagirl_collectParams(objCmd,arrParm)
	Dim i,l,u,v

    If VarType(arrParm) = 8192 or VarType(arrParm) = 8204 or VarType(arrParm) = 8209 then 		'배열여부 확인
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

	    Set dbagirl_collectParams = objCmd
	    Exit Function
    Else
	    Set dbagirl_collectParams = objCmd
    End If
End Function


'---------------------------------------------------
' 배열로 매개변수를 만든다.
'---------------------------------------------------
Function MakeParam(PName,PType,PDirection,PSize,PValue)
  MakeParam = Array(PName, PType, PDirection, PSize, PValue)
End Function


'---------------------------------------------------
' 매개변수 배열 내에서 지정된 이름의 매개변수 값을 반환한다.
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


%>
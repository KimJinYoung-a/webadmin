<%

' ##########################################################################
'// ���ν��� ������ RS���� Output��ȯ
'// fnExecSPReturnRSOutput(������ sp, ��ȯ�迭)
' ##########################################################################
Function db3_fnExecSPReturnRSOutput(ByVal strSql, ByRef params)

	Dim cmd, i
    Set cmd = CreateObject("ADODB.Command")

    cmd.ActiveConnection = db3_dbget
    cmd.CommandText = strSql
    cmd.CommandType = adCmdStoredProc
    Set cmd = db3_collectParams(cmd, params)
    'cmd.Parameters.Refresh

    db3_rsget.CursorLocation = adUseClient
    db3_rsget.Open cmd, ,adOpenForwardOnly, adLockReadOnly

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
    Set db3_rsget.ActiveConnection = Nothing

	'Set fnExecSPReturnRSOutput = rsget

End Function

'---------------------------------------------------------------------------
'Array�� �Ѱܿ��� �Ķ���͸� Parsing �Ͽ� Parameter ��ü��
'�����Ͽ� Command ��ü�� �߰��Ѵ�.
'---------------------------------------------------------------------------
Function db3_collectParams(objCmd,arrParm)
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

	    Set db3_collectParams = objCmd
	    Exit Function
    Else
	    Set db3_collectParams = objCmd
    End If
End Function

%>
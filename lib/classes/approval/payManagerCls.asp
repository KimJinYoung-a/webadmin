<%	
 Class CPayManager
 public FPayManagerIdx
 public Fuserid
 public Fusername
 public Fjob_sn
 public Fjob_name
 public Fpart_sn
 public FPayManagerType
 public FisUsing
 public FisDef
 
	'//manager list
	Function fnGetPayManagerList
		Dim strSql	 
		strSql ="[db_partner].[dbo].sp_Ten_eappPayManager_getList"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetPayManagerList = rsget.getRows()
		END IF
		rsget.close
	End Function
	
	'//manager data
	Function fnGetPayManagerData
		Dim strSql	 
		strSql ="[db_partner].[dbo].sp_Ten_eappPayManager_getData("&FPayManagerIdx&")"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			 Fuserid 			= rsget("userid")
			 FpayManagerType 	= rsget("payManagerType")
			 Fusername  		= rsget("username")
			 Fjob_name 			= rsget("job_name")	
			 Fjob_sn 			= rsget("job_sn")	
			 Fpart_sn			= rsget("part_sn")	
			 FisUsing 			= rsget("isUsing")
			 FisDef				= rsget("isDef")
		END IF
		rsget.close
	End Function 
	
	public Function fnGetPayManager
	Dim strSql	 
		strSql ="[db_partner].[dbo].sp_Ten_eAppPayManager_getPayManager('"&FisDef&"','"&Fuserid&"')"	  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		 fnGetPayManager	= rsget.getRows()
		END IF
		rsget.close 
	End Function
End Class 

	Function fnGetPayManagerTypeDesc(ByVal payManagerType)
		IF payManagerType = "2" THEN
			fnGetPayManagerTypeDesc = "재무회계담당"
		ELSEIF	payManagerType = "1" THEN
			fnGetPayManagerTypeDesc = "최종승인자"
		END IF	
	End Function
%>
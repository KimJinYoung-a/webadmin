<%
CLASS CEappList
public Fcateidx1
public FDatetype
public FStartDate
public FEndDate
public FedmsCode
public FuserName
public FreportName
public FreportState
public Farap_cd
public Farap_nm
public FSPageNo
public FEPageNo
public FTotCnt
public FPageSize
public FCurrPage 
public FOrderType

	public Fdepartment_id
	public FdepartmentNameFull
	public Finc_subdepartment
public FRectNoDepartOnly

		public Function fnGetEappList
		Dim strSql		
		IF Fcateidx1 = "" THEN Fcateidx1 = 0
		IF Farap_cd = "" or Farap_nm ="" THEN Farap_cd =0	
		dim tmpDepartmentID : tmpDepartmentID = Fdepartment_id
		if (FRectNoDepartOnly = "Y") then
			'// 부서 미지정
			tmpDepartmentID = -999
		end if
		strSql ="[db_partner].[dbo].[sp_TEn_eappReport_getAllListCnt]("&Fcateidx1&",'"&FDateType&"','"&FStartDate&"','"&FEndDate&"','"&FedmsCode&"','"&FuserName&"','"&FreportName&"','"&FreportState&"', "&Farap_cd&",'"&Farap_nm&"','"&tmpDepartmentID&"','"&Finc_subdepartment&"')"	  
		 
		 rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		 
		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage		
		
		strSql ="[db_partner].[dbo].sp_Ten_eappReport_getAllList("&Fcateidx1&",'"&FDateType&"','"&FStartDate&"','"&FEndDate&"','"&FedmsCode&"','"&FuserName&"','"&FreportName&"','"&FreportState&"', "&Farap_cd&",'"&Farap_nm&"','"&tmpDepartmentID&"','"&Finc_subdepartment&"','"&FOrderType&"' ,"&FSPageNo&","&FEPageNo&")"	 
	 ' rw strSql
	  rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetEappList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function
End CLASS
%>
<%
Class CBsuiness
public FBusiNo
public FBusiName
public FBusiIdx
public Fuserid		      
public FbusiCEOName	  
public FbusiAddr         
public FbusiType         
public FbusiItem         
public FrepName			
public FrepEmail        
public FrepTel          
public FconfirmYn       
public Fregdate         
public FdelYn           
public FguestOrderserial
public FuseType         
 
 '//리스트
	public Function fnGetBusinessList
		Dim strSql	 
		strSql ="[db_order].[dbo].sp_Ten_busiInfo_getList('"&FBusiNo&"','"&FBusiName&"')"  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBusinessList = rsget.getRows()
		END IF
		rsget.close 
	End Function

'//상세내역	
	public Function fnGetBusinessData
		Dim strSql	 
		strSql ="[db_order].[dbo].[sp_Ten_busiInfo_getData]("&FBusiIdx&")"  
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN  
			Fuserid			= rsget("userid")
			FbusiNo			= rsget("busiNo")
			FbusiName       = rsget("busiName")
			FbusiCEOName	= rsget("busiCEOName")
			FbusiAddr       = rsget("busiAddr")
			FbusiType       = rsget("busiType")
			FbusiItem       = rsget("busiItem")
			FrepName			= rsget("repName")
			FrepEmail           = rsget("repEmail")
			FrepTel             = rsget("repTel") 
			Fregdate            = rsget("regdate")
			FdelYn              = rsget("delYn")
			FguestOrderserial   = rsget("guestOrderserial")
			FuseType            = rsget("useType")
		END IF
		rsget.close
	End Function
	
End Class

%>
<%
'################################################
' 업체어드민 기본정보 
' 2014-05-07 생성
'################################################

Class CPartner
public FRectGroupID
public FRectShiftID

public Fid           
public Fcompany_name 
public Ftel          
public Ffax          
public Furl          
public Femail        
public Fbigo         
public Fuserdiv      
public Fgroupid      
public Fcuserdiv     

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'로그인 업체의 브랜드 리스트 가져오기
	public Function fnGetBrandList
		dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_partnerA_BrandList('"&FRectGroupID&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetBrandList = rsget.getRows()
			END IF
			rsget.close
	End Function
	
	'브랜드변경 재로그인시 체크
	public Function fnGetBrandChangeLogin
		dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_partnerA_changeLogin('"&FRectShiftID&"','"&FRectGroupID&"')" 
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
			Fid           = rsget("id")
			Fcompany_name = rsget("company_name")
			Ftel          = rsget("tel")
			Ffax          = rsget("fax")
			Furl          = rsget("url")
			Femail        = rsget("email")
			Fbigo         = rsget("bigo")
			Fuserdiv      = rsget("userdiv")
			Fgroupid      = rsget("groupid")
			Fcuserdiv     = rsget("cuserdiv")  
			END IF
			rsget.close
	End Function
	

End Class
%>
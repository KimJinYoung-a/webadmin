<%
Class CCategoryMatching

public FRectCateLarge
public FRectCateMid
public FRectCateSmall
public FRectDispCate
public FRectIsNotMatching

public FTotCnt
public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage

public FCateLargeName
public FCateMidName
public FCateSmallName
public FDispCate
public FCateLarge
public FCateMid
public FCateSmall

	Private Sub Class_Initialize()
 		FRectDispCate = 0
 		FRectIsNotMatching = "N"
    End Sub

	Private Sub Class_Terminate()

	End Sub
	
	'카테고리 매칭 리스트
	'/admin/categorymaster/
	public Function fnGetCategoryList
		Dim strSql 
		if FRectDispCate = "" then FRectDispCate = 101 
		strSql ="[db_item].[dbo].sp_Ten_CategoryMatching2_getList("&FRectDispCate&",'"&FRectCateLarge&"','"&FRectCateMid&"','"&FRectCateSmall&"','"&FRectIsNotMatching&"' )" 
		 
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCategoryList = rsget.getRows()
		END IF
		rsget.close
	 
	END Function	
	
	'관리카테고리 명
	'/admin/categorymaster/popMatching.asp
	public Function fnGetDispCateFullName
		Dim strSql 
		strSql ="select [db_item].[dbo].[getCate2CodeFullDepthName]("&FRectDispCate&")" 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetDispCateFullName = rsget(0) 
		END IF
		rsget.close 
	END Function	
	
	'매칭 카테고리 가져오기
	'/admin/categorymaster/popMatching.asp
	public Function fnGetCategoryDisp
		Dim strSql 
		strSql ="[db_item].[dbo].[sp_Ten_CategoryMatching2_GetData]("&FRectDispCate&")"
 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN 
			FCateLarge	= rsget("code_large")
			FCateMid	= rsget("code_mid")
			FCateSmall	= rsget("code_small")
			FCateLargeName = rsget("cdl_nm")
			FCateMidName   = rsget("cdm_nm")
			FCateSmallName = rsget("cds_nm") 
		END IF
		rsget.close 
	END Function	
End Class	
%>
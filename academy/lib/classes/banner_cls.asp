<%
'#############################################
' PageName : /academy/lib/classes/banner_cls.asp	
' Description : �ΰŽ� ��� ����
' History : 2006.11.16 ������ ����
'#############################################

'// ClsBanner : ��� ����Ʈ �����ֱ�
Class ClsBanner 
	public FBannerCnt	'Set Banner �� ����
	public FCPage		'Get ������������ȣ
	public FPSize		'Get �� ���������� �ִ� ������ ���ڵ� ����	
	public FLocation 
	
	public Function fnGetBannerList
		Dim strSql, strSqlCnt, AddSql, AddWSql
		Dim iDelCnt
		
		If FLocation <> "" THEN
			AddSql = " and a.commCd = '"&FLocation&"'"
			AddWSql = " WHERE commCd = '"&FLocation&"'"
		END IF
		
		strSqlCnt = " SELECT COUNT(bannerId) FROM [db_academy].[dbo].[tbl_banner] " & AddWSql
		rsACADEMYget.Open strSqlCnt, dbACADEMYget, 1
			IF not rsACADEMYget.eof THEN
				FBannerCnt = rsACADEMYget(0)
			END IF	
		rsACADEMYget.close		 				 	
		
		IF FBannerCnt > 0 THEN
			iDelCnt =  (FCPage - 1) * FPSize
			
			strSql = " SELECT TOP  "&FPSize&" a.bannerId, a.imgUrl, a.linkUrl, a.commCd, a.isUsing, a.regdate, a.adminID, b.commNm "&_
					 " FROM [db_academy].[dbo].[tbl_banner] as a Inner Join  [db_academy].[dbo].[tbl_commCd] as b "&_
					 " on a.commCd = b.commCd  "&_
					 " WHERE a.bannerId not in ( SELECT TOP "&iDelCnt&" bannerId  FROM [db_academy].[dbo].[tbl_banner] "&AddWSql&" ) " & AddSql
			rsACADEMYget.Open strSql, dbACADEMYget, 1
				IF not rsACADEMYget.eof THEN
					fnGetBannerList = rsACADEMYget.getRows()
				END IF	
			rsACADEMYget.close		 				 	
		END IF	
	End Function	
	
End Class

Class ClsBannerCont
	public FBannerId
	public FImgUrl
	public FLink
	public FCommCd
	public FisUsing
	public FRegdate
	public FAdminId
	public FWidth
	public FHeight
	public FWindow
	
	public Sub sbGetBannerView
		Dim strSql
		strSql = " SELECT a.bannerId, a.imgUrl, a.linkUrl, a.commCd, a.isUsing, a.regdate, a.adminID, b.commNm, a.sWidth, a.sHeight, a.sWindow "&_
		 		" FROM [db_academy].[dbo].[tbl_banner] as a Inner Join  [db_academy].[dbo].[tbl_commCd] as b "&_
		 		" on a.commCd = b.commCd WHERE a.bannerId = "&FBannerId
		rsACADEMYget.Open strSql, dbACADEMYget, 1
			IF not rsACADEMYget.eof THEN
				FImgUrl = rsACADEMYget("imgUrl")
				FLink = rsACADEMYget("linkUrl")
				FCommCd = rsACADEMYget("commCd")
				FisUsing = rsACADEMYget("isUsing")
				FRegdate = rsACADEMYget("regdate")
				FAdminId = rsACADEMYget("adminID")
				FWidth = rsACADEMYget("sWidth")
				FHeight = rsACADEMYget("sHeight")
				FWindow = rsACADEMYget("sWindow")
			END IF	
		rsACADEMYget.close 		
	End Sub
End Class
%>
<%
'################################################
' Webadmin 디자인핑거스 데이터 가져오기
'################################################
 Class CDesignFingers
  public FDFSeq
  public FTitle
  public FCPage
  public FPSize
  public FTotCnt
  public FTotImgCnt
  
  public FDFType 	
  public FDFTypeDesc
    
  public FContents 	
  public FPrizeDate 
  public FComment
  public FIsDisplay 
  public FIsMainDisplay
  public FIsOtherMall 
  public FUserid	
  public FRegDate	
  
  public FProdName
  public FProdSize
  public FProdColor
  public FProdJe
  public FProdGu
  public FProdSpe
  public FTag
  
  public FEDId
  public FEMKTId
  
  public FItemid  
  public FIsMovie
  public FOpenDate
  public FImg(50,50,7)
 
  '//리스트 가져오기
  public Function fnGetDFList
  	Dim strSql, strSqlCnt, strSearch,iDelCnt
  	FTotCnt = 0
  	IF FDFSeq <> "" THEN 	strSearch = " AND DFSeq = "&FDFSeq	
  	IF FTitle <> "" THEN 	strSearch = strSearch & " AND Title like '%"&FTitle&"%'"
  	IF FEDId <> "" THEN 	strSearch = " AND designerid = '" & FEDId & "' "
  	IF FEMKTId <> "" THEN 	strSearch = " AND partMKTid = '" & FEMKTId & "' "

  	strSqlCnt = " SELECT COUNT(DFSeq) FROM [db_sitemaster].[dbo].[tbl_designfingers] Where isUsing =1 "&strSearch
  	rsget.Open strSqlCnt,dbget	
 	IF Not rsget.EOF THEN
 		FTotCnt = rsget(0)
 	END IF
 	rsget.Close	
 	
 	IF FTotCnt > 0 THEN
 		iDelCnt =   ((FCPage - 1) * FPSize) + 1
  	strSql = " SELECT TOP "&FPSize&"  DFSeq, DFType, Title, PrizeDate,isDisplay, RegDate "&_
  			" 		, (SELECT ImgURL FROM [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = A.DFSeq AND  DFCodeSeq = 3 AND  DFImgID =1 ) AS Img "&_
  			"		, OpenDate " & _
  			"  FROM  [db_sitemaster].[dbo].[tbl_designfingers] AS A  WHERE isUsing = 1 AND DFSeq <= ( SELECT MIN(DFSeq) "&_
  			"		  FROM ( SELECT Top "&iDelCnt&" DFSeq FROM [db_sitemaster].[dbo].[tbl_designfingers] WHERE IsUsing =1 "&strSearch&" Order by DFSeq DESC) as DumpTable ) "& strSearch &_	
  			"  Order by DFSeq DESC	 "
  	rsget.Open strSql,dbget	
 	IF Not rsget.EOF THEN
 		fnGetDFList = rsget.getRows()
 	END IF
 	rsget.Close	
	END IF	
  End Function
  
  '//특정id 내용 가져오기
  	public Function fnGetDFCont
  		Dim strSql, arrItemid, arrImage, intLoop
  		strSql = " SELECT  [DFSeq], [DFType], [Title], [Contents], [PrizeDate],[Comment], [IsDisplay], [IsOtherMall], [Userid], [RegDate], [IsMainDisplay] "&_
  			"			, [ProdName], [ProdSize], [ProdColor], [ProdJe], [ProdGu], [ProdSpe], [IsMovie], [OpenDate], [Tag], [designerid], [partMKTid] " & _
  			" FROM [db_sitemaster].[dbo].[tbl_designfingers] WHERE DFSeq = "&FDFSeq
  		rsget.Open strSql,dbget	
 		IF Not rsget.EOF THEN
 			 FDFType 		= rsget(1)
 			 FTitle  		= db2html(rsget(2))
 			 FContents 		= db2html(rsget(3))
 			 FPrizeDate 	= rsget(4)
 			 FComment		= rsget(5)
 			 FIsDisplay 	= rsget(6)
 			 FIsOtherMall 	= rsget(7)
 			 FUserid		= rsget(8)
 			 FRegDate		= rsget(9)
 			 FIsMainDisplay	= rsget(10)
 			 FProdName		= rsget(11)
 			 FProdSize		= rsget(12)
 			 FProdColor		= rsget(13)
 			 FProdJe		= rsget(14)
 			 FProdGu		= rsget(15)
 			 FProdSpe		= rsget(16)
 			 FIsMovie		= rsget(17)
 			 FOpenDate		= rsget(18)
 			 FTag			= db2html(rsget(19))
 			 FEDId			= rsget(20)
 			 FEMKTId		= rsget(21)
 		END IF
 		rsget.close 	
 		
 		strSql = "SELECT itemid from [db_event].[dbo].[tbl_eventitem] where evt_code =1 AND evtgroup_code = "&FDFSeq
 		rsget.Open strSql,dbget	
 		IF Not rsget.EOF THEN
 			FItemid =  rsget.getRows()
 		END IF
 		rsget.close
 		
 		strSql = " SELECT [DFSeq], [DFCodeSeq], [DFImgID], [ImgURL], [Link], [ImgDescCode],[RegDate] "&_
 				" FROM [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFSeq = "&FDFSeq&_
 				" ORDER By DFCodeSeq, DFImgID "  									 				
 		rsget.Open strSql,dbget	
 		IF Not rsget.EOF THEN
 			arrImage=  rsget.getRows()
 		END IF		
 			
 		rsget.close 		
	
		IF isArray(arrImage) THEN
			FTotImgCnt = 0
			For intLoop = 0 To UBound(arrImage,2)			 
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),0)	 = arrImage(0,intLoop)
				
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),1)	 = arrImage(1,intLoop)

				FImg(arrImage(1,intLoop),arrImage(2,intLoop),2)	 = arrImage(2,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),3)	 = arrImage(3,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),4)	 = arrImage(4,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),5)	 = arrImage(5,intLoop)		
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),6)	 = arrImage(6,intLoop)	
				IF arrImage(1,intLoop) = 5 THEN
					FTotImgCnt = FTotImgCnt + 1
				End If						
			Next
		END IF	
	End Function
	
	
	'//최근 3개월 코멘트 수순 리스트
	Function fnGetBestComment
		Dim strSql '// DFCodeSeq 6 : 배너 이미지 코드
		strSql = " SELECT   A.DFSeq, A.DFType, A.Title, A.PrizeDate, A.isDisplay, A.RegDate, B.ImgURL "&_
				"		 , (select count(id) from [db_sitemaster].[dbo].[tbl_zf_comments] where masterid = A.DFSeq  and isdelete ='N') as commcnt "&_
 				" FROM [db_sitemaster].[dbo].[tbl_designfingers] AS A left outer join  [db_sitemaster].[dbo].[tbl_designfingers_image] as B on A.DFSeq = B.DFSeq and B.DFCodeSeq = 6 AND B.DFImgID =1 "&_
				" WHERE  A.isusing = 1  and datediff(m,A.regdate,getdate()) < 3 "&_
				" order by commcnt desc "
		rsget.Open strSql,dbget	
 		IF Not rsget.EOF THEN
 			fnGetBestComment=  rsget.getRows()
 		END IF		
 			
 		rsget.close 		
	End Function
	
	'//요약내용 가져오기
	public Function fnGetDFSummary
  		Dim strSql
  		strSql = " SELECT  [DFSeq], a.[DFType], [Title], [PrizeDate], b.CodeDesc "&_
  			" FROM [db_sitemaster].[dbo].[tbl_designfingers] a, [db_sitemaster].[dbo].tbl_designfingers_code b"&_
  			" WHERE a.DFType= b.DFCodeSeq and b.PCodeSeq = 10 AND DFSeq = "&FDFSeq
  		rsget.Open strSql,dbget	
 		IF Not rsget.EOF THEN
 			 FDFType 	= rsget(1)
 			 FTitle  	= db2html(rsget(2))
 			 FPrizeDate = rsget(3)
 			 FDFTypeDesc = rsget(4)
 		END IF
 		rsget.close 	
 	END Function	
 	
 	
  '//특정id 내용 가져오기
  	public Function fnGetDFContMImage
  		Dim strSql, arrImage, intLoop
 		strSql = " SELECT [DFSeq], [DFCodeSeq], [DFImgID], [ImgURL], [Link], [ImgDescCode],[RegDate] "&_
 				" FROM [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFCodeSeq = '24' AND DFSeq = "&FDFSeq&_
 				" ORDER By DFCodeSeq, DFImgID "  									 				
 		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			arrImage=  rsget.getRows()
 		END IF		
 			
 		rsget.close 
 		
		IF isArray(arrImage) THEN
			FTotImgCnt = 0
			For intLoop = 0 To UBound(arrImage,2)			 
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),0)	 = arrImage(0,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),1)	 = arrImage(1,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),2)	 = arrImage(2,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),3)	 = arrImage(3,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),4)	 = arrImage(4,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),5)	 = arrImage(5,intLoop)		
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),6)	 = arrImage(6,intLoop)	
				IF arrImage(1,intLoop) = 24 THEN
					FTotImgCnt = FTotImgCnt + 1
				End If						
			Next
		END IF	
	END Function
	
	
  '//특정id 내용 가져오기
  	public Function fnGetDFContSourceImage
  		Dim strSql, arrImage, intLoop
 		strSql = " SELECT [DFSeq], [DFCodeSeq], [DFImgID], [ImgURL], [Link], [ImgDescCode],[RegDate] "&_
 				" FROM [db_sitemaster].[dbo].[tbl_designfingers_image] WHERE DFCodeSeq = '25' AND DFSeq = "&FDFSeq&_
 				" ORDER By DFCodeSeq, DFImgID "  									 				
 		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			arrImage=  rsget.getRows()
 		END IF		
 			
 		rsget.close 
 		
		IF isArray(arrImage) THEN
			FTotImgCnt = 0
			For intLoop = 0 To UBound(arrImage,2)			 
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),0)	 = arrImage(0,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),1)	 = arrImage(1,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),2)	 = arrImage(2,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),3)	 = arrImage(3,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),4)	 = arrImage(4,intLoop)
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),5)	 = arrImage(5,intLoop)		
				FImg(arrImage(1,intLoop),arrImage(2,intLoop),6)	 = arrImage(6,intLoop)	
				IF arrImage(1,intLoop) = 25 THEN
					FTotImgCnt = FTotImgCnt + 1
				End If						
			Next
		END IF	
	END Function
	
	
	
  '//추천리스트 가져오기
  public Function fnGetRecommendList
  	Dim strSql, strSqlCnt, strSearch,iDelCnt
  	FTotCnt = 0

	IF FItemID <> "" THEN 	strSearch = strSearch & " AND A.itemid = '" & FItemID & "' "
  	IF FUserid <> "" THEN 	strSearch = strSearch & " AND u.userid = '" & FUserid & "' "

  	strSqlCnt = " SELECT COUNT(IDX) FROM [db_sitemaster].[dbo].[tbl_designfingers_recommend] Where 1=1 " & Replace(Replace(strSearch,"A.",""),"u.","")
  	rsget.Open strSqlCnt,dbget	
 	IF Not rsget.EOF THEN
 		FTotCnt = rsget(0)
 	END IF
 	rsget.Close	
 	
 	IF FTotCnt > 0 THEN
 		iDelCnt =   ((FCPage - 1) * FPSize) + 1
  	strSql = " SELECT TOP "&FPSize&"  A.IDX, A.USERID, A.ITEMID, A.CONTENTS, A.USEYN, A.REGDATE, i.smallimage, u.username "&_
  			"  FROM  [db_sitemaster].[dbo].[tbl_designfingers_recommend] AS A "&_
  			"		LEFT JOIN [db_item].[dbo].[tbl_item] AS i on A.itemid = i.itemid "&_
  			"		LEFT JOIN [db_user].[dbo].[tbl_user_n] AS u on A.USERID = u.userid "&_
  			"  WHERE IDX <= ( SELECT MIN(IDX) "&_
  			"		  FROM ( SELECT Top "&iDelCnt&" IDX FROM [db_sitemaster].[dbo].[tbl_designfingers_recommend] AS AA "&_
  			"			LEFT JOIN [db_item].[dbo].[tbl_item] AS ii on AA.itemid = ii.itemid LEFT JOIN [db_user].[dbo].[tbl_user_n] AS uu on AA.USERID = uu.userid "&_
  			"			WHERE 1=1 " & Replace(Replace(strSearch,"A.","AA."),"u.","uu.") & " Order by AA.IDX DESC) as DumpTable ) "& strSearch &_	
  			"  Order by A.IDX DESC	 "
  	rsget.Open strSql,dbget	
 	IF Not rsget.EOF THEN
 		fnGetRecommendList = rsget.getRows()
 	END IF
 	rsget.Close	
	END IF	
  End Function
End Class



'################################################
' 디자인 핑거스 공통코드
'################################################
Class CDesignFingersCode
	public FDFCodeSeq
	public FPCodeSeq
	public FCodeDesc
	public FCodeSort
	public FIsUsing
	
'// 코드값 가져오기
  public Function fnGetCommCode(ByVal iPCodeSeq)
 	Dim strSql, strSearch
 	IF iPCodeSeq <> "" THEN 	strSearch =  " AND PCodeSeq = "&iPCodeSeq
 	strSql = " SELECT DFCodeSeq, CodeDesc, PCodeSeq, CodeSort, IsUsing FROM  [db_sitemaster].[dbo].tbl_designfingers_code WHERE 1=1 "&strSearch&" order by PCodeSeq, CodeSort" 	 	
 	rsget.Open strSql,dbget	
 	IF Not rsget.EOF THEN
 		fnGetCommCode = rsget.getRows()
 	END IF
 	rsget.Close	
 End Function
 
 public Function fnGetCodeCont(ByVal iCodeSeq)
 	Dim strSql
 	strSql ="SELECT DFCodeSeq, CodeDesc, PCodeSeq, CodeSort, IsUsing FROM [db_sitemaster].[dbo].tbl_designfingers_code  WHERE DFCodeSeq = "&iCodeSeq 	
 	rsget.Open strSql,dbget	
 	IF Not rsget.EOF THEN
 		FDFCodeSeq  = rsget(0)
 		FCodeDesc   = rsget(1)
		FPCodeSeq   = rsget(2)		
		FCodeSort   = rsget(3)
		FIsUsing   	= rsget(4)
 	END IF
 	rsget.Close	
 End Function
End Class

 '// 코드값 셀렉트 박스
  Sub sbOptCommCode(ByVal iPCodeSeq, ByVal selValue)
  	Dim arrList,intLoop
  	Dim CDFCode
  	Set CDFCode = new CDesignFingersCode  	
    arrList = CDFCode.fnGetCommCode(iPCodeSeq)
    Set CDFCode = nothing
    
    If isNull(selValue) Then
    	selValue = ""
    End If
    
    IF isArray(arrList) THEN
    	For intLoop =0 To UBound(arrList,2)
    	%>
    	<option Value="<%=arrList(0,intLoop)%>" <%IF Cstr(arrList(0,intLoop)) = CStr(selValue) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
    	<%
    	Next
    END IF
  End Sub
 
 
'// 특정종류의 코드값의 배열에서 특정값의 코드명 가져오기
	Function fnGetCodeArrDesc(ByVal arrCode, ByVal iCodeValue)
		Dim intLoop		
		IF iCodeValue = "" or isNull(iCodeValue) THEN iCodeValue = -1
		For intLoop =0 To UBound(arrCode,2)		
			IF Cint(iCodeValue) = arrCode(0,intLoop) THEN				
				fnGetCodeArrDesc = arrCode(1,intLoop)
				Exit For
			END IF	
		Next	
	End Function

'-----------------------------------------------------------------------  
' sbAlertMsg : 알림문구 후 페이지 이동 처리
'----------------------------------------------------------------------- 	
	Sub sbAlertMsg(byVal strMsg, ByVal strUrl, ByVal strTarget)
		Dim strLink
		IF strUrl = "close" THEN	'팝업 창 닫을경우
			strLink = strTarget & ".close();"
		ELSEIF strUrl ="back" THEN	'이전 페이지로 이동
			strLink = "history.back(-1);"		
		ELSE
			strLink = strTarget & ".location.href='" & strUrl & "';"
		END IF		
		
		IF strTarget = "opener" THEN strLink = strLink &"self.close();"
%>
	<script language="javascript">
	<!--
		alert("<%=strMsg%>");
		<%=strLink%>;
	//-->
	</script>
<%		dbget.close()	:	response.End
	End Sub
%>

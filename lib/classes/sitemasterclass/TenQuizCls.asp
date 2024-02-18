<%
Class TenQuizObj
'퀴즈
	public Fidx					'시퀀스값
	Public Fchasu				'퀴즈별 차수값(해당년도 월일)
	Public FMonthGroup			'월단위로 묶어서 표시됨
	Public FTopTitle			'상단 타이틀 이미지 경로값
	Public FQuizDescription		'텐퀴즈 설명
	Public FBackGroundImage		'텐퀴즈 배경 이미지
	Public FMWBackGroundImage	'텐퀴즈 모웹 배경 이미지
	Public FPCWBackGroundImage	'텐퀴즈 피씨웹 배경 이미지
	public FQuestionHintNumber	'힌트 문항번호
	public FTotalMileage		'총 지급 마일리지 금액
	public FQuizStartDate		'시작일		
	public FQuizEndDate			'종료일
	public FTotalQuestionCount	'총 문항 수
	public FStartDescription	'하단 대기중 도전하기 밑에 나오는 설명
	public FProductEvtNum
	public FAdminRegister		'등록한 스태프 아이디
	public FAdminName			'등록한 스태프 이름
	public FAdminModifyer		'수정한 스태프 아이디
	public FAdminModifyerName	'수정한 스태프 이름
	public FRegistDate			'등록일
    public FLastUpDate			'수정일
	public FQuizStatus			'퀴즈 상태		1: 등록 대기 	2. 오픈 	3. 종료
	public FEndAlertTxt			' 종료 시 버튼 누르면 나오는 텍스트
	public FWaitingAlertTxt		' 대기 시 버튼 누르면 나오는 텍스트
	public FMileageDiv			'마일리지 배분 상태		1: 분배 완료 	0: 분배 전

'문항
	public FIidx                      '시퀀스값
	public FIchasu                    '퀴즈별 차수값
	public FItype                     '문제 타입 (1,2,3,4)
	public FIquestionNumber           '문제번호
	public FIquestion                 '문항
	public FIquestionType1Image1      'type1번 이미지 1 경로값
	public FIquestionType1Image2      'type1번 이미지 2 경로값
	public FIquestionType1Image3      'type1번 이미지 3 경로값
	public FIquestionType1Image4      'type1번 이미지 4 경로값
	public FIquestionExample1         'type1번일 경우 텍스트 type2번일 경우 이미지 경로값
	public FIquestionExample2         'type1번일 경우 텍스트 type2번일 경우 이미지 경로값
	public FIquestionExample3         'type1번일 경우 텍스트 type2번일 경우 이미지 경로값
	public FIquestionExample4         'type1번일 경우 텍스트 type2번일 경우 이미지 경로값
	public FItype2TextExample1         'type2번 텍스트
	public FItype2TextExample2         'type2번 텍스트
	public FItype2TextExample3         'type2번 텍스트
	public FItype2TextExample4         'type2번 텍스트
	
	public FIanswer                   '문항의 답안
	public FIregistDate               '등록일
	public FIlastUpDate               '수정일
	public FIIsUsing				  '사용 유무
	public FINumOfType1Image		  'typ1 1이미지 갯수
 
'유저 답안 데이터
	public FAquestionNumber				'문항번호
	public FAanswer						'문제 답
	public FAuserAnswer					'유저 답
	public FAresult						'결과		
	public FAuserscore					'유저 점수
	public FAtotalquestioncount			'토탈 문제개수

'기타
	public FUchasu
	public FUuserLevel
	public FUuserId
	public FUage
	public FUuserScore
	public FUbuyDate
	public FUquizCnt
	public FUrecentMileageLog

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class TenQuiz
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
       
    
	Public FmonthGroupOption
	Public FQuizStatusOption
	Public FQuizUserOption
	Public FsdtOption
	Public FedtOption
	public FChasuOption
	public FUserIdOption
	public FWriterOption
	
	public FRectIdx
	Public FRectSubIdx
	Public FRectChasu
	public FRectUserId
	
    public Sub GetContentsList()
        dim sqlStr, i, sqlWhere

		sqlwhere = ""
		if FmonthGroupOption <> "" then
			sqlWhere = sqlWhere + " and monthgroup='" & FmonthGroupOption & "'"
		end if 

		if FQuizStatusOption <> "" then
			sqlWhere = sqlWhere + " and quizStatus='" & FQuizStatusOption & "'"
		end if 

		if FsdtOption <> "" then
			sqlWhere = sqlWhere +  " and quizStartDate >='" & FsdtOption & "'"
		end if 

		if FedtOption <> "" then
			sqlWhere = sqlWhere +  " and quizStartDate <='" & FedtOption & "'"
		end if 		

		if FChasuOption <> "" then
			sqlWhere = sqlWhere +  " and chasu like '%" & FChasuOption & "%'"
		end if 				

		if FWriterOption <> "" then
			sqlWhere = sqlWhere +  " and adminName like '%" & FWriterOption & "%'"
		end if 				

		sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].[tbl_PlayingTenQuizData] "
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere
        

'        if Fisusing<>"" then
'            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
'        end If

'		if Fsdt<>"" then sqlStr = sqlStr & " and StartDate >='" & Fsdt & " 00:00:00' and  EndDate <='" & Fsdt & " 23:59:59' "
		'if Fedt<>"" then sqlStr = sqlStr & " and  EndDate <='" & Fedt & " 23:59:59' "

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
			
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + "  idx "
		sqlStr = sqlStr + " , chasu "
		sqlStr = sqlStr + " , monthGroup "
		sqlStr = sqlStr + " , topTitle "
		sqlStr = sqlStr + " , quizDescription "
		sqlStr = sqlStr + " , backGroundImage "
		sqlStr = sqlStr + " , questionHintNumber "
		sqlStr = sqlStr + " , totalMileage "
		sqlStr = sqlStr + " , quizStartDate "
		sqlStr = sqlStr + " , quizEndDate "
		sqlStr = sqlStr + " , totalQuestionCount "
		sqlStr = sqlStr + " , startDescription "
		sqlStr = sqlStr + " , productEvtNum "
		sqlStr = sqlStr + " , adminRegister "
		sqlStr = sqlStr + " , adminName "
		sqlStr = sqlStr + " , adminModifyer "
		sqlStr = sqlStr + " , adminModifyerName "
		sqlStr = sqlStr + " , registDate "
		sqlStr = sqlStr + " , modifyDate "
		sqlStr = sqlStr + " , quizStatus "
		sqlStr = sqlStr + " , mileageDiv "
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_PlayingTenQuizData "
        sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere
        
		sqlStr = sqlStr + " order by  chasu desc" 

'		response.write sqlStr &"<br>"
		
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new TenQuizObj
				
				FItemList(i).Fidx				= rsget("idx")
				FItemList(i).Fchasu				= rsget("chasu")
				FItemList(i).FMonthGroup		= rsget("monthGroup")
				FItemList(i).FTopTitle			= rsget("topTitle")
				FItemList(i).FQuizDescription	= rsget("quizDescription")
				FItemList(i).FBackGroundImage	= rsget("backGroundImage")
				FItemList(i).FQuestionHintNumber= rsget("questionHintNumber")
				FItemList(i).FTotalMileage		= rsget("totalMileage")
				FItemList(i).FQuizStartDate		= rsget("quizStartDate")
				FItemList(i).FQuizEndDate		= rsget("quizEndDate")
				FItemList(i).FTotalQuestionCount= rsget("totalQuestionCount")
				FItemList(i).FProductEvtNum		= rsget("productEvtNum")				
				FItemList(i).FStartDescription	= rsget("startDescription")
				FItemList(i).FProductEvtNum		= rsget("productEvtNum")
				FItemList(i).FAdminRegister		= rsget("adminRegister")
				FItemList(i).FAdminName			= rsget("adminName")
				FItemList(i).FAdminModifyer		= rsget("adminModifyer")
				FItemList(i).FAdminModifyerName	= rsget("adminModifyerName")
				FItemList(i).FRegistDate		= rsget("registDate")
				FItemList(i).FLastUpDate		= rsget("modifyDate")
				FItemList(i).FQuizStatus		= rsget("quizStatus")		
				FItemList(i).FMileageDiv		= rsget("mileageDiv")									

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetUserInfoList()
        dim sqlStr, i, sqlWhere

		sqlwhere = ""
		if FmonthGroupOption <> "" then
			sqlWhere = sqlWhere + " and b.monthgroup='" & FmonthGroupOption & "'"
		end if 

		if FChasuOption <> "" then
			sqlWhere = sqlWhere +  " and a.chasu like '%" & FChasuOption & "%'"
		end if 				

		if FUserIdOption <> "" then
			sqlWhere = sqlWhere +  " and a.userid like '%" & FUserIdOption & "%'"
		end if 				

		sqlStr = " SELECT count(*) as cnt"
		sqlStr = sqlStr + "  FROM db_sitemaster.dbo.tbl_PlayingTenQuizUserMasterData as a "
		sqlStr = sqlStr + " INNER JOIN db_sitemaster.dbo.tbl_PlayingTenQuizData as b with (nolock) on A.chasu = b.chasu "
		sqlStr = sqlStr + " INNER JOIN db_user.dbo.tbl_logindata as l with (nolock) on A.userid = l.userid "
		sqlStr = sqlStr + " inner join db_user.dbo.tbl_user_n as n with (nolock)on A.userid = n.userid "       
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub

		sqlStr = " select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + " 	  a.chasu										"										
		sqlStr = sqlStr + " 	  ,case when l.userlevel = 0 then 'white'										"										
		sqlStr = sqlStr + " 		 when l.userlevel = 1 then 'red'											"									
		sqlStr = sqlStr + " 		 when l.userlevel = 2 then 'vip'											"									
		sqlStr = sqlStr + " 		 when l.userlevel = 3 then 'vip gold'										"										
		sqlStr = sqlStr + " 		 when l.userlevel = 4 then 'vvip'											"										
		sqlStr = sqlStr + " 		 when l.userlevel = 7 then 'staff'			 	 							"													
		sqlStr = sqlStr + " 		 when l.userlevel = 8 then 'FAMILY'			 	 							"
		sqlStr = sqlStr + " 		 when l.userlevel = 9 then 'BIZ'			 	 							"
		sqlStr = sqlStr + " 	    end as userlevel															"					
		sqlStr = sqlStr + " 	 , A.USERID																		"		
		sqlStr = sqlStr + " 	 , Case																			"	
		sqlStr = sqlStr + " 		 When left(n.juminno,1)='0' Then 2018-Cast('20' + left(n.juminno,2) as int)	"																			
		sqlStr = sqlStr + " 		 Else 2018-Cast('19' + left(n.juminno,2) as int)							"														
		sqlStr = sqlStr + " 	   end as age 																	"			
		sqlStr = sqlStr + " 	 , a.userscore																	"			
		sqlStr = sqlStr + "     , (select top 1 regdate from db_order.dbo.tbl_order_master as m with(nolock) where m.userid =n.userid and m.userid <> '' and m.ipkumdiv>3  AND m.jumundiv<>9 AND m.cancelyn='N' and m.sitename = '10x10' and m.userid <> '' order by regdate desc) buydate "																				
		sqlStr = sqlStr + " 	 , (																			"	
		sqlStr = sqlStr + " 		select count(1)																"				
		sqlStr = sqlStr + " 		  from db_sitemaster.dbo.tbl_PlayingTenQuizUserMasterData with(nolock) 					"															
		sqlStr = sqlStr + " 		 where userid = a.userid		   											"									
		sqlStr = sqlStr + " 	   ) quizcnt																	"			
		sqlStr = sqlStr + " 	, (																				"
		sqlStr = sqlStr + " 			select top 1 regdate													"							
		sqlStr = sqlStr + " 			  from db_user.dbo.tbl_mileagelog with(nolock)										"										
		sqlStr = sqlStr + " 			 where userid = a.userid												"										
		sqlStr = sqlStr + " 			   and jukyo = '상품구매'												 "								
		sqlStr = sqlStr + " 			   and deleteyn = 'N'													"							
		sqlStr = sqlStr + " 			  order by id desc														"						
		sqlStr = sqlStr + " 	  ) recentmileagelog															"							
		sqlStr = sqlStr + "  FROM db_sitemaster.dbo.tbl_PlayingTenQuizUserMasterData as a						"														
		sqlStr = sqlStr + " INNER JOIN db_sitemaster.dbo.tbl_PlayingTenQuizData as b with (nolock) on A.chasu = b.chasu "		
		sqlStr = sqlStr + " INNER JOIN db_user.dbo.tbl_logindata as l with (nolock) on A.userid = l.userid		"																		
		sqlStr = sqlStr + " inner join db_user.dbo.tbl_user_n as n with (nolock)on A.userid = n.userid			"																	
		sqlStr = sqlStr + " where 1=1 "
		sqlStr = sqlStr + sqlWhere 
		sqlStr = sqlStr + " order by a.chasu desc" 

'		response.write sqlStr &"<br>"
		
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new TenQuizObj
				
				FItemList(i).FUchasu				= rsget("chasu")
				FItemList(i).FUuserLevel			= rsget("userlevel")
				FItemList(i).FUuserId				= rsget("userid")
				FItemList(i).FUage					= rsget("age")
				FItemList(i).FUuserScore			= rsget("userscore")
				FItemList(i).FUbuyDate				= rsget("buydate")			
				FItemList(i).FUquizCnt				= rsget("quizcnt")				
				FItemList(i).FUrecentMileageLog		= rsget("recentmileagelog")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub	
    	
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = " Select idx, chasu, type, questionNumber, question, questionType1Image1, questionType1Image2, questionType1Image3, questionType1Image4, type2TextExample1, type2TextExample2, type2TextExample3, type2TextExample4 "
		sqlStr = sqlStr & " , questionExample1, questionExample2, questionExample3, questionExample4, answer, registDate, lastUpDate, isusing, numOfType1Image "
        sqlStr = sqlStr & " From db_sitemaster.dbo.tbl_PlayingTenQuizQuestionData "
        SqlStr = SqlStr & " where idx=" + CStr(FRectSubIdx)

'		response.write sqlStr &"<br>"
'		response.end

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new TenQuizObj

        if Not rsget.Eof then
            FOneItem.FIidx				   = rsget("idx")	
            FOneItem.FIchasu			   = rsget("chasu")	
            FOneItem.FItype				   = rsget("type")
			FOneItem.FIquestionNumber	   = rsget("questionNumber")			
            FOneItem.FIquestion			   = rsget("question")	
            FOneItem.FIquestionType1Image1 = rsget("questionType1Image1")				
            FOneItem.FIquestionType1Image2 = rsget("questionType1Image2")				
			FOneItem.FIquestionType1Image3 = rsget("questionType1Image3")				
			FOneItem.FIquestionType1Image4 = rsget("questionType1Image4")				
			FOneItem.FIquestionExample1	   = rsget("questionExample1")			
			FOneItem.FIquestionExample2	   = rsget("questionExample2")			
			FOneItem.FIquestionExample3	   = rsget("questionExample3")			
			FOneItem.FIquestionExample4	   = rsget("questionExample4")			
			FOneItem.FItype2TextExample1   = rsget("type2TextExample1")			
			FOneItem.FItype2TextExample2   = rsget("type2TextExample2")			
			FOneItem.FItype2TextExample3   = rsget("type2TextExample3")			
			FOneItem.FItype2TextExample4   = rsget("type2TextExample4")						
			FOneItem.FIanswer			   = rsget("answer")	
			FOneItem.FIregistDate		   = rsget("registDate")		
			FOneItem.FIlastUpDate		   = rsget("lastUpDate")
			FOneItem.FIIsUsing			   = rsget("isusing")
			FOneItem.FINumOfType1Image	   = rsget("numOfType1Image")
			
        end if
        rsget.close
	End Sub

    public Sub GetOneContent()
        dim sqlStr
        sqlStr = "select top 1 * "
        sqlStr = sqlStr + " from db_sitemaster.dbo.[tbl_PlayingTenQuizData] "
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new TenQuizObj
        
        if Not rsget.Eof Then	
			FOneItem.Fidx				= rsget("idx")
			FOneItem.Fchasu				= rsget("chasu")
			FOneItem.FMonthGroup		= rsget("monthGroup")
			FOneItem.FTopTitle			= rsget("topTitle")
			FOneItem.FQuizDescription	= rsget("quizDescription")
			FOneItem.FBackGroundImage	= rsget("backGroundImage")
			FOneItem.FMWBackGroundImage	= rsget("MWbackGroundImage")
			FOneItem.FPCWBackGroundImage= rsget("PCWbackGroundImage")
			FOneItem.FQuestionHintNumber= rsget("questionHintNumber")
			FOneItem.FTotalMileage		= rsget("totalMileage")
			FOneItem.FQuizStartDate		= rsget("quizStartDate")
			FOneItem.FQuizEndDate		= rsget("quizEndDate")
			FOneItem.FTotalQuestionCount= rsget("totalQuestionCount")
			FOneItem.FStartDescription	= rsget("startDescription")
			FOneItem.FProductEvtNum		= rsget("productEvtNum")
			FOneItem.FAdminRegister		= rsget("adminRegister")
			FOneItem.FAdminName			= rsget("adminName")
			FOneItem.FAdminModifyer		= rsget("adminModifyer")
			FOneItem.FAdminModifyerName	= rsget("adminModifyerName")
			FOneItem.FRegistDate		= rsget("registDate")
			FOneItem.FLastUpDate		= rsget("modifyDate")		
			FOneItem.FQuizStatus		= rsget("quizStatus")					
			FOneItem.FEndAlertTxt		= rsget("endAlertTxt")					
			FOneItem.FWaitingAlertTxt	= rsget("waitingAlertTxt")					
			
        end If
        
        rsget.Close
    end Sub
    
    public Sub GetContentsItemList()
       dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_PlayingTenQuizquestiondata "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + " and isusing='Y'"

		if FRectChasu <> "" then
		sqlStr = sqlStr + " and chasu='"& FRectChasu &"'"
		end if
		
		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " "        
		sqlStr = sqlStr & "	idx "
		sqlStr = sqlStr & "	,chasu "
		sqlStr = sqlStr & "	,type "
		sqlStr = sqlStr & "	,questionNumber "
		sqlStr = sqlStr & "	,question "
		sqlStr = sqlStr & "	,questionType1Image1 "
		sqlStr = sqlStr & "	,questionType1Image2 "
		sqlStr = sqlStr & "	,questionType1Image3 "
		sqlStr = sqlStr & "	,questionType1Image4 "
		sqlStr = sqlStr & "	,questionExample1 "
		sqlStr = sqlStr & "	,questionExample2 "
		sqlStr = sqlStr & "	,questionExample3 "
		sqlStr = sqlStr & "	,questionExample4 "
		sqlStr = sqlStr & "	,type2TextExample1 "
		sqlStr = sqlStr & "	,type2TextExample2 "
		sqlStr = sqlStr & "	,type2TextExample3 "
		sqlStr = sqlStr & "	,type2TextExample4 "		
		sqlStr = sqlStr & "	,answer "
		sqlStr = sqlStr & "	,registDate "
		sqlStr = sqlStr & "	,lastUpDate "
		sqlStr = sqlStr & " From [db_sitemaster].[dbo].[tbl_PlayingTenQuizquestiondata] "

        sqlStr = sqlStr & "Where 1=1"
		sqlStr = sqlStr & "and isusing='Y'"

		if FRectChasu <> "" then
		sqlStr = sqlStr + " and chasu='"& FRectChasu &"'"
		end if        

		sqlStr = sqlStr + " order by questionNumber asc" 

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new TenQuizObj
				
				FItemList(i).FIidx					= rsget("idx")
	            FItemList(i).FIchasu				= rsget("chasu")
	            FItemList(i).FItype					= rsget("type")
	            FItemList(i).FIquestionNumber		= rsget("questionNumber")
	            FItemList(i).FIquestion				= rsget("question")
	            FItemList(i).FIquestionType1Image1	= rsget("questionType1Image1")
				FItemList(i).FIquestionType1Image2	= rsget("questionType1Image2")
				FItemList(i).FIquestionType1Image3	= rsget("questionType1Image3")
				FItemList(i).FIquestionType1Image4	= rsget("questionType1Image4")
				FItemList(i).FIquestionExample1		= rsget("questionExample1")
				FItemList(i).FIquestionExample2		= rsget("questionExample2")
				FItemList(i).FIquestionExample3		= rsget("questionExample3")
				FItemList(i).FIquestionExample4		= rsget("questionExample4")
				FItemList(i).FItype2TextExample1    = rsget("type2TextExample1")			
				FItemList(i).FItype2TextExample2    = rsget("type2TextExample2")			
				FItemList(i).FItype2TextExample3    = rsget("type2TextExample3")			
				FItemList(i).FItype2TextExample4    = rsget("type2TextExample4")										
				FItemList(i).FIanswer				= rsget("answer")
				FItemList(i).FIregistDate			= rsget("registDate")
				FItemList(i).FIlastUpDate			= rsget("lastUpDate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetUserAnswerList()
       dim sqlStr, i

		sqlStr = " select count(1) as cnt  "
		sqlStr = sqlStr + "  from db_sitemaster.dbo.tbl_PlayingTenQuizQuestionData as a 	"
		sqlStr = sqlStr + "  left join db_sitemaster.dbo.tbl_PlayingTenQuizUserDetailData b 	"
		sqlStr = sqlStr + "    on a.chasu =b.chasu 	"
		sqlStr = sqlStr + "   and a.questionNumber = b.questionNumber 	"
		sqlStr = sqlStr + "   and b.userid = '"& FRectUserId &"'	"
		sqlStr = sqlStr + "  where a.chasu = '"& FRectChasu &"'	"	

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub

        sqlStr = "Select a.questionNumber "        		
		sqlStr = sqlStr + "	 , a.answer	 	"
		sqlStr = sqlStr + "	 , isnull(b.userAnswer, 0) as userAnswer	"
		sqlStr = sqlStr + "	 , case 	"
		sqlStr = sqlStr + "	 when b.userAnswer = a.answer then 'true'	"		
		sqlStr = sqlStr + "	 else 'false'	 	"
		sqlStr = sqlStr + "	 end as 'result'	"
		sqlStr = sqlStr + "	 	 , ("
		sqlStr = sqlStr + "	 		select top 1 userscore"
		sqlStr = sqlStr + "	 		  from db_sitemaster.dbo.tbl_PlayingTenQuizUserMasterData"
		sqlStr = sqlStr + "	 		 where chasu = a.chasu"
		sqlStr = sqlStr + "	 		   and userid = b.userid		"
		sqlStr = sqlStr + "	 	 ) as userscore"
		sqlStr = sqlStr + "	 	 , ("
		sqlStr = sqlStr + "	 		select top 1 totalquestioncount"
		sqlStr = sqlStr + "	 		  from db_sitemaster.dbo.tbl_PlayingTenQuizData"
		sqlStr = sqlStr + "	 		 where chasu = a.chasu"
		sqlStr = sqlStr + "	 	 ) as totalquestioncount		"
		sqlStr = sqlStr + "  from db_sitemaster.dbo.tbl_PlayingTenQuizQuestionData as a 	"
		sqlStr = sqlStr + "  left join db_sitemaster.dbo.tbl_PlayingTenQuizUserDetailData b 	"
		sqlStr = sqlStr + "    on a.chasu =b.chasu 	"
		sqlStr = sqlStr + "   and a.questionNumber = b.questionNumber 	"
		sqlStr = sqlStr + "   and b.userid = '"& FRectUserId &"'	"
		sqlStr = sqlStr + "  where a.chasu = '"& FRectChasu &"'	"		
		sqlStr = sqlStr + " order by questionNumber asc" 

		'response.write sqlStr &"<br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		redim preserve FItemList(FTotalCount)
		if  not rsget.EOF  then
		    i = 0			
			do until rsget.eof
				set FItemList(i) = new TenQuizObj
				
				FItemList(i).FAquestionNumber		= rsget("questionNumber")
	            FItemList(i).FAanswer				= rsget("answer")
	            FItemList(i).FAuserAnswer			= rsget("userAnswer")
	            FItemList(i).FAresult				= rsget("result")	            	            		
	            FItemList(i).FAuserscore			= rsget("userscore")	            	            		
	            FItemList(i).FAtotalquestioncount   = rsget("totalquestioncount")	            	            										

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public Function getQuizCorrectPercent(vChasu)
        dim strSql, i
		
		if vChasu="" or isNull(vChasu) then
			exit Function
		end if	

		strSql = " select A.questionNumber 													"													
		strSql = strSql + " 	 , a.answer													"						 
		strSql = strSql + " 	 , CONVERT(decimal,											"		
		strSql = strSql + " 	   CONVERT(FLOAT, 											"		
		strSql = strSql + " 	   SUM(CASE WHEN A.answer = B.userAnswer					"								
		strSql = strSql + " 			THEN 1												"	
		strSql = strSql + " 			ELSE 0												"	
		strSql = strSql + " 	       END)													"
		strSql = strSql + " 		   ) / CONVERT(float, COUNT(*)) * 100					"								
		strSql = strSql + " 	   ) AS RESULT												"	
		strSql = strSql + " from db_sitemaster.dbo.tbl_PlayingTenQuizQuestionData as a 		"												
		strSql = strSql + " left join db_sitemaster.dbo.tbl_PlayingTenQuizUserDetailData b 	"													
		strSql = strSql + "   on a.chasu =b.chasu 											"			
		strSql = strSql + "  and a.questionNumber = b.questionNumber 						"								
		strSql = strSql + " WHERE A.CHASU = '"& vChasu &"'									"					
		strSql = strSql + " GROUP BY A.questionNumber, a.answer								"			
		strSql = strSql + " ORDER BY A.questionNumber										"			

		rsget.CursorLocation = adUseClient
        rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

 		if not rsget.EOF then
		    getQuizCorrectPercent = rsget.getRows()	
		end if
        
        rsget.Close
    end function

	public Function GetNumberOfWinner(chasu, totalscore)
        dim sqlStr

		if chasu="" or isNull(chasu) or totalscore="" or isNull(totalscore) then
			exit Function
		end if
		
        sqlStr = "SELECT count(userid) cnt "
        sqlStr = sqlStr + " FROM [db_sitemaster].[dbo].[tbl_PlayingTenQuizUserMasterData] WITH (NOLOCK) "
        sqlStr = sqlStr + " where chasu='" + CStr(chasu) + "'"		
		sqlStr = sqlStr + " and userscore=" + CStr(totalscore)

		rsget.Open sqlStr, dbget, 1
		if Not(rsget.EOF or rsget.BOF) then
			GetNumberOfWinner = rsget("cnt")
		end if
        
        rsget.Close
    end function

	public Function GetNumberOfParticipants(chasu)
        dim sqlStr

		if chasu="" or isNull(chasu) then
			exit Function
		end if
		
        sqlStr = "SELECT count(userid) cnt"
        sqlStr = sqlStr + " FROM [db_sitemaster].[dbo].[tbl_PlayingTenQuizUserMasterData] WITH (NOLOCK) "
        sqlStr = sqlStr + " where chasu='" + CStr(chasu) + "'"					
	
		rsget.Open sqlStr, dbget, 1
		if Not(rsget.EOF or rsget.BOF) then
			GetNumberOfParticipants = rsget("cnt")
		end if
        
        rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

    End Sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>
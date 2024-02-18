<%
'####################################################
' Description :  이벤트
' History : 2010.09.24 한용민 수정
'####################################################

class CEventItem
	public fissale
	public fisgift
	public fiscoupon
	public FevtId
	public FevtDivCd
	public FevtTitle
	public FevtCont
	public FevtType
	public FlistImage
	public FcontImage
	public FcontImage2
	public FevtSdate
	public FevtEdate
	public FprizeDate
	public FprtCnt
	public FlecturerID
	public Fregdate
	public FisComment
	public fgift_count
	public fsale_count
	public feventitemcount
	public FELinkType
	public FELinkURL

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEvent
	public FEventList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectevtId
	public FRectevtDivCd
	public FRectTopCnt
	public FRectsearchKey
	public FRectsearchString

	Private Sub Class_Initialize()
		redim preserve FEventList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'// 공지 목록 출력 '/academy/event/Event_list.asp
	public Sub GetNoitceList()
		dim SQL, AddSQL, lp

		'검색 추가 쿼리
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if
		if FRectevtDivCd<>"" then
			AddSQL = AddSQL & " and evtDivCd='" & FRectevtDivCd & "' "
		end if

		'@ 총데이터수
		SQL =	" Select count(evtId) as cnt " &_
				" From db_academy.dbo.tbl_eventInfo " &_
				" Where isusing = 'Y' " & AddSQL
		
		'response.write SQL &"<br>"
		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ 데이터
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				" evtId, evtDivCd, evtTitle, evtCont, listImage, evtSdate, evtEdate, prizeDate, t1.regdate " &_
				" ,issale,iscoupon ,isgift " &_
				" 		,(Select count(*) " &_
				" 		From (Select userid " &_
				" 				From db_academy.dbo.tbl_eventSub as t2 " &_
				" 				Where t2.evtId=t1.evtId and t2.useYN = 'Y' " &_
				" 				Group by userid ) as tmp " &_
				" 		)  as prtCnt " &_
				" ,(SELECT COUNT(gift_code) FROM [db_academy].[dbo].[tbl_gift] WHERE evt_code = t1.evtId and gift_using ='y') as gift_count "&_
				" ,(SELECT COUNT(sale_code) FROM [db_academy].[dbo].[tbl_sale] WHERE evt_code = t1.evtId and sale_using =1) as sale_count " &_
				" ,(SELECT COUNT(itemid) FROM db_academy.dbo.tbl_eventitem WHERE evt_code = t1.evtid) as eventitemcount " &_				
				" From db_academy.dbo.tbl_eventInfo as t1 " &_
				" Where t1.isusing = 'Y' " & AddSQL &_
				" Order by evtId desc "
		
		'response.write SQL &"<br>"
		rsACADEMYget.pagesize = FPageSize
		
		'response.write sql &"<Br>"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FEventList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FEventList(lp) = new CEventItem
				
				FEventList(lp).feventitemcount		= rsACADEMYget("eventitemcount")
				FEventList(lp).fsale_count		= rsACADEMYget("sale_count")
				FEventList(lp).fgift_count		= rsACADEMYget("gift_count")
				FEventList(lp).FevtId		= rsACADEMYget("evtId")
				FEventList(lp).FevtDivCd	= rsACADEMYget("evtDivCd")
				FEventList(lp).FevtTitle	= rsACADEMYget("evtTitle")
				FEventList(lp).FlistImage	= rsACADEMYget("listImage")
				FEventList(lp).FevtSdate	= rsACADEMYget("evtSdate")
				FEventList(lp).FevtEdate	= rsACADEMYget("evtEdate")
				FEventList(lp).FprizeDate	= rsACADEMYget("prizeDate")
				FEventList(lp).FprtCnt		= rsACADEMYget("prtCnt")
				FEventList(lp).Fregdate		= rsACADEMYget("regdate")
				FEventList(lp).fissale	= rsACADEMYget("issale")
				FEventList(lp).fiscoupon	= rsACADEMYget("iscoupon")
				FEventList(lp).fisgift		= rsACADEMYget("isgift")	
			
				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	'// 공지 내용 보기 '/academy/event/Event_modi.asp
	public Sub GetNoitceRead()
		dim SQL

		SQL =	" Select" &_
				" evtId, evtDivCd, evtTitle, evtType, listImage, contImage, contImage2, evtCont, evtSdate, evtEdate, prizeDate " &_
				" ,issale,iscoupon ,isgift, lecturerID, isComment, evt_LinkType ,evt_bannerlink, t1.regdate " &_
				" From db_academy.dbo.tbl_eventInfo as t1 " &_
				" Where t1.isusing = 'Y' " &_
				"	and evtId = " & FRectevtId
		
		'response.write SQL &"<br>"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FEventList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FEventList(0) = new CEventItem

			FEventList(0).FELinkType	= rsACADEMYget("evt_LinkType")
			FEventList(0).FELinkURL	= rsACADEMYget("evt_bannerlink")
			FEventList(0).FevtId		= rsACADEMYget("evtId")
			FEventList(0).FevtDivCd		= rsACADEMYget("evtDivCd")
			FEventList(0).FevtType		= rsACADEMYget("evtType")
			FEventList(0).FevtTitle		= rsACADEMYget("evtTitle")
			FEventList(0).FlistImage	= rsACADEMYget("listImage")
			FEventList(0).FcontImage	= rsACADEMYget("contImage")
			FEventList(0).FcontImage2	= rsACADEMYget("contImage2")
			FEventList(0).FevtCont		= rsACADEMYget("evtCont")
			FEventList(0).FevtSdate		= rsACADEMYget("evtSdate")
			FEventList(0).FevtEdate		= rsACADEMYget("evtEdate")
			FEventList(0).FprizeDate	= rsACADEMYget("prizeDate")
			FEventList(0).FlecturerID	= rsACADEMYget("lecturerID")
			FEventList(0).FisComment	= rsACADEMYget("isComment")
			FEventList(0).Fregdate		= rsACADEMYget("regdate")
			FEventList(0).fissale	= rsACADEMYget("issale")
			FEventList(0).fiscoupon	= rsACADEMYget("iscoupon")
			FEventList(0).fisgift		= rsACADEMYget("isgift")			

		end if
		rsACADEMYget.close

	end sub


	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

'##### 참여자 레코드셋용 클래스 #####
class CPartItem
	public FprtId
	public FprtUserId
	public FprtUserNm
	public FprtUserLevel
	public FprtCont1
	public FprtCont2
	public FprtDate
	public FprtCnt
	public FsixMonthOrder
	public FregDate
	public FprizeCnt

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

'##### 참여자 클래스 #####
Class CPart
	public FPartList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectevtId

	Private Sub Class_Initialize()
		redim preserve FPartList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'// 참여자 목록 출력(화면 출력용) '//academy/event/Event_view.asp
	public Sub GetPartList()
		dim SQL, lp

		SQL = "Select count(t1.prtId) as cnt " &_
				" From db_academy.dbo.tbl_eventSub as t1 " &_
				" left Join db_academy.dbo.tbl_fingers_userlevel as t4 on t1.userid=t4.userid " &_
				" Where t1.evtid = " & FRectevtId & " and t1.useYN = 'Y' "

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		SQL =	" select top " & CStr(FPageSize*FCurrPage) &_
				"		t1.prtId, t1.userid, t1.prtDate, t1.prtCont1, t1.prtCont2, " &_
				"		Case t4.userlevel " &_
				"			When '1' Then 'SEED' "&_
				"			When '2' Then 'BUD' "&_
				"			When '3' Then 'LEAF' "&_
				"			When '4' Then 'BEAN' "&_
				"			When '5' Then 'TREE' "&_
				"			When '6' Then 'STAFF' "&_
				"			Else 'SEED' "&_
				"		End	as userlevel, "&_
				"		(select count(prtId) from db_academy.dbo.tbl_eventSub where userid = t1.userid and useYN = 'Y' and evtId = " & FRectevtId & ") as prtCnt " & _
				" , (select isnull(sum(subtotalprice),0) from  db_academy.dbo.tbl_academy_order_master m, db_academy.dbo.tbl_academy_order_detail d " & _
				" where m.orderserial=d.orderserial and m.userid=t1.userid and m.cancelyn<>'Y' and m.ipkumdiv=8 and m.jumundiv<>9 and d.currstate=7 and dateadd(m,6,m.regdate) < getdate()) as sixmonthorder" & _
				" , isnull(n.regdate,getdate()) as regdate" & _
				" ,(select count(prtId) from db_academy.dbo.tbl_eventSub where userid=t1.userid and isWinner='0') as prizecnt" & _
				" From db_academy.dbo.tbl_eventSub as t1 " &_
				" left Join db_academy.dbo.tbl_fingers_userlevel as t4 on t1.userid=t4.userid " &_
				" left join [DBDATAMART].[db_user].dbo.tbl_user_n n on n.userid=t1.userid" &_
				" Where t1.useYN = 'Y' AND t1.evtId = " & FRectevtId & " " &_
				" Order by t1.prtId desc "
		
		'response.write SQL &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FPartList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FPartList(lp) = new CPartItem

				FPartList(lp).FprtId		= rsACADEMYget("prtId")
				FPartList(lp).FprtUserId	= rsACADEMYget("userId")
				FPartList(lp).FprtUserLevel	= rsACADEMYget("userlevel")			'핑거스 Level 추가	(2009.10.01;허진원)
				FPartList(lp).FprtCont1		= rsACADEMYget("prtCont1")
				FPartList(lp).FprtCont2		= rsACADEMYget("prtCont2")
				FPartList(lp).FprtDate		= rsACADEMYget("prtDate")
				FPartList(lp).FprtCnt		= rsACADEMYget("prtCnt")
				FPartList(lp).FsixMonthOrder	= rsACADEMYget("sixmonthorder")
				FPartList(lp).FregDate	= rsACADEMYget("regdate")
				FPartList(lp).FprizeCnt	= rsACADEMYget("prizecnt")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	'// 참여자 전체 목록 출력(Excel용)
	public Sub GetPartAllList()
		dim SQL, lp

		'@ 데이터
		SQL =	" select t1.prtId, t1.userid, t3.username, t1.prtDate, t1.prtCont1, t1.prtCont2, " &_
				"		Case t4.userlevel " &_
				"			When '1' Then 'SEED' "&_
				"			When '2' Then 'BUD' "&_
				"			When '3' Then 'LEAF' "&_
				"			When '4' Then 'BEAN' "&_
				"			When '5' Then 'TREE' "&_
				"			When '6' Then 'STAFF' "&_
				"			Else 'SEED' "&_
				"		End	as userlevel, "&_
				" (select count(prtId) from db_academy.dbo.tbl_eventSub where userid = t1.userid and useYN = 'Y' and evtId = " & FRectevtId & ") as prtCnt " & _
				" , (select isnull(sum(subtotalprice),0) from  db_academy.dbo.tbl_academy_order_master m, db_academy.dbo.tbl_academy_order_detail d " & _
				" where m.orderserial=d.orderserial and m.userid=t1.userid and m.cancelyn<>'Y' and m.ipkumdiv=8 and m.jumundiv<>9 and d.currstate=7 and dateadd(m,6,m.regdate) < getdate()) as sixmonthorder" & _
				" , isnull(t3.regdate,getdate()) as regdate" & _
				" ,(select count(prtId) from db_academy.dbo.tbl_eventSub where userid=t1.userid and isWinner='0') as prizecnt" & _
				" From db_academy.dbo.tbl_eventSub as t1 " &_
				" Join [TENDB].[db_user].[dbo].tbl_user_n as t3 on t1.userid=t3.userid " &_
				" left Join db_academy.dbo.tbl_fingers_userlevel as t4 on t1.userid=t4.userid " &_
				" Where t1.useYN = 'Y' AND t1.evtId = " & FRectevtId &_
				" Order by t1.prtId "

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FTotalCount = rsACADEMYget.RecordCount

		redim FPartList(FTotalCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			do until rsACADEMYget.eof
				set FPartList(lp) = new CPartItem

				FPartList(lp).FprtId		= rsACADEMYget("prtId")
				FPartList(lp).FprtUserId	= rsACADEMYget("userId")
				FPartList(lp).FprtUserNm	= rsACADEMYget("username")
				FPartList(lp).FprtUserLevel	= rsACADEMYget("userlevel")			'핑거스 Level 추가	(2009.10.01;허진원)
				FPartList(lp).FprtCont1		= rsACADEMYget("prtCont1")
				FPartList(lp).FprtCont2		= rsACADEMYget("prtCont2")
				FPartList(lp).FprtDate		= rsACADEMYget("prtDate")
				FPartList(lp).FprtCnt		= rsACADEMYget("prtCnt")
				FPartList(lp).FsixMonthOrder	= rsACADEMYget("sixmonthorder")
				FPartList(lp).FregDate	= rsACADEMYget("regdate")
				FPartList(lp).FprizeCnt	= rsACADEMYget("prizecnt")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub

	'// 당첨자 목록
	public Sub GetWinnerList()
		dim SQL, lp

		'@ 데이터
		SQL =	" select prtId, userid " &_
				" 	From db_academy.dbo.tbl_eventSub " &_
				" Where useYN = 'Y' AND evtId = " & FRectevtId & " AND isWinner = 'o' " &_
				" Order by prtId "

		rsACADEMYget.Open sql, dbACADEMYget, 1

		FTotalCount = rsACADEMYget.RecordCount

		redim FPartList(FTotalCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			do until rsACADEMYget.eof
				set FPartList(lp) = new CPartItem

				FPartList(lp).FprtId		= rsACADEMYget("prtId")
				FPartList(lp).FprtUserId	= rsACADEMYget("userId")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end Sub
	
	public FPrevID
	public FNextID

	'// 이전 페이지 검사
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// 다음 페이지 검사
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// 첫페이지 산출
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

'-------------------------------------------------------------
'ClsEventSummary : 이벤트 요약 내용 - 사은품, 할인, 쿠폰에 연계 되는 간략한 내용
'-------------------------------------------------------------
Class ClsEventSummary
	public FECode
	public FEName
	public FESDay
	public FEEDay
	public FEState
	public FBrand
	public FEOpenDate
	public FEStateDesc
	public FECloseDate
	public FEScope
	public FPartnerID

	public Function fnGetEventConts
	 Dim strSql
	 strSql = " SELECT " &_ 
	 		" (case when datediff(d,evtsdate,getdate()) < 0 then 5" &_
	 		" when datediff(d,evtSdate,getdate()) >= 0 and datediff(d,evtEdate,getdate()) <=0 then 7 " &_
	 		" when datediff(d,evtEdate,getdate()) > 0 then 9 end) as evt_state "&_
	 		" ,evtId, evtSdate, evtEdate,evtTitle "&_
	 		" FROM [db_academy].dbo.tbl_eventInfo as A "&_
	 		" WHERE A.evtId = "&FECode
	 
	 'response.write strSql &"<Br>"
	 rsACADEMYget.Open strSql,dbACADEMYget
	 
	 IF not rsACADEMYget.EOF THEN
	 	 FEName 	= db2html(rsACADEMYget("evtTitle"))
	 	 FESDay 	= rsACADEMYget("evtSdate")
	 	 FEEDay 	= rsACADEMYget("evtEdate")
	 	 FEState 	= rsACADEMYget("evt_state")
	 	 FEStateDesc= fnSetStatusDesc(FEState,FESDay,FEEDay, "")
	 	 'IF datediff("d",FEEDay,now) > 0  THEN FEState = 9	'종료일이 지난 경우 종료로 표기
	 END IF
	 
	 rsACADEMYget.close
	End Function
End Class

'-------------------------------------------------------------
'ClsEventGroup : 이벤트 그룹
'-------------------------------------------------------------
Class ClsEventGroup
	public FECode
	public FEGCode
	public FGDesc
	public FGSort
	public FGImg
	public FGPCode
	public FGDepth
	public FGPDesc
	public FGlink

	'//academy/event/pop_eventitem_group.asp
	public Function fnGetRootGroup
		Dim strSql
		strSql = " SELECT evtgroup_code, evtgroup_desc FROM [db_academy].[dbo].tbl_eventitem_group "&_
				" WHERE evt_code = "&FECode&" and evtgroup_pcode = 0 and evtgroup_using ='Y' "
		
		'response.write strSql &"<br>"
		rsACADEMYget.Open strSql,dbACADEMYget
			IF not rsACADEMYget.EOF THEN
				fnGetRootGroup = rsACADEMYget.getRows()
			End IF
			rsACADEMYget.Close
	End Function

	'## fnGetEventItemGroup :이벤트화면설정 그룹내용가져오기 ## '/academy/event/iframe_eventitem_group.
	public Function fnGetEventItemGroup
	IF FECode = "" THEN Exit Function
	Dim strSql
	
	strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
			"		(select evtgroup_desc from [db_academy].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode) "&_
			" FROM  [db_academy].[dbo].[tbl_eventitem_group] as a" &_
			"	WHERE evt_code = "&FECode&" and evtgroup_using ='Y' ORDER BY evtgroup_depth, evtgroup_sort "
	
	'response.write strSql &"<br>"
	rsACADEMYget.Open strSql,dbACADEMYget
	
		IF not rsACADEMYget.EOF THEN
			fnGetEventItemGroup = rsACADEMYget.getRows()
		End IF
		rsACADEMYget.Close
	End Function
	
	'//academy/event/pop_eventitem_group.asp
	public Function fnGetEventItemGroupCont
	Dim strSql
	IF FEGCode = "" THEN Exit Function
		
	strSql = " SELECT evtgroup_code,evtgroup_desc, evtgroup_sort, evtgroup_img,evtgroup_link,evtgroup_pcode,evtgroup_depth, "&_
			"		isnull((select evtgroup_desc from [db_academy].[dbo].[tbl_eventitem_group] where evtgroup_code = a.evtgroup_pcode),'최상위') as evtgroup_pdesc"&_
			"	FROM  [db_academy].[dbo].[tbl_eventitem_group] as a " &_
			"	WHERE evt_code = "&FECode&" and evtgroup_code="&FEGCode&" and evtgroup_using ='Y' "

	'response.write strSql &"<br>"	
	rsACADEMYget.Open strSql,dbACADEMYget
		IF not rsACADEMYget.EOF THEN
			
			FGDesc = rsACADEMYget("evtgroup_desc")
			FGSort = rsACADEMYget("evtgroup_sort")
			FGImg  = rsACADEMYget("evtgroup_img")
			FGPCode= rsACADEMYget("evtgroup_pcode")
			FGDepth= rsACADEMYget("evtgroup_depth")
			FGPDesc= rsACADEMYget("evtgroup_pdesc")
			FGlink= rsACADEMYget("evtgroup_link")
			
		End IF
		rsACADEMYget.Close
	End Function
End Class

'------------------------------------------------------
'ClsEvent : 이벤트 내용
'------------------------------------------------------
Class ClsEvent
	public FECode	'해당 이벤트코드
	public FEKind
	public FEManager
	public FEScope
	public FEPartnerID
	public FEName
	public FESDay
	public FEEDay
	public FEPDay
	public FELevel
	public FEState
	public FERegdate
	public FECategory
	public FECateMid
	public FESale
	public FEGift
	public FECoupon
	public FECommnet
	public FEBbs
	public FEItemps
	public FEApply
	public FEBImg
	public FEBImg2010
	public FEGImg
	public FETemp
	public FEMImg
	public FEHtml
	public FEISort
	public FEIAddType
	public FEDId
	public FEMId
	public FEFwd
	public FChkDisp
	public FEBrand
	public FEIcon
	public FECommentTitle
	public FELinkCode
	public FELinkType
	public FELinkURL
	public FCPage	'Set 현재 페이지
	public FPSize	'Set 페이지 사이즈
	public FTotCnt
	public FESGroup	'Set 그룹검색
	public FESSort	'Set 정렬
	public FSfDate
	public FSsDate
	public FSeDate
	public FSfEvt
	public FSeTxt
	public FScategory
	public FScateMid
	public FSstate
	public FSkind
	public FSedid
	public FSemid
	public FSisSale
	public FSisGift
	public FSisCoupon
	public FSisOnlyTen
	public FSisGetBlogURL
	public FEUsing
	public FEOpenDate
	public FECloseDate
	public FRectMakerid
	public FRectItemid
	public FRectItemName
	public FRectSellYN
	public FRectIsUsing
	public FRectDanjongyn
	public FRectLimityn
	public FRectMWDiv
	public FRectDeliveryType
	public FRectSailYn
	public FRectCouponYn
	public FRectVatYn
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FEKindDesc
	public FEStateDesc
	public FEFullYN
	public FEIteminfoYN
	public FETag
	public FPrizeYN
	public fissale
	public fiscoupon
	public fisgift
	
    public Function IsSoldOut(FSellYn,FLimitYn,FLimitNo,FLimitSold)
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa(FLimitNo,FLimitSold)<1))
	end function

    public function GetLimitEa(FLimitNo,FLimitSold)
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

	public Function IsUpcheBeasong(Fdeliverytype)
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function

	public function getMwDivName(FmwDiv)
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "특정"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	'## fnGetEventCont : 이벤트개요 내용 가져오기 ## '/academy/event/eventitem_regist.asp
	public Function fnGetEventCont
	Dim strSql ,lp
	
	IF FECode = "" THEN Exit Function
		
		strSql = "SELECT" &_ 
				" evtId, evtDivCd, evtTitle, evtCont, listImage, evtSdate, evtEdate, prizeDate" &_
				" ,issale,iscoupon ,isgift" &_
	 			" ,(case when datediff(d,evtsdate,getdate()) < 0 then '오픈대기'" &_
	 			" when datediff(d,evtSdate,getdate()) >= 0 and datediff(d,evtEdate,getdate()) <=0 then '오픈'" &_
	 			" when datediff(d,evtEdate,getdate()) > 0 then '종료' end) as evt_statedesc "&_
				" ,(select commnm from db_academy.dbo.tbl_commcd where a.evtDivCd = commcd) as evt_kinddesc" &_
				" FROM [db_academy].dbo.tbl_eventInfo a "&_
				" WHERE evtid = "&FECode
		
		'response.write strSql &"<br>"
		rsacademyget.Open strSql,dbacademyget
		
		IF not rsacademyget.EOF THEN
						
				FEKind	= rsACADEMYget("evtDivCd")
				FEKindDesc	= rsACADEMYget("evt_kinddesc")
				FEName	= rsACADEMYget("evtTitle")			
				FESDay	= rsACADEMYget("evtSdate")
				FEEDay	= rsACADEMYget("evtEdate")
				FEPDay	= rsACADEMYget("prizeDate")				
				fissale	= rsACADEMYget("issale")
				fiscoupon	= rsACADEMYget("iscoupon")
				fisgift		= rsACADEMYget("isgift")
				FEStateDesc = rsACADEMYget("evt_statedesc")
			
		End IF
		rsacademyget.Close
		
	End Function

	'## fnGetEventItem :이벤트상품 가져오기 ## '/academy/event/common/pop_eventitem_addinfo.asp
	public Function fnGetEventItem
	Dim strSql, strSqlCnt,iDelCnt
	Dim strSort,strGroup, striSort,addSql

    '// 추가 쿼리
    if (FRectMakerid <> "") then
        addSql = addSql + " and B.makerid='" + FRectMakerid + "'"
    end if

    if (FRectItemid <> "") then
        addSql = addSql + " and B.itemid in (" + FRectItemid + ")"
    end if

    if (FRectItemName <> "") then
        addSql = addSql + " and B.itemname like '%" + html2db(trim(FRectItemName)) + "%'"
    end if

    if (FRectSellYN <> "") then
        addSql = addSql + " and B.sellyn='" + FRectSellYN + "'"
    end if

    if (FRectIsUsing <> "") then
        addSql = addSql + " and B.isusing='" + FRectIsUsing + "'"
    end if

    if FRectMWDiv="MW" then
        addSql = addSql + " and (B.mwdiv='M' or B.mwdiv='W')"
    elseif FRectMWDiv<>"" then
        addSql = addSql + " and B.mwdiv='" + FRectMwDiv + "'"
    end if

	if FRectLimityn="Y0" then
        addSql = addSql + " and B.limityn='Y' and (B.limitno-B.limitsold<1)"
    elseif FRectLimityn<>"" then
        addSql = addSql + " and B.limityn='" + FRectLimityn + "'"
    end if

    if FRectCate_Large<>"" then
        addSql = addSql + " and B.cate_large='" + FRectCate_Large + "'"
    end if

    if FRectCate_Mid<>"" then
        addSql = addSql + " and B.cate_mid='" + FRectCate_Mid + "'"
    end if

    if FRectCate_Small<>"" then
        addSql = addSql + " and B.cate_small='" + FRectCate_Small + "'"
    end if

    if FRectSailYn<>"" then
        addSql = addSql + " and B.saleyn='" + FRectSailYn + "'"
    end if

    if FRectCouponYn<>"" then
        addSql = addSql + " and B.itemcouponyn='" + FRectCouponYn + "'"
    end if

    if FRectVatYn<>"" then
        addSql = addSql + " and B.vatinclude='" + FRectVatYn + "'"
    end if

    if FRectDeliveryType<>"" then
    	  addSql = addSql + " and B.deliverytype='" + FRectDeliveryType + "'"
    end if

	IF FESGroup <> "" THEN
		IF FESGroup = 0 THEN
			strGroup = " AND (evtgroup_code  is null OR evtgroup_code =0 )"
		ELSE
			strGroup = " AND evtgroup_code = " + FESGroup + ""
		END IF
	END IF

	IF FESSort = "slsell" THEN
		strSort = " ORDER BY evtitem_imgsize desc, sellcash asc "
		striSort =	" ORDER BY evtitem_imgsize desc, sellcash asc "
	ELSEIF FESSort = "shsell" THEN
		strSort = " ORDER BY  evtitem_imgsize desc, sellcash desc "
		striSort = " ORDER BY  evtitem_imgsize desc, sellcash desc "
	ELSEIF FESSort = "sbest" THEN
		strSort = " ORDER BY evtitem_imgsize desc, recentsellcount desc, sellcash desc "
		striSort = " ORDER BY evtitem_imgsize desc, recentsellcount desc, sellcash desc "
	ELSEIF FESSort = "sevtitem" THEN
		strSort = " ORDER BY evtitem_imgsize desc, evtitem_sort ,A.itemid desc"
		striSort = " ORDER BY evtitem_imgsize desc, evtitem_sort ,C.itemid desc"
	ELSEIF FESSort = "sevtgroup" THEN
		strSort = " ORDER BY evtitem_imgsize desc, evtgroup_code "
		striSort = " ORDER BY evtitem_imgsize desc, evtgroup_code "
	ELSEIF FESSort = "sbrand" THEN
		strSort = " ORDER BY evtitem_imgsize desc, makerid "
		striSort = " ORDER BY evtitem_imgsize desc, makerid "
	ELSE
		strSort = " ORDER BY evtitem_imgsize desc, A.itemid DESC "
		striSort = " ORDER BY evtitem_imgsize desc, C.itemid DESC "
	END IF


	strSqlCnt = " SELECT COUNT(A.itemid) as cnt"&_ 
				" FROM [db_academy].[dbo].[tbl_eventitem] AS A "&_
				" INNER JOIN [db_academy].dbo.tbl_diy_item AS B" &_ 
				" ON A.itemid = B.itemid "&_
				" WHERE A.evt_code = " + FECode + strGroup + addSql + " "

	'response.write strSqlCnt &"<br>"
	rsacademyget.Open strSqlCnt,dbacademyget,1
		FTotCnt = rsacademyget("cnt")
	rsacademyget.Close

	IF FTotCnt >0 THEN
		iDelCnt =  (FCPage - 1) * FPSize

		strSql = " SELECT  TOP "&FPSize*FCPage&_ 
				" A.itemid, A.evtgroup_code, A.evtitem_sort,  B.makerid, B.itemname, B.sellcash "&_
				" ,B.buycash,B.orgprice, B.orgsuplycash, B.sailprice, B.sailsuplycash, B.mileage, B.smallimage"&_
				" , B.listimage,   B.sellyn, B.deliverytype ,B.limityn, '', B.saleyn, B.isusing, B.limitno"&_
				"  , B.limitsold, B.itemcouponyn, B.itemcoupontype, B.itemcouponvalue"&_
				" ,Case itemCouponyn When 'Y' then ("&_
				" 									Select top 1 couponbuyprice"&_
				"									From [db_academy].dbo.tbl_diy_item_coupon_detail"&_
				"									Where itemcouponidx=B.curritemcouponidx and itemid=B.itemid) end "&_
				" as couponbuyprice "&_
				" , B.mwdiv, A.evtitem_imgsize	"&_
				" FROM [db_academy].[dbo].[tbl_eventitem] AS A " &_
				" INNER JOIN [db_academy].dbo.tbl_diy_item AS B"&_ 
				" ON A.itemid = B.itemid "&_
				" LEFT OUTER JOIN [db_academy].dbo.tbl_diy_item_Contents AS E"&_
				" ON A.itemid = E.itemid "&_
				" WHERE A.evt_code = "&FECode & strGroup&addSql& strSort

		'response.write strSql &"<br>"		
		rsacademyget.pagesize = FPSize
		rsacademyget.Open strSql,dbacademyget,1
        
        rsacademyget.absolutepage = FCPage
		IF not rsacademyget.EOF THEN
			fnGetEventItem = rsacademyget.getRows()
		End IF
		rsacademyget.Close

	END IF
	End Function
	
End Class	


'-------------------------------------------------------------
'ClsEventSNS : SNS
'-------------------------------------------------------------
Class ClsEventSNS
	public Fidx
	public FECode
	public Ffbtitle
	public Ffbdesc
	public Ffbimage
	public Ftwlink
	public Ftwtag1
	public Ftwtag2
	public Fkatitle
	public Fkaimage
	public Fkalink

	'//academy/event/pop_eventitem_group.asp
	public Function fnGetEventItemSNSCont
	Dim strSql
	IF FECode = "" THEN Exit Function
		
	strSql = " SELECT idx, evtcode, fbtitle, fbdesc, fbimage, twlink, twtag1, twtag2, katitle, kaimage, kalink "&_
			"	FROM  [db_academy].[dbo].[tbl_event_sns] " &_
			"	WHERE evtcode = "&FECode&" "

	'response.write strSql &"<br>"	
	rsACADEMYget.Open strSql,dbACADEMYget
		IF not rsACADEMYget.EOF THEN
			
			Fidx = rsACADEMYget("idx")
			FECode = rsACADEMYget("evtcode")
			Ffbtitle = rsACADEMYget("fbtitle")
			Ffbdesc  = rsACADEMYget("fbdesc")
			Ffbimage= rsACADEMYget("fbimage")
			Ftwlink= rsACADEMYget("twlink")
			Ftwtag1= rsACADEMYget("twtag1")
			Ftwtag2= rsACADEMYget("twtag2")
			Fkatitle= rsACADEMYget("katitle")
			Fkaimage= rsACADEMYget("kaimage")
			Fkalink= rsACADEMYget("kalink")
			
		End IF
		rsACADEMYget.Close
	End Function
End Class

'// Select Tag 생성 (년) //
Function OptYear(s_year,e_year,n_year,fid)
	dim i, strYear

	if n_year="" or isNull(n_year) then
		n_year = Year(date)
	end if

	strYear = "<Select name='" & fid & "'>"
	strYear = strYear & "<option value=''>선택</option>"
	For i = s_year to e_year
		strYear = strYear & "<Option value='" & i & "'"
		if i = Cint(n_year) then 
			strYear = strYear & " Selected"
		end if
		strYear = strYear & ">" & i & "</option>"
	Next
	strYear = strYear & "</Select>년"
	OptYear = strYear
End Function

'// Select Tag 생성 (월) //
Function OptMonth(n_month,fid)
	dim i, strMonth
	
	if n_month="" or isNull(n_month) then
		n_month = Month(date)
	end if

	strMonth = "<Select name='" & fid & "'>"
	strMonth = strMonth & "<option value=''>선택</option>"
	For i = 1 to 12
		if i<10 then
			strMonth = strMonth & "<Option value='0" & i & "'"
		else
			strMonth = strMonth & "<Option value='" & i & "'"
		end if

		if i = Cint(n_month) then 
			strMonth = strMonth & " Selected"
		end if

		strMonth = strMonth & ">" & i & "</option>"
	Next
	strMonth = strMonth & "</Select>월"
	OptMonth = strMonth
End Function

'// Select Tag 생성 (일) //
Function OptDay(n_day,fid)
	dim i, strDay

	if n_day="" or isNull(n_day) then
		n_day = Day(date)
	end if

	strDay = "<Select name='" & fid & "'>"
	strDay = strDay & "<option value=''>선택</option>"
	For i = 1 to 31
		if i<10 then
			strDay = strDay & "<Option value='0" & i & "'"
		else
			strDay = strDay & "<Option value='" & i & "'"
		end if

		if i = Cint(n_day) then 
			strDay = strDay & " Selected"
		end if

		strDay = strDay & ">" & i & "</option>"
	Next
	strDay = strDay & "</Select>일"
	OptDay = strDay
End Function
%>
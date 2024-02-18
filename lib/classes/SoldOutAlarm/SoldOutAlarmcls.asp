<%
'####################################################
' Description :  품절상품입고알림 통계 클래스
' History : 2018.02.27 원승현 생성
'			2020.03.20 한용민 수정(테스트서버 환경셋팅, 회원등급 개편등급으로 변경, 장바구니건수 추가)
'####################################################

'// 품절상품입고알림 통계 관련 클래스
Class CSoldOutAlarm
	Public Fidx	'// 통계 리스트 idx값
	Public FCatecode '// 카테코리 코드
	Public FStartdate '// 기간 시작일
	Public FEnddate '// 기간 종료일
	Public FMakerId '// 브랜드 ID
	Public FCateName1 '// 1차 카테고리 명
	Public FCateName2 '// 2차 카테고리 명
	Public FItemId '// 상품코드
	Public FItemName '// 상품명
	Public FOptionCheck '// 옵션여부 체크(옵션 > 0 이면 리스트 상품명 옆에 옵션신청현황 표시)
	Public FListTotalCount '// 리스트에서 보여지는 총합
	Public FListPCCount '// 리스트에서 보여지는 pc 총합
	Public FListMobileCount '// 리스트에서 보여지는 Mobile 총합
	Public FListAppCount '// 리스트에서 보여지는 App 총합
	Public FListBuyCount '// 리스트에서 보여지는 구매 총합
	Public FGraphUserLevel
	Public FGraphUserCount
	Public FGraphUserPercent
	Public FCategoryCnt
	Public FListImage
	Public FBrandName
	Public FBaesongGubun
	Public FOptionName
	Public FOptionTotalCnt
	Public FOptionPcCnt
	Public FOptionMobileCnt
	Public FOptionAppCnt
	Public FOptionBuyCnt
	public fbagunicnt

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


Class CGetSoldOutAlarm

	public FSoldOutAlarmList()
	Public FUserLevelAlarmList()
	Public FCategoryAlarmList()
	Public FItemBasicInfoAlarm()
	Public FItemOptionInfoAlarm()
	Public FtotalCount
	Public FAlarmCount '// 전체 접수 건
	Public FPcCount '// 해당 기간 내 pc 접수 건
	Public FMobileCount '// 해당 기간 내 mobile 접수 건
	Public FAppCount '// 해당 기간 내App 접수 건
	Public FBuyCount '// 해당 기간 내 구매건
	public fbagunicnt
	public ftendb
	Public FtotalPage

	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FRectSearchGubun '// 검색어구분 (1-상품코드, 2-상품명)
	Public FRectSearchKeyword '// 검색어
	Public FRectStartDate '// 기간 시작일
	Public FRectEndDate '// 기간 종료일
	Public FRectCateCode '// 카테고리 코드
	Public FRectMakerId '// 브랜드ID
	Public FRectItemId '// 상품코드

	'// 품절상품 입고알림 리스트
	public sub GetSoldOutAlarmList()

		dim i, j, sqlStr

		'전체 카운트
		sqlStr = " select " &vbCrLf
		sqlStr = sqlStr & " 	count(sa.idx) as AlarmTotalCnt, CEILING(CAST(COUNT(sa.Idx) AS FLOAT)/20) as totPage, " &vbCrLf
		sqlStr = sqlStr & " 	count(distinct sa.itemid) as TotalCnt, " &vbCrLf
		sqlStr = sqlStr & " 	count(iif(sa.ItemOptionCode<>'0000',sa.idx,null)) as optionCheck, " &vbCrLf
		sqlStr = sqlStr & " 	count(iif(sa.PlatForm='PCWEB',sa.idx,null)) as AlarmPCCnt, " &vbCrLf
		sqlStr = sqlStr & " 	count(iif(sa.PlatForm='MOBILE',sa.idx,null)) as AlarmMobileCnt, " &vbCrLf
		sqlStr = sqlStr & " 	count(iif(sa.PlatForm='APP',sa.idx,null)) as AlarmAppCnt, " &vbCrLf
		sqlStr = sqlStr & " 	0 as AlarmBuyCnt, 0.Alarmbagunicnt " &vbCrLf
		sqlStr = sqlStr & " from "& ftendb &"db_my10x10.dbo.tbl_SoldOutProductAlarm as sa with(noLock) " &vbCrLf
		sqlStr = sqlStr & " 	join db_item.dbo.tbl_item as i with(noLock) " &vbCrLf
		sqlStr = sqlStr & " 		on sa.itemid=i.itemid " &vbCrLf
		sqlStr = sqlStr & " 	left join db_item.dbo.tbl_display_cate_item as ci with(noLock) " &vbCrLf
		sqlStr = sqlStr & " 		on sa.itemid=ci.itemid " &vbCrLf
		sqlStr = sqlStr & " 			and ci.isDefault='y' " &vbCrLf
		sqlStr = sqlStr & " where sa.RegDate between '"&FRectStartDate&"' and '"&FRectEndDate&"' " &vbCrLf

		If Trim(FRectSearchGubun)="1" Then
			sqlStr = sqlStr & " AND SA.ItemId ='"&FRectSearchKeyword&"' " &vbCrLf
		End If
		If Trim(FRectSearchGubun)="2" Then
			sqlStr = sqlStr & " AND I.ItemName LIKE '%"&FRectSearchKeyword&"%' " &vbCrLf
		End If
		If Trim(FRectMakerId)<>"" Then 
			sqlStr = sqlStr & " AND I.makerid = '"&FRectMakerId&"' " &vbCrLf
		End If
		If Trim(FRectCateCode)<>"" Then
			sqlStr = sqlStr & " AND CI.catecode LIKE '"&FRectCateCode&"%' " &vbCrLf
		End If

		db3_rsget.Open sqlstr, db3_dbget, 1
			FtotalCount = db3_rsget("TotalCnt")
			FAlarmCount = db3_rsget("AlarmTotalCnt")
			FPcCount = db3_rsget("AlarmPCCnt")
			FMobileCount = db3_rsget("AlarmMobileCnt")
			FAppCount = db3_rsget("AlarmAppCnt")
			FBuyCount = db3_rsget("AlarmBuyCnt")
			fbagunicnt = db3_rsget("Alarmbagunicnt")
			FtotalPage = db3_rsget("totPage")
		db3_rsget.close


		'카테고리/상품별 카운트
		sqlStr = "select " &vbCrLf
		sqlStr = sqlStr & "	isNull(ci.catecode,'') as catecode, db_item.dbo.getCateCodeFullDepthName(left(ci.catecode,6)) as cateName " &vbCrLf
		sqlStr = sqlStr & "	, i.itemname, i.itemid, i.makerid, " &vbCrLf
		sqlStr = sqlStr & "	count(sa.idx) as totalCnt, " &vbCrLf
		sqlStr = sqlStr & "	count(iif(sa.ItemOptionCode<>'0000',sa.idx,null)) as optionCheck, " &vbCrLf
		sqlStr = sqlStr & "	count(iif(sa.PlatForm='PCWEB',sa.idx,null)) as PcCnt, " &vbCrLf
		sqlStr = sqlStr & "	count(iif(sa.PlatForm='MOBILE',sa.idx,null)) as MobileCnt, " &vbCrLf
		sqlStr = sqlStr & "	count(iif(sa.PlatForm='APP',sa.idx,null)) as AppCnt, " &vbCrLf
		sqlStr = sqlStr & "	0 as buyCnt, 0.baguniCnt " &vbCrLf
		sqlStr = sqlStr & "from "& ftendb &"db_my10x10.dbo.tbl_SoldOutProductAlarm as sa with(noLock) " &vbCrLf
		sqlStr = sqlStr & "	join db_item.dbo.tbl_item as i with(noLock) " &vbCrLf
		sqlStr = sqlStr & "		on sa.itemid=i.itemid " &vbCrLf
		sqlStr = sqlStr & "	left join db_item.dbo.tbl_display_cate_item as ci with(noLock) " &vbCrLf
		sqlStr = sqlStr & "		on sa.itemid=ci.itemid " &vbCrLf
		sqlStr = sqlStr & "			and ci.isDefault='y' " &vbCrLf
		sqlStr = sqlStr & "where sa.RegDate between '"&FRectStartDate&"' and '"&FRectEndDate&"' " &vbCrLf

		If Trim(FRectSearchGubun)="1" Then
			sqlStr = sqlStr & " AND SA.ItemId ='"&FRectSearchKeyword&"' " &vbCrLf
		End If
		If Trim(FRectSearchGubun)="2" Then
			sqlStr = sqlStr & " AND I.ItemName LIKE '%"&FRectSearchKeyword&"%' " &vbCrLf
		End If
		If Trim(FRectMakerId)<>"" Then 
			sqlStr = sqlStr & " AND I.makerid = '"&FRectMakerId&"' " &vbCrLf
		End If
		If Trim(FRectCateCode)<>"" Then
			sqlStr = sqlStr & " AND CI.catecode LIKE '"&FRectCateCode&"%' " &vbCrLf
		End If

		sqlStr = sqlStr & "group by isNull(ci.catecode,''), left(ci.catecode,6), " &vbCrLf
		sqlStr = sqlStr & "	i.itemname, i.itemid, i.makerid " &vbCrLf
		sqlStr = sqlStr & "order by TotalCnt DESC " &vbCrLf
		sqlStr = sqlStr & "offset " & (FRectCurrpage-1)*FRectpagesize & " rows fetch next " & FRectpagesize & " rows only " &vbCrLf

		'rw sqlstr
		db3_rsget.pagesize = FRectpagesize
		db3_rsget.Open sqlstr, db3_dbget, 1

		FResultCount = db3_rsget.RecordCount
        if (FResultCount<1) then FResultCount=0
		redim FSoldOutAlarmList(FResultCount)

		i=0
		if not db3_rsget.EOF  Then
			do until db3_rsget.eof
				set FSoldOutAlarmList(i) = new CSoldOutAlarm
				FSoldOutAlarmList(i).Fcatecode = db3_rsget("catecode")
				FSoldOutAlarmList(i).FCateName1 = db3_rsget("cateName")
				if instr(db3_rsget("cateName"),"^^")>0 then
					FSoldOutAlarmList(i).FCateName1 = split(db3_rsget("cateName"),"^^")(0)
					FSoldOutAlarmList(i).FCateName2 = split(db3_rsget("cateName"),"^^")(1)
				end if
				FSoldOutAlarmList(i).FItemName = db3_rsget("itemname")
				FSoldOutAlarmList(i).FOptionCheck = db3_rsget("OPTIONCHECK")
				FSoldOutAlarmList(i).FMakerId = db3_rsget("makerid")
				FSoldOutAlarmList(i).FItemId = db3_rsget("itemid")
				FSoldOutAlarmList(i).FListTotalCount = db3_rsget("TotalCnt")
				FSoldOutAlarmList(i).FListPCCount = db3_rsget("PCCNT")
				FSoldOutAlarmList(i).FListMobileCount = db3_rsget("MOBILECNT")
				FSoldOutAlarmList(i).FListAppCount = db3_rsget("AppCnt")
				FSoldOutAlarmList(i).FListBuyCount = db3_rsget("BuyCnt")
				FSoldOutAlarmList(i).fbagunicnt = db3_rsget("bagunicnt")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	End Sub

	'// 등급별 신청 그래프
	Public Sub GetUserLevelAlarm()
		dim i, j, sqlStr

		sqlStr = " SELECT  " &vbCrLf
		sqlStr = sqlStr & " 	CASE WHEN L.userlevel=0 THEN 'WHITE' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=1 THEN 'RED' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=2 THEN 'VIP' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=3 THEN 'VIP GOLD' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=4 THEN 'VVIP' " &vbCrLf
		'sqlStr = sqlStr & " 		WHEN L.userlevel=5 THEN 'orange' " &vbCrLf
		'sqlStr = sqlStr & " 		WHEN L.userlevel=6 THEN 'vvip' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=7 THEN 'STAFF' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=8 THEN 'FAMILY' " &vbCrLf
		sqlStr = sqlStr & " 		WHEN L.userlevel=9 THEN 'BIZ' " &vbCrLf
		sqlStr = sqlStr & " 	else '' " &vbCrLf
		sqlStr = sqlStr & " 	END AS userlevel " &vbCrLf
		sqlStr = sqlStr & " , COUNT(L.userlevel) AS CNT " &vbCrLf
		sqlStr = sqlStr & " FROM "& ftendb &"db_my10x10.[dbo].[tbl_SoldOutProductAlarm] SA " &vbCrLf
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item I ON SA.itemid = I.itemid " &vbCrLf
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_display_cate_item CI ON SA.itemid = CI.itemid AND isDefault='y' " &vbCrLf
		sqlStr = sqlStr & " INNER JOIN "& ftendb &"db_user.dbo.tbl_loginData L on SA.UserId = L.UserId " &vbCrLf
		sqlStr = sqlStr & " WHERE SA.Regdate >= '"&FRectStartDate&"' And SA.Regdate < '"&FRectEndDate&"' " &vbCrLf
		If Trim(FRectSearchGubun)="1" Then
			sqlStr = sqlStr & " AND SA.ItemId ='"&FRectSearchKeyword&"' " &vbCrLf
		End If
		If Trim(FRectSearchGubun)="2" Then
			sqlStr = sqlStr & " AND I.ItemName LIKE '%"&FRectSearchKeyword&"%' " &vbCrLf
		End If
		If Trim(FRectMakerId)<>"" Then 
			sqlStr = sqlStr & " AND I.makerid = '"&FRectMakerId&"' " &vbCrLf
		End If
		If Trim(FRectCateCode)<>"" Then
			sqlStr = sqlStr & " AND CI.catecode LIKE '"&FRectCateCode&"%' " &vbCrLf
		End If
		sqlStr = sqlStr & " GROUP BY L.userlevel "

		db3_rsget.Open sqlstr, db3_dbget, 1
		FResultCount = db3_rsget.RecordCount
		redim FUserLevelAlarmList(FResultCount)
		i=0
		if not db3_rsget.EOF  Then
			do until db3_rsget.eof
				set FUserLevelAlarmList(i) = new CSoldOutAlarm
				FUserLevelAlarmList(i).FGraphUserLevel = db3_rsget("userlevel")
				FUserLevelAlarmList(i).FGraphUserCount = db3_rsget("CNT")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	End Sub

	'// 카테고리별 신청 그래프
	Public Sub GetCategoryAlarm()
		dim i, j, sqlStr

		sqlStr = " SELECT  " &vbCrLf
		sqlStr = sqlStr & " 	i.dispcate1 " &vbCrLf
		sqlStr = sqlStr & " 	, ( SELECT TOP 1 catename FROM db_item.dbo.tbl_display_cate WHERE catecode = dispcate1 ) AS CateName " &vbCrLf
		sqlStr = sqlStr & " 	, COUNT(SA.itemid) AS CNT " &vbCrLf
		sqlStr = sqlStr & " FROM "& ftendb &"db_my10x10.[dbo].[tbl_SoldOutProductAlarm] SA " &vbCrLf
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_item I ON SA.itemid = I.itemid " &vbCrLf
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.tbl_display_cate_item CI ON SA.itemid = CI.itemid AND isDefault='y' " &vbCrLf
		sqlStr = sqlStr & " WHERE SA.Regdate >= '"&FRectStartDate&"' And SA.Regdate < '"&FRectEndDate&"' " &vbCrLf
		sqlStr = sqlStr & " 	AND i.dispcate1 is not null " &vbCrLf
		If Trim(FRectSearchGubun)="1" Then
			sqlStr = sqlStr & " AND SA.ItemId ='"&FRectSearchKeyword&"' " &vbCrLf
		End If
		If Trim(FRectSearchGubun)="2" Then
			sqlStr = sqlStr & " AND I.ItemName LIKE '%"&FRectSearchKeyword&"%' " &vbCrLf
		End If
		If Trim(FRectMakerId)<>"" Then 
			sqlStr = sqlStr & " AND I.makerid = '"&FRectMakerId&"' " &vbCrLf
		End If
		If Trim(FRectCateCode)<>"" Then
			sqlStr = sqlStr & " AND CI.catecode LIKE '"&FRectCateCode&"%' " &vbCrLf
		End If
		sqlStr = sqlStr & " GROUP BY dispcate1 " &vbCrLf
		sqlStr = sqlStr & " ORDER BY i.dispcate1 "
		db3_rsget.Open sqlstr, db3_dbget, 1
		FResultCount = db3_rsget.RecordCount
		redim FCategoryAlarmList(FResultCount)
		i=0
		if not db3_rsget.EOF  Then
			do until db3_rsget.eof
				set FCategoryAlarmList(i) = new CSoldOutAlarm
				FCategoryAlarmList(i).FCateName1 = db3_rsget("CateName")
				FCategoryAlarmList(i).FCateCode = db3_rsget("dispcate1")
				FCategoryAlarmList(i).FCategoryCnt = db3_rsget("CNT")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	End Sub

	'// 상품 기본정보
	Public Sub GetItemBasicInfoAlarm()
		dim i, j, sqlStr

		sqlStr = " SELECT  " &vbCrLf
		sqlStr = sqlStr & " 		'http://thumbnail.10x10.co.kr/webimage/image/list/'+  " &vbCrLf
		sqlStr = sqlStr & " 		CASE WHEN LEN(CONVERT(VARCHAR(20),(itemid / 10000)))=1 THEN '0'+convert(VARCHAR(20),(itemid / 10000)) ELSE CONVERT(VARCHAR(20),(itemid / 10000)) END+  " &vbCrLf
		sqlStr = sqlStr & " 		'/'+listimage AS listimage	  " &vbCrLf
		sqlStr = sqlStr & " 		,brandname, itemid, itemname, CASE WHEN mwdiv='M' THEN '매입' WHEN mwdiv='W' THEN '위탁' WHEN mwdiv='U' THEN '업체' END AS baesongGubun " &vbCrLf
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item Where itemid='"&FRectItemId&"' " &vbCrLf
		db3_rsget.Open sqlstr, db3_dbget, 1
		redim FItemBasicInfoAlarm(0)
		if not db3_rsget.EOF  Then
			set FItemBasicInfoAlarm(0) = new CSoldOutAlarm
			FItemBasicInfoAlarm(0).FListImage = db3_rsget("listimage")
			FItemBasicInfoAlarm(0).FBrandName = db3_rsget("brandname")
			FItemBasicInfoAlarm(0).FItemId = db3_rsget("itemid")
			FItemBasicInfoAlarm(0).FItemName = db3_rsget("itemname")
			FItemBasicInfoAlarm(0).FBaesongGubun = db3_rsget("baesongGubun")
		end if
		db3_rsget.Close
	End Sub

	'// 옵션별 신청현황
	Public Sub GetItemOptionInfoAlarm()
		dim i, j, sqlStr

		sqlStr = " SELECT   " &vbCrLf
		sqlStr = sqlStr & " SA.ItemOptionCode  " &vbCrLf
		sqlStr = sqlStr & " , SA.ItemId  " &vbCrLf
		sqlStr = sqlStr & " , CASE WHEN IO.optionname IS NULL THEN '상품자체알림신청' ELSE IO.optionname END AS optionname  " &vbCrLf
		sqlStr = sqlStr & " , COUNT(SA.UserId) AS cnt  " &vbCrLf
		sqlStr = sqlStr & " , (  " &vbCrLf
		sqlStr = sqlStr & "		SELECT COUNT(IDX) FROM db_my10x10.[dbo].[tbl_SoldOutProductAlarm]  " &vbCrLf
		sqlStr = sqlStr & "		WHERE Regdate >= '"&FRectStartDate&"' And Regdate < '"&FRectEndDate&"' AND PlatForm='PCWEB'  " &vbCrLf
		sqlStr = sqlStr & "		AND itemid = SA.itemid AND ItemOptionCode = SA.ItemOptionCode  " &vbCrLf
		sqlStr = sqlStr & " ) AS PCCNT  " &vbCrLf
		sqlStr = sqlStr & " , (  " &vbCrLf
		sqlStr = sqlStr & "		SELECT COUNT(IDX) FROM db_my10x10.[dbo].[tbl_SoldOutProductAlarm]  " &vbCrLf
		sqlStr = sqlStr & "		WHERE Regdate >= '"&FRectStartDate&"' And Regdate < '"&FRectEndDate&"' AND PlatForm='MOBILE'  " &vbCrLf
		sqlStr = sqlStr & "		AND itemid = SA.itemid AND ItemOptionCode = SA.ItemOptionCode  " &vbCrLf
		sqlStr = sqlStr & " ) AS MOBILECNT  " &vbCrLf
		sqlStr = sqlStr & " 	, (  " &vbCrLf
		sqlStr = sqlStr & " 		SELECT COUNT(IDX) FROM db_my10x10.[dbo].[tbl_SoldOutProductAlarm]  " &vbCrLf
		sqlStr = sqlStr & " 		WHERE Regdate >= '"&FRectStartDate&"' And Regdate < '"&FRectEndDate&"' AND PlatForm='APP'  " &vbCrLf
		sqlStr = sqlStr & " 			AND itemid = SA.itemid AND ItemOptionCode = SA.ItemOptionCode  " &vbCrLf
		sqlStr = sqlStr & " 	) AS APPCNT  " &vbCrLf
		sqlStr = sqlStr & " 	, (  " &vbCrLf
		sqlStr = sqlStr & " 		SELECT COUNT(d.itemid) FROM  " &vbCrLf
		sqlStr = sqlStr & " 			db_order.dbo.tbl_order_master m  " &vbCrLf
		sqlStr = sqlStr & " 			INNER JOIN db_order.dbo.tbl_order_detail d on m.orderserial = d.orderserial  " &vbCrLf
		sqlStr = sqlStr & " 		WHERE m.ipkumdiv>3  " &vbCrLf
		sqlStr = sqlStr & " 			AND m.jumundiv not in ('6','9')  " &vbCrLf
		sqlStr = sqlStr & " 			AND m.cancelyn='N'  " &vbCrLf
		sqlStr = sqlStr & " 			AND m.sitename='10x10'  " &vbCrLf
		sqlStr = sqlStr & " 			AND m.regdate >= '"&FRectStartDate&"'  " &vbCrLf
		sqlStr = sqlStr & " 			AND m.regdate < '"&FRectEndDate&"'  " &vbCrLf
		sqlStr = sqlStr & " 			AND d.itemid = SA.itemid   " &vbCrLf
		sqlStr = sqlStr & " 			AND d.itemoption = SA.ItemOptionCode  " &vbCrLf
		sqlStr = sqlStr & " 	  ) AS BuyCnt  " &vbCrLf
		sqlStr = sqlStr & " FROM db_my10x10.[dbo].[tbl_SoldOutProductAlarm] SA  " &vbCrLf
		sqlStr = sqlStr & " LEFT JOIN db_item.[dbo].[tbl_item_option] IO on SA.itemid = IO.itemid And SA.ItemOptionCode = IO.itemoption  " &vbCrLf
		sqlStr = sqlStr & " WHERE SA.itemid='"&FRectItemId&"'  " &vbCrLf
		sqlStr = sqlStr & " 	AND SA.Regdate >= '"&FRectStartDate&"' And SA.Regdate < '"&FRectEndDate&"'  " &vbCrLf
		sqlStr = sqlStr & " GROUP BY SA.ItemId, SA.ItemOptionCode, IO.optionname  " &vbCrLf
		sqlStr = sqlStr & " ORDER BY COUNT(SA.UserId) DESC  "
		db3_rsget.Open sqlstr, db3_dbget, 1
		FResultCount = db3_rsget.RecordCount
		redim FItemOptionInfoAlarm(FResultCount)
		i=0
		if not db3_rsget.EOF  Then
			do until db3_rsget.eof
				set FItemOptionInfoAlarm(i) = new CSoldOutAlarm
				FItemOptionInfoAlarm(i).FOptionName = db3_rsget("optionname")
				FItemOptionInfoAlarm(i).FOptionTotalCnt = db3_rsget("cnt")
				FItemOptionInfoAlarm(i).FOptionPcCnt = db3_rsget("PCCNT")
				FItemOptionInfoAlarm(i).FOptionMobileCnt = db3_rsget("MOBILECNT")
				FItemOptionInfoAlarm(i).FOptionAppCnt = db3_rsget("APPCNT")
				FItemOptionInfoAlarm(i).FOptionBuyCnt = db3_rsget("BuyCnt")
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	End Sub

	Private Sub Class_Initialize()
		IF application("Svr_Info")="Dev" THEN
			ftendb = "tendb."
		end if
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

%>

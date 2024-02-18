<%
'####################################################
' Description :  ǰ����ǰ�԰�˸� ��� Ŭ����
' History : 2018.02.27 ������ ����
'			2020.03.20 �ѿ�� ����(�׽�Ʈ���� ȯ�����, ȸ����� ���������� ����, ��ٱ��ϰǼ� �߰�)
'####################################################

'// ǰ����ǰ�԰�˸� ��� ���� Ŭ����
Class CSoldOutAlarm
	Public Fidx	'// ��� ����Ʈ idx��
	Public FCatecode '// ī���ڸ� �ڵ�
	Public FStartdate '// �Ⱓ ������
	Public FEnddate '// �Ⱓ ������
	Public FMakerId '// �귣�� ID
	Public FCateName1 '// 1�� ī�װ� ��
	Public FCateName2 '// 2�� ī�װ� ��
	Public FItemId '// ��ǰ�ڵ�
	Public FItemName '// ��ǰ��
	Public FOptionCheck '// �ɼǿ��� üũ(�ɼ� > 0 �̸� ����Ʈ ��ǰ�� ���� �ɼǽ�û��Ȳ ǥ��)
	Public FListTotalCount '// ����Ʈ���� �������� ����
	Public FListPCCount '// ����Ʈ���� �������� pc ����
	Public FListMobileCount '// ����Ʈ���� �������� Mobile ����
	Public FListAppCount '// ����Ʈ���� �������� App ����
	Public FListBuyCount '// ����Ʈ���� �������� ���� ����
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
	Public FAlarmCount '// ��ü ���� ��
	Public FPcCount '// �ش� �Ⱓ �� pc ���� ��
	Public FMobileCount '// �ش� �Ⱓ �� mobile ���� ��
	Public FAppCount '// �ش� �Ⱓ ��App ���� ��
	Public FBuyCount '// �ش� �Ⱓ �� ���Ű�
	public fbagunicnt
	public ftendb
	Public FtotalPage

	Public FRectpagesize
	Public FRectcurrpage
	Public FResultCount
	Public FRectSearchGubun '// �˻���� (1-��ǰ�ڵ�, 2-��ǰ��)
	Public FRectSearchKeyword '// �˻���
	Public FRectStartDate '// �Ⱓ ������
	Public FRectEndDate '// �Ⱓ ������
	Public FRectCateCode '// ī�װ� �ڵ�
	Public FRectMakerId '// �귣��ID
	Public FRectItemId '// ��ǰ�ڵ�

	'// ǰ����ǰ �԰�˸� ����Ʈ
	public sub GetSoldOutAlarmList()

		dim i, j, sqlStr

		'��ü ī��Ʈ
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


		'ī�װ�/��ǰ�� ī��Ʈ
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

	'// ��޺� ��û �׷���
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

	'// ī�װ��� ��û �׷���
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

	'// ��ǰ �⺻����
	Public Sub GetItemBasicInfoAlarm()
		dim i, j, sqlStr

		sqlStr = " SELECT  " &vbCrLf
		sqlStr = sqlStr & " 		'http://thumbnail.10x10.co.kr/webimage/image/list/'+  " &vbCrLf
		sqlStr = sqlStr & " 		CASE WHEN LEN(CONVERT(VARCHAR(20),(itemid / 10000)))=1 THEN '0'+convert(VARCHAR(20),(itemid / 10000)) ELSE CONVERT(VARCHAR(20),(itemid / 10000)) END+  " &vbCrLf
		sqlStr = sqlStr & " 		'/'+listimage AS listimage	  " &vbCrLf
		sqlStr = sqlStr & " 		,brandname, itemid, itemname, CASE WHEN mwdiv='M' THEN '����' WHEN mwdiv='W' THEN '��Ź' WHEN mwdiv='U' THEN '��ü' END AS baesongGubun " &vbCrLf
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

	'// �ɼǺ� ��û��Ȳ
	Public Sub GetItemOptionInfoAlarm()
		dim i, j, sqlStr

		sqlStr = " SELECT   " &vbCrLf
		sqlStr = sqlStr & " SA.ItemOptionCode  " &vbCrLf
		sqlStr = sqlStr & " , SA.ItemId  " &vbCrLf
		sqlStr = sqlStr & " , CASE WHEN IO.optionname IS NULL THEN '��ǰ��ü�˸���û' ELSE IO.optionname END AS optionname  " &vbCrLf
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

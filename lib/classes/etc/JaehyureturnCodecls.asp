<%
Class jaehyuitem
	Public FReturnName
	Public FTempReturnCode
	Public FAddrChk
	Public FNormalSellFee
	Public FId
	Public FMapPid
	Public FMakerid
	Public FCNT
	Public FDeliver_phone
	Public FReturnCode
	Public FMapAddress
	Public FReturnAddress
	Public FRegCnt
	Public FNotMakerId
	Public FIsusing
End Class

Class RtCodeList
	Public FItemList()
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FMakerid
	Public FLotteSellyn
	Public FRegCntYN
	Public FAddrChk
	Public FNotMakerId
	Public FIsusing

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub


	Public Sub RtCodeList
		Dim strSql, strSqlAdd, strSqlAdd2, strSqlAdd3, i

		If FMakerid <> "" Then
			strSqlAdd = strSqlAdd & " AND Rt.makerid = '"&FMakerid&"' "
		End If

		If FIsusing = "Y" Then
			strSqlAdd = strSqlAdd & " AND R.isusing='Y' "
		ElseIf FIsusing = "N" Then
			strSqlAdd = strSqlAdd & " AND R.isusing='N' "
		End If

		If FAddrChk = "O" Then
			strSqlAdd2 = strSqlAdd2 & " AND replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2=T.returnAddress "
		ElseIf FAddrChk = "X" Then
			strSqlAdd2 = strSqlAdd2 & " AND replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2<>T.returnAddress "
		End If

		If FRegCntYN = "Y" Then
			strSqlAdd2 = strSqlAdd2 & " AND T.RegCnt = 0 "
		End If

		If FNotMakerId = "Y" Then
			strSqlAdd2 = strSqlAdd2 & " AND T.notMakerId is not null "
		End If

		If FLotteSellyn = "Y" Then
			strSqlAdd3 = strSqlAdd3 & " AND lt.lotteSellyn='Y' "
		ElseIf FLotteSellyn = "N" Then
			strSqlAdd3 = strSqlAdd3 & " AND lt.lotteSellyn='N' "
		End If

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) AS cnt, CEILING(CAST(COUNT(*) AS FLOAT)/" & FPageSize & ") AS totPg " & VBCRLF
		strSql = strSql & " FROM(  " & VBCRLF
		strSql = strSql & " 	SELECT R.ReturnName, R.returnCode as tempReturnCode  " & VBCRLF
		strSql = strSql & " 	,R.mapPid,rt.* " & VBCRLF
		strSql = strSql & " 	,r.returnAddress " & VBCRLF
		strSql = strSql & " 	,(SELECT count(*) from db_item.dbo.tbl_lotte_regItem AS lt " & VBCRLF
		strSql = strSql & " 		INNER Join db_item.dbo.tbl_item i " & VBCRLF
		strSql = strSql & " 		ON lt.itemid=i.itemid " & VBCRLF
		strSql = strSql & " 		AND i.mwdiv='U'" & VBCRLF                   '''업체배송상품만
		strSql = strSql & " 		AND i.makerid=rt.makerid " & VBCRLF
		strSql = strSql & " 		AND lt.Lottestatcd='30' " & VBCRLF
		strSql = strSql & " 		"&strSqlAdd3&" " & VBCRLF
		strSql = strSql & " 	) as RegCnt " & VBCRLF
		strSql = strSql & " 	,Nm.makerid as notMakerId, R.isusing " & VBCRLF
		strSql = strSql & " 	FROM db_temp.dbo.tbl_jaehyumall_returnInfo AS R " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_item.dbo.tbl_OutMall_BrandReturnCode AS Rt ON r.returncode=Rt.returncode " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid AS Nm ON rt.makerid=Nm.makerid and Nm.mallgubun = 'lotte' " & VBCRLF
		strSql = strSql & " 	WHERE rt.mallid = 'lotteCom' " & strSqlAdd & VBCRLF
		strSql = strSql & " ) AS T " & VBCRLF
		strSql = strSql & " INNER JOIN db_partner.dbo.tbl_partner as p on T.makerid = p.id " & VBCRLF
		strSql = strSql & " WHERE 1=1 " & strSqlAdd2
'		strSql = ""
'		strSql = strSql & " SELECT COUNT(*) AS cnt, CEILING(CAST(COUNT(*) AS FLOAT)/" & FPageSize & ") AS totPg From( " & VBCRLF
'		strSql = strSql & " 	SELECT " & VBCRLF
'		strSql = strSql & " 	R.ReturnName, R.returnCode as tempReturnCode " & VBCRLF
'		strSql = strSql & " 	,(CASE WHEN replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2=R.returnAddress then 1 else 0 end ) AS addrChk " & VBCRLF
'		strSql = strSql & " 	,p.id,R.mapPid,rt.* " & VBCRLF
'		strSql = strSql & " 	,replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2 as mapAddress, r.returnAddress " & VBCRLF
'		strSql = strSql & " 	,(SELECT count(*) from db_item.dbo.tbl_lotte_regItem AS lt  " & VBCRLF
'		strSql = strSql & " 		INNER Join db_item.dbo.tbl_item i " & VBCRLF
'		strSql = strSql & " 		ON lt.itemid=i.itemid " & VBCRLF
'		strSql = strSql & " 		AND i.mwdiv='U'" & VBCRLF                   '''업체배송상품만
'		strSql = strSql & " 		AND i.makerid=rt.makerid " & VBCRLF
'		strSql = strSql & " 		AND lt.Lottestatcd='30' " & VBCRLF
'		strSql = strSql & " 		"&strSqlAdd3&" " & VBCRLF
'		strSql = strSql & " 	) as RegCnt " & VBCRLF
'		strSql = strSql & " 	,Nm.makerid as notMakerId, R.isusing " & VBCRLF
'		strSql = strSql & " 	FROM db_temp.dbo.tbl_jaehyumall_returnInfo AS R " & VBCRLF
'		strSql = strSql & " 	LEFT JOIN db_partner.dbo.tbl_partner AS p ON R.mapPid=p.id " & VBCRLF
'		strSql = strSql & " 	LEFT JOIN db_item.dbo.tbl_OutMall_BrandReturnCode AS Rt ON p.id=Rt.makerid " & VBCRLF
'		strSql = strSql & " 	LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid AS Nm ON p.id=Nm.makerid and Nm.mallgubun = 'lotte' " & VBCRLF
'		strSql = strSql & " 	WHERE rt.mallid = 'lotteCom' " & strSqlAdd & VBCRLF
'		strSql = strSql & " ) AS T " & VBCRLF
'		strSql = strSql & " WHERE 1=1 " & strSqlAdd2
		rsget.Open strSql,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage)>Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If


		strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage) &" T.ReturnName, T.tempReturnCode " & VBCRLF
		strSql = strSql & " ,(CASE WHEN replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2=T.returnAddress then 1 else 0 end ) AS addrChk  " & VBCRLF
		strSql = strSql & " , T.mallid, T.makerid, T.returnCode " & VBCRLF
		strSql = strSql & " ,replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2 as mapAddress " & VBCRLF
		strSql = strSql & " ,T.returnaddress, T.RegCnt, T.notMakerid, T.isusing " & VBCRLF
		strSql = strSql & " FROM(  " & VBCRLF
		strSql = strSql & " 	SELECT R.ReturnName, R.returnCode as tempReturnCode  " & VBCRLF
		strSql = strSql & " 	,R.mapPid,rt.* " & VBCRLF
		strSql = strSql & " 	,r.returnAddress " & VBCRLF
		strSql = strSql & " 	,(SELECT count(*) from db_item.dbo.tbl_lotte_regItem AS lt " & VBCRLF
		strSql = strSql & " 		INNER Join db_item.dbo.tbl_item i " & VBCRLF
		strSql = strSql & " 		ON lt.itemid=i.itemid " & VBCRLF
		strSql = strSql & " 		AND i.mwdiv='U'" & VBCRLF                   '''업체배송상품만
		strSql = strSql & " 		AND i.makerid=rt.makerid " & VBCRLF
		strSql = strSql & " 		AND lt.Lottestatcd='30' " & VBCRLF
		strSql = strSql & " 		"&strSqlAdd3&" " & VBCRLF
		strSql = strSql & " 	) as RegCnt " & VBCRLF
		strSql = strSql & " 	,Nm.makerid as notMakerId, R.isusing " & VBCRLF
		strSql = strSql & " 	FROM db_temp.dbo.tbl_jaehyumall_returnInfo AS R " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_item.dbo.tbl_OutMall_BrandReturnCode AS Rt ON r.returncode=Rt.returncode " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid AS Nm ON rt.makerid=Nm.makerid and Nm.mallgubun = 'lotte' " & VBCRLF
		strSql = strSql & " 	WHERE rt.mallid = 'lotteCom' " & strSqlAdd & VBCRLF
		strSql = strSql & " ) AS T " & VBCRLF
		strSql = strSql & " INNER JOIN db_partner.dbo.tbl_partner as p on T.makerid = p.id " & VBCRLF
		strSql = strSql & " WHERE 1=1 " & strSqlAdd2
		strSql = strSql & " ORDER BY T.makerid ASC " 
'		strSql = ""
'		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage) &" T.* FROM( " & VBCRLF
'		strSql = strSql & " 	SELECT " & VBCRLF
'		strSql = strSql & " 	R.ReturnName, R.returnCode as tempReturnCode " & VBCRLF
'		strSql = strSql & " 	,(CASE WHEN replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2=R.returnAddress then 1 else 0 end ) AS addrChk " & VBCRLF
'		strSql = strSql & " 	,p.id,R.mapPid,rt.* " & VBCRLF
'		strSql = strSql & " 	,replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2 as mapAddress, r.returnAddress " & VBCRLF
'		strSql = strSql & " 	,(SELECT count(*) from db_item.dbo.tbl_lotte_regItem AS lt  " & VBCRLF
'		strSql = strSql & " 		INNER Join db_item.dbo.tbl_item i " & VBCRLF
'		strSql = strSql & " 		ON lt.itemid=i.itemid " & VBCRLF
'		strSql = strSql & " 		AND i.mwdiv='U'" & VBCRLF                   '''업체배송상품만
'		strSql = strSql & " 		AND i.makerid=rt.makerid " & VBCRLF
'		strSql = strSql & " 		AND lt.Lottestatcd='30' " & VBCRLF
'		strSql = strSql & " 		"&strSqlAdd3&" " & VBCRLF
'		strSql = strSql & " 	) as RegCnt " & VBCRLF
'		strSql = strSql & " 	,Nm.makerid as notMakerId, R.isusing " & VBCRLF
'		strSql = strSql & " 	FROM db_temp.dbo.tbl_jaehyumall_returnInfo AS R " & VBCRLF
'		strSql = strSql & " 	LEFT JOIN db_partner.dbo.tbl_partner AS p ON R.mapPid=p.id " & VBCRLF
'		strSql = strSql & " 	LEFT JOIN db_item.dbo.tbl_OutMall_BrandReturnCode AS Rt ON p.id=Rt.makerid " & VBCRLF
'		strSql = strSql & " 	LEFT JOIN db_temp.dbo.tbl_jaehyumall_not_in_makerid AS Nm ON p.id=Nm.makerid and Nm.mallgubun = 'lotte' " & VBCRLF
'		strSql = strSql & " 	WHERE rt.mallid = 'lotteCom' " & strSqlAdd & VBCRLF
'		strSql = strSql & " ) AS T " & VBCRLF
'		strSql = strSql & " WHERE 1=1 " & strSqlAdd2 & VBCRLF
'		strSql = strSql & " ORDER BY id ASC "
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1

		Redim preserve FItemList(FResultCount)

		'FPageCount = FCurrPage - 1

		i=0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				SET FItemList(i) = new jaehyuitem
					FItemList(i).FReturnName		= rsget("ReturnName")
					FItemList(i).FTempReturnCode	= rsget("tempReturnCode")
					FItemList(i).FAddrChk			= rsget("addrChk")
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FReturnCode		= rsget("returnCode")
					FItemList(i).FMapAddress		= rsget("mapAddress")
					FItemList(i).FReturnAddress		= rsget("returnAddress")
					FItemList(i).FRegCnt			= rsget("RegCnt")
					FItemList(i).FNotMakerId		= rsget("notMakerId")
					FItemList(i).FIsusing			= rsget("isusing")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub
	
	Public Sub NotRtCodeList
		Dim strSqlAdd,strSqlAdd2, i

		If FMakerid <> "" Then
			strSqlAdd = strSqlAdd & " AND i.makerid = '"&FMakerid&"' "
		End If

		If FAddrChk = "O" Then
			strSqlAdd2 = strSqlAdd2 & " AND rt.returnCode IS NOT NULL "
		ElseIf FAddrChk = "X" Then
			strSqlAdd2 = strSqlAdd2 & " AND rt.returnCode IS NULL "
		End If

		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) AS cnt, CEILING(CAST(COUNT(*) AS FLOAT)/" & FPageSize & ") AS totPg From( " & VBCRLF
		strSql = strSql & " 	SELECT i. makerid,count(*) as CNT , rt.returnCode, p.deliver_phone  " & VBCRLF
		strSql = strSql & " 	,replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2 as mapAddress, r.returnAddress  " & VBCRLF
		strSql = strSql & " 	FROM db_item.dbo.tbl_lotte_regItem AS lt  " & VBCRLF
		strSql = strSql & " 		INNER JOIN db_item.dbo.tbl_item AS i  " & VBCRLF
		strSql = strSql & " 		ON lt.itemid=i.itemid  " & VBCRLF
		strSql = strSql & " 		LEFT JOIN db_item.dbo.tbl_OutMall_BrandReturnCode AS Rt " & VBCRLF
		strSql = strSql & " 		ON rt.mallid = 'lotteCom' " & VBCRLF
		strSql = strSql & " 		AND rt.makerid=i.makerid " & VBCRLF
		strSql = strSql & "			LEFT JOIN db_partner.dbo.tbl_partner AS p ON i.makerid=p.id " & VBCRLF
		strSql = strSql & " 		LEFT JOIN db_temp.dbo.tbl_jaehyumall_returnInfo AS R on R.mapPid=p.id " & VBCRLF
		strSql = strSql & " 	WHERE lt.Lottestatcd='30' "&strSqlAdd&" " & VBCRLF
		strSql = strSql & " 	AND i.mwdiv='U' " & VBCRLF
		strSql = strSql & " 	AND lt.lotteSellyn='Y' " & VBCRLF
		strSql = strSql & " 	"&strSqlAdd2&" " & VBCRLF
		strSql = strSql & " 	GROUP BY i.makerid , rt.returnCode , p.id, p.deliver_phone, p.return_zipCode, p.return_address, R.returnAddress, p.return_address2 " & VBCRLF
		strSql = strSql & " 	HAVING count(*) > 0 " & VBCRLF
		strSql = strSql & " ) AS T " & VBCRLF
		strSql = strSql & " WHERE 1=1 "

		rsget.Open strSql,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage)>Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & CStr(FPageSize*FCurrPage) &" T.* FROM( " & VBCRLF
		strSql = strSql & " 	SELECT i.makerid, count(*) as CNT , rt.returnCode, p.deliver_phone  " & VBCRLF
		strSql = strSql & " 	,replace(p.return_zipCode,'-','') + ' ' + p.return_address + ' ' + p.return_address2 as mapAddress, r.returnAddress  " & VBCRLF
		strSql = strSql & " 	FROM db_item.dbo.tbl_lotte_regItem AS lt  " & VBCRLF
		strSql = strSql & " 		INNER JOIN db_item.dbo.tbl_item AS i  " & VBCRLF
		strSql = strSql & " 		ON lt.itemid=i.itemid  " & VBCRLF
		strSql = strSql & " 		LEFT JOIN db_item.dbo.tbl_OutMall_BrandReturnCode AS Rt " & VBCRLF
		strSql = strSql & " 		ON rt.mallid = 'lotteCom' " & VBCRLF
		strSql = strSql & " 		AND rt.makerid=i.makerid " & VBCRLF
		strSql = strSql & "			LEFT JOIN db_partner.dbo.tbl_partner AS p ON i.makerid=p.id " & VBCRLF
		strSql = strSql & " 		LEFT JOIN db_temp.dbo.tbl_jaehyumall_returnInfo AS R on R.mapPid=p.id " & VBCRLF
		strSql = strSql & " 	WHERE lt.Lottestatcd='30' "&strSqlAdd&" " & VBCRLF
		strSql = strSql & " 	AND i.mwdiv='U' " & VBCRLF
		strSql = strSql & " 	AND lt.lotteSellyn='Y' " & VBCRLF
		strSql = strSql & " 	"&strSqlAdd2&" " & VBCRLF
		strSql = strSql & " 	GROUP BY i.makerid , rt.returnCode , p.id, p.deliver_phone, p.return_zipCode, p.return_address, R.returnAddress, p.return_address2 " & VBCRLF
		strSql = strSql & " 	HAVING count(*) > 0 " & VBCRLF
		strSql = strSql & " ) AS T " & VBCRLF
		strSql = strSql & " WHERE 1=1 " & VBCRLF
		strSql = strSql & " ORDER BY T.makerid ASC "
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1

		Redim preserve FItemList(FResultCount)

		i=0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				SET FItemList(i) = new jaehyuitem
					FItemList(i).FMakerid			= rsget("makerid")
					FItemList(i).FCNT				= rsget("CNT")
					FItemList(i).FReturnCode		= rsget("returnCode")
					FItemList(i).FDeliver_phone		= rsget("deliver_phone")
					FItemList(i).FMapAddress		= rsget("mapAddress")
					FItemList(i).FReturnAddress		= rsget("returnAddress")
				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close

	End Sub
End Class
%>


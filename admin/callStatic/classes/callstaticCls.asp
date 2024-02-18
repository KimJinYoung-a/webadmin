<%
Class ccall_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fuserid
	public fcomtelno
	public fteltime
	public ftelterm
	public fclienttelno
	public fdisposition
	public fwavlink
	public finout
	public fdate
	public fcstrcalldate

End Class


Class ccall_list

	public FItemList()
	
End Class


Class ClsCall

	public FItemList()
	public FOneItem
	public FGubun
	public FUserID
	public FSDate
	public FEDate
	public FInOut
	public FDisposi
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount


	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	public Sub FUserCallList
		Dim sqlStr, i, vSubQuery
		
		'######## 상담원 콜센터 통계 조건들 ########
		' AND C.disposition = 'ANSWERED' AND C.billsec > 0
		' If FInOut = "in" Then --> vSubQuery = vSubQuery & " AND C.dcontext = 'inbound' AND (C.lastapp = 'Hangup' or C.lastapp = 'Dial') AND C.accountcode <> 'asterisk' "
		' ElseIf FInOut = "out" Then --> vSubQuery = vSubQuery & " AND C.dcontext = 'outbound' AND C.dst <> 's' "
		'######## 상담원 콜센터 통계 조건들 ########
		
		vSubQuery = " AND C.yyyymmdd Between '" & FSDate & "' AND '" & FEDate & "' "
		
		If FInOut = "in" Then
			vSubQuery = vSubQuery & " AND C.dcontext = 'inbound' "
		ElseIf FInOut = "out" Then
			vSubQuery = vSubQuery & " AND C.dcontext = 'outbound' "
		End If
		
		If FDisposi <> "" AND FDisposi <> "all" Then
			vSubQuery = vSubQuery & " AND C.disposition = '" & FDisposi & "' "
		End If
		
		sqlStr = "SELECT COUNT(*) " & _
				 "		FROM [db_datamart].[dbo].[tbl_call_cdr] AS C " & _
				 "	WHERE " & _
				 "		C.tenUserID = '" & FUserID & "' AND Left(C.extension,1) = '9' " & _
				 "	" & vSubQuery & " "
		db3_rsget.Open sqlStr, db3_dbget, 1
		ftotalcount = db3_rsget(0)
		db3_rsget.Close
		
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " " & _
				 "			C.yyyymmdd, C.extension, C.tenUserID, C.calldate, C.billsec, C.src, C.dst, C.disposition, C.userfield, Convert(varchar(30),C.calldate,120) AS cstrcalldate " & _
				 "	FROM [db_datamart].[dbo].[tbl_call_cdr] AS C " & _
				 "	WHERE " & _
				 "		C.tenUserID = '" & FUserID & "' AND Left(C.extension,1) = '9' " & _
				 "	" & vSubQuery & " " & _
				 "	ORDER BY C.yyyymmdd DESC "
		db3_rsget.Open sqlStr, db3_dbget ,1
		'response.write sqlStr
		
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		db3_rsget.PageSize= FPageSize
		If  not db3_rsget.EOF  then
			db3_rsget.absolutepage = FCurrPage
			Do Until db3_rsget.Eof
				set FItemList(i) = new ccall_oneitem
					FItemList(i).fdate			= db3_rsget("yyyymmdd")
					FItemList(i).fcomtelno		= db3_rsget("extension")
					FItemList(i).fuserid		= db3_rsget("tenUserID")
					FItemList(i).fteltime		= db3_rsget("calldate")
					FItemList(i).fcstrcalldate	= db3_rsget("cstrcalldate")
					FItemList(i).ftelterm		= db3_rsget("billsec")
					If FInOut = "all" Then
						FItemList(i).fclienttelno	= db3_rsget("src")
					ElseIf FInOut = "in" Then
						FItemList(i).fclienttelno	= db3_rsget("src")
					ElseIf FInOut = "out" Then
						FItemList(i).fclienttelno	= db3_rsget("dst")
					End If
					FItemList(i).fdisposition	= db3_rsget("disposition")
					If db3_rsget("userfield") = "" Then
						FItemList(i).fwavlink	= "x"
					Else
						FItemList(i).fwavlink	= "o"
					End IF
				i=i+1
				db3_rsget.moveNext
			Loop
		End If
		db3_rsget.Close
	end Sub
	

	public Sub FCallWavPlay
		Dim sqlStr
		sqlStr = "SELECT C.userfield FROM [db_datamart].[dbo].[tbl_call_cdr] AS C " & _
				 "	WHERE C.yyyymmdd = '" & FSDate & "' AND C.tenUserID = '" & FUserID & "' AND C.calldate = '" & FEDate & "' "
        db3_rsget.Open SqlStr, db3_dbget, 1
        
        set FOneItem = new ccall_oneitem

        If Not db3_rsget.Eof Then
			If db3_rsget("userfield") = "" Then
				FOneItem.fwavlink	= "x"
			Else
				FOneItem.fwavlink	= "http://203.84.251.210" & db3_rsget("userfield")
			End IF
        End If
        db3_rsget.Close
	end Sub





	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class

' 초데이터를 시분초 형식으로 변환
Function sec2time(ByVal sec)
	sec2time = Int(sec / 3600) & ":" & Right("0"&(Int(sec/60) Mod 60),2) & ":" & Right("0"&(sec Mod 60),2)
End Function 
%>
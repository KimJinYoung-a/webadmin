<%
Class hitchhiker_item
	Public FIdx
	Public FHvol
	Public Fevt_code
	Public Fmevt_code
	Public Fstartdate
	Public Fenddate
	Public Fregdate
	Public Fisusing
	public Fdelidate
End Class

Class viphitchhker
	Public FhitchList()
	Public FIdx
	Public FIsusing
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public Sub fnhitchlist
		Dim strSql, where

		If FIsusing <> "" Then
			where = where & " and isusing = '" & FIsusing & "' "
		End IF

		' ÃÑ °¹¼ö ±¸ÇÏ±â '
		strSql = "select count(*) as cnt " + vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_vip_hitchhiker where 1=1 "& where &" " & vbcrlf
		'response.write strSql
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
		Else
			FTotalCount = 0
		End If
		rsget.Close
		' ÃÑ °¹¼ö ±¸ÇÏ±â ³¡'
		
		'¸®½ºÆ® ±¸ÇÏ±â'
		strSql = ""
		strSql = strSql & " select top "& Cstr(FPageSize * FCurrPage) &" idx, Hvol, evt_code, mevt_code, startdate, enddate, regdate, isusing, delidate " & vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_vip_hitchhiker where 1=1 "& where &"  " & vbcrlf
		strSql = strSql & " order by idx desc"
		'response.write strSql
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1
		
		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FhitchList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FhitchList(i) = new hitchhiker_item
					FhitchList(i).FIdx 			= rsget("idx")
					FhitchList(i).FHvol			= rsget("Hvol")
					FhitchList(i).Fevt_code		= rsget("evt_code")
					FhitchList(i).Fmevt_code	= rsget("mevt_code")
					FhitchList(i).Fstartdate	= rsget("startdate")
					FhitchList(i).Fenddate 		= rsget("enddate")
					FhitchList(i).Fregdate 		= rsget("regdate")
					FhitchList(i).Fisusing 		= rsget("isusing")
					FhitchList(i).Fdelidate 		= rsget("delidate")
				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Sub

	Public Sub fnhitchmodify()
		dim SqlStr,i
		SqlStr = "SELECT idx, Hvol, evt_code, mevt_code, startdate, enddate, regdate, isusing, delidate " & _
				 "	FROM db_event.dbo.tbl_vip_hitchhiker WHERE idx = '" & FIdx & "' "
		rsget.Open sqlStr,dbget,1
		
		Set FOneItem = new hitchhiker_item
			FOneItem.FIdx 			= rsget("idx")
			FOneItem.FHvol			= rsget("Hvol")
			FOneItem.Fevt_code		= rsget("evt_code")
			FOneItem.Fmevt_code		= rsget("mevt_code")
			FOneItem.Fstartdate		= rsget("startdate")
			FOneItem.Fenddate 		= rsget("enddate")
			FOneItem.Fregdate 		= rsget("regdate")
			FOneItem.Fisusing 		= rsget("isusing")
			FOneItem.Fdelidate 		= rsget("delidate")
		rsget.Close
	End Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

End Class
%>

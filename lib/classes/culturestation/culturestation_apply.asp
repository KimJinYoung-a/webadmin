<%
	Class CCultureApplyItem
		public Fidx
		public Fuserid
		public Fusername
		public Fusercell
		public Fusermail
		public Flinkurl
		public Fwhyapply
		public Fregdate
		public Fuserlevel
		
	End Class
	
	Class CCultureApply
		public FItemList()
		public Foneitem
		public FCurrPage
		public FPageSize
		public FTotalCount
		public Fidx
		public Fuserid
		public Fusername
		public Fusercell
		public Fusermail
		public Flinkurl
		public Fwhyapply
		public Fregdate
		public Fuserlevel
		public FTotalPage

	    public FPageCount
		public FResultCount
	    public FScrollCount
		
		
	public sub getApplyList()
		dim sqlStr , i

		sqlStr = "exec [db_culture_station].[dbo].[sp_Ten_CultureStation_Apply_Cnt] '" & Fuserid & "', '" & Fusername & "' "
		'Response.write sqlStr &"<br>"
        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic  
		rsget.Open sqlStr,dbget,1

		if not rsget.EOF then
			FTotalCount = rsget(0)
		end if
		rsget.Close

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		
		If FTotalCount > 0 Then
			sqlStr = "EXEC [db_culture_station].[dbo].[sp_Ten_CultureStation_Apply] 'l', '', '" & Fuserid & "', '" & Fusername & "', '', '', '', '', '" & FPageSize*FCurrPage & "' "
			'Response.write sqlStr &"<br>"
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic  
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1
			
			FResultCount = rsget.recordcount
			redim preserve FItemList(FResultCount)
	
			i=0
			if  not rsget.EOF  then
				do until rsget.EOF
					set FItemList(i) = new CCultureApplyItem
	
					FItemList(i).Fidx			= rsget("idx")
					FItemList(i).Fuserid		= rsget("userid")
					FItemList(i).Fusername		= rsget("username")
					FItemList(i).Fusercell		= rsget("usercell")
					FItemList(i).Fusermail		= rsget("usermail")
					FItemList(i).Flinkurl		= rsget("linkurl")
					FItemList(i).Fwhyapply		= db2html(rsget("whyapply"))
					FItemList(i).Fregdate		= rsget("regdate")
					FItemList(i).Fuserlevel		= rsget("userlevel")
					
					rsget.movenext
					i=i+1
				loop
			end if
			rsget.Close
		End If
	end sub
	
	
	public sub getApplyView()
		dim sqlStr , i

		If FIdx <> "" Then
			sqlStr = "EXEC [db_culture_station].[dbo].[sp_Ten_CultureStation_Apply] 'v', '" & Fidx & "', '" & Fuserid & "', '" & Fusername & "', '', '', '', '', '' "
			'Response.write sqlStr &"<br>"
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic  
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1

			if  not rsget.EOF  then
				set FItemList(i) = new CCultureApplyItem

				Foneitem.Fidx			= rsget("idx")
				Foneitem.Fuserid		= rsget("userid")
				Foneitem.Fusername		= rsget("username")
				Foneitem.Fusercell		= rsget("usercell")
				Foneitem.Fusermail		= rsget("usermail")
				Foneitem.Flinkurl		= rsget("linkurl")
				Foneitem.Fwhyapply		= db2html(rsget("whyapply"))
				Foneitem.Fregdate		= rsget("regdate")
				Foneitem.Fuserlevel		= rsget("userlevel")
				
			end if
			rsget.Close
		End If
	end sub


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
	

	Private Sub Class_Initialize()
		FScrollCount = 10
	End Sub

	Private Sub Class_Terminate()
	End Sub
	
	End Class
%>
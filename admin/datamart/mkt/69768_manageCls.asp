<%
'###########################################################
' Description :  [2016 S/S 웨딩] Wedding Membership cls
' History : 2016-03-16 유태욱 생성
'###########################################################
%>
<%
Class CMagaZineItem
    public Fidx
    public Fsub_opt1
	Public Fsub_opt1_userid
	Public Fsub_opt1_suserid
	public Fsub_opt1_state
	public Fsub_opt2
	public Fsub_opt3
	public Fregdate
	public Fevt_code
	public Fuserid

end Class


Class CMagaZineContents
    public FItemList()

	public FPageSize
	public FCurrPage
	public FTotalPage
	public FTotalCount
	public FResultCount
	public FScrollCount

	Public FRectstate
	Public FRectevt_code

	''---------------------------------------------------------------------------------
	public function fnGetMagaZineList()
        dim sqlStr, sqlsearch, i

		If FRectstate <> "" THEN
			sqlsearch  = sqlsearch & " and RIGHT(sub_opt1, 1) = '" &FRectstate & "' "
		End If

		If FRectevt_code <> "" THEN
			sqlsearch  = sqlsearch & " and evt_code = " &FRectevt_code & ""
		End If
		
		'// 총 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_event.dbo.tbl_event_subscript "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

'		response.write sqlStr &"<Br>"
'		response.end
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " sub_idx , evt_code , userid , sub_opt1 , sub_opt2 , sub_opt3 , regdate"
        sqlStr = sqlStr & " from db_event.dbo.tbl_event_subscript "
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by sub_idx DESC"

'		response.write sqlStr &"<Br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
				set FItemList(i) = new CMagaZineItem
					
					FItemList(i).Fidx						= rsget("sub_idx")
					FItemList(i).Fevt_code				= rsget("evt_code")
					FItemList(i).Fuserid					= rsget("userid")
					FItemList(i).Fsub_opt1				= rsget("sub_opt1")
					if isarray(split(FItemList(i).Fsub_opt1,"/!/")) then
						FItemList(i).Fsub_opt1_userid	= split(rsget("sub_opt1"),"/!/")(0)
						FItemList(i).Fsub_opt1_suserid	= split(rsget("sub_opt1"),"/!/")(1)
						FItemList(i).Fsub_opt1_state	= split(rsget("sub_opt1"),"/!/")(2)
					end if
					FItemList(i).Fsub_opt2				= rsget("sub_opt2")
					FItemList(i).Fsub_opt3				= rsget("sub_opt3")
					FItemList(i).Fregdate					= rsget("regdate")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

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
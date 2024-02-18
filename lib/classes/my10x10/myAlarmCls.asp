<%
'###########################################################
' Description : MY알림
' Hieditor : 2009.04.17 허진원 생성
'			 2016.07.19 한용민 수정
'###########################################################

Class CMyAlarmByLevelItem
    public FlevelAlarmIdx
    public Fyyyymmdd
    public Fmsgdiv
    public Ftitle
    public Fsubtitle
    public Fcontents
    public Fuserlevel
    public FwwwTargetURL
    public FopenYN
    public FuseYN
    public Fregdate
    public Freguserid
    public Flastupdate

	'/사용금지		'/공용펑션에 공용함수 쓸것.		'/2016.07.20 한용민
    public function GetUserLevelName()
        select case CStr(Fuserlevel)
            case "100"
                GetUserLevelName = "우수회원 전체"
            case "2"
                GetUserLevelName = "BLUE"
            case "3"
                GetUserLevelName = "VIP SILVER"
            case "4"
                GetUserLevelName = "VIP GOLD"
            case else
                GetUserLevelName = Fuserlevel
        end select
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMyAlarm
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectIDX
    public FRectUseYN

    public Sub GetMyAlarmByLevel()
        dim sqlStr, addSql

		addSql = " where 1 = 1 "
		if (FRectUseYN <> "") then
			addSql = addSql + " and useYN = 'Y' "
		end if

        sqlStr = "select count(levelAlarmIdx) as cnt from db_my10x10.dbo.tbl_myAlarm_by_level "
		sqlStr = sqlStr + addSql
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * from db_my10x10.dbo.tbl_myAlarm_by_level "
		sqlStr = sqlStr + addSql
        sqlStr = sqlStr + " order by levelAlarmIdx desc"

		''rw sqlStr

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
				set FItemList(i) = new CMyAlarmByLevelItem

				FItemList(i).FlevelAlarmIdx	= rsget("levelAlarmIdx")
				FItemList(i).Fyyyymmdd		= rsget("yyyymmdd")
				FItemList(i).Fmsgdiv		= rsget("msgdiv")
				FItemList(i).Ftitle			= db2html(rsget("title"))
				FItemList(i).Fsubtitle		= db2html(rsget("subtitle"))
				FItemList(i).Fcontents		= db2html(rsget("contents"))
				FItemList(i).Fuserlevel		= rsget("userlevel")
				FItemList(i).FwwwTargetURL	= rsget("wwwTargetURL")
				FItemList(i).FopenYN		= rsget("openYN")
				FItemList(i).FuseYN			= rsget("useYN")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Freguserid		= rsget("reguserid")
				FItemList(i).Flastupdate	= rsget("lastupdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

    public Sub GetMyAlarmByLevelOne()
        dim sqlStr
        sqlStr = "select top 1 a.* "
        sqlStr = sqlStr + " from db_my10x10.dbo.tbl_myAlarm_by_level a"
        sqlStr = sqlStr + " where levelAlarmIdx = " + CStr(FRectIdx)

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CMyAlarmByLevelItem

        if Not rsget.Eof then
			FOneItem.FlevelAlarmIdx	= rsget("levelAlarmIdx")
			FOneItem.Fyyyymmdd		= rsget("yyyymmdd")
			FOneItem.Fmsgdiv		= rsget("msgdiv")
			FOneItem.Ftitle			= db2html(rsget("title"))
			FOneItem.Fsubtitle		= db2html(rsget("subtitle"))
			FOneItem.Fcontents		= db2html(rsget("contents"))
			FOneItem.Fuserlevel		= rsget("userlevel")
			FOneItem.FwwwTargetURL	= rsget("wwwTargetURL")
			FOneItem.FopenYN		= rsget("openYN")
			FOneItem.FuseYN			= rsget("useYN")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.Freguserid		= rsget("reguserid")
			FOneItem.Flastupdate	= rsget("lastupdate")
        end if
        rsget.Close
    end Sub

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
		HasPreScroll = StarScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>

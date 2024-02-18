<%

 IF application("Svr_Info")="Dev" THEN
 	staticImgUrl	= "http://testimgstatic.10x10.co.kr"	'Å×½ºÆ®
 	uploadUrl		= "http://testimgstatic.10x10.co.kr"
 	wwwUrl			= "http://test.10x10.co.kr"
 ELSE
 	staticImgUrl	= "http://imgstatic.10x10.co.kr"	
 	uploadUrl		= "http://imgstatic.10x10.co.kr"
 	wwwUrl			= "http://www.10x10.co.kr"
 END IF
 
class CMainMdChoiceRotateItem
	public Fidx
	public Fphotoimg
	public Flinkinfo
	public Fisusing
	public Fregdate
	public FDispOrder
	public Flinkitemid
	
	public FSellyn
	public FDispyn
	public FLimityn
	public FLimitno
	public FLimitsold
	
	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FDispyn="N") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CMainEventRotateItem
	public Fidx
	public Fphotoimg
	public Flinkinfo
	public Fisusing
	public Fregdate
	public FDispOrder

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMainEventRotate
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FrectMallType
	public FRectCDL
	public FRectIsusing
    

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub list()
		dim sql, i

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_contents].[dbo].tbl_main_rotate_event "
        sql = sql + " where 1=1"
        
		if FRectIsusing<>"" then
			sql = sql + " and isusing = '" + FRectIsusing + "'"
		end if
        
        
        
		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close


		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_contents].[dbo].tbl_main_rotate_event "
		sql = sql + " where 1=1"
        
		if FRectIsusing<>"" then
			sql = sql + " and isusing = '" + FRectIsusing + "'"
		end if
        
        

		sql = sql + " order by disporder , idx desc "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

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
				set FItemList(i) = new CMainEventRotateItem

				FItemList(i).Fidx          	= rsget("idx")
		        FItemList(i).Fphotoimg       = staticImgUrl&"/contents/othermall_maincontents/" + rsget("photoimg")
				FItemList(i).Flinkinfo   	= rsget("linkinfo")
				FItemList(i).Fisusing   	= rsget("isusing")
				FItemList(i).Fregdate      	= rsget("regdate")
				FItemList(i).FDispOrder		= rsget("disporder")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub read(byVal v)
		dim sql, i

		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_contents].[dbo].tbl_main_rotate_event "
		sql = sql + " where (idx = " + CStr(v) + ") "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FTotalCount = rsget.RecordCount
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
				set FItemList(i) = new CMainEventRotateItem

				FItemList(i).Fidx          = rsget("idx")
		        FItemList(i).Fphotoimg       = staticImgUrl&"/contents/othermall_maincontents/" + rsget("photoimg")
				FItemList(i).Flinkinfo   = rsget("linkinfo")
				FItemList(i).Fisusing   = rsget("isusing")
				FItemList(i).Fregdate      = rsget("regdate")
				FItemList(i).FDispOrder		= rsget("disporder")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

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


Class CMainMdChoiceRotate
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FrectMallType
	public FRectCDL
	public FRectIsusing
    public FRectItemId

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub list()
		dim sql, i

		sql = "select count(idx) as cnt "
		sql = sql + " from [db_contents].[dbo].tbl_othermall_main_mdchoice_flash "
        sql = sql + " where 1=1"
        
		if FRectIsusing<>"" then
			sql = sql + " and isusing = '" + FRectIsusing + "'"
		end if
        
        if FRectItemId<>"" then
            sql = sql + " and linkitemid=" + CStr(itemid)
        end if
        
		rsget.Open sql, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close

		sql = "select top " + CStr(FPageSize * FCurrPage) 
		sql = sql + " f.* "
		sql = sql + " , i.sellyn, i.limityn, i.limitno, i.limitsold" 
		sql = sql + " from [db_contents].[dbo].tbl_othermall_main_mdchoice_flash f"
		sql = sql + " left join [db_item].dbo.tbl_item i on f.linkitemid=i.itemid"
		sql = sql + " where 1=1"
        
		if FRectIsusing<>"" then
			sql = sql + " and f.isusing = '" + FRectIsusing + "'"
		end if
        
        if FRectItemId<>"" then
            sql = sql + " and f.linkitemid=" + CStr(itemid)
        end if
        
		sql = sql + " order by f.disporder ,f.idx desc "
		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

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
				set FItemList(i) = new CMainMdChoiceRotateItem

				FItemList(i).Fidx          	= rsget("idx")
		        FItemList(i).Fphotoimg       = staticImgUrl&"/contents/othermall_maincontents/" + rsget("photoimg")
				FItemList(i).Flinkinfo   	= rsget("linkinfo")
				FItemList(i).Fisusing   	= rsget("isusing")
				FItemList(i).Fregdate      	= rsget("regdate")
				FItemList(i).FDispOrder		= rsget("disporder")
				
				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

	public Sub read(byVal v)
		dim sql, i

		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_contents].[dbo].tbl_othermall_main_mdchoice_flash "
		sql = sql + " where (idx = " + CStr(v) + ") "

		rsget.pagesize = FPageSize
		rsget.Open sql, dbget, 1

		FTotalCount = rsget.RecordCount
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
				set FItemList(i) = new CMainMdChoiceRotateItem

				FItemList(i).Fidx          = rsget("idx")
		        FItemList(i).Fphotoimg       = staticImgUrl&"/contents/othermall_maincontents/" + rsget("photoimg")
				FItemList(i).Flinkinfo   = rsget("linkinfo")
				FItemList(i).Fisusing   = rsget("isusing")
				FItemList(i).Fregdate      = rsget("regdate")
				FItemList(i).FDispOrder		= rsget("disporder")
				FItemList(i).Flinkitemid	= rsget("linkitemid")
				
				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end sub

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
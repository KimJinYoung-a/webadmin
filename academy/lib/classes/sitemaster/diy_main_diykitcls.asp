<%

class CMainMdChoiceRotateItem
	public Fidx
	public Fphotoimg
	public Flinkinfo
	public Fisusing
	public Fregdate
	public FDispOrder
	public Flinkitemid
	
	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold
	public Fsmallimage
	
	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
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
		sql = sql + " from [db_academy].[dbo].tbl_diymain_diykit "
        sql = sql + " where 1=1"
        
		if FRectIsusing<>"" then
			sql = sql + " and isusing = '" + FRectIsusing + "'"
		end if
        
        if FRectItemId<>"" then
            sql = sql + " and linkitemid=" + CStr(itemid)
        end if
        
		rsACADEMYget.Open sql, dbACADEMYget, 1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close

		sql = "select top " + CStr(FPageSize * FCurrPage) 
		sql = sql + " f.* "
		sql = sql + " , i.sellyn, i.limityn, i.limitno, i.limitsold, i.smallimage" 
		sql = sql + " from [db_academy].[dbo].tbl_diymain_diykit f"
		sql = sql + " left join [db_academy].[dbo].tbl_diy_item i on f.linkitemid=i.itemid"
		sql = sql + " where 1=1"
        
		if FRectIsusing<>"" then
			sql = sql + " and f.isusing = '" + FRectIsusing + "'"
		end if
        
        if FRectItemId<>"" then
            sql = sql + " and f.linkitemid=" + CStr(itemid)
        end if
        
		sql = sql + " order by f.disporder ,f.idx desc "
		rsACADEMYget.pagesize = FPageSize
		'response.Write sql
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
		        i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CMainMdChoiceRotateItem

				FItemList(i).Fidx          	= rsACADEMYget("idx")
		        'FItemList(i).Fphotoimg       = staticImgUrl & "/contents/maincontents/" + rsACADEMYget("photoimg")
				FItemList(i).Flinkinfo   	= rsACADEMYget("linkinfo")
				FItemList(i).Fisusing   	= rsACADEMYget("isusing")
				FItemList(i).Fregdate      	= rsACADEMYget("regdate")
				FItemList(i).FDispOrder		= rsACADEMYget("disporder")
				
				FItemList(i).FSellyn		= rsACADEMYget("sellyn")
				FItemList(i).Flimityn		= rsACADEMYget("limityn")
				FItemList(i).Flimitno		= rsACADEMYget("limitno")
				FItemList(i).Flimitsold		= rsACADEMYget("limitsold")
				FItemList(i).Flinkitemid	= rsACADEMYget("linkitemid")
				FItemList(i).Fsmallimage	= imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("linkitemid")) + "/"  + rsACADEMYget("smallimage")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	end sub

	public Sub read(byVal v)
		dim sql, i

		sql = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_academy].[dbo].tbl_diymain_diykit "
		sql = sql + " where (idx = " + CStr(v) + ") "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FTotalCount = rsACADEMYget.RecordCount
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		if  not rsACADEMYget.EOF  then
		        i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CMainMdChoiceRotateItem

				FItemList(i).Fidx          = rsACADEMYget("idx")
		        'FItemList(i).Fphotoimg       = staticImgUrl & "/contents/maincontents/" + rsACADEMYget("photoimg")
				FItemList(i).Flinkinfo   = rsACADEMYget("linkinfo")
				FItemList(i).Fisusing   = rsACADEMYget("isusing")
				FItemList(i).Fregdate      = rsACADEMYget("regdate")
				FItemList(i).FDispOrder		= rsACADEMYget("disporder")
				FItemList(i).Flinkitemid	= rsACADEMYget("linkitemid")
				
				
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
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
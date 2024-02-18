<%

Class CMDChoiceItem

	public Fidx
	Public Fcdl
	Public Fcdm
	public Fitemid
	public Fisusing
	public Fcode_nm
	public Fmcode_nm
	public FitemName
	public FImageSmall

	public FSellyn
	public FLimityn
	public FLimitno
	public FLimitsold

	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CMDChoice
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Public FRectCDL
	public FRectCDM
	Public FRectIsUsing
	public FRectStyleSerail

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
	end function

	public Function GetMDChoiceList()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf		
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice as c "+ vbcrlf
		sqlStr = sqlStr + " 	inner join [db_item].[dbo].tbl_item as i on c.itemid = i.itemid "+ vbcrlf	 	
		sqlStr = sqlStr + " WHERE c.cdl = '" + FRectCDL + "'" + vbcrlf
		
		if FRectCDM <> "" then
			sqlStr = sqlStr + " and c.cdm = '" + FRectCDM + "'" + vbcrlf
		else
			sqlStr = sqlStr + " and c.cdm <> 0 " + vbcrlf
		end if	
		
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.cdl, c.cdm, c.itemid, c.isusing, i.itemname, i.smallimage" + vbcrlf
		sqlStr = sqlStr + " ,l.code_nm as code_nm , m.code_nm as mcode_nm " + vbcrlf
		sqlStr = sqlStr + " ,i.sellyn, i.limityn, i.limitno, i.limitsold " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_MDChoice as c "+ vbcrlf
		sqlStr = sqlStr + " inner join [db_item].[dbo].tbl_item as i on c.itemid = i.itemid "+ vbcrlf
	 	sqlStr = sqlStr + "		inner join [db_item].[dbo].tbl_cate_large as l on l.code_large= c.cdl "+ vbcrlf
		sqlStr = sqlStr + "		inner join [db_item].[dbo].tbl_cate_mid as m on  m.code_mid =c.cdm  and m.code_large = c.cdl "+ vbcrlf		
		sqlStr = sqlStr + " where  c.cdl = '" + FRectCDL + "'" + vbcrlf
		
		if FRectCDM <> "" then
			sqlStr = sqlStr + " and c.cdm = '" + FRectCDM + "'" + vbcrlf
		else
			sqlStr = sqlStr + " and c.cdm <> 0 " + vbcrlf	
		end if	
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		sqlStr = sqlStr + " order by c.idx desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMDChoiceItem

				FItemList(i).Fidx		= rsget("idx")
				FItemList(i).Fcdl		= rsget("cdl")
				FItemList(i).Fcdm		= rsget("cdm")
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fmcode_nm	= rsget("mcode_nm")
				FItemList(i).Fitemid	= rsget("itemid")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).FitemName	= db2html(rsget("itemname"))
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

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
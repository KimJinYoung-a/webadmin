<%
'###########################################################
' Description :  brandstreet
' History : 2009.03.24 한용민 생성
'###########################################################
%>
<% 
Class cbrandstreet_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public ftype
	public fmakerid
	public fitemid
	public fregdate
	public fisusing
	public fitemname
	public fSellYn
	public fLimitYn
	public fLimitNo
	public fLimitSold
	public fdanjongyn
	public fsellcash
	public fbuycash
	public fmainimage
	public flistimage
	public fbasicimage
	public fsmallimage
	
end class

class cbrandstreet_list

	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem

	public FRectIsusing
	public frecttype
	public frectitemid
	public frectmakerid
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	'//admin/designer/brandstreet/brandstreet.asp 
	public sub fcontents_list			
		dim sql , i 

		sql = "select "
		sql = sql & " count(a.idx) as cnt "
		sql = sql & " from db_brand.dbo.tbl_upche_brandstreet a "
		sql = sql & " join [db_item].[dbo].tbl_item d "
		sql = sql & " on a.itemid = d.itemid "
		sql = sql & " where d.isusing ='Y' and d.sellyn = 'Y' "		
			
		if frectmakerid =  "" then
			response.write "<script>"
			response.write "alert('브랜디아이디가 없습니다. 시스템팀에 문의하세요.');"
			response.write "history.go(-1);"
			response.write "</script>"
			dbget.close()	:	response.End			
		else
			sql = sql & " and d.makerid = '"&frectmakerid&"'"				
		end if
		if frectisusing = "Y" then
			sql = sql & " and a.isusing ='Y' "			
		elseif frectisusing = "N" then
			sql = sql & " and a.isusing ='N' "			
		end if
		if frecttype = "1" then
			sql = sql & " and a.type =1 "	
		end if
		
		'response.write sqlcount&"<br>"
		rsget.open sql,dbget,1
		FTotalCount = rsget("cnt")			
		rsget.close

		sql = "select  top "& FPageSize*FCurrpage&"" 
		sql = sql & " a.idx, a.type, a.makerid, a.itemid, a.regdate, a.isusing"
		sql = sql & " , d.itemname, d.SellYn, d.LimitYn, d.LimitNo, d.LimitSold"
		sql = sql & " ,d.danjongyn, d.sellcash, d.buycash , d.mainimage, d.listimage"
		sql = sql & " ,d.basicimage ,d.smallimage"
		sql = sql & " from db_brand.dbo.tbl_upche_brandstreet a"
		sql = sql & " join [db_item].[dbo].tbl_item d"
		sql = sql & " on a.itemid = d.itemid"
		sql = sql & " where d.isusing ='Y' and d.sellyn = 'Y' "
		
		if frectmakerid =  "" then
			response.write "<script>"
			response.write "alert('브랜디아이디가 없습니다. 시스템팀에 문의하세요.');"
			response.write "history.go(-1);"
			response.write "</script>"
			dbget.close()	:	response.End			
		else
			sql = sql & " and d.makerid = '"&frectmakerid&"'"				
		end if
		if frectisusing = "Y" then
			sql = sql & " and a.isusing ='Y' "			
		elseif frectisusing = "N" then
			sql = sql & " and a.isusing ='N' "			
		end if
		if frecttype = "1" then
			sql = sql & " and a.type =1 "	
		end if
		
		sql = sql & " order by a.regdate desc"
		
		'response.write sql&"<br>"
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1
		
		FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1	
		redim FItemList(FResultCount)
		i = 0
		
			if not rsget.eof then				
				rsget.absolutepage = FCurrPage
				do until rsget.eof				
					set FItemList(i) = new cbrandstreet_item 		
						
						FItemList(i).fidx = rsget("idx")
						FItemList(i).ftype = rsget("type")
						FItemList(i).fmakerid = rsget("makerid")
						FItemList(i).fitemid = rsget("itemid")
						FItemList(i).fregdate = rsget("regdate")
						FItemList(i).fisusing = rsget("isusing")
						FItemList(i).fitemname = rsget("itemname")
						FItemList(i).fSellYn = rsget("SellYn")
						FItemList(i).fLimitYn = rsget("LimitYn")
						FItemList(i).fLimitNo = rsget("LimitNo")
						FItemList(i).fLimitSold = rsget("LimitSold")
						FItemList(i).fdanjongyn = rsget("danjongyn")
						FItemList(i).fsellcash = rsget("sellcash")
						FItemList(i).fbuycash = rsget("buycash")																								
						FItemList(i).fmainimage = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
						FItemList(i).flistimage = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
						FItemList(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")
						FItemList(i).fbasicimage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage")									
						rsget.movenext
					i = i+1
					
				loop
			end if
		rsget.close
	end Sub
 
	'//admin/designer/brandstreet/brandstreet_upcheitem.asp 
	public sub fupche_item			
		dim sql , i 

		sql = "select "
		sql = sql & " count(itemid) as cnt "
		sql = sql & " from [db_item].[dbo].tbl_item "
		sql = sql & " where isusing ='Y' and sellyn = 'Y' "			

		if frectmakerid =  "" then
			response.write "<script>"
			response.write "alert('브랜디아이디가 없습니다. 시스템팀에 문의하세요.');"
			response.write "self.close()"
			response.write "</script>"
			dbget.close()	:	response.End			
		else
			sql = sql & " and makerid = '"&frectmakerid&"'"				
		end if
		
		if frectitemid <> "" then
			sql = sql & " and itemid = "&frectitemid&""				
		end if
		
		'response.write sqlcount&"<br>"
		rsget.open sql,dbget,1
		FTotalCount = rsget("cnt")			
		rsget.close

		sql = "select  top "& FPageSize*FCurrpage&"" 
		sql = sql & " itemid, itemname, SellYn, LimitYn, LimitNo, LimitSold"
		sql = sql & " ,danjongyn, sellcash, buycash , mainimage, listimage"
		sql = sql & " ,basicimage ,smallimage"
		sql = sql & " from [db_item].[dbo].tbl_item "
		sql = sql & " where isusing ='Y' and sellyn = 'Y' "	

		if frectmakerid =  "" then
			response.write "<script>"
			response.write "alert('브랜디아이디가 없습니다. 시스템팀에 문의하세요.');"
			response.write "self.close()"
			response.write "</script>"
			dbget.close()	:	response.End			
		else
			sql = sql & " and makerid = '"&frectmakerid&"'"				
		end if
		
		if frectitemid <> "" then
			sql = sql & " and itemid = "&frectitemid&""				
		end if			
	
		sql = sql & " order by itemid desc"
		
		'response.write sql&"<br>"
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1
		
		FResultCount = rsget.RecordCount - (FPageSize*(FCurrPage-1))
		FTotalPage = CInt(FTotalCount\FPageSize) + 1	
		redim FItemList(FResultCount)
		i = 0
		
			if not rsget.eof then				
				rsget.absolutepage = FCurrPage
				do until rsget.eof				
					set FItemList(i) = new cbrandstreet_item 		

						FItemList(i).fitemid = rsget("itemid")						
						FItemList(i).fitemname = rsget("itemname")
						FItemList(i).fSellYn = rsget("SellYn")
						FItemList(i).fLimitYn = rsget("LimitYn")
						FItemList(i).fLimitNo = rsget("LimitNo")
						FItemList(i).fLimitSold = rsget("LimitSold")
						FItemList(i).fdanjongyn = rsget("danjongyn")
						FItemList(i).fsellcash = rsget("sellcash")
						FItemList(i).fbuycash = rsget("buycash")																								
						FItemList(i).fmainimage = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
						FItemList(i).flistimage = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
						FItemList(i).fsmallimage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")
						FItemList(i).fbasicimage = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage")									
						rsget.movenext
					i = i+1
					
				loop
			end if
		rsget.close
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

end class

%>
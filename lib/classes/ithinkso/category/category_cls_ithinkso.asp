<%
'###########################################################
' Description : 아이띵소 카테고리 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################

class ccategory_ithinkso_item
	public fitemid
	public fmakerid
	public Fitemname
	public fsmallimage
	public Flistimage
	public Flistimage120
	public fCateTypeSeq
	public fCateTypeName
	public fCateTypeOrder
	public fIsUsing
	public fRegdate
	public fCateSeq
	public fCateName
	public fDepth
	public fCateOrder
	public forgprice
	public fsailyn
	public fitemCouponYn
	public fsellyn
	public fsailprice
	public fitemCouponType
	public fCateSeq1
	public fCateSeq2
	public fCateSeq3
	public fCatename1
	public fCatename2
	public fCatename3
	public fCateDispSeq
	public fsubCateSeq1
	public fsubCateSeq2
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class ccategory_ithinkso
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	
	public frectisusing
	public frectCateTypeSeq
	public frectDepth
	public frectsubCateSeq1
	public frectsubCateSeq2
	public FRectMakerid
	public FRectSellYN
	public FRectItemid	
	public FRectItemName	
	public FRectCateSeq1
	public FRectCateSeq2
	public FRectCateSeq3
	public frectcountryCd
	
	'/admin/ithinkso/category/category_item_reg_ithinkso.asp
	public function getitemlist()
        dim sqlStr, addSql, i

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if
        
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if

		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item as i"
        
        if frectCateTypeSeq = "1" then
	        sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang as mi"
	        sqlStr = sqlStr & " 	on i.itemid=mi.itemid"
	        sqlStr = sqlStr & " 	and mi.useyn='Y'"
	        sqlStr = sqlStr & " 	and mi.countryCd='" & frectcountryCd & "'"
    	end if
    
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr & " 	on i.makerid=c.userid"
        sqlStr = sqlStr & " where 1=1 " & addSql

		'response.write sqlStr & "<br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
		
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.itemid,i.itemname as itemname, i.makerid, i.smallimage, i.listimage, i.listimage120"
        sqlStr = sqlStr & " ,i.orgprice, i.sailprice, i.sailyn, i.itemCouponYn, i.sellyn, i.isusing, i.itemCouponType"
        sqlStr = sqlStr & " , IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay"
        sqlStr = sqlStr & " , IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item as i"

        if frectCateTypeSeq = "1" then
	        sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang as mi"
	        sqlStr = sqlStr & " 	on i.itemid=mi.itemid"
	        sqlStr = sqlStr & " 	and mi.useyn='Y'"
	        sqlStr = sqlStr & " 	and mi.countryCd='" & frectcountryCd & "'"
    	end if
    	
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr & " 	on i.makerid=c.userid"                
        sqlStr = sqlStr & " where 1=1 " & addSql
       	sqlStr = sqlStr & " order by i.itemid desc"

		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
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
                set FItemList(i) = new ccategory_ithinkso_item
				
				FItemList(i).fitemCouponType            = rsget("itemCouponType")
				FItemList(i).fisusing            = rsget("isusing")
				FItemList(i).fsellyn            = rsget("sellyn")
				FItemList(i).forgprice            = rsget("orgprice")
				FItemList(i).fsailprice            = rsget("sailprice")
				FItemList(i).fitemCouponYn            = rsget("itemCouponYn")
				FItemList(i).fsailyn            = rsget("sailyn")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
	'/admin/ithinkso/category/category_item_ithinkso.asp
	public function getCategoryitem()
        dim sqlStr, addSql, i

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if
        
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if frectCateTypeSeq<>"" then
            addSql = addSql + " and ci.CateTypeSeq='" + frectCateTypeSeq + "'"
        end if

        if FRectCateSeq1<>"" then
            addSql = addSql + " and ci.CateSeq1='" + FRectCateSeq1 + "'"
        end if

        if FRectCateSeq2<>"" then
            addSql = addSql + " and ci.CateSeq2='" + FRectCateSeq2 + "'"
        end if

        if FRectCateSeq3<>"" then
            addSql = addSql + " and ci.CateSeq3='" + FRectCateSeq3 + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and ci.isusing='" + FRectIsUsing + "'"
        end if

		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_item.dbo.tbl_ithinkso_Categoryitem ci"
        sqlStr = sqlStr & " join [db_item].[dbo].tbl_item as i"
        sqlStr = sqlStr & " 	on ci.itemid=i.itemid"   
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr & " 	on i.makerid=c.userid"
        sqlStr = sqlStr & " where 1=1 " & addSql

		'response.write sqlStr & "<br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " ci.CateDispSeq, ci.CateTypeSeq, ci.Itemid, ci.CateSeq1, ci.CateSeq2, ci.CateSeq3, ci.IsUsing"
        sqlStr = sqlStr & " ,i.itemid,i.itemname as itemname, i.makerid, i.smallimage, i.listimage, i.listimage120"
        sqlStr = sqlStr & " ,i.orgprice, i.sailprice, i.sailyn, i.itemCouponYn, i.sellyn, i.itemCouponType"
        sqlStr = sqlStr & " , IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit, IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay"
        sqlStr = sqlStr & " , IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " ,( select CateTypeName from db_item.dbo.tbl_ithinkso_CategoryType"
        sqlStr = sqlStr & " 	where isusing='Y' and CateTypeSeq = ci.CateTypeSeq) as CateTypename"        
        sqlStr = sqlStr & " ,( select CateName from db_item.dbo.tbl_ithinkso_CategoryInfo"
        sqlStr = sqlStr & " 	where isusing='Y' and CateSeq = ci.CateSeq1) as Catename1"
        sqlStr = sqlStr & " ,( select CateName from db_item.dbo.tbl_ithinkso_CategoryInfo"
        sqlStr = sqlStr & " 	where isusing='Y' and CateSeq = ci.CateSeq2) as Catename2"
        sqlStr = sqlStr & " ,( select CateName from db_item.dbo.tbl_ithinkso_CategoryInfo"
        sqlStr = sqlStr & " 	where isusing='Y' and CateSeq = ci.CateSeq3) as Catename3"
        sqlStr = sqlStr & " from db_item.dbo.tbl_ithinkso_Categoryitem ci"
        sqlStr = sqlStr & " join [db_item].[dbo].tbl_item as i"
        sqlStr = sqlStr & " 	on ci.itemid=i.itemid"   
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr & " 	on i.makerid=c.userid"  
        sqlStr = sqlStr & " where 1=1 " & addSql
       	sqlStr = sqlStr & " order by ci.CateDispSeq desc"
		
		'response.write sqlStr & "<br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
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
                set FItemList(i) = new ccategory_ithinkso_item
				
				FItemList(i).fCateDispSeq        = rsget("CateDispSeq")
                FItemList(i).fCateTypename        = db2html(rsget("CateTypename"))
                FItemList(i).fCatename1        = db2html(rsget("Catename1"))
                FItemList(i).fCatename2        = db2html(rsget("Catename2"))
                FItemList(i).fCatename3        = db2html(rsget("Catename3"))                                               
                FItemList(i).fCateTypeSeq        = rsget("CateTypeSeq")
                FItemList(i).fCateSeq1        = rsget("CateSeq1")
                FItemList(i).fCateSeq2          = rsget("CateSeq2")
                FItemList(i).fCateSeq3        = rsget("CateSeq3")
				FItemList(i).fitemCouponType            = rsget("itemCouponType")
				FItemList(i).fisusing            = rsget("isusing")
				FItemList(i).fsellyn            = rsget("sellyn")
				FItemList(i).forgprice            = rsget("orgprice")
				FItemList(i).fsailprice            = rsget("sailprice")
				FItemList(i).fitemCouponYn            = rsget("itemCouponYn")
				FItemList(i).fsailyn            = rsget("sailyn")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsmallimage        = webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = webImgUrl & "/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = webImgUrl & "/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
	'///admin/ithinkso/category/category_ithinkso.asp
	public sub getCategory_notpaging()
		dim sqlStr,i , sqlsearch
		
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'"
		end if
		
		if frectDepth <> "" then
			sqlsearch = sqlsearch & " and Depth = '"&frectDepth&"'"
		end if
		
		if frectCateTypeSeq <> "" then
			sqlsearch = sqlsearch & " and CateTypeSeq = '"&frectCateTypeSeq&"'"	
		end if
		
		'/대카테
		if frectsubCateSeq1 = "" and frectsubCateSeq2 = "" then
			sqlsearch = sqlsearch & " and subCateSeq1=0 and subCateSeq2=0"
		else
			'/중카테
			if frectsubCateSeq1 <> "" and frectsubCateSeq2 = "" then
				sqlsearch = sqlsearch & " and subCateSeq1 = '"&frectsubCateSeq1&"' and subCateSeq2=0"	
			end if
			'/소카테
			if frectsubCateSeq1 <> "" and frectsubCateSeq2 <> "" then
				sqlsearch = sqlsearch & " and subCateSeq1 = '"&frectsubCateSeq1&"' and subCateSeq2 = '"&frectsubCateSeq2&"'"	
			end if
		end if
		
		sqlStr = "select"
		sqlStr = sqlStr & " CateSeq, CateTypeSeq, subCateSeq1, subCateSeq2, CateName, Depth, CateOrder, IsUsing, Regdate"
		sqlStr = sqlStr & " from db_item.dbo.tbl_ithinkso_CategoryInfo"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by CateOrder asc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount
		FTotalCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ccategory_ithinkso_item
				
				FItemList(i).fCateSeq = rsget("CateSeq")
				FItemList(i).fsubCateSeq1 = rsget("subCateSeq1")
				FItemList(i).fsubCateSeq2 = rsget("subCateSeq2")
				FItemList(i).fCateTypeSeq = rsget("CateTypeSeq")
				FItemList(i).fCateName = db2html(rsget("CateName"))
				FItemList(i).fDepth = rsget("Depth")
				FItemList(i).fCateOrder = rsget("CateOrder")
				FItemList(i).fIsUsing = rsget("IsUsing")
				FItemList(i).fRegdate = rsget("Regdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	

	'///admin/ithinkso/category/category_ithinkso.asp
	public sub getCategoryType_notpaging()
		dim sqlStr,i , sqlsearch
		
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'"
		end if
		
		sqlStr = "select"
		sqlStr = sqlStr & " CateTypeSeq, CateTypeName, CateTypeOrder, IsUsing, Regdate"
		sqlStr = sqlStr & " from db_item.dbo.tbl_ithinkso_CategoryType"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by CateTypeOrder asc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount
		FTotalCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new ccategory_ithinkso_item

				FItemList(i).fCateTypeSeq = rsget("CateTypeSeq")
				FItemList(i).fCateTypeName = db2html(rsget("CateTypeName"))
				FItemList(i).fCateTypeOrder = rsget("CateTypeOrder")
				FItemList(i).fIsUsing = rsget("IsUsing")
				FItemList(i).fRegdate = rsget("Regdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
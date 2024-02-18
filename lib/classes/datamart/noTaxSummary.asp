<%
'###########################################################
' Description :  면세 매출 클래스
' History : 2011.06.02 eastone 생성
'			2012.07.11 한용민 수정
'###########################################################

'tbl_noTax_placeName
'tbl_DM_PlaceCommCD

function drawBoxNormalPlaceGubun(boxName,boxval, dtGubun)
    Dim retVal, sqlStr
    Dim irows
    sqlStr = "select top 1000 placegubun,placesub,placename from db_datamart.dbo.tbl_DM_PlaceCommCD"
    sqlStr = sqlStr & " where dtGubun='"&dtGubun&"'"
    sqlStr = sqlStr & " order by orderSEq"
    
    db3_rsget.open sqlStr,db3_dbget,1
    if not db3_rsget.Eof then
        irows = db3_rsget.getRows()
    end if
    db3_rsget.close
    
    Dim i,cnt : cnt =  UBound(irows,2)
    
    retVal = "<select name='"&boxName&"'>"
    retVal = retVal & "<option value=''>전체"
    FOR i=0 to cnt
        retVal = retVal & "<option value='"&irows(0,i)&irows(1,i)&"' "&CHKIIF(boxval=irows(0,i)&irows(1,i),"selected","")&">"&irows(2,i)
    Next
    retVal = retVal & "</select>"
    
    response.write retVal
end function

function drawBoxNotaxPlaceGubun(boxName,boxval)
   call drawBoxNormalPlaceGubun(boxName,boxval,"NOTAX")
end function

function drawBoxMileSellPlaceGubun(boxName,boxval)
   call drawBoxNormalPlaceGubun(boxName,boxval,"SELML")
end function

Class CMonthMileItem
    public FYYYYMM       
    public FplaceGubun   
    public FplaceSub     
    public FspendMile    
    public FGainMile     
                  
    public FplaceSubName 

    function getPlaceSubName()
        IF Not IsNULL(FplaceSubName) then
            getPlaceSubName = FplaceSubName
            Exit function
        end if
    end function
    
    
    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CNoTaxItem
    public FYYYYMM
    public FplaceGubun
    public FplaceSub
    public FOrderNo
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fmakerid
    public Fitemname
    public Fitemoptionname
    public FnotaxPrice
    public Fitemno
    public FplaceSubName
	public fsitename
	
	public FCurVatinclude
	public FOffimgSmall
    public FimageSmall

    function getPlaceName()
        IF (FplaceGubun="TENON") then
            getPlaceName = "온라인"
        ELSEIF (FplaceGubun="TENOF") then
            getPlaceName = "오프라인"
        ELSE
            getPlaceName = FplaceGubun
        END IF
    end function
    
    function getPlaceSubName()
        IF Not IsNULL(FplaceSubName) then
            getPlaceSubName = FplaceSubName
            Exit function
        end if
    end function
    
    
    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CNoTaxList
    public FItemList()
	public FOneItem
	
    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	
	public FRectMakerid
	public FRectYYYYMM
	public FRectStYYYYMM
	public FRectEdYYYYMM
	public FRectplaceGubun
    public FRectplaceSub
    public FRectSellSite
    public FRectGrpSum
    
    public sub getMileSellListMonth
        Dim sqlStr
        
		
		sqlStr = "SELECT "
		sqlStr = sqlStr & " A.YYYYMM"
        sqlStr = sqlStr & " ,A.placeGubun"
        sqlStr = sqlStr & " ,A.placeSub"
        sqlStr = sqlStr & " ,IsNULL(D.spendMile,0) as spendMile"
        sqlStr = sqlStr & " ,IsNULL(D.GainMile,0) as GainMile"
		sqlStr = sqlStr & " ,A.placeName as placeSubName "
		sqlStr = sqlStr & " from (select Distinct S.YYYYMM, D.placegubun, D.placeSub, D.placeName, D.orderseq"
        sqlStr = sqlStr & "     from db_datamart.dbo.tbl_DM_PlaceCommCD D"
	    sqlStr = sqlStr & "     Join db_datamart.dbo.tbl_mng_MileageSummary S"
	    sqlStr = sqlStr & "     on D.dtGubun='SELML'"
	    if (FRectStYYYYMM<>"") then
            sqlStr = sqlStr & " and S.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'" 
        end if
        sqlStr = sqlStr & "     ) A"
		sqlStr = sqlStr & "     left join db_datamart.dbo.tbl_mng_MileageSummary D"
		sqlStr = sqlStr & "     on A.YYYYMM=D.YYYYMM"
		sqlStr = sqlStr & "     and A.placeGubun=D.placeGubun"
		sqlStr = sqlStr & "     and A.placeSub=D.placeSub"
		if (FRectStYYYYMM<>"") then
            sqlStr = sqlStr & " and D.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'" 
        end if
        if (FRectplaceGubun<>"") then
            sqlStr = sqlStr & " and D.placeGubun='"&FRectplaceGubun&"'"
        end if
        if (FRectplaceSub<>"") then
            sqlStr = sqlStr & " and D.placeSub='"&FRectplaceSub&"'"
        end if
        sqlStr = sqlStr & " where 1=1"
        
        sqlStr = sqlStr & " order by A.YYYYMM desc, A.orderSeq, A.placeSub"
''rw  sqlStr      
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount
		FTotalCount  = FResultCount
        if (FResultCount<1) then FResultCount=0
        
	    redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CMonthMileItem
				FItemList(i).FYYYYMM          = db3_rsget("YYYYMM")
                FItemList(i).FplaceGubun      = db3_rsget("placeGubun")
                FItemList(i).FplaceSub        = db3_rsget("placeSub")
                FItemList(i).FspendMile          = db3_rsget("spendMile")
                FItemList(i).FGainMile      = db3_rsget("GainMile")
                                 
                FItemList(i).FplaceSubName    = db3_rsget("placeSubName")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end Sub

	'//admin/datamart/mng/popmonthNoTaxDetail.asp
    public sub getMonthNoTaxDetailGroup
        Dim sqlStr , sqlsearch
        
        if (FRectMakerid<>"") then
            sqlsearch = sqlsearch & " and D.makerid='"&FRectMakerid&"'" 
        end if
        
        if (FRectStYYYYMM<>"") then
            sqlsearch = sqlsearch & " and D.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'" 
        end if
        
        if (FRectplaceGubun<>"") then
            sqlsearch = sqlsearch & " and D.placeGubun='"&FRectplaceGubun&"'"
        end if
        
        if (FRectplaceSub<>"") then
            sqlsearch = sqlsearch & " and D.placeSub='"&FRectplaceSub&"'"
        end if

        if (FRectSellSite<>"") then
            sqlsearch = sqlsearch & " and D.sitename='"&FRectSellSite&"'"
        end if
        
        sqlStr = "select count(*) as CNT"
        sqlStr = sqlStr & " From (select itemid"
        sqlStr = sqlStr & " from db_datamart.dbo.tbl_NoTax_Detail d"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " group by  placeGubun, placeSub, itemgubun, itemid, itemoption, makerid, itemname ,itemoptionname "
        if (FRectGrpSum<>"on") then 
            sqlStr = sqlStr & " ,YYYYMM,d.sitename"
        end if
        sqlStr = sqlStr & " ) T"
        
        'response.write sqlStr &"<Br>"
        db3_rsget.open sqlStr,db3_dbget,1
        
		IF not db3_rsget.EOF THEN
			FTotalCount = db3_rsget("cnt")
		END IF
		
		db3_rsget.close
''rw 	sqlStr	
		
		sqlStr = "SELECT TOP "& (FPageSize * FCurrPage) 
		if (FRectGrpSum<>"on") then 
		    sqlStr = sqlStr & " YYYYMM,d.sitename,"
	    end if
        sqlStr = sqlStr & " D.placeGubun"
        sqlStr = sqlStr & " ,D.placeSub"
        sqlStr = sqlStr & " ,D.itemgubun"
        sqlStr = sqlStr & " ,D.itemid"
        sqlStr = sqlStr & " ,D.itemoption"
        sqlStr = sqlStr & " ,D.makerid"
        sqlStr = sqlStr & " ,D.itemname"
        sqlStr = sqlStr & " ,D.itemoptionname "
        sqlStr = sqlStr & " ,isNULL(i.vatinclude,si.vatinclude) as vatinclude"
        sqlStr = sqlStr & " ,sum(D.notaxPrice*D.itemno) as notaxSum"
        sqlStr = sqlStr & " ,sum(D.itemno) as cnt"
		sqlStr = sqlStr & " ,N.placeName as placeSubName "
		sqlStr = sqlStr & " ,i.smallimage, si.offimgsmall"
		sqlStr = sqlStr & " from db_datamart.dbo.tbl_NoTax_Detail D"
		sqlStr = sqlStr & " left Join db_item.dbo.tbl_item i"
		sqlStr = sqlStr & "     on D.itemgubun='10' and D.itemid=i.itemid"
		sqlStr = sqlStr & " left Join db_shop.dbo.tbl_shop_item si"
		sqlStr = sqlStr & "     on D.itemgubun=si.itemgubun and D.itemid=si.shopitemid and D.itemoption=si.itemoption"
		sqlStr = sqlStr & " left Join  db_datamart.dbo.tbl_DM_PlaceCommCD N"
		sqlStr = sqlStr & "     on dtGubun='NOTAX'"
		sqlStr = sqlStr & "     and D.placeGubun=N.placeGubun"
		sqlStr = sqlStr & "     and D.placeSub=N.placeSub"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " group by  D.placeGubun, D.placeSub, D.itemgubun, D.itemid, D.itemoption, D.makerid, D.itemname ,D.itemoptionname, N.placeName, N.orderSeq,isNULL(i.vatinclude,si.vatinclude),i.smallimage, si.offimgsmall"
        if (FRectGrpSum<>"on") then 
            sqlStr = sqlStr & " ,D.YYYYMM,d.sitename"
            sqlStr = sqlStr & " order by D.YYYYMM, N.orderSeq, D.itemgubun,D.itemid ,D.itemoption"
        else
            sqlStr = sqlStr & " order by N.orderSeq, D.itemgubun,D.itemid ,D.itemoption"
        end if
        
        
       'response.write sqlStr &"<Br>"
        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CNoTaxItem
				FItemList(i).Fitemid          = db3_rsget("itemid")
				if (FRectGrpSum<>"on") then 
				    FItemList(i).FYYYYMM          = db3_rsget("YYYYMM")
				    FItemList(i).fsitename		  = db3_rsget("sitename")
			    end if
			    
                FItemList(i).FplaceGubun      = db3_rsget("placeGubun")
                FItemList(i).FplaceSub        = db3_rsget("placeSub")
                FItemList(i).Fitemgubun       = db3_rsget("itemgubun")
                FItemList(i).Fitemid          = db3_rsget("itemid")
                FItemList(i).Fitemoption      = db3_rsget("itemoption")
                FItemList(i).Fmakerid         = db3_rsget("makerid")
                FItemList(i).Fitemname        = db2HTML(db3_rsget("itemname"))
                FItemList(i).Fitemoptionname  = db2HTML(db3_rsget("itemoptionname"))
                FItemList(i).FnotaxPrice      = db3_rsget("notaxSum")
                FItemList(i).Fitemno          = db3_rsget("cnt")
                FItemList(i).FplaceSubName    = db3_rsget("placeSubName")
				
				FItemList(i).FCurVatinclude   = db3_rsget("vatinclude")
				
				FItemList(i).FOffimgSmall	= db3_rsget("offimgsmall")
				if FItemList(i).FOffimgSmall<>"" then 
				    FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FOffimgSmall
				end if
				
    			FItemList(i).FimageSmall     = db3_rsget("smallimage")
    			if FItemList(i).FimageSmall<>"" then
    				FItemList(i).FimageSmall     = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FimageSmall
    			end if
			
				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end sub

	public Sub getNoTaxListMonth
	    Dim sqlStr
	    
        sqlStr = " select A.YYYYMM,A.placegubun,A.placesub, A.placeName"
        sqlStr = sqlStr & " ,isNULL(B.TTLCNT,0) as TTLCNT, isNULL(B.TTLSUM,0) as TTLSUM"
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & " 	select distinct D.YYYYMM,N.placegubun,N.placesub, N.placeName,N.orderseq"
        sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_DM_PlaceCommCD N"
        sqlStr = sqlStr & " 		Join db_datamart.dbo.tbl_noTax_Detail D"
        sqlStr = sqlStr & " 		on N.dtGubun='NOTAX'"
        if (FRectStYYYYMM<>"") then
            sqlStr = sqlStr & " and D.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'" 
        end if
        if (FRectplaceGubun<>"") then
            sqlStr = sqlStr & " and D.placeGubun='"&FRectplaceGubun&"'"
        end if
        if (FRectplaceSub<>"") then
            sqlStr = sqlStr & " and D.placeSub='"&FRectplaceSub&"'"
        end if
        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and D.makerid='"&FRectMakerid&"'" 
        end if
        sqlStr = sqlStr & " ) A"
        sqlStr = sqlStr & " left  Join "
        sqlStr = sqlStr & " ("
        sqlStr = sqlStr & " 	select D.YYYYMM, D.placegubun,D.placesub"
        sqlStr = sqlStr & " 	, sum(D.itemno) as TTLCNT, sum(D.notaxprice*D.itemno) as TTLSUM"
        sqlStr = sqlStr & " 	from db_datamart.dbo.tbl_noTax_Detail D"
        sqlStr = sqlStr & " 	where 1=1" 
        if (FRectStYYYYMM<>"") then
            sqlStr = sqlStr & " and D.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'" 
        end if
        if (FRectplaceGubun<>"") then
            sqlStr = sqlStr & " and D.placeGubun='"&FRectplaceGubun&"'"
        end if
        if (FRectplaceSub<>"") then
            sqlStr = sqlStr & " and D.placeSub='"&FRectplaceSub&"'"
        end if
        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and D.makerid='"&FRectMakerid&"'" 
        end if
        sqlStr = sqlStr & " 	group by  D.YYYYMM, D.placegubun,D.placesub"
        sqlStr = sqlStr & " ) B"
        sqlStr = sqlStr & " on A.YYYYMM=B.YYYYMM"
        sqlStr = sqlStr & " and A.placegubun=B.placegubun"
        sqlStr = sqlStr & " and A.placesub=B.placesub"
        sqlStr = sqlStr & " order by A.YYYYMM, A.orderseq"
''
''
''	    sqlStr = "select D.YYYYMM, D.placeGubun, D.placeSub, N.placeName as placeSubName"
''	    sqlStr = sqlStr & " , sum(D.itemno) as CNT, sum(D.notaxPrice*D.itemno) as notaxSum"
''        sqlStr = sqlStr & " from db_datamart.dbo.tbl_NoTax_Detail D"
''        sqlStr = sqlStr & "     left Join db_datamart.dbo.tbl_noTax_placeName N"
''	    sqlStr = sqlStr & "     on D.placeGubun=N.placeGubun"
''	    sqlStr = sqlStr & "     and D.placesub=N.placesub"
''        sqlStr = sqlStr & " where 1=1"
''        if (FRectMakerid<>"") then
''            sqlStr = sqlStr & " and D.makerid='"&FRectMakerid&"'" 
''        end if
''        if (FRectStYYYYMM<>"") then
''            sqlStr = sqlStr & " and D.YYYYMM between '"&FRectStYYYYMM&"' and '"&FRectEdYYYYMM&"'" 
''        end if
''        if (FRectplaceGubun<>"") then
''            sqlStr = sqlStr & " and D.placeGubun='"&FRectplaceGubun&"'"
''        end if
''        if (FRectplaceSub<>"") then
''            sqlStr = sqlStr & " and D.placeSub='"&FRectplaceSub&"'"
''        end if
''        sqlStr = sqlStr & " group by D.YYYYMM, D.placeGubun, D.placeSub, N.placeName, N.orderseq"
''        sqlStr = sqlStr & " order by D.YYYYMM, N.orderseq"
''rw  sqlStr 
        db3_rsget.Open sqlStr,db3_dbget,1
		FResultCount = db3_rsget.RecordCount
		FTotalCount  = FResultCount
        if (FResultCount<1) then FResultCount=0
        
	    redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
			set FItemList(i) = new CNoTaxItem
			FItemList(i).FYYYYMM         = db3_rsget("YYYYMM")
            FItemList(i).FplaceGubun     = db3_rsget("placeGubun")
            FItemList(i).FplaceSub       = db3_rsget("placeSub")
            FItemList(i).FnotaxPrice     = db3_rsget("TTLSUM")
            FItemList(i).Fitemno         = db3_rsget("TTLCNT")
            FItemList(i).FplaceSubName   = db3_rsget("placeName")
			db3_rsget.MoveNext
			i = i + 1
		loop
		db3_rsget.close
		
    end sub

    Private Sub Class_Initialize()

		Redim FItemList(0)

		FCurrPage =1
		FPageSize = 20
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

End Class


%>
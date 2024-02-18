<%
'' [db_sitemaster].[dbo].tbl_category_contents_poscode
'' poscode, posname, linktype, fixtype, isusing


function DrawLinktypeCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select name='linktype' " + changeFlag + " class='select'>" + VbCrlf
    bufStr = bufStr + " <option value='' > 선택" + VbCrlf
    if selectedId="L" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='L' " + tmp_str + " >링크 (a href)" + VbCrlf
    if selectedId="M" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='M' " + tmp_str + " >맵   (#Map)" + VbCrlf
    bufStr = bufStr & " </select>" + VbCrlf
    
	response.write bufStr
end function


Class CCateContentsCodeItem
    public Fposcode
    public Fposname
    public FposVarname
    public Flinktype
    public Ffixtype
    public Fimagewidth
    public Fimageheight
    public FuseSet			'한페이지에 사용될 이미지수
    public Fisusing

    
    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case else
                getlinktypeName = Flinktype
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
end Class 

Class CCateContentsCode
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectPoscode
    

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

Class CCateContentsItem
    public Fidx
    public Fcdl
    public Fcdm
	public Fcds '2012-04-03 이종화 추가
    public Fcodename
    public Fcdmname
	public Fcdsname '2012-04-03 이종화 추가
    public Fposcode
    public FposVarname
    public Fposname
    public Flinktype
    public Ffixtype
    public Fimageurl
    public Flinkurl
    public Fimagewidth
    public Fimageheight
    public FuseSet
    public Fstartdate
    public Fenddate
    public Fregdate
    public Freguserid
    public Fisusing
    public FsortNo
    public Fdesc
	public Fitemseq ' 2012-04-04 이종화 추가
	public Flistimage ' 2012-04-04 이종화 추가
	public Fimg1 ' 2012-04-04 이종화 추가
	public Fimg2 ' 2012-04-04 이종화 추가
	public Fimg3 ' 2012-04-04 이종화 추가
    
    
    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fenddate)
    end function

    public function GetImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetImageUrl = ""
        else
            GetImageUrl = staticImgUrl & "/category/" + Fimageurl
        end if
    end function

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case else
                getlinktypeName = Flinktype
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CCateContents
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
       
    public FRectIdx
    public FRectIsusing
    public FRectPoscode
    public FRectfixtype
    public FRectValiddate
    public FRectCdl
    public FRectCdm
	public FRectCds '2012-04-03 추가 이종화
    
    public Sub GetOneCateContents()
        dim sqlStr
        sqlStr = "select top 1 c.* "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_hot c"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CCateContentsItem
        
        if Not rsget.Eof then
    
    		FOneItem.Fidx			= rsget("idx")
            FOneItem.Flinktype		= rsget("linktype")
            FOneItem.Fimageurl		= db2html(rsget("imageurl"))
            FOneItem.Flinkurl		= db2html(rsget("linkurl"))
            FOneItem.Fstartdate		= rsget("startdate")
            FOneItem.Fenddate		= rsget("enddate")
            FOneItem.Fregdate		= rsget("regdate")
            FOneItem.Fcdl			= rsget("cdl")
            FOneItem.Fcdm			= rsget("cdm")
		    FOneItem.Fcds			= rsget("cds")
            FOneItem.Fisusing		= rsget("isusing")

        end if
        rsget.Close
    end Sub

	'hot cate item 2012-04-04 이종화
	 public Sub GetOneCateiIemContents()
        dim sqlStr
        sqlStr = "select top 1 c.* "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_hot_item c"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CCateContentsItem
        
        if Not rsget.Eof then
    
    		FOneItem.Fidx			= rsget("idx")
            FOneItem.Fcdl			= rsget("cdl")
            FOneItem.Fcdm		= rsget("cdm")
			FOneItem.Fcds			= rsget("cds")
			FOneItem.Fitemseq	= rsget("itemseq")
            FOneItem.Fstartdate	= rsget("startdate")
            FOneItem.Fenddate	= rsget("enddate")
            FOneItem.Fregdate	= rsget("regdate")
            FOneItem.Fisusing	= rsget("isusing")

        end if
        rsget.Close
    end Sub

    public Sub GetCateContentsList()
        dim sqlStr, i, addStr
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

        if FRectIdx<>"" then
            addStr = addStr + " and c.idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            addStr = addStr + " and enddate>getdate()"
        end if
        
        if FRectfixtype<>"" then
            addStr = addStr + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if
        
        if FRectIsusing<>"" then
            addStr = addStr + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectCdl<>"" then
            addStr = addStr + " and c.cdl='" + CStr(FRectCdl) + "'"
        end if

        if FRectCdm<>"" then
            addStr = addStr + " and c.cdm='" + CStr(FRectCdm) + "'"
        end If
        
		if FRectCds<>"" Then '2012-04-03 이종화 추가
            addStr = addStr + " and c.cds='" + CStr(FRectCds) + "'"
        end if

        '// 전체 카운트 //
        sqlStr = " select count(c.idx) as cnt from [db_sitemaster].[dbo].tbl_category_hot as c "
        sqlStr = sqlStr + " where 1=1 " & addStr
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        '// 목록 접수 //
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, l.code_nm "
        sqlStr = sqlStr + " ,(select top 1 code_nm from [db_item].[dbo].tbl_cate_mid where code_large=c.cdl and code_mid=c.cdm ) as cdm_nm "
        sqlStr = sqlStr + " ,(select top 1 code_nm from [db_item].[dbo].tbl_cate_small where code_large=c.cdl and code_mid=c.cdm and code_small=c.cds ) as cds_nm  "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_hot c"
        sqlStr = sqlStr + " join [db_item].[dbo].tbl_cate_large l "
        sqlStr = sqlStr + " 	on c.cdl=l.code_large "
        sqlStr = sqlStr + " where 1=1" & addStr
        
        sqlStr = sqlStr + " order by c.idx desc"
        
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
				set FItemList(i) = new CCateContentsItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Fimageurl		= db2html(rsget("imageurl"))
                FItemList(i).Flinkurl		= db2html(rsget("linkurl"))
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Fcdl			= rsget("cdl")
                FItemList(i).Fcodename		= rsget("code_nm")
                FItemList(i).Fcdm			= rsget("cdm")
                FItemList(i).Fcdmname		= rsget("cdm_nm")
                FItemList(i).Fisusing		= rsget("isusing")
				FItemList(i).Fcds			= rsget("cds")
                FItemList(i).Fcdsname		= rsget("cds_nm")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	public Sub GetHotCateItemList()
        dim sqlStr, i, addStr
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

        if FRectIdx<>"" then
            addStr = addStr + " and c.idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            addStr = addStr + " and enddate>getdate()"
        end if
        
        if FRectfixtype<>"" then
            addStr = addStr + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if
        
        if FRectIsusing<>"" then
            addStr = addStr + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if

        if FRectCdl<>"" then
            addStr = addStr + " and c.cdl='" + CStr(FRectCdl) + "'"
        end if

        if FRectCdm<>"" then
            addStr = addStr + " and c.cdm='" + CStr(FRectCdm) + "'"
        end If
        
		 if FRectCds<>"" Then '2012-04-03 이종화 추가
            addStr = addStr + " and c.cds='" + CStr(FRectCds) + "'"
        end if

        '// 전체 카운트 //
        sqlStr = " select count(c.idx) as cnt from [db_sitemaster].[dbo].tbl_category_hot_item as c "
        sqlStr = sqlStr + " where 1=1 " & addStr
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        '// 목록 접수 //
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, l.code_nm "
        sqlStr = sqlStr + " ,(select top 1 code_nm from [db_item].[dbo].tbl_cate_mid where code_large=c.cdl and code_mid=c.cdm ) as cdm_nm "
        sqlStr = sqlStr + " ,(select top 1 code_nm from [db_item].[dbo].tbl_cate_small where code_large=c.cdl and code_mid=c.cdm and code_small=c.cds ) as cds_nm  "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_category_hot_item c"
        sqlStr = sqlStr + " join [db_item].[dbo].tbl_cate_large l "
        sqlStr = sqlStr + " 	on c.cdl=l.code_large "
        sqlStr = sqlStr + " where 1=1" & addStr
        
        sqlStr = sqlStr + " order by c.idx desc"

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
				set FItemList(i) = new CCateContentsItem

				FItemList(i).Fidx				= rsget("idx")
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate			= rsget("regdate")
                FItemList(i).Fcdl					= rsget("cdl")
                FItemList(i).Fcodename		= rsget("code_nm")
                FItemList(i).Fcdm				= rsget("cdm")
                FItemList(i).Fcdmname		= rsget("cdm_nm")
                FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fcds				= rsget("cds")
                FItemList(i).Fcdsname		= rsget("cds_nm")
				FItemList(i).Fitemseq			= rsget("itemseq")
				Dim arritemseq
				arritemseq	= Split(FItemList(i).Fitemseq,",")
				
				FItemList(i).Fimg1 = "http://webimage.10x10.co.kr/image/List/" & GetImageSubFolderByItemid(arritemseq(0)) & "/" & rsget("img1")
				FItemList(i).Fimg2 = "http://webimage.10x10.co.kr/image/List/" & GetImageSubFolderByItemid(arritemseq(1)) & "/" & rsget("img2")
				FItemList(i).Fimg3 = "http://webimage.10x10.co.kr/image/List/" & GetImageSubFolderByItemid(arritemseq(2)) & "/" & rsget("img3")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
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

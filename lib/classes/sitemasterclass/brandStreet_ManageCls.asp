<%
'' [db_sitemaster].[dbo].tbl_StreetMain_poscode
'' poscode, posname, linktype, fixtype, isusing

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select top 100 * from [db_sitemaster].[dbo].tbl_StreetMain_poscode where isusing='Y' order by poscode"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("poscode")&"' "&tmp_str&">" + db2html(rsget("posname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function


function DrawFixTypeCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select class='select' name='" + selectBoxName + "' " + changeFlag + " >" + VbCrlf
    bufStr = bufStr + " <option value=''> 선택" + VbCrlf
    if selectedId="K" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='K' " + tmp_str + " >관리자확정시" + VbCrlf
    if selectedId="R" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <option value='R' " + tmp_str + " >실시간" + VbCrlf
	if selectedId="D" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <option value='D' " + tmp_str + " >일별" + VbCrlf
	if selectedId="W" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <!-- <option value='W' " + tmp_str + " >주별 -->" + VbCrlf
	bufStr = bufStr + " </select>" + VbCrlf
	
	response.write bufStr
end function

function DrawLinktypeCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select class='select' name='linktype' " + changeFlag + " >" + VbCrlf
    bufStr = bufStr + " <option value='' > 선택" + VbCrlf
    if selectedId="L" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='L' " + tmp_str + " >링크 (a href)" + VbCrlf
    if selectedId="M" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='M' " + tmp_str + " >맵   (#Map)" + VbCrlf
    if selectedId="F" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='F' " + tmp_str + " >플래시" + VbCrlf
    bufStr = bufStr + " </select>" + VbCrlf
    
	response.write bufStr
end function


Class CStreetMainCodeItem
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
            case "F"
                getlinktypeName = "플래시"
            case else
                getlinktypeName = Flinktype
        end select
    end function
    
    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "관리자확정시"
            case "R"
                getfixtypeName = "실시간"
            case "D"
                getfixtypeName = "일별"
            case "W"
                getfixtypeName = "주별"
            case else
                getfixtypeName = Flinktype
        end select
    end function
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
end Class 

Class CStreetMainCode
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectPoscode
    
    public Sub GetOneContentsCode()
        dim SqlStr
        SqlStr = "select top 1 * "
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_StreetMain_poscode"
        SqlStr = SqlStr + " where poscode=" + CStr(FRectPoscode)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CStreetMainCodeItem
        if Not rsget.Eof then
            
            FOneItem.Fposcode		= rsget("poscode")
            FOneItem.Fposname		= db2html(rsget("posname"))
            FOneItem.FposVarname	= rsget("posVarname")
            FOneItem.Flinktype		= rsget("linktype")
            FOneItem.Ffixtype		= rsget("fixtype")
            FOneItem.Fimagewidth	= rsget("imagewidth")
            FOneItem.FuseSet		= rsget("useSet")
            FOneItem.Fisusing		= rsget("isusing")
            
            FOneItem.Fimageheight = rsget("imageheight")
        end if
        rsget.close
    end Sub
    
    public Sub GetposcodeList()
        dim sqlStr
        sqlStr = "select count(poscode) as cnt from [db_sitemaster].[dbo].tbl_StreetMain_poscode"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_StreetMain_poscode "
        sqlStr = sqlStr + " order by poscode desc"
        
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
				set FItemList(i) = new CStreetMainCodeItem

				FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).Fposname		= db2html(rsget("posname"))
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).FuseSet		= rsget("useSet")
                FItemList(i).Fisusing		= rsget("isusing")

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

Class CStreetMainContItem
    public Fidx
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
    
    
    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fenddate)
    end function

    public function GetImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetImageUrl = ""
        else
            GetImageUrl = uploadUrl & "/Street/" + Fimageurl
        end if
    end function

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "링크"
            case "M"
                getlinktypeName = "맵"
            case "F"
                getlinktypeName = "플래시"
            case else
                getlinktypeName = Flinktype
        end select
    end function
    
    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "관리자확정시"
            case "R"
                getfixtypeName = "실시간"
            case "D"
                getfixtypeName = "일별"
            case "W"
                getfixtypeName = "주별"
            case else
                getfixtypeName = Flinktype
        end select
    end function
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CStreetMainCont
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
    
    public Sub GetOneMainContents()
        dim sqlStr
        sqlStr = "select top 1 c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_StreetMain_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_StreetMain_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CStreetMainContItem
        
        if Not rsget.Eof then
    
    		FOneItem.Fidx			= rsget("idx")
            FOneItem.Fposcode		= rsget("poscode")
            FOneItem.Fposname		= db2html(rsget("posname"))
            FOneItem.FposVarname	= rsget("posVarname")
            FOneItem.Flinktype		= rsget("linktype")
            FOneItem.Ffixtype		= rsget("fixtype")
            FOneItem.Fimageurl		= db2html(rsget("imageurl"))
            FOneItem.Flinkurl		= db2html(rsget("linkurl"))
            FOneItem.Fimagewidth	= rsget("imagewidth")
            FOneItem.Fimageheight	= rsget("imageheight")
            FOneItem.FuseSet		= rsget("useSet")
            FOneItem.Fstartdate		= rsget("startdate")
            FOneItem.Fenddate		= rsget("enddate")
            FOneItem.Fregdate		= rsget("regdate")
            FOneItem.Freguserid		= rsget("reguserid")
            FOneItem.Fisusing		= rsget("isusing")
            
        end if
        rsget.Close
    end Sub

    public Sub GetMainContentsList()
        dim sqlStr, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)
        
        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_StreetMain_contents"
        sqlStr = sqlStr + " where 1=1"
        
        if FRectIdx<>"" then
            sqlStr = sqlStr + " and idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            sqlStr = sqlStr + " and enddate>getdate()"
        end if
        
        if FRectfixtype<>"" then
            sqlStr = sqlStr + " and fixtype='" + CStr(FRectfixtype) + "'"
        end if
        
        if FRectIsusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(FRectIsusing) + "'"
        end if
        
        if FRectPoscode<>"" then
            sqlStr = sqlStr + " and poscode='" + CStr(FRectPoscode) + "'"
        end if
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_StreetMain_contents c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_StreetMain_poscode p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where 1=1"
        if FRectIdx<>"" then
            sqlStr = sqlStr + " and c.idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            sqlStr = sqlStr + " and enddate>getdate()"
        end if
        
        if FRectfixtype<>"" then
            sqlStr = sqlStr + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if
        
        if FRectIsusing<>"" then
            sqlStr = sqlStr + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if
        
        if FRectPoscode<>"" then
            sqlStr = sqlStr + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if
        
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
				set FItemList(i) = new CStreetMainContItem

				FItemList(i).Fidx			= rsget("idx")
                FItemList(i).Fposcode		= rsget("poscode")
                FItemList(i).FposVarname	= rsget("posVarname")
                FItemList(i).Fposname		= db2html(rsget("posname"))
                FItemList(i).Flinktype		= rsget("linktype")
                FItemList(i).Ffixtype		= rsget("fixtype")
                FItemList(i).Fimageurl		= db2html(rsget("imageurl"))
                FItemList(i).Flinkurl		= db2html(rsget("linkurl"))
                FItemList(i).Fimagewidth	= rsget("imagewidth")
                FItemList(i).Fimageheight	= rsget("imageheight")
                FItemList(i).FuseSet		= rsget("useSet")
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fenddate		= rsget("enddate")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Freguserid		= rsget("reguserid")
                FItemList(i).Fisusing		= rsget("isusing")


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
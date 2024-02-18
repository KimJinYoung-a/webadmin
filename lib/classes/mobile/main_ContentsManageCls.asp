<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'           2009.12.21 ������ : ���ں� �÷��� ���� ��� �߰�
'###############################################

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >��ü</option>
   <%
   query1 = " select top 100 * from [db_sitemaster].[dbo].tbl_mobile_mainCont_code where isusing='Y' and Left(posname,5) <> 'POINT' order by poscode"
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


function DrawPoint1010PosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >��ü</option>
   <%
   query1 = " select top 100 * from [db_sitemaster].[dbo].tbl_mobile_mainCont_code where isusing='Y' and Left(posname,5) = 'POINT' order by poscode"
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
    bufStr = bufStr + " <option value=''> ����" + VbCrlf
    if selectedId="K" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='K' " + tmp_str + " >������Ȯ����" + VbCrlf
	if selectedId="R" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <option value='R' " + tmp_str + " >�ǽð�" + VbCrlf
	if selectedId="D" then tmp_str = "selected" else tmp_str = ""
	    bufStr = bufStr + " <option value='D' " + tmp_str + " >�Ϻ�" + VbCrlf
	bufStr = bufStr + " </select>" + VbCrlf
	
	response.write bufStr
end function

function DrawLinktypeCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select class='select' name='linktype' " + changeFlag + " >" + VbCrlf
    bufStr = bufStr + " <option value='' > ����" + VbCrlf
    if selectedId="L" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='L' " + tmp_str + " >��ũ (a href)" + VbCrlf
    if selectedId="M" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='M' " + tmp_str + " >��   (#Map)" + VbCrlf
    if selectedId="X" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='X' " + tmp_str + " >XML   " + VbCrlf
    if selectedId="J" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='J' " + tmp_str + " >JS   " + VbCrlf
    bufStr = bufStr + " </select>" + VbCrlf
    
	response.write bufStr
end function


Class CMainContentsCodeItem
    public Fposcode
    public Fposname
    public FposVarname
    public Flinktype
    public Ffixtype
    public Fimagewidth
    public Fimageheight
    public FuseSet			'���������� ���� �̹�����
    public Fisusing

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "��ũ"
            case "M"
                getlinktypeName = "��"
            case "X"
                getlinktypeName = "XML"
            case else
                getlinktypeName = Flinktype
        end select
    end function
    
    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "������Ȯ����"
			case "R"
                getfixtypeName = "�ǽð�"
            case "D"
                getfixtypeName = "�Ϻ�"
            case else
                getfixtypeName = Flinktype
        end select
    end function
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
	
end Class 

Class CMainContentsCode
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
        SqlStr = SqlStr + " from [db_sitemaster].[dbo].tbl_mobile_mainCont_code"
        SqlStr = SqlStr + " where poscode=" + CStr(FRectPoscode)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainContentsCodeItem
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
        sqlStr = "select count(poscode) as cnt from [db_sitemaster].[dbo].tbl_mobile_mainCont_code"
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " * from [db_sitemaster].[dbo].tbl_mobile_mainCont_code "
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
				set FItemList(i) = new CMainContentsCodeItem

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

Class CMainContentsItem
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
	public forderidx
	public fbackColor
	Public faltname
	Public Fordertext
	Public Fmakerid 'brand makerid
	Public Fmaincopy
	Public Fsubcopy
	Public Fcgubun
	Public Fculopt

	Public Fmaincopy2
    public Ftag_only
	public Ftag_gift
	public Ftag_plusone
	public Ftag_launching
	public Ftag_actively
	public Fsale_per
	public Fcoupon_per
	public Fevt_code
	public Fsalediv

    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fenddate)
    end function

    public function GetImageUrl()
        if (IsNULL(Fimageurl) or (Fimageurl="")) then
            GetImageUrl = ""
        else
            if Fcgubun="2" then

                if instr(Fimageurl,"webimage.10x10.co.kr/eventIMG/") > 0 then
                    GetImageUrl	= Fimageurl
                else
                    GetImageUrl =  staticImgUrl & "/mobile/" + Fimageurl
                end if
            else
                GetImageUrl =  staticImgUrl & "/mobile/" + Fimageurl
            end if
        end if
    end function

    public function getlinktypeName()
        select case Flinktype
            case "L"
                getlinktypeName = "��ũ"
            case "M"
                getlinktypeName = "��"
            case "X"
                getlinktypeName = "XML"
            case else
                getlinktypeName = Flinktype
        end select
    end function
    
    public function getfixtypeName()
        select case Ffixtype
            case "K"
                getfixtypeName = "������Ȯ����"
			case "R"
                getfixtypeName = "�ǽð�"
            case "D"
                getfixtypeName = "�Ϻ�"
            case else
                getfixtypeName = Flinktype
        end select
    end function
    
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMainContents
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
    public FRectSelDate
    public FRectSelDateTime
	public Flinktype
	public frectorderidx
	Public FRectsedatechk
	    
    public Sub GetOneMainContents()
        dim sqlStr
        sqlStr = "select top 1 c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_mobile_mainCont c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_mobile_mainCont_code p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where idx=" + CStr(FRectIdx)
        
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainContentsItem
        
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
			FOneItem.forderidx		= rsget("orderidx")
			FOneItem.fbackColor		= rsget("backColor") 
			FOneItem.faltname		= rsget("altname")
			FOneItem.Fordertext		= db2html(rsget("ordertext"))
			FOneItem.Fmakerid		= rsget("makerid")
			FOneItem.Fmaincopy		= rsget("maincopy")
			FOneItem.Fsubcopy		= rsget("subcopy")
			FOneItem.Fcgubun		= rsget("cgubun")
			FOneItem.Fculopt		= rsget("culopt")

			FOneItem.Fmaincopy2		= rsget("maincopy2")
            FOneItem.Ftag_only		= rsget("tag_only")            
			FOneItem.Ftag_gift		= rsget("tag_gift")
			FOneItem.Ftag_plusone	= rsget("tag_plusone")
			FOneItem.Ftag_launching	= rsget("tag_launching")
			FOneItem.Ftag_actively	= rsget("tag_actively")
			FOneItem.Fsale_per		= rsget("sale_per")
			FOneItem.Fcoupon_per	= rsget("coupon_per")
			FOneItem.Fevt_code	= rsget("evt_code")
			FOneItem.Fsalediv	= rsget("salediv")
        end if
        rsget.Close
    end Sub

    public Sub GetMainContentsList()
        dim sqlStr, addSql, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)

        if FRectIdx<>"" then
            addSql = addSql + " and c.idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            addSql = addSql + " and c.enddate>getdate()"
        end if
        
        if FRectfixtype<>"" then
            addSql = addSql + " and c.fixtype='" + CStr(FRectfixtype) + "'"
        end if
        
        if FRectIsusing<>"" then
            addSql = addSql + " and c.isusing='" + CStr(FRectIsusing) + "'"
        end if
        
        if FRectPoscode<>"" then
            addSql = addSql + " and c.poscode='" + CStr(FRectPoscode) + "'"
        end if

        If FRectsedatechk <> "" And FRectSelDate<>"" Then
            addSql = addSql + " and c.startdate = '" & FRectSelDate & "'"
		ElseIf FRectsedatechk = "" And  FRectSelDate<> "" Then 
			addSql = addSql + " and '" & FRectSelDate & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
		End If 

        if FRectSelDate<> "" and FRectSelDateTime <> "00" then 
            addSql = addSql + " and datepart(hh , c.startdate) >=" &FRectSelDateTime
        end if 

        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_mobile_mainCont c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_mobile_mainCont_code p on c.poscode=p.poscode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(p.posname,5) <> 'POINT' " & addSql
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_mobile_mainCont c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_mobile_mainCont_code p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(posname,5) <> 'POINT' " & addSql
        
        '//�켱���� ���� ����
		sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc"
       	
       	'response.write sqlStr &"<br>"
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
				set FItemList(i) = new CMainContentsItem

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
                FItemList(i).forderidx		= rsget("orderidx")
				FItemList(i).faltname		= rsget("altname")
				FItemList(i).FCgubun		= rsget("cgubun")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub
    
    
    public Sub GetPoint1010ContentsList()
        dim sqlStr, i
        dim yyyymmdd
        yyyymmdd = Left(now(),10)
        
        sqlStr = " select count(idx) as cnt from [db_sitemaster].[dbo].tbl_mobile_mainCont c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_mobile_mainCont_code p on c.poscode=p.poscode "
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(p.posname,5) = 'POINT' "
        
        if FRectIdx<>"" then
            sqlStr = sqlStr + " and c.idx=" + CStr(FRectIdx)
        end if
        
        if FRectValiddate<>"" then
            sqlStr = sqlStr + " and c.enddate>getdate()"
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
        
        rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget("cnt")
		rsget.close
        
        
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " c.*, p.posname, p.useSet "
        sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_mobile_mainCont c"
        sqlStr = sqlStr + " left join [db_sitemaster].[dbo].tbl_mobile_mainCont_code p"
        sqlStr = sqlStr + " on c.poscode=p.poscode"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and Left(posname,5) = 'POINT' "
        
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
        
        '//�÷����� �켱���� ���� ����
    	Select Case FRectPoscode
    		Case "400", "401", "402", "403", "404", "405", "420", "421", "428", "430"
    			sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc"
    		Case Else
    			sqlStr = sqlStr + " order by c.idx desc"
    	end Select
       	
       	'response.write sqlStr &"<br>"
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
				set FItemList(i) = new CMainContentsItem

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
                FItemList(i).forderidx		= rsget("orderidx")

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
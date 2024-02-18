<%
'###########################################################
' Description :  play 모바일 class
' History : 2013.09.03 이종화 생성
'			2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<%
Class CPlayMoContentsItem
    public Fidx
    public Fviewno
    public Fviewnotxt
    public Ftype
    public Ftypename
    public Fstate
    public Ftitle
    public Fsubcopy
    public Fstartdate
    public Fisusing
    public Flistimg
    public Fcontents
    public Fcolorcd
    public Fregdate
    public Flastupdate
    public Flastadminid
    public FpartWDname
    public FpartMDname
    public FpartPBname
    public FpartWDID
    public FpartMDID
    public FpartPBID
    public Ftagname
    public Ftagurl
    public Fiscomment
    public fitemid
    public forderno
    public FitemName
    public FImageSmall
    public FSellyn
    public Flimityn
    public Flimitno
    public Flimitsold
    public Fplayidx
    public Fstyle
    public Fworkcomm
    public Fcontents_idx
    public Fsortno
    public Ffavcnt


	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function
	
    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CPlayMoContents
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx
    public FRectType
    public FRectIsusing
    public FRectTitle
    public FRectState
    public FRectMDID
    public FRectWDID
    public FRectPBID
    public FRectitemid
    public FRectitemname
    public FRectPlayIdx

	public Sub sbPlayMoDetail()
	dim sqlStr, addsql
	
	If FRectIdx <> "" Then
		addsql = addsql & " and p.idx = '" & FRectIdx & "'"
	End If
	
	sqlStr = "select viewno, viewnotxt, type, title, subcopy, convert(varchar(10),startdate,120) as startdate, state, isusing, iscomment, partwdid, "
	sqlStr = sqlStr & " partmdid, partpbid, listimg, contents, colorcd, stylecd, workcomment, regdate, lastupdate, lastadminid, contents_idx, sortno "
	sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo as p with (nolock)"
	sqlStr = sqlStr & " where 1=1 " & addsql

	'response.write sqlStr &"<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	FResultCount = rsget.RecordCount

	set FOneItem = new CPlayMoContentsItem

	if Not rsget.Eof then

		FOneItem.Fviewno		= rsget("viewno")
		FOneItem.Fviewnotxt		= db2html(rsget("viewnotxt"))
		FOneItem.Ftype			= rsget("type")
		FOneItem.Fstate			= rsget("state")
        FOneItem.Ftitle			= db2html(rsget("title"))
        FOneItem.Fsubcopy		= db2html(rsget("subcopy"))
        FOneItem.Fstartdate		= rsget("startdate")
        FOneItem.Fisusing		= rsget("isusing")
        FOneItem.Flistimg		= rsget("listimg")
        FOneItem.Fcontents		= db2html(rsget("contents"))
        FOneItem.Fcolorcd		= rsget("colorcd")
        FOneItem.Fregdate		= rsget("regdate")
        FOneItem.Flastupdate	= rsget("lastupdate")
        FOneItem.Flastadminid	= rsget("lastadminid")
        FOneItem.FpartWDID		= rsget("partwdid")
        FOneItem.FpartMDID		= rsget("partmdid")
        FOneItem.FpartPBID		= rsget("partpbid")
        FOneItem.Fiscomment		= rsget("iscomment")
        FOneItem.Fstyle			= rsget("stylecd")
        FOneItem.Fworkcomm		= db2html(rsget("workcomment"))
        FOneItem.Fcontents_idx	= rsget("contents_idx")
        FOneItem.Fsortno		= rsget("sortno")

	end if
	rsget.Close
	end Sub

	public function fnPlayMoList()
        dim sqlStr, sqlsearch, i

		If FRectIsusing <> "" Then
			sqlsearch = sqlsearch & " and p.isusing = '" & FRectIsusing & "' "
		End If
		
		If FRectState <> "" Then
			sqlsearch = sqlsearch & " and p.state = '" & FRectState & "' "
		End If
		
		If FRectType <> "" Then
			sqlsearch = sqlsearch & " and p.type = '" & FRectType & "' "
		End If
		
		If FRectTitle <> "" Then
			sqlsearch = sqlsearch & " and p.title like '%" & FRectTitle & "%' "
		End If
		
		If FRectMDID <> "" Then
			sqlsearch = sqlsearch & " and p.partmdid = '" & FRectMDID & "' "
		End If
		
		If FRectWDID <> "" Then
			sqlsearch = sqlsearch & " and p.partwdid = '" & FRectWDID & "' "
		End If
		
		If FRectPBID <> "" Then
			sqlsearch = sqlsearch & " and p.partpbid = '" & FRectPBID & "' "
		End If


		'// 결과수 카운트
		sqlStr = "select count(p.idx) as cnt, CEILING(CAST(Count(p.idx) AS FLOAT)/" & FPageSize & ") AS totPg"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo as p with (nolock)"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            FTotalCount = rsget("cnt")
            FTotalPage	= rsget("totPg")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " p.idx, p.viewno, p.viewnotxt, p.type, c.typename, p.state, p.title, p.startdate, p.isusing, iscomment, p.listimg, isNull(p.colorcd,0) as colorcd, isNull(p.stylecd,0) as stylecd, p.contents, p.regdate, p.lastupdate, p.lastadminid, p.contents_idx, p.sortno, p.favcnt "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten with (nolock) WHERE userid = p.partmdid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as partMDname "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten with (nolock) WHERE userid = p.partWDid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as partWDname "
		sqlStr = sqlStr & " , (SELECT top 1 username from db_partner.dbo.tbl_user_tenbyten with (nolock) WHERE userid = p.partpbid and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '') as partPBname "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo as p with (nolock)"
        sqlStr = sqlStr & " 	left join db_sitemaster.dbo.tbl_play_mo_code as c with (nolock) on p.type = c.type"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by p.sortno DESC, p.idx DESC, p.startdate desc"

		'response.write sqlStr &"<Br>"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FtotalPage =  CInt(FTotalCount\FPageSize)
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
                set FItemList(i) = new CPlayMoContentsItem

                FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fviewno		= rsget("viewno")
				FItemList(i).Fviewnotxt		= db2html(rsget("viewnotxt"))
				FItemList(i).Ftype			= rsget("type")
                FItemList(i).Ftypename		= rsget("typename")
				FItemList(i).Fstate			= rsget("state")
                FItemList(i).Ftitle			= db2html(rsget("title"))
                FItemList(i).Fstartdate		= rsget("startdate")
                FItemList(i).Fisusing		= rsget("isusing")
                FItemList(i).Fiscomment		= rsget("iscomment")
                FItemList(i).Flistimg		= rsget("listimg")
                FItemList(i).Fcontents		= db2html(rsget("contents"))
                FItemList(i).Fcolorcd		= rsget("colorcd")
                FItemList(i).Fregdate		= rsget("regdate")
                FItemList(i).Flastupdate	= rsget("lastupdate")
                FItemList(i).Flastadminid	= rsget("lastadminid")
                FItemList(i).FpartWDname	= rsget("partWDname")
                FItemList(i).FpartMDname	= rsget("partMDname")
                FItemList(i).FpartPBname	= rsget("partPBname")
                FItemList(i).Fstyle			= rsget("stylecd")
                FItemList(i).Fcontents_idx	= rsget("contents_idx")
                FItemList(i).Fsortno		= rsget("sortno")
                FItemList(i).Ffavcnt		= rsget("favcnt")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function
    
    
	public function GetRowTagContent()
		dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and playidx="& FRectIdx &""
		end If

		if FRectType <> "" then
			sqlsearch = sqlsearch & " and playcate='"& FRectType &"'"
		end if

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo_tag"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit function

		'// 본문 내용 접수
		sqlStr = "select "
		sqlStr = sqlStr & " tagname , tagurl "
		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo_tag"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by tagidx asc "

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
		redim preserve FItemList(FTotalCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CPlayMoContentsItem

				FItemList(i).Ftagname        = rsget("tagname")
				FItemList(i).Ftagurl            = rsget("tagurl")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
    end Function
    
    
    public function fnGetCodeList()
    	dim sqlStr
    	
    	set FOneItem = new CPlayMoContentsItem
    	If FRectType <> "" Then
	    	sqlStr = "select * from db_sitemaster.dbo.tbl_play_mo_code where type = '" & FRectType & "'"
	    	rsget.Open sqlStr,dbget,1
	    	IF not rsget.EOF THEN
	    		FOneItem.Ftypename = rsget("typename")
	    		FOneItem.Fisusing = rsget("isusing")
	    	End IF
	    	rsget.Close
    	End If
    	
    	sqlStr = "select * from db_sitemaster.dbo.tbl_play_mo_code"
    	rsget.Open sqlStr,dbget,1
			IF not rsget.EOF THEN
				fnGetCodeList = rsget.getRows()
			End IF
    	rsget.Close
    end Function
    
    
    public function fnGetStyleCodeList()
    	dim sqlStr
    	
    	set FOneItem = new CPlayMoContentsItem
    	If FRectType <> "" Then
	    	sqlStr = "select * from db_sitemaster.dbo.tbl_play_mo_stylecode where type = '" & FRectType & "'"
	    	rsget.Open sqlStr,dbget,1
	    	IF not rsget.EOF THEN
	    		FOneItem.Ftypename = rsget("typename")
	    		FOneItem.Fisusing = rsget("isusing")
	    	End IF
	    	rsget.Close
    	End If
    	
    	sqlStr = "select * from db_sitemaster.dbo.tbl_play_mo_stylecode"
    	rsget.Open sqlStr,dbget,1
			IF not rsget.EOF THEN
				fnGetStyleCodeList = rsget.getRows()
			End IF
    	rsget.Close
    end Function
    
    
	public Function fnPlayItemList()
		dim sqlStr ,sqlsearch , i

		If FRectPlayIdx <> "" Then
			sqlsearch = sqlsearch + " and ti.playidx = '" + FRectPlayIdx + "'" + vbcrlf
		End If

		if FRectIsUsing <> "" then
			sqlsearch = sqlsearch + " and ti.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		if FRectitemid <> "" then
			sqlsearch = sqlsearch + " and ti.itemid = " + FRectitemid + "" + vbcrlf
		end if
		if FRectitemname <> "" then
			sqlsearch = sqlsearch + " and i.itemname like '%"+ FRectitemname + "%'" + vbcrlf
		end if
				
		'총수 접수
		sqlStr = "select count(*), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")"
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_mo_item ti with (nolock)"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on ti.itemid = i.itemid"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<BR>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		if FTotalCount < 1 then exit Function
		
		'내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " ti.idx ,ti.playidx ,ti.itemid ,ti.orderno ,ti.isusing" + vbcrlf
		sqlStr = sqlStr + " ,i.itemname, i.smallimage ,i.sellyn ,i.limityn ,i.limitno ,i.limitsold" + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_play_mo_item ti with (nolock)"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on ti.itemid = i.itemid"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by ti.orderno asc, ti.idx desc"

		'response.write sqlStr &"<BR>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CPlayMoContentsItem

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fplayidx		= rsget("playidx")
				FItemList(i).fitemid		= rsget("itemid")
				FItemList(i).forderno		= rsget("orderno")
				FItemList(i).fisusing		= rsget("isusing")
				FItemList(i).FitemName		= db2html(rsget("itemname"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


Function fnTypeSelectBox(gubun,ttype,isusing)
	Dim sqlStr, vBody
	sqlStr = "select type, typename "
	sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo_code"
	sqlStr = sqlStr & " where isusing = '" & isusing & "' "
	If gubun = "one" Then
		sqlStr = sqlStr & " and type = '" & ttype & "'"
	End If
	'response.write sqlStr
	rsget.Open SqlStr, dbget, 1
	
	If gubun = "select" Then
		vBody = vBody & "<option value="""" " & CHKIIF(ttype="","selected","") & "> - 선택 - </option>"
		
		if  not rsget.EOF  then
			do until rsget.EOF
				vBody = vBody & "<option value=""" & rsget("type") & """"
				If CStr(rsget("type")) = CStr(ttype) Then
					vBody = vBody & " selected"
				End If
				vBody = vBody & ">" & rsget("typename") & "</option>"
				
				rsget.movenext
			loop
		end if
		rsget.Close
	ElseIf gubun = "one" Then
		if  not rsget.EOF  then
		vBody = rsget("typename")
		end if
		rsget.Close
	End If
	
	fnTypeSelectBox = vBody
End Function


Function fnStyleSelectBox(gubun,ttype,isusing)
	Dim sqlStr, vBody
	sqlStr = "select type, typename "
	sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_play_mo_stylecode"
	sqlStr = sqlStr & " where isusing = '" & isusing & "' "
	If gubun = "one" Then
		sqlStr = sqlStr & " and type = '" & ttype & "'"
	End If
	'response.write sqlStr
	rsget.Open SqlStr, dbget, 1
	
	If gubun = "select" Then
		vBody = vBody & "<option value="""" " & CHKIIF(ttype="","selected","") & "> - 선택 - </option>"
		
		if  not rsget.EOF  then
			do until rsget.EOF
				vBody = vBody & "<option value=""" & rsget("type") & """"
				If CStr(rsget("type")) = CStr(ttype) Then
					vBody = vBody & " selected"
				End If
				vBody = vBody & ">" & rsget("typename") & "</option>"
				
				rsget.movenext
			loop
		end if
		rsget.Close
	ElseIf gubun = "one" Then
		if  not rsget.EOF  then
		vBody = rsget("typename")
		end if
		rsget.Close
	End If
	
	fnStyleSelectBox = vBody
End Function


Function fnStateSelectBox(gubun,sstate)
	Dim sqlStr, vBody
	
	If gubun = "one" Then
		SELECT Case sstate
			Case "0" : vBody = "등록대기"
			Case "3" : vBody = "이미지등록요청"
			Case "4" : vBody = "코딩등록요청"
			Case "5" : vBody = "오픈요청"
			Case "7" : vBody = "오픈"
			Case "9" : vBody = "종료"
			Case Else : vBody = ""
		END SELECT
	Else
		vBody = vBody & "<option value="""" " & CHKIIF(sstate="","selected","") & "> - 선택 - </option>" & vbCrLf
		vBody = vBody & "<option value=""0"" " & CHKIIF(sstate="0","selected","") & ">등록대기</option>" & vbCrLf
		vBody = vBody & "<option value=""3"" " & CHKIIF(sstate="3","selected","") & ">이미지등록요청</option>" & vbCrLf
		vBody = vBody & "<option value=""4"" " & CHKIIF(sstate="4","selected","") & ">코딩등록요청</option>" & vbCrLf
		vBody = vBody & "<option value=""5"" " & CHKIIF(sstate="5","selected","") & ">오픈요청</option>" & vbCrLf
		vBody = vBody & "<option value=""7"" " & CHKIIF(sstate="7","selected","") & ">오픈</option>" & vbCrLf
		vBody = vBody & "<option value=""9"" " & CHKIIF(sstate="9","selected","") & ">종료</option>"
	End If

	fnStateSelectBox = vBody
End Function
%>
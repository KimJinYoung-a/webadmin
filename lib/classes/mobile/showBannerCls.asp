<%
'//모바일 관리 sub class
Class CShowBanner
    public Fidx
	Public Fidxsub
	public Fstitle
	Public Freservationdate
	Public Fstate
	Public Fworktext

	Public FRegdate

	Public Ftagname
	Public Ftagurl

	Public Fsimg1
	Public Fsimg2
	Public Fsimg3
	Public Fsimg4
	Public Fsimg5

	Public Fsurl1
	Public Fsurl2
	Public Fsurl3
	Public Fsurl4
	Public Fsurl5

	Public Fsalt1
	Public Fsalt2
	Public Fsalt3
	Public Fsalt4
	Public Fsalt5

	Public FpartMDid
	Public FpartWDid
	Public FpartMKid
	public FpartMKname
	Public FpartMDname
	Public FpartWDname

	Public Fitemcnt

	Public Fitemid
	Public Fsortnum
	Public FIsUsing
	Public FitemName
	Public FsmallImage

	Public Fcolorcode
	Public Fviewno
	Public Fsubtitle

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CShowBannerContents
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

    public FRectIdx
	Public FRectsubidx
    public FRectgIdx
    public FRectPlaycate

	Public FRecttitle
	Public FRectstate

	Public FRPlaycate
	Public FRectTag

	Public FRectNo

	Public FRectpartMDid
	Public FRectpartWDid

	Public FRectIsusing
	Public FRectstitle
	Public Fisusing

	public Sub GetOneRowShowBanner()
	dim sqlStr
	sqlStr = "select * "
	sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_showbanner_list"
	sqlStr = sqlStr + " where showidx=" + CStr(FRectIdx)
	'response.write sqlStr
	rsget.Open SqlStr, dbget, 1
	FResultCount = rsget.RecordCount

	set FOneItem = new CShowBanner

	if Not rsget.Eof then

		FOneItem.Fidx				= rsget("showidx")
		FOneItem.Fstitle			= rsget("stitle")
		FOneItem.Fsimg1				= rsget("simg1")
		FOneItem.Fsimg2				= rsget("simg2")
		FOneItem.Fsimg3				= rsget("simg3")
		FOneItem.Fsimg4				= rsget("simg4")
		FOneItem.Fsimg5				= rsget("simg5")
		FOneItem.Freservationdate	= rsget("reservationdate")
		FOneItem.Fstate				= rsget("state")
		FOneItem.Fworktext			= rsget("worktext")
		FOneItem.FpartMDid			= rsget("partMDid")
		FOneItem.FpartWDid			= rsget("partWDid")
		FOneItem.Fsurl1				= rsget("surl1")
		FOneItem.Fsurl2				= rsget("surl2")
		FOneItem.Fsurl3				= rsget("surl3")
		FOneItem.Fsurl4				= rsget("surl4")
		FOneItem.Fsurl5				= rsget("surl5")
		FOneItem.Fsalt1				= rsget("salt1")
		FOneItem.Fsalt2				= rsget("salt2")
		FOneItem.Fsalt3				= rsget("salt3")
		FOneItem.Fsalt4				= rsget("salt4")
		FOneItem.Fsalt5				= rsget("salt5")
		FOneItem.Fcolorcode			= rsget("colorcode")
		FOneItem.Fviewno			= rsget("viewno")
		FOneItem.Fsubtitle			= rsget("subtitle")

	end if
	rsget.Close
	end Sub

	public function fnGetShowBannerList()
        dim sqlStr, sqlsearch, i

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and showidx = '"&FRectIdx&"'"
		end If

		if FRecttitle <> "" then
			sqlsearch = sqlsearch & " and stitle like '%"&FRecttitle&"%'"
		end if

		If FRectstate <> "" THEN
			IF FRectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and state = 7 and convert(varchar(10),getdate(),120) <= reservationdate"
			ELSE
				sqlsearch  = sqlsearch & " and state = " &FRectstate & ""
			END IF
		End If

		If FRectpartMDid <> "" Then
			sqlsearch = sqlsearch & " and partMDid = '"& FRectpartMDid  &"'"
		End If

		If FRectpartWDid <> "" Then
			sqlsearch = sqlsearch & " and partwDid = '"& FRectpartWDid  &"'"
		End If

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_showbanner_list"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

		if FTotalCount < 1 then exit function

        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " showidx , stitle , reservationdate , state , simg1 "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partMDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partMDname "
		sqlStr = sqlStr & " , (SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = partWDid and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0)) and p.userid <> '') as partWDname "
		sqlStr = sqlStr & " ,  isnull((select count(itemid) from db_sitemaster.dbo.tbl_mobile_showbanner_subitem as S where S.showidx = l.showidx and S.isusing = 'Y'),0) as itemcnt "
		sqlStr = sqlStr & " ,  viewno "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_showbanner_list as l"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by viewno asc , showidx DESC, reservationdate desc"

		'response.write sqlStr &"<Br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

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
                set FItemList(i) = new CShowBanner

                FItemList(i).Fidx						= rsget("showidx")
                FItemList(i).Fstitle					= rsget("stitle")
                FItemList(i).Freservationdate			= rsget("reservationdate")
                FItemList(i).Fstate						= rsget("state")
                FItemList(i).Fsimg1						= rsget("simg1")
                FItemList(i).FpartWDname				= rsget("partWDname")
                FItemList(i).FpartMDname				= rsget("partMDname")
                FItemList(i).Fitemcnt					= rsget("itemcnt")
				FItemList(i).Fviewno					= rsget("viewno")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end Function

	public Sub GetContentsItemList()
       dim sqlStr, addSql, i

		sqlStr = " select count(showidx) as cnt from db_sitemaster.dbo.tbl_mobile_showbanner_subitem "
		sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr & " and  showidx='" & FRectIdx & "'"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		'response.write sqlStr &"<br>"
        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "Select top " + CStr(FPageSize * FCurrPage) + " s.showitemidx , s.showidx , s.itemid , s.isusing as itemusing , s.sortnum , i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_mobile_showbanner_subitem as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        sqlStr = sqlStr & "Where showidx='" & FRectIdx & "'"

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If
        
		sqlStr = sqlStr + " order by sortnum asc" 

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
				set FItemList(i) = new CShowBanner
				
				FItemList(i).Fidxsub				= rsget("showitemidx")
	            FItemList(i).Fidx					= rsget("showidx")
	            FItemList(i).Fitemid				= rsget("itemid")
	            FItemList(i).Fsortnum				= rsget("sortnum")
	            FItemList(i).FIsUsing				= rsget("itemusing")
	            FItemList(i).FitemName				= rsget("itemname")
	            FItemList(i).FsmallImage			= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & rsget("smallImage"),"")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub

	'//subitem
	public Sub GetOneSubItem()
		dim SqlStr
        sqlStr = "Select top 1 s.*, i.itemname, i.smallImage "
        sqlStr = sqlStr & "From [db_sitemaster].[dbo].tbl_mobile_showbanner_subitem as s "
        sqlStr = sqlStr & "	left join db_item.dbo.tbl_item as i "
        sqlStr = sqlStr & "		on s.Itemid=i.itemid "
        sqlStr = sqlStr & "			and i.itemid<>0 "
        SqlStr = SqlStr + " where showitemidx=" + CStr(FRectSubIdx)

		'rw SqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CShowBanner
        if Not rsget.Eof then
            FOneItem.FIdxsub			= rsget("showitemidx")
            FOneItem.Fidx				= rsget("showidx")
            FOneItem.FItemid			= rsget("Itemid")
            FOneItem.Fsortnum			= rsget("sortnum")
            FOneItem.Fisusing			= rsget("isusing")
            FOneItem.FitemName			= rsget("itemname")
            FOneItem.FsmallImage		= chkIIF(Not(rsget("smallImage")="" or isNull(rsget("smallImage"))),webImgUrl & "/image/small/" & GetImageSubFolderByItemid(FOneItem.FItemid) & "/" & rsget("smallImage"),"")
        end if
        rsget.close
	End Sub

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


'//메인페이지 , 이벤트 공통함수		'/오픈예정 노출함 , 검색페이지용
function Draweventstate2(selectBoxName,selectedId,changeFlag)
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value="" <%if selectedId="" then response.write " selected"%>>선택</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>등록대기</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %>>이미지등록요청</option>
		<option value="5" <% if selectedId="5" then response.write "selected" %>>오픈요청</option>
		<option value="6" <% if selectedId="6" then response.write "selected" %>>오픈예정</option>
		<option value="7" <% if selectedId="7" then response.write "selected" %>>오픈</option>
		<option value="9" <% if selectedId="9" then response.write "selected" %>>종료</option>
	</select>
<%
end Function

'//메인페이지 , 이벤트 모두 공통
function geteventstate(v)
	if v = "0" then
		geteventstate = "등록대기"
	elseif v = "3" then
		geteventstate = "이미지등록요청"
	elseif v = "5" then
		geteventstate = "오픈요청"
	elseif v = "6" then
		geteventstate = "오픈예정"
	elseif v = "7" then
		geteventstate = "오픈"
	elseif v = "9" then
		geteventstate = "종료"
	end if
end Function

'/담당MD 리스트가져오기 (팀장 미만,직원 이상)
Sub sbGetpartid(ByVal selName, ByVal sIDValue, ByVal sScript,part_sn)
	Dim strSql, arrList, intLoop

	if part_sn = "" then exit sub

	strSql = " SELECT userid, username"
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
	strSql = strSql & " WHERE part_sn IN("&part_sn&") and  posit_sn>='4' and  posit_sn<='12' and   isUsing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	strSql = strSql & " and userid <> '' order by posit_sn, empno"

	'response.write strSql &"<Br>"
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
	arrList = rsget.getRows()
	End IF
	rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">선택</option>
	<%
	If isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
	%>
	<option value="<%=arrList(0,intLoop)%>" <%if arrList(0,intLoop) = sIDValue then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
	<%
		Next
	End IF
	%>
	</select>
<%
End Sub
%>
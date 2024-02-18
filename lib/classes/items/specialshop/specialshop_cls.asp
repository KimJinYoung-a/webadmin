<%
'###########################################################
' Description : 우수회원샵
' Hieditor : 2009.12.28 한용민 생성
'###########################################################
%>
<%
Class cspecialshop_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fid
	public fopenDate
	public fstatus
	public fstatusstr
	public fregdate
	public fidx
	public fitemid
	public fisusing	
	public fitemcount
	public FItemName
	public Fmakerid
	public FSellCash
	public FBuyCash
	public FOrgPrice
	public FSellyn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public FSailYN
	public FImageSmall
	public FImageList
	public FSpecialuseritem
	public Ftitle
	public FendDate

	public function getRealPrice()
		getRealPrice = FSellCash - FSellCash*0.25
	end function

	public function IsSail()
	    IsSail =  (FSailYN="Y") and (FOrgPrice>FSellCash) 
	end function
	
	public function getMargin()
		getMargin = 100- FBuyCash/getRealPrice*100 
	end function
	
	public function IsSailItem()
	    IsSailItem = ((FSailYN="Y") and (FOrgPrice>FSellCash)) or (FSpecialuseritem>0)
	end function

	public function getSailPro()
		if FOrgPrice=0 then
			getSailPro = 0
		else
			getSailPro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100)
		end if
	end function


	public function IsFreeBeasong()
		if FItemGubun="04" then
			if FSellCash>=getFreeBeasongLimitByUserLevel() then
				IsFreeBeasong = true
			else
				IsFreeBeasong = false
			end if
		else
			if FSellCash>=getFreeBeasongLimitByUserLevel() then
				IsFreeBeasong = true
			else
				IsFreeBeasong = false
			end if
		end if

		if (FDeliverytype="4") or (FDeliverytype="5") then
			IsFreeBeasong = true
		end if
	end function

	public function getFreeBeasongLimitByUserLevel()
		dim ulevel
		ulevel = getUserLevel()
		if ulevel>2 then
			getFreeBeasongLimitByUserLevel = 0
		elseif ulevel>1 then
			getFreeBeasongLimitByUserLevel = 30000
		elseif ulevel>0 then
			getFreeBeasongLimitByUserLevel = 40000
		else
			getFreeBeasongLimitByUserLevel = 50000
		end if
	end function

	public function IsSoldOut()
		IsSoldOut = (FDispyn="N") or (FSailyn="N") or ((FLimityn="Y") and (FLimitno-Limitsold<1))
	end function
		

	public function getUserLevel()
		getUserLevel = request.cookies("uinfo")("userlevel")
	end function
	
end class	

class cspecialshop_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem
	
	public frectstatus
	public frectid
	public frectisusing

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
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

	'//aadmin/shopmaster/specialshop/specialshop.asp
	public sub fspecialshop_list()
		dim sqlStr,i , sqlsearch

		sqlsearch=""
		'//검색처리
		if frectid <> "" then
			sqlsearch = sqlsearch & " and id = "&frectid&"" + vbcrlf
		elseif frectisusing = "Y" then
			sqlsearch = sqlsearch & " and isusing = '"&frectisusing&"'" + vbcrlf	
		end if
		'//대기
		if frectstatus = "0" then
			sqlsearch = sqlsearch & " and openDate > getdate() and status = 0" + vbcrlf
		'//오픈
		elseif 	frectstatus = "1" then
			sqlsearch = sqlsearch & " and openDate <= getdate() and status = 1" + vbcrlf
		'//종료
		elseif 	frectstatus = "2" then
			sqlsearch = sqlsearch & " and openDate <= getdate() and status = 2" + vbcrlf					
		end if		
		
		'총 갯수 구하기
		sqlStr = "select count(id) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_specialShop with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where id <> 0 "& sqlsearch &"" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf		
		sqlStr = sqlStr & " 	(case when openDate > getdate() or status = 0 then '대기' " + vbcrlf
		sqlStr = sqlStr & " 	when openDate <= getdate() and status = 1 then '오픈'" + vbcrlf
		sqlStr = sqlStr & " 	when openDate <= getdate() and status = 2  then '종료' end) as statusstr" + vbcrlf	
		sqlStr = sqlStr & " ,id,openDate,status,regdate , isusing" + vbcrlf
		sqlStr = sqlStr & " , (select count(idx) from db_item.dbo.tbl_specialShopitem with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	where a.id = id and isusing ='Y') as itemcount" + vbcrlf
		sqlStr = sqlStr & " , a.title, a.endDate " + vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_specialShop a with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where id <> 0 "& sqlsearch &"" + vbcrlf	
		sqlStr = sqlStr & " order by id desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cspecialshop_oneitem
				
				FItemList(i).fstatusstr = rsget("statusstr")
				FItemList(i).fid = rsget("id")	
				FItemList(i).fopenDate = rsget("openDate")
				FItemList(i).fstatus = rsget("status")	
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fitemcount = rsget("itemcount")	
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).Ftitle = db2html(rsget("title"))
				If isNull(rsget("endDate")) Then
					FItemList(i).FendDate = ""
				Else
					FItemList(i).FendDate = rsget("endDate")
				End If
																	
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//aadmin/shopmaster/specialshop/specialshop_edit.asp
	public sub fspecialshop_oneitem()
        dim sqlStr , sqlsearch, i

		sqlsearch=""
		'//검색처리
		if frectid <> "" then
			sqlsearch = sqlsearch & " and a.id = "&frectid&"" + vbcrlf
		end if
        
		'데이터 리스트 
		sqlStr = "select top 1" + vbcrlf		
		sqlStr = sqlStr & " 	(case when openDate > getdate() or status = 0 then '대기' " + vbcrlf
		sqlStr = sqlStr & " 	when openDate <= getdate() and status = 1 then '오픈'" + vbcrlf
		sqlStr = sqlStr & " 	when openDate <= getdate() and status = 2  then '종료' end) as statusstr" + vbcrlf	
		sqlStr = sqlStr & " ,id,openDate,status,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " , (select count(idx) from db_item.dbo.tbl_specialShopitem with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	where a.id = id and isusing ='Y') as itemcount" + vbcrlf
		sqlStr = sqlStr & " , a.title, a.endDate " + vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_specialShop a with (nolock)" & vbcrlf
		sqlStr = sqlStr & " where id <> 0 "& sqlsearch &"" + vbcrlf	
		sqlStr = sqlStr & " order by id desc" + vbcrlf

        'response.write sqlStr&"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cspecialshop_oneitem
        
        if Not rsget.Eof then
    
			FOneItem.fstatusstr = rsget("statusstr")
			FOneItem.fid = rsget("id")	
			FOneItem.fopenDate = rsget("openDate")
			FOneItem.fstatus = rsget("status")	
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fitemcount = rsget("itemcount")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.ftitle = rsget("title")
			If isNull(rsget("endDate")) Then
				FOneItem.fendDate = ""
			Else
				FOneItem.fendDate = rsget("endDate")
			End If
			       
        end if
        rsget.Close
	end sub

	'//aadmin/shopmaster/specialshop/specialshop_edit.asp
	public sub fspecialshop_itemlist()
		dim sqlStr,i , search
		
		'//검색처리
		if frectid <> "" then
			search = search & " and a.id = "&frectid&"" + vbcrlf
		end if
			
		'데이터 리스트 
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " a.idx,a.id,a.itemid,a.isusing,a.regdate" + vbcrlf
		sqlStr = sqlStr & " ,i.itemid, i.itemname, i.makerid,i.sellcash,i.sailyn" + vbcrlf
		sqlStr = sqlStr & " ,i.orgprice,i.sellyn, i.limityn, i.limitno,i.limitsold" + vbcrlf
		sqlStr = sqlStr & " ,i.sailyn,i.specialuseritem,i.smallimage , i.listimage, i.buycash" + vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_specialShopitem a" + vbcrlf
		sqlStr = sqlStr & " left join db_item.[dbo].tbl_item i" + vbcrlf
		sqlStr = sqlStr & " on a.itemid = i.itemid" + vbcrlf
		sqlStr = sqlStr & " where a.isusing='Y' "&search&"" + vbcrlf
		sqlStr = sqlStr & " order by a.regdate desc ,a.itemid desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		ftotalcount = rsget.recordcount
		redim preserve FItemList(ftotalcount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cspecialshop_oneitem
				
				FItemList(i).FSpecialuseritem = rsget("specialuseritem")								
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fid = rsget("id")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).FItemId       = rsget("itemid")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).Fmakerid     = rsget("makerid")
				FItemList(i).FSellCash     = rsget("sellcash")
				FItemList(i).FBuyCash     = rsget("buycash")
				FItemList(i).FOrgPrice = rsget("orgprice")
				FItemList(i).FSellyn       = rsget("sellyn")
				FItemList(i).FLimitYn      = rsget("limityn")
				FItemList(i).FLimitNo      = rsget("limitno")
				FItemList(i).FLimitSold    = rsget("limitsold")
				FItemList(i).FSailYN		= rsget("sailyn")				
				FItemList(i).FImageSmall   = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallimage")
				FItemList(i).FImageList   = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("listimage")																
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
end Class

function drawstatus(byval boxname, byval status , byval id)

	'//신규등록
	if id = "" then
%>
		<select name="<%=boxname%>">
			<option value="" <% if status ="" then response.write " selected" %>>선택</option>
			<option value="0" <% if status ="0" then response.write " selected" %>>대기</option>
			<option value="1" <% if status ="1" then response.write " selected" %>>오픈</option>
			<option value="2" <% if status ="2" then response.write " selected" %>>종료</option>		
		</select>
<%
	'//수정모드
	else
		'//상태가 오픈이상일경우 대기는 선택못함
		if status >= "1" then
%>
		<select name="<%=boxname%>">
			<option value="" <% if status ="" then response.write " selected" %>>선택</option>
			<option value="1" <% if status ="1" then response.write " selected" %>>오픈</option>
			<option value="2" <% if status ="2" then response.write " selected" %>>종료</option>		
		</select>
<%	
		else
%>
		<select name="<%=boxname%>">
			<option value="" <% if status ="" then response.write " selected" %>>선택</option>
			<option value="0" <% if status ="0" then response.write " selected" %>>대기</option>
			<option value="1" <% if status ="1" then response.write " selected" %>>오픈</option>
			<option value="2" <% if status ="2" then response.write " selected" %>>종료</option>		
		</select>
<%		
		end if
	end if
end function
%>
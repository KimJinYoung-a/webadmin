<%
'####################################################
' Description :  이벤트
' History : 2016.08.09 김진영 생성
'####################################################
Class CEventItem
	Public FIdx
	Public FGubun
	Public FActid
	Public FEvt_startdate
	Public FEvt_enddate
	Public FContentsCode
	Public FEvt_name
	Public FIsusing
	Public FRegid
	Public FRegdate
	Public FLastupdateid
	Public FLastupdate

	Public FId
	Public FCompany_name
	Public FSocname

	Public FLecLimitSold
	Public FLecIdx
	Public FReg_yn
	Public FLecCost
	Public FMatCost
	Public FLectCount
	Public FReg_endday
	Public FLecperiod
	Public Flecturer_id
	Public FLecLimitCount
	Public Flec_outline
	Public FReg_startday
	Public FMatincludeYN
	Public FLecDate
	Public FLecturer_name
	Public Flecturer_regdate
	Public flecturercouponyn
	Public Flecturercoupontype
	Public FLecTitle
	Public FlecImgProfile75
	Public Flecturercouponvalue
	Public fcurrlecturercouponidx
	Public Ficon1
	Public Fbasicimg
	Public FSmallimg
	Public Foblong_img1
	Public Foblong_img2
	Public Foblong_img3
	Public Fmorollingimg1
	Public FOptionCnt
	Public FOptLimitCnt
	
	Public FItemid
	Public FItemName
	Public FSellCash
	Public FOrgPrice
	Public FSellyn
	Public FSaleyn
	Public FItemcouponyn
	Public FItemcouponvalue
	Public FItemcoupontype

	Public FUserid
	Public FLec_yn
	Public FDiy_yn

	'// 세일 상품 여부 '! 
	public Function IsSaleItem() 
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0))
	end Function

	'// 상품 쿠폰 여부  '!
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격  '!
	public Function getRealPrice()
		getRealPrice = FSellCash
	end Function

	'// 할인율 '!
	public Function getSalePro() 
		if FOrgprice=0 then
			getSalePro = 0 & "%"
		else
			getSalePro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100) & "%"
		end if
	end Function

	'// 쿠폰 적용가
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() 
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	end Function

	'// 상품 쿠폰 내용
	public function GetCouponDiscountStr()
		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr = formatNumber(Fitemcouponvalue,0) + "원 할인"
			Case "3"
				GetCouponDiscountStr ="무료배송"
			Case Else
				GetCouponDiscountStr = Fitemcoupontype
		End Select
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CEvent
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectStartdate
	Public FRectEnddate
	Public FRectGubun
	Public FRectIsusing
	Public FRectEvting

	Public FRectCharcd
	Public FRectSearchKey
	Public FRectSearchString
	Public FRectLecturerid
	Public FRectMakerid

	Public FRectIdx

	'작가/강사 이벤트 리스트
	Public Sub getEventItemList
		Dim i, sqlStr, addSql
		'기간 검색
		If FRectStartdate <> "" and FRectEnddate <> "" Then
			addSql = addSql & " and evt_startdate <= '"&FRectStartdate&"' and evt_enddate >= '"&FRectEnddate&"' "
		End If
		
		'셀렉트박스로 검색
		If FRectSearchKey <> "" and FRectSearchString <> "" Then
			Select Case FRectSearchKey
				Case "eCode"			addSql = addSql & " and idx = '"&FRectSearchString&"' "
				Case "contentsCode"		addSql = addSql & " and contentsCode = '"&FRectSearchString&"' "
				Case "evt_name"			addSql = addSql & " and evt_name like '%"&FRectSearchString&"%' "
			End Select
		End If

		If FRectEvting <> "" Then
			Select Case FRectEvting
				Case "ing"
					addSql = addSql & " and evt_startdate <= getdate() "
					addSql = addSql & " and evt_enddate >= getdate() "
				Case "end"
					addSql = addSql & " and evt_enddate < getdate() "
				Case "will"
					addSql = addSql & " and evt_startdate > getdate() "
			End Select
		End If

		'등록위치 검색
		If FRectGubun <> "" Then
			addSql = addSql & " and gubun = '"&FRectGubun&"' "
		End If

		'사용여부 검색
		If FRectIsusing <> "" Then
			addSql = addSql & " and isusing = '"&FRectIsusing&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_event] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and actid = '"&session("ssBctID")&"'  "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	idx, gubun, actid, evt_startdate, evt_enddate, contentsCode, evt_name, isusing, regid, regdate, lastupdateid, lastupdate "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_event] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and actid = '"&session("ssBctID")&"'  "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new CEventItem
					FItemList(i).FIdx			= rsACADEMYget("idx")
					FItemList(i).FGubun			= rsACADEMYget("gubun")
					FItemList(i).FActid			= rsACADEMYget("actid")
					FItemList(i).FEvt_startdate	= rsACADEMYget("evt_startdate")
					FItemList(i).FEvt_enddate	= rsACADEMYget("evt_enddate")
					FItemList(i).FContentsCode	= rsACADEMYget("contentsCode")
					FItemList(i).FEvt_name		= db2html(rsACADEMYget("evt_name"))
					FItemList(i).FIsusing		= rsACADEMYget("isusing")
					FItemList(i).FRegid			= rsACADEMYget("regid")
					FItemList(i).FRegdate		= rsACADEMYget("regdate")
					FItemList(i).FLastupdateid	= rsACADEMYget("lastupdateid")
					FItemList(i).FLastupdate	= rsACADEMYget("lastupdate")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	'진행중인 강좌 리스트
	Public Sub getLecList
		Dim i, sqlStr, addSql

		If FRectSearchKey <> "" AND FRectSearchString <> "" Then
			If FRectSearchKey = "lecidx" Then
				addsql = addsql & " and i.idx = '"&FRectSearchString&"' "
			ElseIf FRectSearchKey = "lectitle" Then
				addsql = addsql & " and i.lec_title like '%"&FRectSearchString&"%' "
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as g on i.lecturer_id = g.lecturer_id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.disp_yn = 'Y' "
		sqlStr = sqlStr & " and i.lecturer_id = '"&FRectlecturerID&"' "
		sqlStr = sqlStr & " and i.reg_startday <= getdate() "
		sqlStr = sqlStr & " and i.reg_endday >= getdate() "
		sqlStr = sqlStr & addsql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.idx, i.lec_title, i.lec_cost, i.lec_count, i.lec_period, i.lec_startday1, i.limit_count, i.limit_sold, i.lecturer_id, morollingimg1 "
		sqlStr = sqlStr & " ,i.mainimg, i.storyimg, i.smallimg, i.basicimg, i.icon1, i.oblongImg1, i.oblongImg2, i.oblongImg3, i.mat_cost, i.matinclude_yn "
		sqlStr = sqlStr & " ,i.reg_yn, i.reg_startday, i.reg_endday, i.lecturercouponyn, i.currlecturercouponidx, i.lecturercoupontype "
		sqlStr = sqlStr & " ,i.lecturer_regdate, i.lecturercouponvalue, i.lec_outline, i.cate_large, g.image_profile_75x75, g.lecturer_name, i.optioncnt "
		sqlStr = sqlStr & " ,(SELECT TOP 1 limit_count - limit_sold FROM [db_academy].[dbo].tbl_lec_item_option WHERE lecIdx = i.idx and lecOptionName = i.lec_period) as optLimitCnt "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_item as i "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as g on i.lecturer_id = g.lecturer_id"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.disp_yn = 'Y' "
		sqlStr = sqlStr & " and i.lecturer_id = '"&FRectlecturerID&"' "
		sqlStr = sqlStr & " and i.reg_startday <= getdate() "
		sqlStr = sqlStr & " and i.reg_endday >= getdate() "
		sqlStr = sqlStr & addsql
		sqlStr = sqlStr & " ORDER BY optLimitCnt DESC"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				set FItemList(i) = new CEventItem
					if rsACADEMYget("limit_sold") > rsACADEMYget("limit_count") then
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_count")
					else
						FItemList(i).FLecLimitSold		= rsACADEMYget("limit_sold")
					end if
					FItemList(i).FLecIdx					= rsACADEMYget("idx")
					FItemList(i).FReg_yn					= rsACADEMYget("reg_yn")
					FItemList(i).FLecCost					= rsACADEMYget("lec_cost")
					FItemList(i).FMatCost					= rsACADEMYget("mat_cost")
					FItemList(i).FLectCount				= rsACADEMYget("lec_count")
					FItemList(i).FReg_endday				= rsACADEMYget("reg_endday")
					FItemList(i).FLecperiod				= rsACADEMYget("lec_period")
					FItemList(i).Flecturer_id			= rsACADEMYget("lecturer_id")
					FItemList(i).FLecLimitCount	  		= rsACADEMYget("limit_count")
					FItemList(i).Flec_outline			= rsACADEMYget("lec_outline")
					FItemList(i).FReg_startday			= rsACADEMYget("reg_startday")
					FItemList(i).FMatincludeYN			= rsACADEMYget("matinclude_yn")
					FItemList(i).FLecDate					= rsACADEMYget("lec_startday1")
					FItemList(i).FLecturer_name	  		= rsACADEMYget("lecturer_name")			'강사 이름
					FItemList(i).Flecturer_regdate		= rsACADEMYget("lecturer_regdate")
					FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
					FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
					FItemList(i).FLecTitle				= db2html(rsACADEMYget("lec_title"))
					FItemList(i).FlecImgProfile75	  	= rsACADEMYget("image_profile_75x75")	'강사 아이콘(75x75)
					FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")
					FItemList(i).fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
					FItemList(i).Ficon1					= uploadUrl & "/lectureitem/icon1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("icon1")
					FItemList(i).Fbasicimg	  			= uploadUrl & "/lectureitem/basic/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("basicimg")
					FItemList(i).FSmallimg				= uploadUrl & "/lectureitem/small/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("smallimg")
					FItemList(i).Foblong_img1			= uploadUrl & "/lectureitem/obl1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg1")
					FItemList(i).Foblong_img2			= uploadUrl & "/lectureitem/obl2/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg2")
					FItemList(i).Foblong_img3			= uploadUrl & "/lectureitem/obl3/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("oblongImg3")
					FItemList(i).Fmorollingimg1			= uploadUrl & "/lectureitem/morolling1/"+ GetImageSubFolderByItemid(FItemList(i).FLecIdx) + "/" + rsACADEMYget("morollingimg1")
					FItemList(i).FOptionCnt 			= rsACADEMYget("optionCnt")
					FItemList(i).FOptLimitCnt 			= rsACADEMYget("optLimitCnt")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub

	'판매중인 작품 리스트
	Public Sub getArtList
		Dim i, sqlStr, addSql

		If FRectSearchKey <> "" AND FRectSearchString <> "" Then
			If FRectSearchKey = "itemid" Then
				addsql = addsql & " and i.itemid = '"&FRectSearchString&"' "
			ElseIf FRectSearchKey = "itemname" Then
				addsql = addsql & " and i.itemname like '%"&FRectSearchString&"%' "
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_diy_item as i "
		sqlStr = sqlStr & " JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] as ci on i.itemid = ci.itemid and ci.isDefault = 'y' "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as g on i.makerid = g.lecturer_id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.sellyn = 'Y' "
		sqlStr = sqlStr & " and ((i.limityn = 'N') or ((i.limityn = 'Y') and (i.limitno - i.limitsold > 0))) "
		sqlStr = sqlStr & " and i.makerid = '"&FRectMakerid&"' "
		sqlStr = sqlStr & addsql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.itemid, i.itemname, i.sellcash, i.orgPrice "
		sqlStr = sqlStr & " ,i.Sellyn, i.SaleYn, i.ItemCouponYn, i.ItemCouponValue, i.ItemCouponType, i.regdate "
		sqlStr = sqlStr & " ,g.lecturer_name "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_diy_item as i "
		sqlStr = sqlStr & " JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] as ci on i.itemid = ci.itemid and ci.isDefault = 'y' "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as g on i.makerid = g.lecturer_id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.sellyn = 'Y' "
		sqlStr = sqlStr & " and ((i.limityn = 'N') or ((i.limityn = 'Y') and (i.limitno - i.limitsold > 0))) "
		sqlStr = sqlStr & " and i.makerid = '"&FRectMakerid&"' "
		sqlStr = sqlStr & addsql
		sqlStr = sqlStr & " ORDER BY i.itemid DESC"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				set FItemList(i) = new CEventItem
					FItemList(i).FItemid			= rsACADEMYget("itemid")
					FItemList(i).FItemName			= db2html(rsACADEMYget("itemName"))
					FItemList(i).FSellCash			= rsACADEMYget("sellCash")
					FItemList(i).FOrgPrice			= rsACADEMYget("orgPrice")
					FItemList(i).FSellyn			= rsACADEMYget("SellYn")
					FItemList(i).FSaleyn			= rsACADEMYget("SaleYn")
					FItemList(i).FItemcouponyn		= rsACADEMYget("itemcouponYn")
					FItemList(i).FItemcouponvalue	= rsACADEMYget("itemCouponValue")
					FItemList(i).FItemcoupontype	= rsACADEMYget("itemCouponType")
					FItemList(i).Flecturer_name		= rsACADEMYget("lecturer_name")
					FItemList(i).FRegdate		= rsACADEMYget("regdate")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub

	'이벤트 수정
	Public Sub getEventOneItem
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, gubun, actid, company_name, evt_startdate, evt_enddate, contentsCode, evt_name, isusing, regid, regdate, lastupdateid, lastupdate "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_event] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and idx = '"&FRectIdx&"' "
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FTotalCount = rsACADEMYget.RecordCount
		Set FOneItem = new CEventItem

		If Not rsACADEMYget.Eof Then
			FOneItem.FIdx				= rsACADEMYget("idx")
			FOneItem.FGubun				= rsACADEMYget("gubun")
			FOneItem.FActid				= rsACADEMYget("actid")
			FOneItem.FCompany_name		= rsACADEMYget("company_name")
			FOneItem.FEvt_startdate		= rsACADEMYget("evt_startdate")
			FOneItem.FEvt_enddate		= rsACADEMYget("evt_enddate")
			FOneItem.FContentsCode		= rsACADEMYget("contentsCode")
			FOneItem.FEvt_name			= rsACADEMYget("evt_name")
			FOneItem.FIsusing			= rsACADEMYget("isusing")
			FOneItem.FRegid				= rsACADEMYget("regid")
			FOneItem.FRegdate			= rsACADEMYget("regdate")
			FOneItem.FLastupdateid		= rsACADEMYget("lastupdateid")
			FOneItem.FLastupdate		= rsACADEMYget("lastupdate")
		End If
		rsACADEMYget.Close
	End Sub

	'내가 작가인지 강사인지 판별하기
	Public Function getWhatMyJob()
		Dim sqlStr
		sqlStr = ""
		sqlStr =  sqlStr & " SELECT TOP 1 "
		sqlStr =  sqlStr & " Case WHEN u.lec_yn = 'Y' and u.diy_yn = 'N' THEN 'L' "
		sqlStr =  sqlStr & "	  WHEN ((u.lec_yn = 'Y' and u.diy_yn = 'Y') OR (u.lec_yn = 'N' and u.diy_yn = 'Y')) THEN 'D' Else 'X' End as gubun "
		sqlStr =  sqlStr & " FROM [db_user].[dbo].tbl_user_c c "
		sqlStr =  sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner p on c.userid=p.id "
		sqlStr =  sqlStr & " LEFT JOIN [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U on c.userid = U.lecturer_id "
		sqlStr =  sqlStr & " WHERE c.userid<>'' "
		sqlStr =  sqlStr & " and c.userdiv ='14' "
		sqlStr =  sqlStr & " and c.isusing = 'Y' "
		sqlStr =  sqlStr & " and c.userid= '"&session("ssBctID")&"' "
		rsget.Open sqlStr, dbget, 1
		FTotalCount = rsget.RecordCount
		Set FOneItem = new CEventItem
		If Not rsget.Eof Then
			getWhatMyJob = rsget("gubun")
		Else
			getWhatMyJob = "X"
		End If
		rsget.Close
	End Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
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
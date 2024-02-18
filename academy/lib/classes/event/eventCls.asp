<%
'####################################################
' Description :  ÀÌº¥Æ®
' History : 2016.08.09 ±èÁø¿µ »ý¼º
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
	Public FSocname_kor
	Public FLec_yn
	Public FDiy_yn

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

	'// ¼¼ÀÏ »óÇ° ¿©ºÎ '! 
	public Function IsSaleItem() 
	    IsSaleItem = ((FSaleYn="Y") and (FOrgPrice-FSellCash>0))
	end Function

	'// »óÇ° ÄíÆù ¿©ºÎ  '!
	public Function IsCouponItem()
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// ¼¼ÀÏÆ÷ÇÔ ½ÇÁ¦°¡°Ý  '!
	public Function getRealPrice()
		getRealPrice = FSellCash
	end Function

	'// ÇÒÀÎÀ² '!
	public Function getSalePro() 
		if FOrgprice=0 then
			getSalePro = 0 & "%"
		else
			getSalePro = CLng((FOrgPrice-getRealPrice)/FOrgPrice*100) & "%"
		end if
	end Function

	'// ÄíÆù Àû¿ë°¡
	public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function

	'// ÄíÆù ÇÒÀÎ°¡
	public Function GetCouponDiscountPrice() 
		Select case Fitemcoupontype
			case "1" ''% ÄíÆù
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''¿ø ÄíÆù
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''¹«·á¹è¼Û ÄíÆù
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

	end Function

	'// »óÇ° ÄíÆù ³»¿ë
	public function GetCouponDiscountStr()
		Select Case Fitemcoupontype
			Case "1"
				GetCouponDiscountStr =CStr(Fitemcouponvalue) + "%"
			Case "2"
				GetCouponDiscountStr = formatNumber(Fitemcouponvalue,0) + "¿ø ÇÒÀÎ"
			Case "3"
				GetCouponDiscountStr ="¹«·á¹è¼Û"
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

	'ÀÛ°¡/°­»ç ÀÌº¥Æ® ¸®½ºÆ®
	Public Sub getEventItemList
		Dim i, sqlStr, addSql
		'±â°£ °Ë»ö
		If FRectStartdate <> "" and FRectEnddate <> "" Then
			addSql = addSql & " and evt_startdate <= '"&FRectStartdate&"' and evt_enddate >= '"&FRectEnddate&"' "
		End If
		
		'¼¿·ºÆ®¹Ú½º·Î °Ë»ö
		If FRectSearchKey <> "" and FRectSearchString <> "" Then
			Select Case FRectSearchKey
				Case "eCode"			addSql = addSql & " and idx = '"&FRectSearchString&"' "
				Case "contentsCode"		addSql = addSql & " and contentsCode = '"&FRectSearchString&"' "
				Case "evt_name"			addSql = addSql & " and evt_name like '%"&FRectSearchString&"%' "
				Case "actid"			addSql = addSql & " and actid = '"&FRectSearchString&"' "
				Case "company_name"		addSql = addSql & " and company_name like '%"&FRectSearchString&"%' "
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

		'µî·ÏÀ§Ä¡ °Ë»ö
		If FRectGubun <> "" Then
			addSql = addSql & " and gubun = '"&FRectGubun&"' "
		End If

		'»ç¿ë¿©ºÎ °Ë»ö
		If FRectIsusing <> "" Then
			addSql = addSql & " and isusing = '"&FRectIsusing&"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_event] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		'ÁöÁ¤ÆäÀÌÁö°¡ ÀüÃ¼ ÆäÀÌÁöº¸´Ù Å¬ ¶§ ÇÔ¼öÁ¾·á
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	idx, gubun, actid, evt_startdate, evt_enddate, contentsCode, evt_name, isusing, regid, regdate, lastupdateid, lastupdate "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_academy_event] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
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

	'ÀÛ°¡/°­»ç ¸®½ºÆ®
	Public Sub getPartnerList
		Dim i, sqlStr, addSql

		If FRectGubun <> "" Then
			Select Case FRectGubun
				Case "L"		'°­»ç
					addsql = addsql & " and u.lec_yn = 'Y' and u.diy_yn = 'N' "
				Case "D"		'ÀÛ°¡
					addsql = addsql & " and ((u.lec_yn = 'Y' and u.diy_yn = 'Y') OR (u.lec_yn = 'N' and u.diy_yn = 'Y')) "
			End Select
		End If

		If FRectSearchKey <> "" AND FRectSearchString <> "" Then
			If FRectSearchKey = "id" Then
				addsql = addsql & " and p.id = '"&FRectSearchString&"' "
			ElseIf FRectSearchKey = "name" Then
				addsql = addsql & " and (p.company_name like '%"&FRectSearchString&"%' OR c.socname like '%"&FRectSearchString&"%' )"
			End If
		End If

		If FRectCharcd <> "" Then
			Select Case FRectCharcd
				Case "all"		addSql = addSql & " and 1=1 "
				Case "°¡"		addSql = addSql & " and p.company_name between '¤¡' and '¤¤' "
				Case "³ª"		addSql = addSql & " and p.company_name between '¤¤' and '¤§' "
				Case "´Ù"		addSql = addSql & " and p.company_name between '¤§' and '¤©' "
				Case "¶ó"		addSql = addSql & " and p.company_name between '¤©' and '¤±' "
				Case "¸¶"		addSql = addSql & " and p.company_name between '¤±' and '¤²' "
				Case "¹Ù"		addSql = addSql & " and p.company_name between '¤²' and '¤µ' "
				Case "»ç"		addSql = addSql & " and p.company_name between '¤µ' and '¤·' "
				Case "¾Æ"		addSql = addSql & " and p.company_name between '¤·' and '¤¸' "
				Case "ÀÚ"		addSql = addSql & " and p.company_name between '¤¸' and '¤º' "
				Case "Â÷"		addSql = addSql & " and p.company_name between '¤º' and '¤»' "
				Case "Ä«"		addSql = addSql & " and p.company_name between '¤»' and '¤¼' "
				Case "Å¸"		addSql = addSql & " and p.company_name between '¤¼' and '¤½' "
				Case "ÆÄ"		addSql = addSql & " and p.company_name between '¤½' and '¤¾' "
				Case "ÇÏ"		addSql = addSql & " and p.company_name between '¤¾' and 'ÆR' "
				Case "etc"		addSql = addSql & " and not p.company_name between '¤¡' and 'ÆR' "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(c.userid) as cnt, CEILING(CAST(Count(c.userid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner p on c.userid = p.id  "
		sqlStr = sqlStr & " LEFT JOIN [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U on c.userid=U.lecturer_id  "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and c.userdiv='14' "
		sqlStr = sqlStr & " and c.isusing='Y' "
		sqlStr = sqlStr & " and c.userid <> '' "
		sqlStr = sqlStr & " and isnull(U.lec_yn, '') <> '' "
		sqlStr = sqlStr & " and isnull(U.diy_yn, '') <> '' "
		sqlStr = sqlStr & addSql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'ÁöÁ¤ÆäÀÌÁö°¡ ÀüÃ¼ ÆäÀÌÁöº¸´Ù Å¬ ¶§ ÇÔ¼öÁ¾·á
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & "	p.id, p.company_name, c.socname, c.socname_kor, U.lec_yn, U.diy_yn "
		sqlStr = sqlStr & " FROM [db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner p on c.userid = p.id  "
		sqlStr = sqlStr & " LEFT JOIN [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U on c.userid=U.lecturer_id  "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and p.userdiv='9999'  "
		sqlStr = sqlStr & " and c.userdiv='14' "
		sqlStr = sqlStr & " and c.isusing='Y' "
		sqlStr = sqlStr & " and isnull(U.lec_yn, '') <> '' "
		sqlStr = sqlStr & " and isnull(U.diy_yn, '') <> '' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER by c.userid ASC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CEventItem
					FItemList(i).FId			= rsget("id")
					FItemList(i).FCompany_name	= db2html(rsget("company_name"))
					FItemList(i).FSocname		= db2html(rsget("socname"))
					FItemList(i).FSocname_kor	= db2html(rsget("socname_kor"))
					FItemList(i).FLec_yn		= rsget("lec_yn")
					FItemList(i).FDiy_yn		= rsget("diy_yn")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'ÁøÇàÁßÀÎ °­ÁÂ ¸®½ºÆ®
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
		sqlStr = sqlStr & " i.idx, i.lec_title, i.lec_cost, i.lec_count, i.lec_period, i.lec_startday1, i.limit_count, i.limit_sold, i.lecturer_id "
		sqlStr = sqlStr & " ,i.mat_cost, i.matinclude_yn "
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
					FItemList(i).FLecturer_name	  		= rsACADEMYget("lecturer_name")			'°­»ç ÀÌ¸§
					FItemList(i).Flecturer_regdate		= rsACADEMYget("lecturer_regdate")
					FItemList(i).flecturercouponyn      = rsACADEMYget("lecturercouponyn")
					FItemList(i).Flecturercoupontype    = rsACADEMYget("lecturercoupontype")
					FItemList(i).FLecTitle				= db2html(rsACADEMYget("lec_title"))
					FItemList(i).FlecImgProfile75	  	= rsACADEMYget("image_profile_75x75")	'°­»ç ¾ÆÀÌÄÜ(75x75)
					FItemList(i).Flecturercouponvalue   = rsACADEMYget("lecturercouponvalue")
					FItemList(i).fcurrlecturercouponidx = rsACADEMYget("currlecturercouponidx")
					FItemList(i).FOptionCnt 			= rsACADEMYget("optionCnt")
					FItemList(i).FOptLimitCnt 			= rsACADEMYget("optLimitCnt")
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close
	End Sub

	'ÆÇ¸ÅÁßÀÎ ÀÛÇ° ¸®½ºÆ®
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

	'ÀÌº¥Æ® ¼öÁ¤
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
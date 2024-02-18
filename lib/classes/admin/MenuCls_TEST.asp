<%

function fnGetMenuPos(byval menuid, byref menuposnotice, byref menuposhelp)
	dim sqlStr,menustr
	dim pid
    if (menuid<2) then Exit function

	sqlStr = "select id,menuname,parentid,menuposnotice,menuposhelp from [db_partner].[dbo].tbl_partner_menu"
	sqlStr = sqlStr + " where id=" + CStr(menuid)

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		menustr         = db2html(rsget("menuname"))
		pid             = rsget("parentid")
		menuposnotice  = db2html(rsget("menuposnotice"))
		menuposhelp    = db2html(rsget("menuposhelp"))

		if IsNULL(menuposnotice) then Fmenuposnotice=""
		if IsNULL(menuposhelp) then Fmenuposhelp=""
	end if
	rsget.Close

	if pid=0 then
		menustr = "&gt;&gt;" + menustr
	else
		sqlStr = "select id,menuname,parentid, menuposnotice from [db_partner].[dbo].tbl_partner_menu"
		sqlStr = sqlStr + " where id=" + CStr(pid)

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			menustr = db2html(rsget("menuname")) + "&gt;&gt;" + menustr
			pid     = rsget("parentid")
		end if
		rsget.Close
	end if

	fnGetMenuPos = menustr
end function



Class CMenuListItem
	public Fmenu_id
	public Fmenu_name
	public Fmenu_linkurl
	public Fmenu_parentid
	public Fmenu_color
	public Fmenu_isusing
	public Fmenu_viewidx
	public Fmenu_cnt
    public Fmenu_name_En

	public FLastMenu
	public FChildCount
	public FChildItem()

	public Fcid
	public Fpid

    ''기존권한 - 업체, 가맹점, Zoom등
    public Fmenu_divcd

	Private Sub Class_Initialize()
		FLastMenu = true
		FChildCount = 0
		redim FChildItem(0)
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public function GetMenuColor()
		if IsNULL(Fmenu_color) or (Fmenu_color="") then
			GetMenuColor = "#000000"
		else
			GetMenuColor = Fmenu_color
		end if
	end function

	public sub AddChild(byval ichild)
		dim cnt
		cnt = UBound(FChildItem)
		if FChildCount<1 then
			set FChildItem(0) = ichild
		else
			redim preserve FChildItem(cnt+1)
			'redim  FChildItem(cnt+1)
			FChildItem(cnt).FlastMenu=false
			set FChildItem(cnt+1) = ichild
		end if
		FChildCount = FChildCount+1
	end sub


	public function IsHasChild()
		IsHasChild = FChildCount>0
	end function

end Class


Class CMenuList
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectshopdiv
	public FRectPid
	public FRectMid
	public FRectsearchKey
	public FRectsearchString
	public FRectisUsing
	public FRectPart_sn
	public FRectLevel_sn
	public FRectUserDiv

	public FMenuitemlist()
	public FMenuCount

    public FRectUserID
    public FRectUsingEnMenuName
	public FRectGroup_sn

	public FRectIsFavorite
	public FRectHasAdminAuth

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

		FMenuCount =0
		redim  FMenuitemlist(0)
		FMenuitemlist(0) = null
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'##### 사용자 목록 접수 #####
	public Sub GetMenuList()
		dim SQL, AddSQL, i, strTemp

		'// 검색어 쿼리 //
		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and t1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectisUsing<>"" and FRectisUsing<>"all" then
			AddSQL = AddSQL & " and t1.isUsing = '" & FRectisUsing & "' "
		end if

		'// 개수 파악 //
		SQL =	"Select count(id), CEILING(CAST(Count(id) AS FLOAT)/" & FPageSize & ") " &_
				"From db_partner.[dbo].tbl_partner_menu as t1 " &_
				"where parentid=" & FRectPid & AddSQL
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	"select top " & CStr(FPageSize*FCurrPage) &_
				"	 t1.id, t1.menuname, t1.linkurl, t1.parentid " &_
				"	, t1.menucolor, t1.isusing, t1.viewidx, t1.divcd, t1.menuname_En " &_
				"	, t2.menu_cnt " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"		Left Join (Select parentid, count(*) as menu_cnt " &_
				"					from db_partner.[dbo].tbl_partner_menu " &_
				"					where isusing='Y' Group by parentid) as t2 " &_
				"			on t1.id=t2.parentid " &_
				"where t1.parentid=" & FRectPid & AddSQL &_
				"Order by t1.parentid, t1.viewidx "
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMenuListItem

				FItemList(i).Fmenu_id		= rsget("id")
				FItemList(i).Fmenu_name		= rsget("menuname")
				FItemList(i).Fmenu_linkurl	= rsget("linkurl")
				FItemList(i).Fmenu_parentid	= rsget("parentid")
				FItemList(i).Fmenu_color	= rsget("menucolor")
				FItemList(i).Fmenu_isusing	= rsget("isusing")
				FItemList(i).Fmenu_viewidx	= rsget("viewidx")
				FItemList(i).Fmenu_cnt		= rsget("menu_cnt")

                FItemList(i).Fmenu_divcd    = rsget("divcd")
                FItemList(i).Fmenu_name_En   = rsget("menuname_En")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function


	'##### 사용자 내용 접수 #####
	public Sub GetMenuCont()
		dim SQL

		'// 목록 접수 //
		SQL =	"select " &_
				"	 t1.id, t1.menuname, t1.linkurl, t1.parentid " &_
				"	, t1.menucolor, t1.isusing, t1.viewidx, t1.divcd, t1.menuname_En " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"where t1.id=" & FRectMid
		rsget.Open SQL,dbget,1

		if Not(rsget.EOF or rsget.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CMenuListItem

			FItemList(1).Fmenu_id		= rsget("id")
			FItemList(1).Fmenu_name		= rsget("menuname")
			FItemList(1).Fmenu_linkurl	= rsget("linkurl")
			FItemList(1).Fmenu_parentid	= rsget("parentid")
			FItemList(1).Fmenu_color	= rsget("menucolor")
			FItemList(1).Fmenu_isusing	= rsget("isusing")
			FItemList(1).Fmenu_viewidx	= rsget("viewidx")

			FItemList(1).Fmenu_divcd    = rsget("divcd")
			FItemList(1).Fmenu_name_En		= rsget("menuname_En")
		else
			FResultCount = 0
		end if

		rsget.Close

	end Sub


	'##### 왼쪽 메뉴 목록 #####
	'// GetLeftMenuListNew() 로 대체합니다.
	public Sub GetLeftMenuList()
		dim SQL, AddSQL, i, strTemp
		dim onemenuitem, tmp

		'관리자등급(Level:1)은 모든 메뉴 출력
		if FRectLevel_sn=1 then
			AddSQL = "1=1"
		else
			if NOT(FRectLevel_sn="" or isNull(FRectLevel_sn)) then
				AddSQL = "part_sn in (1, '" & FRectPart_sn & "')" &_
						" and level_sn>='" & FRectLevel_sn & "'"
				'※ 부서번호 1 : 부서전체

				''추가 권한 관련 2011-09-19
				''특정부서의 특정권한을 추가한다. : 특정부서+파트선임권한 을 추가해도, 부서전체 파트선임권한은 제외된다.
				if (FRectUserID<>"") then
				    AddSQL = AddSQL & " OR menu_id in ("
				    AddSQL = AddSQL & "     select menu_id from db_partner.dbo.tbl_menu_part p"
                	AddSQL = AddSQL & "     Join db_partner.dbo.tbl_partner_AddLevel L"
                	AddSQL = AddSQL & "     on L.userid='"&FRectUserID&"'"
                	AddSQL = AddSQL & "     and L.isDefault<>'Y'"
                	AddSQL = AddSQL & "     and p.part_sn=L.part_sn"
                	AddSQL = AddSQL & "     and p.level_sn>=L.level_sn"
				    AddSQL = AddSQL & " )"
				end if
			else
				'권한이 없으면 메뉴 표시 없음
				AddSQL = " level_sn>9999 "
			end if
		end if

		'// 목록 접수 //
		'' SQL =	"Select " &_
		'' 		"	 t1.id, t1.menuname, t1.linkurl, t1.parentid, t1.menucolor, t1.menuname_En " &_
		'' 		"from db_partner.[dbo].tbl_partner_menu as t1 " &_
		'' 		"		Join (Select distinct menu_id " &_
		'' 		"				from db_partner.dbo.tbl_menu_part " &_
		'' 		"				where " & AddSQL &_
		'' 		"			) as t2 " &_
		'' 		"			on t1.id=t2.menu_id " &_
		'' 		"Where t1.isusing='Y' " &_
		'' 		"Order by t1.parentid, t1.viewidx "

		SQL = " Select "
		SQL = SQL + " 	 t1.id, t1.menuname, t1.linkurl, t1.parentid, t1.menucolor, t1.menuname_En "
		SQL = SQL + " from db_partner.[dbo].tbl_partner_menu as t1 "
		SQL = SQL + " 		Join (Select distinct menu_id "
		SQL = SQL + " 				from db_partner.dbo.tbl_menu_part "
		SQL = SQL + " 				where " + AddSQL
		SQL = SQL + " 			) as t2 "
		SQL = SQL + " 			on t1.id=t2.menu_id "

		if (FRectSearchString <> "") then
			SQL = SQL + " 		left join ( "
			SQL = SQL + " 			select p.id, count(*) as cnt "
			SQL = SQL + " 			from "
			SQL = SQL + " 			db_partner.[dbo].tbl_partner_menu p "
			SQL = SQL + " 			left join db_partner.[dbo].tbl_partner_menu c "
			SQL = SQL + " 			on "
			SQL = SQL + " 				p.id = c.parentid "
			SQL = SQL + " 			where "
			SQL = SQL + " 				1 = 1 "
			SQL = SQL + " 				and p.parentid = 0 "
			SQL = SQL + " 				and p.isusing = 'Y' "
			SQL = SQL + " 				and c.isusing = 'Y' "
			SQL = SQL + " 				and c.menuname like '%" + CStr(FRectSearchString) + "%' "
			SQL = SQL + " 			group by "
			SQL = SQL + " 				p.id "
			SQL = SQL + " 		) TT "
			SQL = SQL + " 		on TT.id = t1.id "
			SQL = SQL + " left join db_partner.[dbo].tbl_partner_menu p "
			SQL = SQL + " on "
			SQL = SQL + " 	p.id = t1.parentid "
		end if

		SQL = SQL + " Where t1.isusing='Y' "

		if (FRectSearchString <> "") then
			SQL = SQL + " and ( "
			SQL = SQL + " 	(IsNull(TT.cnt, 0) > 0) "
			SQL = SQL + " 	or "
			SQL = SQL + " 	(IsNull(p.menuname, '') like '%" + CStr(FRectSearchString) + "%') "
			SQL = SQL + " 	or "
			SQL = SQL + " 	(t1.menuname like '%" + CStr(FRectSearchString) + "%') "
			SQL = SQL + " ) "
		end if

		SQL = SQL + " Order by t1.parentid, t1.viewidx "

		''rw SQL
		rsget.Open SQL,dbget,1

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set onemenuitem = new CMenuListItem

				onemenuitem.Fmenu_id		= rsget("id")
				onemenuitem.Fmenu_name		= db2html(rsget("menuname"))
				onemenuitem.Fmenu_linkurl	= db2html(rsget("linkurl"))
				onemenuitem.Fmenu_parentid	= rsget("parentid")
				onemenuitem.Fmenu_color		= db2html(rsget("menucolor"))
                onemenuitem.Fmenu_name_En    = db2html(rsget("menuname_En"))

				'// 하위메뉴 목록 저장
				if onemenuitem.Fmenu_parentid=0 then
					AddChild onemenuitem
				else
					set tmp = getParentMenu ( onemenuitem.Fmenu_parentid )
					if Not(tmp is Nothing) then
						tmp.addChild onemenuitem
					end if
				end if

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Sub GetLeftMenuListNew()
		dim strSql, i, strTemp
		dim onemenuitem, tmp
		dim tmpPart_sn, tmpLevel_sn, tmpGroup_sn

		tmpPart_sn 		= FRectPart_sn
		tmpLevel_sn 	= FRectLevel_sn
		tmpGroup_sn 	= "9999"

		if IsNull(tmpPart_sn) then
			tmpPart_sn = "0"
		end if

		if tmpPart_sn = "" then
			tmpPart_sn = "0"
		end if

		if IsNull(tmpLevel_sn) then
			tmpLevel_sn = "0"
		end if

		if tmpLevel_sn = "" then
			tmpLevel_sn = "0"
		end if

		if IsNull(tmpGroup_sn) then
			tmpGroup_sn = "0"
		end if

		if tmpGroup_sn = "" then
			tmpGroup_sn = "0"
		end if

		''그룹코드를 이용해 메뉴를 표시하는 것은 일단 보류
		''GetLeftMenuList() 참조(skyer9)

		strSql = " exec db_partner.dbo.usp_Ten_LeftMenu " + CStr(tmpPart_sn) + ", " + CStr(tmpLevel_sn) + ", " + CStr(tmpGroup_sn) + ", '" + CStr(FRectUserID) + "', '" + CStr(FRectSearchString) + "', '" + CStr(FRectIsFavorite) + "', '" + CStr(FRectHasAdminAuth) + "' "
		''rw strSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = rsget.RecordCount

			if FResultCount<1 then FResultCount=0

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				do until rsget.eof
					set FItemList(i) = new CMenuListItem

					FItemList(i).Fmenu_id		= rsget("id")
					FItemList(i).Fmenu_name		= db2html(rsget("menuname"))
					FItemList(i).Fmenu_linkurl	= db2html(rsget("linkurl"))
					FItemList(i).Fmenu_parentid	= rsget("parentid")
					FItemList(i).Fmenu_color	= db2html(rsget("menucolor"))
					FItemList(i).Fmenu_name_En  = db2html(rsget("menuname_En"))

					FItemList(i).Fcid			= rsget("cid")
					FItemList(i).Fpid			= rsget("pid")

					rsget.moveNext
					i=i+1
				loop
			end if
		rsget.Close
	end Sub

''2013-03-11 진영 생성
	Public Function GROUP_PSN()
		Dim strSQL
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 p.part_sn, G.group_sn, G.part_sn as groupPart_sn " & VBCRLF
		strSQL = strSQL & " FROM db_partner.dbo.tbl_partInfo as P " & VBCRLF
		strSQL = strSQL & " LEFT JOIN db_partner.dbo.tbl_partInfoGroup as G " & VBCRLF
		strSQL = strSQL & " ON P.part_sn = G.part_sn  " & VBCRLF
		strSQL = strSQL & " WHERE p.part_sn = '"&FRectPart_sn&"' "
		rsget.Open strSQL,dbget,1
		IF not rsget.EOF Then
			GROUP_PSN = rsget("part_sn") & "," & rsget("group_sn")
		End If
		rsget.Close
	End Function

	'##### 왼쪽 메뉴 목록(OFFLine용) #####
	public Sub GetLeftMenuList_offLine()
		dim SQL, AddSQL, i, strTemp
		dim onemenuitem, tmp

		if CStr(FRectUserDiv)="509" then
			'// 매출조회용이라면, 매장통계, 기간별매출통계, 시간대별매출분석, 요일별매출분석만 나오게
			AddSQL = AddSQL & " and id in (501, 508, 511, 512)"
		end if

		'//도매일경우 매장통계는 제낌
		if FRectshopdiv = "5" then
			AddSQL = AddSQL & " and id not in (501)"
		end if

		'//해외가 아닐경우, [해외] 출고조회, 해외매장 상품설정은 제낌
		if FRectshopdiv <> "7" then
			IF application("Svr_Info")="Dev" THEN
				AddSQL = AddSQL & " and id not in (1477,1211)"
			else
				AddSQL = AddSQL & " and id not in (1391,1210)"
			end if
		end if

		'//도매,해외,its일 경우 매장게시판관리는 제낌
		if FRectshopdiv = "5" or FRectshopdiv = "7" or FRectshopdiv = "11" or FRectshopdiv = "13" then
			AddSQL = AddSQL & " and id not in (524)"
		end if

		'// 목록 접수 //
		SQL =	"Select " &_
				"	 t1.id, t1.menuname, t1.linkurl, t1.parentid, t1.menucolor, t1.menuname_en " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"Where t1.isusing='Y' and divCD in ('500','" & CStr(FRectUserDiv) & "')" & AddSQL &_
				"Order by t1.parentid, t1.viewidx "
		rsget.Open SQL,dbget,1

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set onemenuitem = new CMenuListItem

				onemenuitem.Fmenu_id		= rsget("id")
				onemenuitem.Fmenu_name		= db2html(rsget("menuname"))
				onemenuitem.Fmenu_linkurl	= db2html(rsget("linkurl"))
				onemenuitem.Fmenu_parentid	= rsget("parentid")
				onemenuitem.Fmenu_color		= db2html(rsget("menucolor"))
                onemenuitem.Fmenu_name_En		= db2html(rsget("menuname_en"))

                ''영문메뉴 추가.
                if (FRectUsingEnMenuName="on") then
                    if Not isNULL(onemenuitem.Fmenu_name_En) and (onemenuitem.Fmenu_name_En<>"") then
                        onemenuitem.Fmenu_name = onemenuitem.Fmenu_name_En
                    end if
                end if

				'// 하위메뉴 목록 저장
				if onemenuitem.Fmenu_parentid=0 then
					AddChild onemenuitem
				else
					set tmp = getParentMenu ( onemenuitem.Fmenu_parentid )
					if Not(tmp is Nothing) then
						tmp.addChild onemenuitem
					end if
				end if

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close
	end Sub

	public sub AddChild(byval ichild)
		dim cnt
		cnt = UBound(FMenuitemlist)
		if FMenuCount<1 then
			set FMenuitemlist(0) = ichild
		else
			redim preserve FMenuitemlist(cnt+1)
			FMenuitemlist(cnt).FlastMenu=false
			set FMenuitemlist(cnt+1) = ichild
		end if
		FMenuCount = FMenuCount+1

	end sub

	public function getParentMenu(byval iid)
		dim i
		set getParentMenu = Nothing

		for i=0 to Ubound(FMenuitemlist)
			if (CStr(FMenuitemlist(i).Fmenu_id) = CStr(iid) )then
				set getParentMenu  = FMenuitemlist(i)
				Exit for
			end if
		next

	end function

end Class

'/// 부서 옵션 생성 함수 ///
public function printPartOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select class='select' name='" & fnm & "'>" &_
				"<option value=''>::부서선택::</option>"

	SQL =	"Select part_sn, part_name " &_
			"From db_partner.dbo.tbl_partInfo " &_
			"Where part_isDel='N' " &_
			"Order by part_sort"
	rsget.Open SQL,dbget,1
if IsNULL(psn) then psn=""
	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("part_sn") & "'"
			if Cstr(rsget("part_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("part_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printPartOption = strOpt
end function

'/// 직급 옵션 생성 함수 ///
public function printPositOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select class='select' name='" & fnm & "'>" &_
				"<option value=''>::직급선택::</option>"

	SQL =	"Select posit_sn, posit_name " &_
			"From db_partner.dbo.tbl_positInfo " &_
			"Where posit_isDel='N' "
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("posit_sn") & "'"
			if Cstr(rsget("posit_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("posit_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printPositOption = strOpt
end function

'/// 등급 옵션 생성 함수 ///
public function printLevelOption(fnm, psn)
	dim SQL, i, strOpt

	strOpt =	"<select class='select' name='" & fnm & "'>" &_
				"<option value=''>::등급선택::</option>"

	SQL =	"Select level_sn, level_name " &_
			"From db_partner.dbo.tbl_level " &_
			"Where level_isDel='N' " &_
			"Order by level_no"
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("level_sn") & "'"
			if Cstr(rsget("level_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("level_name") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printLevelOption = strOpt
end function

'// 지정 부서/등급 정보 접수 //
public function getPartLevelInfo(pid,mode)
	dim SQL, i, strPrt

	SQL =	"Select t2.part_name, t1.part_sn, t3.level_name, t1.level_sn " &_
			"From db_partner.dbo.tbl_menu_part as t1 " &_
			"	join db_partner.dbo.tbl_partInfo as t2 " &_
			"		on t1.part_sn=t2.part_sn " &_
			"	join db_partner.dbo.tbl_level as t3 " &_
			"		on t1.level_sn=t3.level_sn " &_
			"Where t2.part_isDel='N' and t3.level_isDel='N' " &_
			"	and t1.menu_id=" & pid & " " &_
			"Order by t2.part_sort"
	rsget.Open SQL,dbget,1

	strPrt = ""
	if mode="modi" then strPrt = strPrt & "<table name='tbl_auth' id='tbl_auth' class=a>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			if mode="list" then
				strPrt = strPrt & rsget(0) & "/" & rsget(2) & "<br>"
			elseif mode="modi" then
				strPrt = strPrt &_
					"<tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>" &_
						"<td>" & rsget(0) & "<input type='hidden' name='part_sn' value='" & rsget(1) & "'></td>" &_
						"<td>" & rsget(2) & "<input type='hidden' name='level_sn' value='" & rsget(3) & "'></td>" &_
						"<td><img src='http://photoimg.10x10.co.kr/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle></td>" &_
					"</tr>"
			end if
			i = i + 1
		rsget.MoveNext
		Loop
	else
		if mode="modi" then
			strPrt = strPrt &_
					"<tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>" &_
						"<td><input type='hidden' name='part_sn' value=''></td>" &_
						"<td><input type='hidden' name='level_sn' value=''></td>" &_
						"<td></td>" &_
					"</tr>"
		end if
	end if
	if mode="modi" then strPrt = strPrt & "</table>"

	'결과값 반환
	getPartLevelInfo = strPrt

	rsget.Close
end Function

'// 루트메뉴 옵션생성 접수 //
public function printRootMenuOption(fnm,pid,atn)
	dim SQL, i, strOpt

	strOpt = "<select class='select' name='" & fnm & "'"
	if atn="Action" then strOpt = strOpt & " onchange='form.submit()'"
	strOpt = strOpt & "><option value=''>::메뉴선택::</option>"

	'루트메뉴 추가
	if pid="0" and atn="NoAction" then
		strOpt = strOpt & "<option value='0' selected>루트메뉴</option>"
	else
		strOpt = strOpt & "<option value='0'>메뉴루트</option>"
	end if

	SQL =	"Select id, menuname " &_
			"From db_partner.[dbo].tbl_partner_menu " &_
			"Where parentid=0 and isusing='Y' " &_
			"Order by viewidx"
	rsget.Open SQL,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			strOpt = strOpt & "<option value='" & rsget("id") & "'"
			if cLng(rsget("id"))=cLng(pid) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget("menuname") & "</option>"
		rsget.MoveNext
		Loop
	end if

	rsget.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printRootMenuOption = strOpt
end Function
%>

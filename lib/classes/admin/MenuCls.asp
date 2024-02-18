<%

function fnGetMenuPos(byval menuid, byref menuposnotice, byref menuposhelp)
	dim sqlStr,menustr
	dim pid
	menuid = LEFT(menuid,9)  ''2017/04/20 추가
    if (menuid<2) then Exit function

	sqlStr = "select id,menuname,parentid,menuposnotice,menuposhelp from [db_partner].[dbo].tbl_partner_menu WITH(NOLOCK)"
	sqlStr = sqlStr + " where id=" + CStr(menuid)

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
		sqlStr = "select id,menuname,parentid, menuposnotice from [db_partner].[dbo].tbl_partner_menu WITH(NOLOCK)"
		sqlStr = sqlStr + " where id=" + CStr(pid)

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			menustr = db2html(rsget("menuname")) + "&gt;&gt;" + menustr
			pid     = rsget("parentid")
		end if
		rsget.Close
	end if

	fnGetMenuPos = menustr
end function

function fnGetMenuFavoriteAdded(userid, menuid)
	dim sqlStr

	fnGetMenuFavoriteAdded = False

	if Not IsNumeric(menuid) then
		Exit function
	end if

	sqlStr = "select top 1 menu_id from db_partner.dbo.tbl_partner_menu_favorite WITH(NOLOCK) "
	sqlStr = sqlStr + " where userid = '" + CStr(userid) + "' and menu_id = " + CStr(menuid) + " and useYN = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		fnGetMenuFavoriteAdded = True
	end if
	rsget.Close

end function

Class CMenuListItem
    public Fmenu_id
    public Fmenu_name
	public Fmenu_name_parent
    public Fmenu_linkurl
    public Fmenu_parentid
    public Fmenu_color
	public Fmenu_isusing
	public Fmenu_viewidx
	public Fmenu_cnt
    public Fmenu_name_En
    public Fmenu_useSslYN
    public Fmenu_criticinfo
	public Fmenu_saveLog
	public Flv1customerYN
	public Flv2partnerYN
	public Flv3InternalYN

	public FLastMenu
	public FChildCount
	public FChildItem()

	public Fcid
	public Fpid

    ''기존권한 - 업체, 가맹점, Zoom등
    public Fmenu_divcd

    public function getOldMenuDivStr
        SELECT CASE Fmenu_divcd
            CASE "9999" : getOldMenuDivStr = "업체"
            CASE "999" : getOldMenuDivStr = "제휴사"
            CASE "9000" : getOldMenuDivStr = "강사"
            CASE "9","7","5","4","2","1" : getOldMenuDivStr = "직원"

            CASE "500" : getOldMenuDivStr = "매장공통"
            CASE "501" : getOldMenuDivStr = "직영매장"
            CASE "502" : getOldMenuDivStr = "수수료매장"
            CASE "503" : getOldMenuDivStr = "대리점"
            CASE "101" : getOldMenuDivStr = "오프샵"
            CASE "111" : getOldMenuDivStr = "오프샵점장"
            CASE "112" : getOldMenuDivStr = "오프샵부점장"
            CASE "509" : getOldMenuDivStr = "오프매출조회"
            CASE "201" : getOldMenuDivStr = "Zoom"
            CASE "301" : getOldMenuDivStr = "College"
            CASE ELSE : getOldMenuDivStr = "?"
        END SELECT
    end function

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

Class CMenuPrivListItem
    public Fmenu_id
    public Fmenu_name
	public Fmenu_name_parent
    public Fmenu_parentid
	public Fmenu_isusing
	public Fmenu_viewidx
    public Fmenu_useSslYN
    public Fmenu_criticinfo
	public Fmenu_saveLog
	public Flv1customerYN
	public Flv2partnerYN
	public Flv3InternalYN

	public Fmenu_part_sn1
	public Fmenu_part_sn16
	public Fmenu_part_sn14
	public Fmenu_part_sn11
	public Fmenu_part_sn21
	public Fmenu_part_sn12
	public Fmenu_part_sn23
	public Fmenu_part_sn13
	public Fmenu_part_sn24
	public Fmenu_part_sn30
	public Fmenu_part_sn7
	public Fmenu_part_sn9
	public Fmenu_part_sn10
	public Fmenu_part_sn8
	public Fmenu_part_sn20
	public Fmenu_part_sn17
	public Fmenu_part_sn22
	public Fmenu_part_sn33
	public Fmenu_part_sn25
	public Fmenu_part_sn26
	public Fmenu_part_sn27
	public Fmenu_part_sn28
	public Fmenu_part_sn29
	public Fmenu_part_sn34
	public Fmenu_part_sn_etc

    ''기존권한 - 업체, 가맹점, Zoom등
    public Fmenu_divcd

    public function getOldMenuDivStr
        SELECT CASE Fmenu_divcd
            CASE "9999" : getOldMenuDivStr = "업체"
            CASE "999" : getOldMenuDivStr = "제휴사"
            CASE "9000" : getOldMenuDivStr = "강사"
            CASE "9","7","5","4","2","1" : getOldMenuDivStr = "직원"

            CASE "500" : getOldMenuDivStr = "매장공통"
            CASE "501" : getOldMenuDivStr = "직영매장"
            CASE "502" : getOldMenuDivStr = "수수료매장"
            CASE "503" : getOldMenuDivStr = "대리점"
            CASE "101" : getOldMenuDivStr = "오프샵"
            CASE "111" : getOldMenuDivStr = "오프샵점장"
            CASE "112" : getOldMenuDivStr = "오프샵부점장"
            CASE "509" : getOldMenuDivStr = "오프매출조회"
            CASE "201" : getOldMenuDivStr = "Zoom"
            CASE "301" : getOldMenuDivStr = "College"
            CASE ELSE : getOldMenuDivStr = "?"
        END SELECT
    end function

	Private Sub Class_Initialize()
		''
	End Sub

	Private Sub Class_Terminate()
		''
	End Sub

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

    public FRectuseSslYN
    public FRectcriticinfo
	public FRectSaveLog

	public FRectlv1customerYN
	public FRectlv2partnerYN
	public FRectlv3InternalYN

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

	'##### 기타>>메뉴관리 #####
	'// !!!! GetMenuListNew 사용할 것
' 	public Sub GetMenuList()
' 		dim SQL, AddSQL, i, strTemp
' 		dim addSqlJoin

' 		'// 검색어 쿼리 //
' 		if (FRectPid="0") then
' 		   '' AddSQL = AddSQL & " and t1.parentid="&FRectPid&""
' 		else
' 		    AddSQL = AddSQL & " and t1.parentid="&FRectPid&""
' 		end if

' 		if FRectsearchString<>"" then
' 			AddSQL = AddSQL & " and t1." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
' 		end if

' 		if FRectisUsing<>"" and FRectisUsing<>"all" then
' 			AddSQL = AddSQL & " and t1.isUsing = '" & FRectisUsing & "' "
' 		end if

'         if FRectuseSslYN<>"" then
'             AddSQL = AddSQL & " and t1.useSslYN = '" & FRectuseSslYN & "' "
'         end if

'         if FRectcriticinfo<>"" then
'             AddSQL = AddSQL & " and t1.criticinfo = '" & FRectcriticinfo & "' "
'         end if
' 		if FRectlv1customerYN<>"" then
'             AddSQL = AddSQL & " and t1.lv1customerYN = '"&FRectlv1customerYN&"' "
' 		End If
' 		if FRectlv2partnerYN<>"" then
'             AddSQL = AddSQL & " and t1.lv2partnerYN = '"&FRectlv2partnerYN&"' "
' 		End If		
' 		if FRectlv3InternalYN<>"" then
'             AddSQL = AddSQL & " and t1.lv3InternalYN = '"&FRectlv3InternalYN&"' "
' 		End If

' 		addSqlJoin = ""
' 		if (FRectPart_sn <> "" or FRectLevel_sn <> "") then
' 			addSqlJoin = addSqlJoin + " 	join ( "
' 			addSqlJoin = addSqlJoin + " 		Select t1.menu_id, count(t1.part_sn) as part_snCnt, count(t1.level_sn) as level_snCnt "
' 			addSqlJoin = addSqlJoin + " 		From "
' 			addSqlJoin = addSqlJoin + " 			db_partner.dbo.tbl_menu_part as t1 WITH(NOLOCK) "
' 			addSqlJoin = addSqlJoin + " 			join db_partner.dbo.tbl_partInfo as t2 WITH(NOLOCK) "
' 			addSqlJoin = addSqlJoin + " 			on  "
' 			addSqlJoin = addSqlJoin + " 				t1.part_sn=t2.part_sn "
' 			addSqlJoin = addSqlJoin + " 			join db_partner.dbo.tbl_level as t3 WITH(NOLOCK) "
' 			addSqlJoin = addSqlJoin + " 			on "
' 			addSqlJoin = addSqlJoin + " 				t1.level_sn=t3.level_sn "
' 			addSqlJoin = addSqlJoin + " 		Where "
' 			addSqlJoin = addSqlJoin + " 			1 = 1 "
' 			addSqlJoin = addSqlJoin + " 			and t2.part_isDel='N' and t3.level_isDel='N' "

' 			if (FRectPart_sn <> "") then
' 				addSqlJoin = addSqlJoin + " 			and t1.part_sn = " + CStr(FRectPart_sn) + " "
' 			end if

' 			if (FRectLevel_sn <> "") then
' 				addSqlJoin = addSqlJoin + " 			and t1.level_sn = " + CStr(FRectLevel_sn) + " "
' 			end if

' 			addSqlJoin = addSqlJoin + " 		group by t1.menu_id "
' 			addSqlJoin = addSqlJoin + " 	) S "
' 			addSqlJoin = addSqlJoin + " 	on "
' 			addSqlJoin = addSqlJoin + " 		t1.id = S.menu_id "
' 		end if

' 		'// 개수 파악 //
' 		SQL =	"Select count(id), CEILING(CAST(Count(id) AS FLOAT)/" & FPageSize & ") " & Vbcrlf
' 		SQL = SQL &" From db_partner.[dbo].tbl_partner_menu as t1 WITH(NOLOCK) " &  Vbcrlf
' 		SQL = SQL & addSqlJoin &  Vbcrlf
' 		SQL = SQL &" where 1=1 " & AddSQL
' 		rsget.CursorLocation = adUseClient
' 		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
' 			FTotalCount = rsget(0)
' 			FtotalPage = rsget(1)
' 		rsget.Close

' 		'// 목록 접수 //
' 		SQL =	"select top " & CStr(FPageSize*FCurrPage) & Vbcrlf
' 		SQL = SQL &"	 t1.id, t1.menuname, t1.linkurl, t1.parentid " & Vbcrlf
' 		SQL = SQL &"	, t1.menucolor, t1.isusing, t1.viewidx, t1.divcd, t1.menuname_En " & Vbcrlf
' 		SQL = SQL &"	, t2.menu_cnt, t1.useSslYN, isNULL(t1.criticinfo,0) as criticinfo, IsNull(p.menuname, '') as parentmenuname " & Vbcrlf
' 		SQL = SQL &"	, ISNULL(t1.lv1customerYN,'N') as lv1customerYN, ISNULL(t1.lv2partnerYN,'N') as lv2partnerYN, ISNULL(t1.lv3InternalYN,'N') as lv3InternalYN" & Vbcrlf		
' 		SQL = SQL &" from db_partner.[dbo].tbl_partner_menu as t1 WITH(NOLOCK) " & Vbcrlf
' 		SQL = SQL &" 		left join db_partner.[dbo].tbl_partner_menu p WITH(NOLOCK) " & Vbcrlf
' 		SQL = SQL &" 		on " & Vbcrlf
' 		SQL = SQL &" 			1 = 1 " & Vbcrlf
' 		SQL = SQL &" 			and t1.parentid = p.id " & Vbcrlf
' 		SQL = SQL &" 			and p.parentid = 0 " & Vbcrlf
' 		SQL = SQL &"		Left Join (Select parentid, count(*) as menu_cnt " & Vbcrlf
' 		SQL = SQL &"					from db_partner.[dbo].tbl_partner_menu WITH(NOLOCK) " & Vbcrlf
' 		SQL = SQL &"					where isusing='Y' Group by parentid) as t2 " & Vbcrlf
' 		SQL = SQL &"			on t1.id=t2.parentid " & Vbcrlf
' 		SQL = SQL & addSqlJoin &  Vbcrlf
' 		SQL = SQL &" where 1=1 " & AddSQL & Vbcrlf
' 		SQL = SQL &" Order by IsNull(p.viewidx, 0), IsNull(p.id, 0), t1.viewidx, t1.id "
' 		rsget.pagesize = FPageSize
' 		rsget.CursorLocation = adUseClient
' 		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
' ''rw SQL
' 		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

' 		if FResultCount<1 then FResultCount=0

' 		redim preserve FItemList(FResultCount)
' 		i=0
' 		if  not rsget.EOF  then
' 			rsget.absolutepage = FCurrPage
' 			do until rsget.eof
' 				set FItemList(i) = new CMenuListItem

' 				FItemList(i).Fmenu_id		= rsget("id")
' 				FItemList(i).Fmenu_name			= rsget("menuname")
' 				FItemList(i).Fmenu_name_parent	= rsget("parentmenuname")

' 				FItemList(i).Fmenu_linkurl	= rsget("linkurl")
' 				FItemList(i).Fmenu_parentid	= rsget("parentid")
' 				FItemList(i).Fmenu_color	= rsget("menucolor")
' 				FItemList(i).Fmenu_isusing	= rsget("isusing")
' 				FItemList(i).Fmenu_viewidx	= rsget("viewidx")
' 				FItemList(i).Fmenu_cnt		= rsget("menu_cnt")

'                 FItemList(i).Fmenu_divcd    = rsget("divcd")
'                 FItemList(i).Fmenu_name_En   = rsget("menuname_En")
'                 FItemList(i).Fmenu_useSslYN		= rsget("useSslYN")
'                 FItemList(i).Fmenu_criticinfo    = rsget("criticinfo")
' 				FItemList(i).Flv1customerYN    = rsget("lv1customerYN")
' 				FItemList(i).Flv2partnerYN    = rsget("lv2partnerYN")
' 				FItemList(i).Flv3InternalYN    = rsget("lv3InternalYN")

' 				if (FItemList(i).Fmenu_parentid = 0) then
' 					FItemList(i).Fmenu_name_parent = FItemList(i).Fmenu_name
' 					FItemList(i).Fmenu_name = ""
' 				end if

' 				rsget.moveNext
' 				i=i+1
' 			loop
' 		end if

' 		rsget.Close

' 	end Sub

	public Sub GetMenuListNew()
		dim SQL, AddSQL, i, strTemp
		dim addSqlJoin

		'// 검색어 쿼리 //
		if (FRectPid="0") then
		   '' AddSQL = AddSQL & " and v.parentid="&FRectPid&""
		else
		    AddSQL = AddSQL & " and v.parentid="&FRectPid&""
		end if

		if FRectsearchString<>"" and FRectsearchKey <> "" then
			AddSQL = AddSQL & " and v." & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if

		if FRectisUsing<>"" and FRectisUsing<>"all" then
			AddSQL = AddSQL & " and v.isUsing = '" & FRectisUsing & "' "
		end if

        if FRectuseSslYN<>"" then
            AddSQL = AddSQL & " and v.useSslYN = '" & FRectuseSslYN & "' "
        end if

        if FRectcriticinfo<>"" then
            AddSQL = AddSQL & " and v.criticinfo = '" & FRectcriticinfo & "' "
        end if

        if FRectSaveLog<>"" then
            AddSQL = AddSQL & " and v.saveLog = '" & FRectSaveLog & "' "
        end if

		if FRectlv1customerYN<>"" then
            AddSQL = AddSQL & " and v.lv1customerYN = '"&FRectlv1customerYN&"' "
		End If
		if FRectlv2partnerYN<>"" then
            AddSQL = AddSQL & " and v.lv2partnerYN = '"&FRectlv2partnerYN&"' "
		End If		
		if FRectlv3InternalYN<>"" then
            AddSQL = AddSQL & " and v.lv3InternalYN = '"&FRectlv3InternalYN&"' "
		End If

		addSqlJoin = ""
		if (FRectPart_sn <> "" or FRectLevel_sn <> "") then
			addSqlJoin = addSqlJoin + " 	join ( "
			addSqlJoin = addSqlJoin + " 		Select t1.menu_id, count(t1.part_sn) as part_snCnt, count(t1.level_sn) as level_snCnt "
			addSqlJoin = addSqlJoin + " 		From "
			addSqlJoin = addSqlJoin + " 			db_partner.dbo.tbl_menu_part as t1 WITH(NOLOCK) "
			addSqlJoin = addSqlJoin + " 			join db_partner.dbo.tbl_partInfo as t2 WITH(NOLOCK) "
			addSqlJoin = addSqlJoin + " 			on  "
			addSqlJoin = addSqlJoin + " 				t1.part_sn=t2.part_sn "
			addSqlJoin = addSqlJoin + " 			join db_partner.dbo.tbl_level as t3 WITH(NOLOCK) "
			addSqlJoin = addSqlJoin + " 			on "
			addSqlJoin = addSqlJoin + " 				t1.level_sn=t3.level_sn "
			addSqlJoin = addSqlJoin + " 		Where "
			addSqlJoin = addSqlJoin + " 			1 = 1 "
			addSqlJoin = addSqlJoin + " 			and t2.part_isDel='N' and t3.level_isDel='N' "

			if (FRectPart_sn <> "") then
				addSqlJoin = addSqlJoin + " 			and t1.part_sn = " + CStr(FRectPart_sn) + " "
			end if

			if (FRectLevel_sn <> "") then
				addSqlJoin = addSqlJoin + " 			and t1.level_sn = " + CStr(FRectLevel_sn) + " "
			end if

			addSqlJoin = addSqlJoin + " 		group by t1.menu_id "
			addSqlJoin = addSqlJoin + " 	) S "
			addSqlJoin = addSqlJoin + " 	on "
			addSqlJoin = addSqlJoin + " 		v.id = S.menu_id "
		end if

		'// 개수 파악 //
		SQL =	"Select count(id), CEILING(CAST(Count(id) AS FLOAT)/" & FPageSize & ") " & Vbcrlf
		SQL = SQL &" From db_partner.[dbo].[vw_partner_menu] AS v WITH(NOLOCK) " &  Vbcrlf
		SQL = SQL & addSqlJoin &  Vbcrlf
		SQL = SQL &" where 1=1 " & AddSQL

		'response.write SQL & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	"select top " & CStr(FPageSize*FCurrPage) & Vbcrlf
		SQL = SQL &"	 v.id, v.menuname, v.linkurl, v.parentid " & Vbcrlf
		SQL = SQL &"	, v.menucolor, v.isusing, v.viewidx, v.divcd, v.menuname_En " & Vbcrlf
		SQL = SQL &"	, t2.menu_cnt, v.useSslYN, isNULL(v.criticinfo,0) as criticinfo, isNULL(v.saveLog,0) as saveLog, IsNull(v.menuname1, '') as parentmenuname " & Vbcrlf
		SQL = SQL &"	, ISNULL(v.lv1customerYN,'N') as lv1customerYN, ISNULL(v.lv2partnerYN,'N') as lv2partnerYN, ISNULL(v.lv3InternalYN,'N') as lv3InternalYN" & Vbcrlf				
		SQL = SQL &" from db_partner.[dbo].[vw_partner_menu] AS v WITH(NOLOCK) " & Vbcrlf
		SQL = SQL &"		Left Join (Select parentid, count(*) as menu_cnt " & Vbcrlf
		SQL = SQL &"					from db_partner.[dbo].tbl_partner_menu WITH(NOLOCK) " & Vbcrlf
		SQL = SQL &"					where isusing='Y' Group by parentid) as t2 " & Vbcrlf
		SQL = SQL &"			on v.id=t2.parentid " & Vbcrlf
		SQL = SQL & addSqlJoin &  Vbcrlf
		SQL = SQL &" where 1=1 " & AddSQL & Vbcrlf
		SQL = SQL &" Order by IsNull(v.viewidx1, 0) ,IsNull(v.id1, 0) desc ,IsNull(v.viewidx2, 0) ,IsNull(v.id2, 0) desc "

		'response.write SQL & "<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMenuListItem

				FItemList(i).Fmenu_id		= rsget("id")
				FItemList(i).Fmenu_name			= rsget("menuname")
				FItemList(i).Fmenu_name_parent	= rsget("parentmenuname")

				FItemList(i).Fmenu_linkurl	= rsget("linkurl")
				FItemList(i).Fmenu_parentid	= rsget("parentid")
				FItemList(i).Fmenu_color	= rsget("menucolor")
				FItemList(i).Fmenu_isusing	= rsget("isusing")
				FItemList(i).Fmenu_viewidx	= rsget("viewidx")
				FItemList(i).Fmenu_cnt		= rsget("menu_cnt")

                FItemList(i).Fmenu_divcd    	= rsget("divcd")
                FItemList(i).Fmenu_name_En   	= rsget("menuname_En")
                FItemList(i).Fmenu_useSslYN		= rsget("useSslYN")
                FItemList(i).Fmenu_criticinfo   = rsget("criticinfo")
				FItemList(i).Fmenu_saveLog    	= rsget("saveLog")
				
				FItemList(i).Flv1customerYN    = rsget("lv1customerYN")
				FItemList(i).Flv2partnerYN    = rsget("lv2partnerYN")
				FItemList(i).Flv3InternalYN    = rsget("lv3InternalYN")

				''if (FItemList(i).Fmenu_criticinfo = 1) then
				''	FItemList(i).Fmenu_criticinfo = "Y"
				''else
				''	FItemList(i).Fmenu_criticinfo = "N"
				''end if

				if (FItemList(i).Fmenu_saveLog = 1) then
					FItemList(i).Fmenu_saveLog = "Y"
				else
					FItemList(i).Fmenu_saveLog = "N"
				end if

				if (FItemList(i).Fmenu_parentid = 0) then
					''FItemList(i).Fmenu_name_parent = FItemList(i).Fmenu_name
					FItemList(i).Fmenu_name = ""
				end if

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	'##### 기타>>메뉴권한관리 #####
	public Sub GetMenuPrivList()
		dim SQL, AddSQL, i, strTemp, strSql
		dim addSqlJoin

		'// 검색어 쿼리 //
		if (FRectPid <> "0") and (FRectPid <> "") then
		    AddSQL = AddSQL & " and c.parentid=" & FRectPid & ""
		end if

		if FRectSearchKey <> "" and FRectSearchString<>"" then
			AddSQL = AddSQL & " and c." & FRectSearchKey & " like '%" & FRectSearchString & "%' "
		end if

		if FRectisUsing<>"" and FRectisUsing<>"all" then
			AddSQL = AddSQL & " and c.isUsing = '" & FRectisUsing & "' "
			AddSQL = AddSQL & " and IsNull(p.isUsing, '" + CStr(FRectisUsing) + "') = '" & FRectisUsing & "' "
		end if

        if FRectuseSslYN<>"" then
            AddSQL = AddSQL & " and c.useSslYN = '" & FRectuseSslYN & "' "
        end if

        if FRectcriticinfo<>"" then
            AddSQL = AddSQL & " and c.criticinfo = '" & FRectcriticinfo & "' "
        end if

		if FRectlv1customerYN<>"" then
            AddSQL = AddSQL & " and c.lv1customerYN = '"&FRectlv1customerYN&"' "
		End If
		if FRectlv2partnerYN<>"" then
            AddSQL = AddSQL & " and c.lv2partnerYN = '"&FRectlv2partnerYN&"' "
		End If		
		if FRectlv3InternalYN<>"" then
            AddSQL = AddSQL & " and c.lv3InternalYN = '"&FRectlv3InternalYN&"' "
		End If

		addSqlJoin = ""
		if (FRectPart_sn <> "" or FRectLevel_sn <> "") then
			addSqlJoin = addSqlJoin + " 	join ( "
			addSqlJoin = addSqlJoin + " 		Select t1.menu_id, count(t1.part_sn) as part_snCnt, count(t1.level_sn) as level_snCnt "
			addSqlJoin = addSqlJoin + " 		From "
			addSqlJoin = addSqlJoin + " 			db_partner.dbo.tbl_menu_part as t1 WITH(NOLOCK) "
			addSqlJoin = addSqlJoin + " 			join db_partner.dbo.tbl_partInfo as t2 WITH(NOLOCK) "
			addSqlJoin = addSqlJoin + " 			on  "
			addSqlJoin = addSqlJoin + " 				t1.part_sn=t2.part_sn "
			addSqlJoin = addSqlJoin + " 			join db_partner.dbo.tbl_level as t3 WITH(NOLOCK) "
			addSqlJoin = addSqlJoin + " 			on "
			addSqlJoin = addSqlJoin + " 				t1.level_sn=t3.level_sn "
			addSqlJoin = addSqlJoin + " 		Where "
			addSqlJoin = addSqlJoin + " 			1 = 1 "
			addSqlJoin = addSqlJoin + " 			and t2.part_isDel='N' and t3.level_isDel='N' "

			if (FRectPart_sn <> "") then
				addSqlJoin = addSqlJoin + " 			and t1.part_sn = " + CStr(FRectPart_sn) + " "
			end if

			if (FRectLevel_sn <> "") then
				addSqlJoin = addSqlJoin + " 			and t1.level_sn = " + CStr(FRectLevel_sn) + " "
			end if

			addSqlJoin = addSqlJoin + " 		group by t1.menu_id "
			addSqlJoin = addSqlJoin + " 	) S "
			addSqlJoin = addSqlJoin + " 	on "
			addSqlJoin = addSqlJoin + " 		c.id = S.menu_id "
		end if


		'// 개수 파악 //
		SQL =	"Select count(c.id), CEILING(CAST(Count(c.id) AS FLOAT)/" & FPageSize & ") " & Vbcrlf
		SQL = SQL &" From db_partner.[dbo].tbl_partner_menu as c WITH(NOLOCK) " &  Vbcrlf
		SQL = SQL  + "		LEFT JOIN db_partner.[dbo].tbl_partner_menu p WITH(NOLOCK) "
		SQL = SQL  + "		ON "
		SQL = SQL  + "			1 = 1 "
		SQL = SQL  + "			AND c.parentid = p.id "
		SQL = SQL  + "			AND p.parentid = 0 "
		SQL = SQL & addSqlJoin &  Vbcrlf
		SQL = SQL &" where 1=1 " & AddSQL
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close


		strSql = " SELECT TOP " + CStr(FPageSize*FCurrPage) + " "
		strSql = strSql + "		c.id "
		strSql = strSql + "		,IsNull(p.menuname, '') AS parentmenuname "
		strSql = strSql + "		,c.menuname "
		strSql = strSql + "		,c.parentid "
		strSql = strSql + "		,c.menucolor "
		strSql = strSql + "		,c.isusing "
		strSql = strSql + "		,c.viewidx "
		strSql = strSql + "		,c.divcd "
		strSql = strSql + "		,c.useSslYN "
		strSql = strSql + "		,isNULL(c.criticinfo, 0) AS criticinfo "
		strSql = strSql & "		,isNULL(c.lv1customerYN, 'N') AS lv1customerYN"		
		strSql = strSql & "		,isNULL(c.lv2partnerYN, 'N') AS lv2partnerYN "
		strSql = strSql & "		,isNULL(c.lv3InternalYN, 'N') AS lv3InternalYN "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 1 and menu_id = c.id) as part_sn1 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 16 and menu_id = c.id) as part_sn16 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 14 and menu_id = c.id) as part_sn14 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 11 and menu_id = c.id) as part_sn11 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 21 and menu_id = c.id) as part_sn21 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 12 and menu_id = c.id) as part_sn12 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 23 and menu_id = c.id) as part_sn23 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 13 and menu_id = c.id) as part_sn13 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 24 and menu_id = c.id) as part_sn24 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 30 and menu_id = c.id) as part_sn30 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 7 and menu_id = c.id) as part_sn7 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 9 and menu_id = c.id) as part_sn9 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 10 and menu_id = c.id) as part_sn10 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 8 and menu_id = c.id) as part_sn8 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 20 and menu_id = c.id) as part_sn20 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 17 and menu_id = c.id) as part_sn17 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 22 and menu_id = c.id) as part_sn22 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 33 and menu_id = c.id) as part_sn33 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 25 and menu_id = c.id) as part_sn25 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 26 and menu_id = c.id) as part_sn26 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 27 and menu_id = c.id) as part_sn27 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 28 and menu_id = c.id) as part_sn28 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 29 and menu_id = c.id) as part_sn29 "
		strSql = strSql + "		, (select top 1 level_sn from db_partner.dbo.tbl_menu_part where part_sn = 34 and menu_id = c.id) as part_sn34 "
		strSql = strSql + "		, (select top 1 part_sn from db_partner.dbo.tbl_menu_part where part_sn not in (1, 16, 14, 11, 21, 12, 23, 13, 24, 30, 7, 9, 10, 8, 20, 17, 22, 33, 25) and menu_id = c.id) as part_sn_etc "
		strSql = strSql + "	FROM "
		strSql = strSql + "		db_partner.[dbo].tbl_partner_menu AS c WITH(NOLOCK) "
		strSql = strSql + "		LEFT JOIN db_partner.[dbo].tbl_partner_menu p WITH(NOLOCK) "
		strSql = strSql + "		ON "
		strSql = strSql + "			1 = 1 "
		strSql = strSql + "			AND c.parentid = p.id "
		strSql = strSql + "			AND p.parentid = 0 "
		strSql = strSql & addSqlJoin &  Vbcrlf
		strSql = strSql + "	WHERE "
		strSql = strSql + "		1 = 1 "

		strSql = strSql + AddSQL

		strSql = strSql + "	ORDER BY "
		strSql = strSql + "		IsNull(p.viewidx, 0) "
		strSql = strSql + "		,IsNull(p.id, 0) "
		strSql = strSql + "		,c.viewidx "
		strSql = strSql + "		,c.id "

		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		''rw strSql

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CMenuPrivListItem

				FItemList(i).Fmenu_id			= rsget("id")
				FItemList(i).Fmenu_name			= rsget("menuname")
				FItemList(i).Fmenu_name_parent	= rsget("parentmenuname")

				FItemList(i).Fmenu_parentid		= rsget("parentid")
				FItemList(i).Fmenu_isusing		= rsget("isusing")
				FItemList(i).Fmenu_viewidx		= rsget("viewidx")
                FItemList(i).Fmenu_divcd    	= rsget("divcd")
                FItemList(i).Fmenu_useSslYN		= rsget("useSslYN")
                FItemList(i).Fmenu_criticinfo   = rsget("criticinfo")

				FItemList(i).flv1customerYN    = rsget("lv1customerYN")
				FItemList(i).flv2partnerYN    = rsget("lv2partnerYN")
				FItemList(i).flv3InternalYN    = rsget("lv3InternalYN")
				FItemList(i).Fmenu_part_sn1   	= rsget("part_sn1")
				FItemList(i).Fmenu_part_sn16   	= rsget("part_sn16")
				FItemList(i).Fmenu_part_sn14   	= rsget("part_sn14")
				FItemList(i).Fmenu_part_sn11   	= rsget("part_sn11")
				FItemList(i).Fmenu_part_sn21   	= rsget("part_sn21")
				FItemList(i).Fmenu_part_sn12   	= rsget("part_sn12")
				FItemList(i).Fmenu_part_sn23   	= rsget("part_sn23")
				FItemList(i).Fmenu_part_sn13   	= rsget("part_sn13")
				FItemList(i).Fmenu_part_sn24   	= rsget("part_sn24")
				FItemList(i).Fmenu_part_sn30   	= rsget("part_sn30")
				FItemList(i).Fmenu_part_sn7   	= rsget("part_sn7")
				FItemList(i).Fmenu_part_sn9   	= rsget("part_sn9")
				FItemList(i).Fmenu_part_sn10   	= rsget("part_sn10")
				FItemList(i).Fmenu_part_sn8   	= rsget("part_sn8")
				FItemList(i).Fmenu_part_sn20   	= rsget("part_sn20")
				FItemList(i).Fmenu_part_sn17	= rsget("part_sn17")
				FItemList(i).Fmenu_part_sn22	= rsget("part_sn22")
				FItemList(i).Fmenu_part_sn33	= rsget("part_sn33")
				FItemList(i).Fmenu_part_sn25	= rsget("part_sn25")
				FItemList(i).Fmenu_part_sn26	= rsget("part_sn26")
				FItemList(i).Fmenu_part_sn27	= rsget("part_sn27")
				FItemList(i).Fmenu_part_sn28	= rsget("part_sn28")
				FItemList(i).Fmenu_part_sn29	= rsget("part_sn29")
				FItemList(i).Fmenu_part_sn34	= rsget("part_sn34")
				FItemList(i).Fmenu_part_sn_etc	= rsget("part_sn_etc")

				if (FItemList(i).Fmenu_parentid = 0) then
					FItemList(i).Fmenu_name_parent = FItemList(i).Fmenu_name
					FItemList(i).Fmenu_name = ""
				end if

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
		SQL =	"select " & vbCRLF
		SQL = SQL & "	 t1.id, t1.menuname, t1.linkurl, t1.parentid " & vbCRLF
		SQL = SQL & "	, t1.menucolor, t1.isusing, t1.viewidx, t1.divcd, t1.menuname_En, t1.useSslYN, isNULL(t1.criticinfo,0) as criticinfo, isNULL(t1.saveLog,0) as saveLog " & vbCRLF
		SQL = SQL & "	, IsNULL(t1.lv1customerYN,'N') AS lv1customerYN, IsNULL(t1.lv2partnerYN,'N') AS lv2partnerYN, IsNULL(t1.lv3InternalYN,'N') AS lv3InternalYN" & vbCRLF
		SQL = SQL & " from db_partner.[dbo].tbl_partner_menu as t1 WITH(NOLOCK) " & vbCRLF
		SQL = SQL & " where t1.id=" & FRectMid

		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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
			FItemList(1).Fmenu_useSslYN		= rsget("useSslYN")
			FItemList(1).Fmenu_criticinfo   = rsget("criticinfo")
			FItemList(1).Fmenu_saveLog    	= rsget("saveLog")
			FItemList(1).Flv1customerYN    	= rsget("lv1customerYN")
			FItemList(1).Flv2partnerYN	= rsget("lv2partnerYN")
			FItemList(1).Flv3InternalYN    	= rsget("lv3InternalYN")
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
			    if (FRectPart_sn="17") then ''관계사
			    AddSQL = "part_sn in ('" & FRectPart_sn & "')" & VbCRLF
				AddSQL = AddSQL & " and level_sn>='" & FRectLevel_sn & "'"
			    else
				AddSQL = "part_sn in (1, '" & FRectPart_sn & "')" & VbCRLF
				AddSQL = AddSQL & " and level_sn>='" & FRectLevel_sn & "'"
				end if
				'※ 부서번호 1 : 부서전체

				''추가 권한 관련 2011-09-19
				''특정부서의 특정권한을 추가한다. : 특정부서+파트선임권한 을 추가해도, 부서전체 파트선임권한은 제외된다.
				if (FRectUserID<>"") then
				    AddSQL = AddSQL & " OR menu_id in ("
				    AddSQL = AddSQL & "     select menu_id from db_partner.dbo.tbl_menu_part p WITH(NOLOCK) "
                	AddSQL = AddSQL & "     Join db_partner.dbo.tbl_partner_AddLevel L WITH(NOLOCK) "
                	AddSQL = AddSQL & "     on L.userid='"&FRectUserID&"'" & VbCRLF
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
		SQL = SQL + " from db_partner.[dbo].tbl_partner_menu as t1 WITH(NOLOCK) "
		SQL = SQL + " 		Join (Select distinct menu_id "
		SQL = SQL + " 				from db_partner.dbo.tbl_menu_part WITH(NOLOCK) "
		SQL = SQL + " 				where " + AddSQL
		SQL = SQL + " 			) as t2 "
		SQL = SQL + " 			on t1.id=t2.menu_id "

		if (FRectSearchString <> "") then
			SQL = SQL + " 		left join ( "
			SQL = SQL + " 			select p.id, count(*) as cnt "
			SQL = SQL + " 			from "
			SQL = SQL + " 			db_partner.[dbo].tbl_partner_menu p WITH(NOLOCK) "
			SQL = SQL + " 			left join db_partner.[dbo].tbl_partner_menu c WITH(NOLOCK) "
			SQL = SQL + " 			on "
			SQL = SQL + " 				p.id = c.parentid "
			SQL = SQL + " 			where "
			SQL = SQL + " 				1 = 1 "
			SQL = SQL + " 				and p.parentid = 0 "
			SQL = SQL + " 				and p.isusing = 'Y' "
			SQL = SQL + " 				and c.isusing = 'Y' "
			SQL = SQL + " 				and c.menuname like '%" + CStr(FRectSearchString) + "%' " &VBCRLF
			SQL = SQL + " 			group by "
			SQL = SQL + " 				p.id "
			SQL = SQL + " 		) TT "
			SQL = SQL + " 		on TT.id = t1.id "
			SQL = SQL + " left join db_partner.[dbo].tbl_partner_menu p WITH(NOLOCK) "
			SQL = SQL + " on "
			SQL = SQL + " 	p.id = t1.parentid "
		end if

		SQL = SQL + " Where t1.isusing='Y' "

		if (FRectSearchString <> "") then
			SQL = SQL + " and ( "
			SQL = SQL + " 	(IsNull(TT.cnt, 0) > 0) "
			SQL = SQL + " 	or "
			SQL = SQL + " 	(IsNull(p.menuname, '') like '%" + CStr(FRectSearchString) + "%') " &VBCRLF
			SQL = SQL + " 	or "
			SQL = SQL + " 	(t1.menuname like '%" + CStr(FRectSearchString) + "%') " &VBCRLF
			SQL = SQL + " ) "
		end if

		SQL = SQL + " Order by t1.parentid, t1.viewidx "

		''rw SQL
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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
					FItemList(i).Fmenu_useSslYN	= rsget("useSslYN")

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
		strSQL = strSQL & " FROM db_partner.dbo.tbl_partInfo as P WITH(NOLOCK) " & VBCRLF
		strSQL = strSQL & " LEFT JOIN db_partner.dbo.tbl_partInfoGroup as G WITH(NOLOCK)" & VBCRLF
		strSQL = strSQL & " ON P.part_sn = G.part_sn  " & VBCRLF
		strSQL = strSQL & " WHERE p.part_sn = '"&FRectPart_sn&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
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
        '// 항상표시, 2020-11-19, skyer9
		''if FRectshopdiv = "5" then
		''	AddSQL = AddSQL & " and id not in (501)"
		''end if

		'//해외가 아닐경우, [해외] 출고조회, 해외매장 상품설정은 제낌
        '// 해외가 아니어도 표시, 2020-11-19, skyer9
		''if FRectshopdiv <> "7" then
		''	IF application("Svr_Info")="Dev" THEN
		''		AddSQL = AddSQL & " and id not in (1477,1211)"
		''	else
		''		AddSQL = AddSQL & " and id not in (1391,1210)"
		''	end if
		''end if

		'//도매,해외,its일 경우 매장게시판관리는 제낌
		if FRectshopdiv = "5" or FRectshopdiv = "7" or FRectshopdiv = "11" or FRectshopdiv = "13" then
			AddSQL = AddSQL & " and id not in (524)"
		end if

		'// 리테일매장
		if FRectshopdiv = "15" then
			AddSQL = AddSQL & " and id not in (1351,1352,1353,1376,1377)"
		end if

		'// 목록 접수 //
		SQL =	"Select " & vbCRLF
		SQL = SQL & "	 t1.id, t1.menuname, t1.linkurl, t1.parentid, t1.menucolor, t1.menuname_en " & vbCRLF
		SQL = SQL & " from db_partner.[dbo].tbl_partner_menu as t1 WITH(NOLOCK) " & vbCRLF
		SQL = SQL & " Where t1.isusing='Y' and divCD in ('500','" & CStr(FRectUserDiv) & "')" & AddSQL & vbCRLF
		SQL = SQL & " Order by t1.parentid, t1.viewidx "
		rsget.CursorLocation = adUseClient
		rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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

	strOpt =	"<select class='select' name='" & fnm & "'>" & vbCRLF
	strOpt = strOpt & "<option value=''>::부서선택::</option>"

	SQL =	"Select part_sn, part_name " & vbCRLF
	SQL = SQL & " From db_partner.dbo.tbl_partInfo " & vbCRLF
	SQL = SQL & " Where part_isDel='N' " & vbCRLF
	SQL = SQL & " Order by part_sort"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly
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

	strOpt =	"<select class='select' name='" & fnm & "'>" & vbCRLF
	strOpt = strOpt & "<option value=''>::직급선택::</option>"

	SQL =	"Select posit_sn, posit_name " & vbCRLF
	SQL = SQL & " From db_partner.dbo.tbl_positInfo WITH(NOLOCK) " & vbCRLF
	SQL = SQL & " Where posit_isDel='N' "
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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

	strOpt =	"<select class='select' name='" & fnm & "'>" & vbCRLF
	strOpt = strOpt & "<option value=''>::등급선택::</option>"

	SQL =	"Select level_sn, level_name " & vbCRLF
	SQL = SQL & " From db_partner.dbo.tbl_level WITH(NOLOCK)" & vbCRLF
	SQL = SQL & " Where level_isDel='N' " & vbCRLF
	SQL = SQL & " Order by level_no"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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

	SQL =	"Select t2.part_name, t1.part_sn, t3.level_name, t1.level_sn " & vbCRLF
	SQL = SQL & " From db_partner.dbo.tbl_menu_part as t1 WITH(NOLOCK) " & vbCRLF
	SQL = SQL & "	join db_partner.dbo.tbl_partInfo as t2 WITH(NOLOCK) " & vbCRLF
	SQL = SQL & "		on t1.part_sn=t2.part_sn " & vbCRLF
	SQL = SQL & "	join db_partner.dbo.tbl_level as t3 WITH(NOLOCK)" & vbCRLF
	SQL = SQL & "		on t1.level_sn=t3.level_sn " & vbCRLF
	SQL = SQL & " Where t2.part_isDel='N' and t3.level_isDel='N' " & vbCRLF
	SQL = SQL & " 	and t1.menu_id=" & pid & " " & vbCRLF
	SQL = SQL & " Order by t2.part_sort"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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
						"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle></td>" &_
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

	SQL =	"Select id, menuname " & vbCRLF
	SQL = SQL & " From db_partner.[dbo].tbl_partner_menu WITH(NOLOCK) " & vbCRLF
	SQL = SQL & " Where parentid=0 and (isusing='Y' or id = " + CStr(pid) + ") " & vbCRLF
	SQL = SQL & " Order by viewidx"
	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

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

'// 사용자 권한과 메뉴등급을 일치시켜야 함
Sub DrawSelectBoxCriticInfoMenu(selectedname, selectedId)
	dim tmp_str, query1
%>
<select class='select' name="<%= selectedname %>" >
    <option value='' <%if selectedId="" then response.write " selected"%> >선택</option>
	<option value='500' <%if selectedId="500" then response.write " selected"%> >LV1(개인정보)</option>
	<option value='100' <%if selectedId="100" then response.write " selected"%> >LV2(배송정보)</option>
	<option value='1' <%if selectedId="1" then response.write " selected"%> >LV3(주문정보)</option>
	<option value='0' <%if selectedId="0" then response.write " selected"%> >일반</option>
</select>
<%
End Sub

'// 사용자 권한과 메뉴등급을 일치시켜야 함
Function GetCriticInfoMenuLevelName(selectedId)
	Select Case selectedId
		Case "500"
			GetCriticInfoMenuLevelName = "LV1(개인정보)"
		Case "100"
			GetCriticInfoMenuLevelName = "LV2(배송정보)"
		Case "1"
			GetCriticInfoMenuLevelName = "LV3(주문정보)"
		Case "0"
			GetCriticInfoMenuLevelName = "일반"
		Case Else
			GetCriticInfoMenuLevelName = selectedId
	End Select
End Function

%>

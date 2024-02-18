<%

'텐텐 원본파일을 복사해온 후 dbget_CS -> dbget_CS_CS, rsget_CS -> rsget_CS_CS 로 수정해준다.

function fnGetMenuPos(byval menuid, byref menuposnotice, byref menuposhelp)
	dim sqlStr,menustr
	dim pid

    if (menuid<2) then Exit function

	sqlStr = "select id,menuname,parentid,menuposnotice,menuposhelp from [db_partner].[dbo].tbl_partner_menu"
	sqlStr = sqlStr + " where id=" + CStr(menuid)

	rsget_CS.Open sqlStr,dbget_CS,1
	if Not rsget_CS.Eof then
		menustr         = db2html(rsget_CS("menuname"))
		pid             = rsget_CS("parentid")
		menuposnotice  = db2html(rsget_CS("menuposnotice"))
		menuposhelp    = db2html(rsget_CS("menuposhelp"))

		if IsNULL(menuposnotice) then Fmenuposnotice=""
		if IsNULL(menuposhelp) then Fmenuposhelp=""
	end if
	rsget_CS.Close

	if pid=0 then
		menustr = "&gt;&gt;" + menustr
	else
		sqlStr = "select id,menuname,parentid, menuposnotice from [db_partner].[dbo].tbl_partner_menu"
		sqlStr = sqlStr + " where id=" + CStr(pid)

		rsget_CS.Open sqlStr,dbget_CS,1
		if Not rsget_CS.Eof then
			menustr = db2html(rsget_CS("menuname")) + "&gt;&gt;" + menustr
			pid     = rsget_CS("parentid")
		end if
		rsget_CS.Close
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

	public FLastMenu
	public FChildCount
	public FChildItem()

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
		rsget_CS.Open SQL,dbget_CS,1
			FTotalCount = rsget_CS(0)
			FtotalPage = rsget_CS(1)
		rsget_CS.Close

		'// 목록 접수 //
		SQL =	"select top " & CStr(FPageSize*FCurrPage) &_
				"	 t1.id, t1.menuname, t1.linkurl, t1.parentid " &_
				"	, t1.menucolor, t1.isusing, t1.viewidx, t1.divcd " &_
				"	, t2.menu_cnt " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"		Left Join (Select parentid, count(*) as menu_cnt " &_
				"					from db_partner.[dbo].tbl_partner_menu " &_
				"					where isusing='Y' Group by parentid) as t2 " &_
				"			on t1.id=t2.parentid " &_
				"where t1.parentid=" & FRectPid & AddSQL &_
				"Order by t1.parentid, t1.viewidx "
		rsget_CS.pagesize = FPageSize
		rsget_CS.Open SQL,dbget_CS,1

		FResultCount = rsget_CS.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget_CS.EOF  then
			rsget_CS.absolutepage = FCurrPage
			do until rsget_CS.eof
				set FItemList(i) = new CMenuListItem

				FItemList(i).Fmenu_id		= rsget_CS("id")
				FItemList(i).Fmenu_name		= rsget_CS("menuname")
				FItemList(i).Fmenu_linkurl	= rsget_CS("linkurl")
				FItemList(i).Fmenu_parentid	= rsget_CS("parentid")
				FItemList(i).Fmenu_color	= rsget_CS("menucolor")
				FItemList(i).Fmenu_isusing	= rsget_CS("isusing")
				FItemList(i).Fmenu_viewidx	= rsget_CS("viewidx")
				FItemList(i).Fmenu_cnt		= rsget_CS("menu_cnt")

                FItemList(i).Fmenu_divcd    = rsget_CS("divcd")

				rsget_CS.moveNext
				i=i+1
			loop
		end if

		rsget_CS.Close

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
				"	, t1.menucolor, t1.isusing, t1.viewidx, t1.divcd " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"where t1.id=" & FRectMid
		rsget_CS.Open SQL,dbget_CS,1

		if Not(rsget_CS.EOF or rsget_CS.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CMenuListItem

			FItemList(1).Fmenu_id		= rsget_CS("id")
			FItemList(1).Fmenu_name		= rsget_CS("menuname")
			FItemList(1).Fmenu_linkurl	= rsget_CS("linkurl")
			FItemList(1).Fmenu_parentid	= rsget_CS("parentid")
			FItemList(1).Fmenu_color	= rsget_CS("menucolor")
			FItemList(1).Fmenu_isusing	= rsget_CS("isusing")
			FItemList(1).Fmenu_viewidx	= rsget_CS("viewidx")

			FItemList(1).Fmenu_divcd    = rsget_CS("divcd")
		else
			FResultCount = 0
		end if

		rsget_CS.Close

	end Sub


	'##### 왼쪽 메뉴 목록 #####
	public Sub GetLeftMenuList()
		dim SQL, AddSQL, i, strTemp
		dim onemenuitem, tmp

		'관리자등급(Level:1)은 모든 메뉴 출력
		if FRectLevel_sn=1 then
			AddSQL = "1=1"
		else
			if NOT(FRectLevel_sn="" or isNull(FRectLevel_sn)) then
				AddSQL = "part_sn in (1, " & FRectPart_sn & ")" &_
						" and level_sn>=" & FRectLevel_sn
				'※ 부서번호 1 : 부서전체
			else
				'권한이 없으면 메뉴 표시 없음
				AddSQL = " level_sn>9999 "
			end if
		end if

		'// 목록 접수 //
		SQL =	"Select " &_
				"	 t1.id, t1.menuname, t1.linkurl, t1.parentid, t1.menucolor " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"		Join (Select distinct menu_id " &_
				"				from db_partner.dbo.tbl_menu_part " &_
				"				where " & AddSQL &_
				"			) as t2 " &_
				"			on t1.id=t2.menu_id " &_
				"Where t1.isusing='Y' " &_
				"Order by t1.parentid, t1.viewidx "

		'Response.Write SQL
		rsget_CS.Open SQL,dbget_CS,1

		i=0
		if  not rsget_CS.EOF  then
			do until rsget_CS.eof
				set onemenuitem = new CMenuListItem

				onemenuitem.Fmenu_id		= rsget_CS("id")
				onemenuitem.Fmenu_name		= db2html(rsget_CS("menuname"))
				onemenuitem.Fmenu_linkurl	= db2html(rsget_CS("linkurl"))
				onemenuitem.Fmenu_parentid	= rsget_CS("parentid")
				onemenuitem.Fmenu_color		= db2html(rsget_CS("menucolor"))

				'// 하위메뉴 목록 저장
				if onemenuitem.Fmenu_parentid=0 then
					AddChild onemenuitem
				else
					set tmp = getParentMenu ( onemenuitem.Fmenu_parentid )
					if Not(tmp is Nothing) then
						tmp.addChild onemenuitem
					end if
				end if

				rsget_CS.moveNext
				i=i+1
			loop
		end if

		rsget_CS.Close
	end Sub

	'##### 왼쪽 메뉴 목록(OFFLine용) #####
	public Sub GetLeftMenuList_offLine()
		dim SQL, AddSQL, i, strTemp
		dim onemenuitem, tmp

		if CStr(FRectUserDiv)="509" then
			'// 매출조회용이라면
			AddSQL = AddSQL & " and id in (501, 508, 511, 512)"
		end if

		'// 목록 접수 //
		SQL =	"Select " &_
				"	 t1.id, t1.menuname, t1.linkurl, t1.parentid, t1.menucolor " &_
				"from db_partner.[dbo].tbl_partner_menu as t1 " &_
				"Where t1.isusing='Y' and divCD in ('500','" & CStr(FRectUserDiv) & "')" & AddSQL &_
				"Order by t1.parentid, t1.viewidx "
		rsget_CS.Open SQL,dbget_CS,1

		i=0
		if  not rsget_CS.EOF  then
			do until rsget_CS.eof
				set onemenuitem = new CMenuListItem

				onemenuitem.Fmenu_id		= rsget_CS("id")
				onemenuitem.Fmenu_name		= db2html(rsget_CS("menuname"))
				onemenuitem.Fmenu_linkurl	= db2html(rsget_CS("linkurl"))
				onemenuitem.Fmenu_parentid	= rsget_CS("parentid")
				onemenuitem.Fmenu_color		= db2html(rsget_CS("menucolor"))

				'// 하위메뉴 목록 저장
				if onemenuitem.Fmenu_parentid=0 then
					AddChild onemenuitem
				else
					set tmp = getParentMenu ( onemenuitem.Fmenu_parentid )
					if Not(tmp is Nothing) then
						tmp.addChild onemenuitem
					end if
				end if

				rsget_CS.moveNext
				i=i+1
			loop
		end if

		rsget_CS.Close
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
	rsget_CS.Open SQL,dbget_CS,1

	if Not(rsget_CS.EOF or rsget_CS.BOF) then
		Do Until rsget_CS.EOF
			strOpt = strOpt & "<option value='" & rsget_CS("part_sn") & "'"
			if Cstr(rsget_CS("part_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget_CS("part_name") & "</option>"
		rsget_CS.MoveNext
		Loop
	end if

	rsget_CS.Close

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
	rsget_CS.Open SQL,dbget_CS,1

	if Not(rsget_CS.EOF or rsget_CS.BOF) then
		Do Until rsget_CS.EOF
			strOpt = strOpt & "<option value='" & rsget_CS("posit_sn") & "'"
			if Cstr(rsget_CS("posit_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget_CS("posit_name") & "</option>"
		rsget_CS.MoveNext
		Loop
	end if

	rsget_CS.Close

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
	rsget_CS.Open SQL,dbget_CS,1

	if Not(rsget_CS.EOF or rsget_CS.BOF) then
		Do Until rsget_CS.EOF
			strOpt = strOpt & "<option value='" & rsget_CS("level_sn") & "'"
			if Cstr(rsget_CS("level_sn"))=Cstr(psn) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget_CS("level_name") & "</option>"
		rsget_CS.MoveNext
		Loop
	end if

	rsget_CS.Close

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
	rsget_CS.Open SQL,dbget_CS,1

	strPrt = ""
	if mode="modi" then strPrt = strPrt & "<table name='tbl_auth' id='tbl_auth' class=a>"
	if Not(rsget_CS.EOf or rsget_CS.BOf) then
		i = 0
		Do Until rsget_CS.EOF
			if mode="list" then
				strPrt = strPrt & rsget_CS(0) & "/" & rsget_CS(2) & "<br>"
			elseif mode="modi" then
				strPrt = strPrt &_
					"<tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>" &_
						"<td>" & rsget_CS(0) & "<input type='hidden' name='part_sn' value='" & rsget_CS(1) & "'></td>" &_
						"<td>" & rsget_CS(2) & "<input type='hidden' name='level_sn' value='" & rsget_CS(3) & "'></td>" &_
						"<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle></td>" &_
					"</tr>"
			end if
			i = i + 1
		rsget_CS.MoveNext
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

	rsget_CS.Close
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
	rsget_CS.Open SQL,dbget_CS,1

	if Not(rsget_CS.EOF or rsget_CS.BOF) then
		Do Until rsget_CS.EOF
			strOpt = strOpt & "<option value='" & rsget_CS("id") & "'"
			if cLng(rsget_CS("id"))=cLng(pid) then
				strOpt = strOpt & " selected"
			end if
			strOpt = strOpt & ">" & rsget_CS("menuname") & "</option>"
		rsget_CS.MoveNext
		Loop
	end if

	rsget_CS.Close

	strOpt = strOpt & "</select>"

	'값 반환
	printRootMenuOption = strOpt
end Function
%>
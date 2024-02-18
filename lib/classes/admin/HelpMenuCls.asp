<%
 

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


Class CMenuItem
	public FMenuID
	public FMenuName
	public FLinkURL
	public FViewIndex
	public FDivCD

	public FHasChild
	public FChildItem()
	public FChildCount

	public FParentID
	public FLastMenu
	public FIsUsing
	public Fmenucolor
	public Fmenuposnotice
	public Fmenuposhelp

	public FMenuStr
	public FuseSslYN
 

	Private Sub Class_Initialize()
		FLastMenu = true
		FChildCount =0
		'redim preserve FChildItem(0)
		redim  FChildItem(0)
	End Sub

	Private Sub Class_Terminate()

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
 
  public FRectID
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
 public FOneItem
 
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

  
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
 

	public Sub GetLeftMenuListNew()
		dim strSql, i  
 
		strSql = " select  id, menuname,  linkurl,  parentid,  menucolor,  menuname_En ,useSslYN, id as cid, parentid as pid  "&vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_partner_menu "&vbcrlf
		if FRectUserDiv ="2" then
			strSql = strSql & " where divcd in (9,7,5,4,2,1) and isusing ='Y'"&vbcrlf
		else
			strSql = strSql & " where divcd = '"&FRectUserDiv&"' and isusing ='Y'"&vbcrlf
		end if
		strSql = strSql & " order by parentid, viewidx "
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
	
	
	public function getOneMenu()
		dim sqlStr, menustr

		sqlStr = "select top 1 * from [db_partner].[dbo].tbl_partner_menu"
		sqlStr = sqlStr + " where id=" + CStr(FRectID)

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		set FOneItem = new CMenuItem
		if Not rsget.Eof then

			FOneItem.FMenuID       = rsget("id")
			FOneItem.FMenuName     = db2html(rsget("menuname"))
			FOneItem.FLinkURL      = db2html(rsget("linkurl"))
			FOneItem.FViewIndex    = rsget("viewidx")
			FOneItem.FDivCD        = rsget("divcd")

			FOneItem.FHasChild     = rsget("divcd")
			FOneItem.Fparentid     = rsget("parentid")

			FOneItem.FIsUsing      = rsget("isusing")
			FOneItem.Fmenucolor    = db2html(rsget("menucolor"))
			FOneItem.Fmenuposnotice= db2html(rsget("menuposnotice"))
			FOneItem.Fmenuposhelp  = db2html(rsget("menuposhelp"))

		end if
		rsget.Close

		if FOneItem.Fparentid=0 then
			menustr = "&gt;&gt;" + FOneItem.FMenuName
		else
			sqlStr = "select id,menuname,parentid, menuposnotice from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where id=" + CStr(FOneItem.Fparentid)

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				menustr = "&gt;&gt;" + db2html(rsget("menuname")) + "&gt;&gt;" + FOneItem.FMenuName
			end if
			rsget.Close
		end if

		FOneItem.FMenuStr = menustr
	end function
	
	
 End Class

%>

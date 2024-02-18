<%
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

	public function GetMenuColor()
		if IsNULL(Fmenucolor) or (Fmenucolor="") then
			GetMenuColor = "#000000"
		else
			GetMenuColor = Fmenucolor
		end if
	end function

	public function HasLinkURL()
		if (FLinkURL=null) or (FLinkURL="") then
			HasLinkURL = false
		else
			HasLinkURL = true
		end if
	end function

	public function getLinkURL()
		if (FLinkURL=null) or (FLinkURL="") then
			getLinkURL = "#"
		else
			getLinkURL = FLinkURL

			if Not IsNull(FuseSslYN) then
				if (FuseSslYN = "Y") then
					getLinkURL = "https://webadmin.10x10.co.kr" + FLinkURL
				end if
			end if
		end if
	end function

	public function IsHasChild()
		IsHasChild = FChildCount>0
	end function

	public function IsLastMenu()
		IsLastMenu = FLastMenu
	end function

	public function getCloseIimageUrl()
		if isLastMenu then
			getCloseIimageUrl = "/images/blank.png"
		else
			getCloseIimageUrl = "/images/I.png"
		end if
	end function

	public function getCloseTimageUrl()
		if IsHasChild then
			if isLastMenu then
				getCloseTimageUrl = "/images/Lplus.png"
			else
				getCloseTimageUrl = "/images/Tplus.png"
			end if
		else
			if isLastMenu then
				getCloseTimageUrl = "/images/L.png"
			else
				getCloseTimageUrl = "/images/T.png"
			end if
		end if
	end function

	public function getOpenTimageUrl()
		if IsHasChild then
			if isLastMenu then
				getOpenTimageUrl = "/images/Lminus.png"
			else
				getOpenTimageUrl = "/images/Tminus.png"
			end if
		else
			if isLastMenu then
				getOpenTimageUrl = "/images/L.png"
			else
				getOpenTimageUrl = "/images/T.png"
			end if
		end if
	end function

	public function getOpenIconURL()
		if IsHasChild then
			getOpenIconURL = "/images/openfolder.png"
		else
			getOpenIconURL = "/images/paper2.gif"
		end if

	end function

	public function getCloseIconURL()
		if IsHasChild then
			getCloseIconURL = "/images/closedfolder.png"
		else
			getCloseIconURL = "/images/paper2.gif"
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

	Private Sub Class_Initialize()
		FLastMenu = true
		FChildCount =0
		'redim preserve FChildItem(0)
		redim  FChildItem(0)
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CMenu
	public FOneItem
	public FMenuitemlist()
	public FMenuCount
	public FrectUsingOnly

	public FRectID

	public Fmenuposnotice
	public Fmenuposhelp

	public FResultCount

	Private Sub Class_Initialize()
		FMenuCount =0
		redim  FMenuitemlist(0)
		FMenuitemlist(0) = null
	End Sub

	Private Sub Class_Terminate()

	End Sub

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

	public function getMenuPos(byval menuid)
		dim sqlStr,menustr
		dim pid

		sqlStr = "select id,menuname,parentid,menuposnotice,menuposhelp from [db_partner].[dbo].tbl_partner_menu"
		sqlStr = sqlStr + " where id=" + CStr(menuid)

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			menustr = db2html(rsget("menuname"))
			pid = rsget("parentid")
			Fmenuposnotice = db2html(rsget("menuposnotice"))
			Fmenuposhelp = db2html(rsget("menuposhelp"))

			if IsNULL(Fmenuposnotice) then Fmenuposnotice=""
			if IsNULL(Fmenuposhelp) then Fmenuposhelp=""
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
				pid = rsget("parentid")
			end if
			rsget.Close
		end if

		getMenuPos = menustr
	end function

	public function UpdateMenu(byval imenuid, imenuname, iparentid, iviewidx, ilinkurl, idivcd, isusing, menucolor)
		dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner_menu" + VBCrlf
		sqlStr = sqlStr + " set menuname='" + imenuname + "'," + VBCrlf
		sqlStr = sqlStr + " linkurl='" + ilinkurl + "'," + VBCrlf
		sqlStr = sqlStr + " viewidx=" + CStr(iviewidx) + "," + VBCrlf
		sqlStr = sqlStr + " parentid=" + CStr(iparentid) + "," + VBCrlf
		sqlStr = sqlStr + " divcd=" + CStr(idivcd) + "," + VBCrlf
		sqlStr = sqlStr + " isusing='" + CStr(isusing) + "'," + VBCrlf
		sqlStr = sqlStr + " menucolor='" + CStr(menucolor) + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(imenuid)
'response.write sqlStr
		rsget.Open sqlStr,dbget,1
	end function

	public function AddMenu(byval imenuname, iparentid, iviewidx, ilinkurl, idivcd, menucolor)
		dim sqlStr
		sqlStr = "insert into [db_partner].[dbo].tbl_partner_menu(menuname,linkurl,viewidx,parentid,divcd,menucolor)"
		sqlStr = sqlStr + " values('" + imenuname + "',"
		sqlStr = sqlStr + " '" + ilinkurl + "',"
		sqlStr = sqlStr + " " + CStr(iviewidx) + ","
		sqlStr = sqlStr + " " + CStr(iparentid) + ","
		sqlStr = sqlStr + " " + CStr(idivcd) + ","
		sqlStr = sqlStr + " '" + CStr(menucolor) + "'"
		sqlStr = sqlStr + ")"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
	end function

	public sub AddChild(byval ichild)
		dim cnt
		cnt = UBound(FMenuitemlist)
		if FMenuCount<1 then
			set FMenuitemlist(0) = ichild
		else
			redim preserve FMenuitemlist(cnt+1)
			FMenuitemlist(cnt).FLastMenu = false
			set FMenuitemlist(cnt+1) = ichild
		end if
		FMenuCount = FMenuCount+1

	end sub

	public function getParentMenu(byval iid)
		dim i
		set getParentMenu = Nothing

		for i=0 to Ubound(FMenuitemlist)
			if (CStr(FMenuitemlist(i).FMenuID) = CStr(iid) )then
				set getParentMenu  = FMenuitemlist(i)
				Exit for
			end if
		next

	end function

	function getMenuItems(byval userdiv)
		dim onemenuitem,tmp
		dim sqlStr
		''##############################################
		''  관리자.
		''##############################################
		if CInt(userdiv)=9 then
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " where isusing='Y'"
			end if
			sqlStr = sqlStr + " order by parentid,viewidx"
		elseif CInt(userdiv)=1 then
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd=" + CStR(userdiv)
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if

			sqlStr = sqlStr + " order by parentid,viewidx"
		elseif CInt(userdiv)<9 then
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd<=" + CStR(userdiv)
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if

			sqlStr = sqlStr + " order by parentid,viewidx"
		elseif CStr(userdiv)="999" then
			''파트너
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd=999 "
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if

			sqlStr = sqlStr + " order by parentid,viewidx"


			if (LCASE(session("ssBctId"))="mbchpd") then                '''201203추가;; 해품달 imbc매출조회
			    sqlStr = "select id, menuname, linkurl, haschild, parentid, viewidx, divcd, isusing, menucolor, useSslYN "
			    sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_menu"
			    sqlStr = sqlStr + " where divcd=999"
			    sqlStr = sqlStr + " and isusing='Y'"
                sqlStr = sqlStr + " and id in (48,99)"
                sqlStr = sqlStr + " Union select 100,'기간별매출집계','/company/report/sellreportHPD.asp','N',99,3,999,'Y','','N'"
                sqlStr = sqlStr + " Union select 1413,'상품별매출집계','/company/report/sellreportHPDitem.asp','N',99,4,999,'Y','','N'"
                sqlStr = sqlStr + " order by parentid,viewidx"
            end if
		elseif CStr(userdiv)="9999" then

			''디자이너
			''############## Check OffShop ##############
			dim isOffShopOpen, isOffUpBeaExists
			isOffShopOpen = false
			isOffUpBeaExists = false

			''계약 되어있는 브랜드 또는 매장 업체 배송 브랜드 (2011-03)''2011-06-28 수정 eastone
			sqlStr = "select top 1 makerid, IsNULL(defaultbeasongdiv,0) as defaultbeasongdiv from [db_shop].[dbo].tbl_shop_designer"
			sqlStr = sqlStr + " where (comm_cd in ('B011','B012','B022') or (IsNULL(defaultbeasongdiv,0)=2) or (adminopen='Y'))" ''adminopen='Y' ''2011-06-29 추가
			sqlStr = sqlStr + " and makerid='" + session("ssBctId") + "'"
            sqlStr = sqlStr + " order by defaultbeasongdiv desc"

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				isOffShopOpen = true
				isOffUpBeaExists = rsget("defaultbeasongdiv")>0
			end if
			rsget.close


			''############## Check OffShop ##############

            ''텐바이텐관리 아이디를 업체에 오픈할 경우
            dim isTenIDUpcheOpen
            isTenIDUpcheOpen = (session("ssBctId")="its2pm") or (session("ssBctId")="haepumdal")

			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd=9999"
			sqlStr = sqlStr + " and id < 1838 " '2016-11-16 정윤정 추가 new 메뉴 안보이게
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if
			if (Not isOffShopOpen) then
				sqlStr = sqlStr + " and (id<>190) and (parentid<>190)"
			end if
			''2011-06-28 수정 eastone
			if (NOt isOffUpBeaExists) then
			    IF (application("Svr_Info")	= "Dev") then
			        sqlStr = sqlStr + " and id not in (1291,1292,1293,1307,1312)"
			    ELSE
			        sqlStr = sqlStr + " and id not in (1301,1302,1303,1305,1313)"
			    END IF
			end if

			if (isTenIDUpcheOpen) then
			    sqlStr = sqlStr + " and id in (52, 403, 54,112, 969, 96, 113)"
			end if
			sqlStr = sqlStr + " order by parentid,viewidx"

        elseif (CStr(userdiv)="501") or (CStr(userdiv)="502") or (CStr(userdiv)="503") then
			''파트너
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd in ('500','" + CStr(userdiv) + "')"
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if

			sqlStr = sqlStr + " order by parentid,viewidx"
		elseif (CStr(userdiv)="509") then
		    ''매출조회
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd in ('500','" + CStr(userdiv) + "')"
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if
            sqlStr = sqlStr + " and id in (501, 508, 511,512)"
			sqlStr = sqlStr + " order by parentid,viewidx"
		elseif CStr(userdiv)="111" then
			''오프샵점장메뉴
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd in ('101','" + CStr(userdiv) + "')"
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if

			sqlStr = sqlStr + " order by parentid,viewidx"
		elseif CStr(userdiv)="112" then
			''오프샵부점장메뉴
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd in ('101','" + CStr(userdiv) + "')"
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if

			sqlStr = sqlStr + " order by parentid,viewidx"
		else
			''파트너
			sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
			sqlStr = sqlStr + " where divcd=" + Cstr(userdiv)
			if FrectUsingOnly="Y" then
				sqlStr = sqlStr + " and isusing='Y'"
			end if
		 
			sqlStr = sqlStr + " order by parentid,viewidx"
		end if

		if sqlStr="" then Exit function
'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		do until rsget.eof
			set onemenuitem = New CMenuItem
			onemenuitem.FMenuID = rsget("id")
			onemenuitem.FMenuName = db2html(rsget("menuname"))
			onemenuitem.FLinkURL = db2html(rsget("linkurl"))
			onemenuitem.FHasChild = rsget("haschild")
			onemenuitem.FParentID = rsget("parentid")
			onemenuitem.FViewIndex = rsget("viewidx")
			onemenuitem.FDivCD = rsget("divcd")
			onemenuitem.FIsUsing = rsget("isusing")
			onemenuitem.Fmenucolor = db2html(rsget("menucolor"))
			onemenuitem.FuseSslYN = rsget("useSslYN")

			if onemenuitem.FParentID=0 then
				AddChild onemenuitem
			else
				'response.write onemenuitem.FMenuID & "<br>"
				set tmp = getParentMenu ( onemenuitem.FParentID )
				if Not(tmp is Nothing) then
					tmp.addChild onemenuitem
				end if
			end if

			rsget.movenext
		loop
		rsget.close

		''##############################################
		''  일반사원.
		''##############################################

	end function


	Public Function getMenuNewItems
		Dim sqlStr, i
		''파트너
		sqlStr = "select * from [db_partner].[dbo].tbl_partner_menu"
		sqlStr = sqlStr + " where divcd=999 "
		if FrectUsingOnly="Y" then
			sqlStr = sqlStr + " and isusing='Y'"
		end if
		sqlStr = sqlStr + " order by parentid,viewidx"
		rsget.Open sqlStr, dbget, 1
		''rw strSql
		FResultCount = rsget.RecordCount
		if FResultCount<1 then FResultCount=0
		redim preserve FMenuitemlist(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FMenuitemlist(i) = new CMenuItem
					FMenuitemlist(i).FMenuID = rsget("id")
					FMenuitemlist(i).FMenuName = db2html(rsget("menuname"))
					FMenuitemlist(i).FLinkURL = db2html(rsget("linkurl"))
					FMenuitemlist(i).FHasChild = rsget("haschild")
					FMenuitemlist(i).FParentID = rsget("parentid")
					FMenuitemlist(i).FViewIndex = rsget("viewidx")
					FMenuitemlist(i).FDivCD = rsget("divcd")
					FMenuitemlist(i).FIsUsing = rsget("isusing")
					FMenuitemlist(i).Fmenucolor = db2html(rsget("menucolor"))
					FMenuitemlist(i).FuseSslYN = rsget("useSslYN")
				rsget.moveNext
				i=i+1
			loop
		end if
		rsget.Close
	End Function

end Class
%>

<%
'###############################################
' PageName :enjoyevent
' Discription : 사이트 메인 공지 배너 관리
' History : 2014.06.09 이종화 생성
'###############################################

Class CMainbannerItem
	public fidx
	Public Fevtimg
	Public Fevtalt
	Public Flinktype
	Public Flinkurl
	Public Fevttitle
	Public Fevttitle2
	Public Fissalecoupon
	Public Fissalecoupontxt
	Public Fevtstdate
	Public Fevteddate
	Public Fstartdate
	Public Fenddate 
	Public Fadminid
	Public Flastadminid
	public Fisusing 
	Public Fregdate
	Public Fusername
	Public Flastupdate
	Public Fordertext
	Public Fsortnum
	Public Ftodaybanner
	Public Fevt_code

	Public Fevtmolistbanner

	Public Fxmlregdate

	Public Fetc_opt

	public Ftag_only
	public Ftag_gift
	public Ftag_plusone
	public Ftag_launching
	public Ftag_actively
	public Fsale_per
	public Fcoupon_per

	Public Fitemid1
	Public Fitemid2
	Public Fitemid3
	Public Faddtype

	Public Fiteminfo
	public FcontentType
	
	public FESale
	public FEGift
	public FECoupon
	public FECommnet
	public FSisOnlyTen
	public FEOneplusOne
	public FEFreedelivery
	public FENew
	public FESalePer
	public FECsalePer	

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CMainbanner
    public FOneItem
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	Public FRectvaliddate
	public FRectSelDateTime
	public FRectDispOption

    public FRectIdx
    public Fisusing
	Public Fsdt
	Public Fedt
	Public FRectsedatechk
	Public FRecttype
	
	'//admin/appmanage/today/enjoyevent/enjoy_insert.asp
    public Sub GetOneContents()
        dim sqlStr
        sqlStr = "select top 1 t.* , d.evt_todaybanner , d.evt_mo_listbanner "
        sqlStr = sqlStr & " , STUFF((   "
        sqlStr = sqlStr & " SELECT '^^' + cast(i.itemid as varchar(120)) +'|'+ cast(i.itemname as varchar(120)) +'|'+ cast(i.smallimage as varchar(50))"
        sqlStr = sqlStr & " FROM db_item.dbo.tbl_item as i"
        sqlStr = sqlStr & " WHERE i.itemid in (t.itemid1 , t.itemid2 , t.itemid3) and i.itemid<>0"
        sqlStr = sqlStr & " FOR XML PATH('')"
        sqlStr = sqlStr & " ), 1, 1, '') AS iteminfo"
		sqlStr = sqlStr & " , d.issale, d.isgift, d.iscoupon, d.isOnlyTen, d.isoneplusone, d.isfreedelivery, d.isbookingsell, d.iscomment, d.ISNEW, d.SALEPER, d.SALECPER "
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_main_enjoyevent_new as t "
		sqlStr = sqlStr & " left outer join db_event.dbo.tbl_event_display as d "
        sqlStr = sqlStr & " on t.evt_code = d.evt_code"
        sqlStr = sqlStr & " where idx=" & CStr(FRectIdx)

		'rw sqlStr & "<Br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMainbannerItem
        
        if Not rsget.Eof then
    		FOneItem.fidx				= rsget("idx")
			FOneItem.Fevtalt			= rsget("evtalt")
			FOneItem.Flinkurl			= rsget("linkurl")
			FOneItem.Fevttitle			= rsget("evttitle")
			FOneItem.Fissalecoupon		= rsget("issalecoupon")
			FOneItem.Fevtstdate			= rsget("evtstdate")
			FOneItem.Fevteddate			= rsget("evteddate")
			FOneItem.Fissalecoupontxt	= rsget("issalecoupontxt")
			FOneItem.Fstartdate			= rsget("startdate")
			FOneItem.Fenddate			= rsget("enddate")
			FOneItem.Fadminid			= rsget("adminid")
			FOneItem.Flastadminid		= rsget("lastadminid")
			FOneItem.Fisusing			= rsget("isusing")
			FOneItem.Fordertext			= rsget("ordertext")
			FOneItem.Flinktype			= rsget("linktype")
			FOneItem.Fsortnum			= rsget("sortnum")
			FOneItem.Ftodaybanner		= rsget("evt_todaybanner")
			FOneItem.Fevt_code			= rsget("evt_code")
			FOneItem.Fevtmolistbanner	= rsget("evt_mo_listbanner")
			FOneItem.Fevttitle2			= rsget("evttitle2")
			FOneItem.Fetc_opt			= rsget("etc_opt")

			FOneItem.Ftag_only			= rsget("tag_only")	'2018-08-08 단독 태그
			FOneItem.Ftag_gift			= rsget("tag_gift")	'2017-07-27 기프트 태그
			FOneItem.Ftag_plusone		= rsget("tag_plusone") '2017-07-27 1+1 태그
			FOneItem.Ftag_launching		= rsget("tag_launching") '2017-07-27 런칭 태그
			FOneItem.Ftag_actively		= rsget("tag_actively") '2017-07-27 참여관련 태그
			FOneItem.Fsale_per			= rsget("sale_per") '2017-07-27 세일 text
			FOneItem.Fcoupon_per		= rsget("coupon_per") '2017-07-27 쿠폰 text

			FOneItem.FESale				 = rsget("issale")
			FOneItem.FEGift				 = rsget("isgift")
			FOneItem.FECoupon			 = rsget("iscoupon")
			FOneItem.FECommnet			 = rsget("iscomment")	
			FOneItem.FSisOnlyTen		 = rsget("isOnlyTen")
			FOneItem.FEOneplusOne		 = rsget("isoneplusone") 
			FOneItem.FEFreedelivery		 = rsget("isfreedelivery")
			FOneItem.FENew 				 = rsget("isnew")
			FOneItem.FECsalePer 		 = rsget("SALECPER")
			FOneItem.FESalePer   	 	 = rsget("SALEPER")			

			FOneItem.Fitemid1			= rsget("itemid1") '2017-07-27 itemid1
			FOneItem.Fitemid2			= rsget("itemid2") '2017-07-27 itemid2
			FOneItem.Fitemid3			= rsget("itemid3") '2017-07-27 itemid3
			FOneItem.Faddtype			= rsget("addtype") '2017-07-27 addtype

			FOneItem.Fiteminfo			= rsget("iteminfo") '2017-07-27 addtype
			if FOneItem.Faddtype = 3 or FOneItem.Faddtype = 4 then
				FOneItem.Fevtimg			= rsget("evtimg")
			else	
				FOneItem.Fevtimg			= staticImgUrl & "/mobile/enjoyevent" & rsget("evtimg")
			end if
			FOneItem.FcontentType			= rsget("contenttype") '2018-11-29 addtype 추가
        end If
        
        rsget.Close
    end Sub
	
	'//admin/appmanage/today/enjoyevent/index.asp
    public Sub GetContentsList()
        dim sqlStr, i

		sqlStr = " select count(idx) as cnt from db_sitemaster.dbo.tbl_mobile_main_enjoyevent_new "
		sqlStr = sqlStr + " where 1=1"
        
        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If
'//날짜 조건 수정 201081203
		if fsdt <> "" then
			sqlStr = sqlStr + " and '"&fsdt&"' between CONVERT(CHAR(10), startdate, 23) AND CONVERT(CHAR(10), enddate, 23) "
		end if		
'노출 위치 옵션 추가 20181127 최종원
		if FRectDispOption <> "" then 
            sqlStr = sqlStr + " and dispOption =" &FRectDispOption
        end if 

		If FRecttype = "" Then
			sqlStr = sqlStr & " and (addtype is null or addtype = '') "
		else
			sqlStr = sqlStr & " and addtype = '"& FRecttype &"'"
		End If 

		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr + " and enddate > getdate() "
		End If 

		'response.write sqlStr &"<br>"
		'response.end

        rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
        
        if FTotalCount < 1 then exit Sub
        	
        sqlStr = "select top " + CStr(FPageSize * FCurrPage) + " "
		sqlStr = sqlStr + " t.idx, t.evtimg, t.evtalt, t.linkurl, t.evttitle, t.issalecoupon, t.startdate, t.enddate"
		sqlStr = sqlStr + ", t.adminid, t.lastadminid, t.isusing, t.regdate, t.lastupdate, t.xmlregdate, t.linktype"
		sqlStr = sqlStr + ", t.sortnum, d.evt_todaybanner, d.evt_mo_listbanner, t.evttitle2, t.addtype, t.evt_code"
        sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_mobile_main_enjoyevent_new as t "
        sqlStr = sqlStr + " left outer join db_event.dbo.tbl_event_display as d "
        sqlStr = sqlStr + " on t.evt_code = d.evt_code"
        sqlStr = sqlStr + " where 1=1"

		'Response.write sqlStr

        if Fisusing<>"" then
            sqlStr = sqlStr + " and isusing='" + CStr(Fisusing) + "'"
        end If

		If FRecttype = "" Then
			sqlStr = sqlStr & " and (addtype is null or addtype = '') "
		else
			sqlStr = sqlStr & " and addtype = '"& FRecttype &"'"
		End If  
        
		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr + " and t.enddate > getdate() "
		Else
			sqlStr = sqlStr + " and t.enddate < getdate() "
		End If 
'//날짜 조건 수정 201081203
		if fsdt <> "" then
			sqlStr = sqlStr + " and '"&fsdt&"' between CONVERT(CHAR(10), startdate, 23) AND CONVERT(CHAR(10), enddate, 23) "
		end if		

		If FRectvaliddate = "on" Then 
			sqlStr = sqlStr + " order by t.sortnum asc, t.startdate asc" 
		Else
			sqlStr = sqlStr + " order by t.sortnum asc " 
		End If 
		'response.write sqlStr &"<br>"
		'response.end
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
				set FItemList(i) = new CMainbannerItem
				
				FItemList(i).fidx				= rsget("idx")				
				FItemList(i).Fevtalt			= rsget("evtalt")
				FItemList(i).Flinkurl			= rsget("linkurl")
				FItemList(i).Fevttitle			= rsget("evttitle")
				FItemList(i).Fissalecoupon		= rsget("issalecoupon")
				FItemList(i).Fstartdate			= rsget("startdate")
				FItemList(i).Fenddate			= rsget("enddate")
				FItemList(i).Fadminid			= rsget("adminid")
				FItemList(i).Flastadminid		= rsget("lastadminid")
				FItemList(i).Fisusing			= rsget("isusing")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")
				FItemList(i).Fxmlregdate		= rsget("xmlregdate")
				FItemList(i).Flinktype			= rsget("linktype")
				FItemList(i).Fsortnum			= rsget("sortnum")
				FItemList(i).Ftodaybanner		= rsget("evt_todaybanner")
				FItemList(i).Fevtmolistbanner	= rsget("evt_mo_listbanner")
				FItemList(i).Fevttitle2			= rsget("evttitle2")
				FItemList(i).Faddtype			= rsget("addtype")
				if FItemList(i).Faddtype = 3 or FItemList(i).Faddtype = 4 then
					FItemList(i).Fevtimg			= rsget("evtimg")				
				else
					FItemList(i).Fevtimg			= staticImgUrl & "/mobile/enjoyevent" & rsget("evtimg")				
				end if
				FItemList(i).Fevt_code			= rsget("evt_code")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
    end Sub
    

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

'// STAFF 이름 접수
public Function getStaffUserName(uid)
	if uid="" or isNull(uid) then
		exit Function
	end if

	Dim strSql
	strSql = "Select top 1 username From db_partner.dbo.tbl_user_tenbyten Where userid='" & uid & "'"
	rsget.Open strSql, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		getStaffUserName = rsget("username")
	end if
	rsget.Close
End Function
%>
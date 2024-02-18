<%
'###########################################################
' Description :  오프라인이벤트 클래스
' History : 2010.03.09 한용민 생성
'###########################################################
Class cevent_AddShop
    public fevt_Code 
    public fshopid   
    public fshopName 
    
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class cevent_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fisracknum
	public fevt_code
	public fevt_kind
	public fevt_name
	public fevt_startdate
	public fevt_enddate
	public fevt_state
	public fevt_regdate
	public fevt_using
	public fevt_prizedate
	public fopendate
	public fclosedate
	public fevt_lastupdate
	public fadminid
	public fprizeyn
	public fcategoryname
	public fevt_category
	public fbrand
	public fisgift
	public fisrack
	public fisprize
	public fgift_count
	public fcode_nm
	public fmdname
	public fevt_kinddesc
	public fevt_statedesc
	public fevt_cateMid
	public fpartMDid
	public fevt_forward
	public fevt_icon
	public fevt_comment
	public FChkDisp
	public fshopid
	public fshopname
	public fcentermwdiv
	public fitemid
	public FMakerid
	public fshopitemname
	public fshopitemoptionname
	public fshopitemprice
	public fshopsuplycash
	public forgsellprice
	public fdiscountsellprice
	public fshopbuyprice
	public foffimgmain
	public foffimglist
	public foffimgsmall
	public Fitemgubun
	public Fitemoption
	public Fisusing
	public Fregdate
	public Fextbarcode
	public FOnLineItemprice
	public FOnlineitemorgprice
	public FOnlineOptaddprice
	public FOnlineOptaddbuyprice
	public FShopItemOrgprice
	public FimageSmall
	public Fshopitemid
	public Fsellyn
	public FOnlinedanjongyn
	public fitem_count
	public fprize_count
	public fevtprize_code
	public fevt_ranking
	public fevt_rankname
	public fevt_giftname
	public fevt_winner
	public fevt_winner_name
	public fevtprize_startdate
	public fevtprize_enddate
	public fevtprize_status
	public fgiftkind_code
	public fgiftkind_name
	public fevtprize_type
	public fgive_evtprizecode
	public fimg_basic
	public fissale
	public fsale_count
	public faddShopCnt

	public function GetImageSmall()
		if Fitemgubun="10" then
			GetImageSmall = FimageSmall
		else
			GetImageSmall = FOffImgSmall
		end if
	end function

end class

class cevent_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public frectevt_code
	public FrectSfDate
	public frectevt_startdate
	public frectevt_enddate
	public FrectSfEvt
	public FrectSeTxt
	public FrectScategory
	public FrectScateMid
	public frectevt_state
	public frectpartMDid
	public frectevt_kind
	public frectbrand
	public frectshopid
	public FRectMakerid
	public FRectItemid
	public FRectItemName
	public FRectIsUsing
	public frectcentermwdiv
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public frectisgift
	public frectisrack
	public frectisprize
	public frectissale
	public FRectDesigner
	public FRectItemgubun
	public FRectOnlyUsing
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectOnlineExpiredItem

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

    public Sub getAddedShopList()
        dim sqlStr, i
        sqlStr = "select A.evt_Code,A.AssignShopid , U.shopName"
        sqlStr = sqlStr& " from db_shop.dbo.tbl_event_off_AssignedShop A"
        sqlStr = sqlStr& " 	Left Join db_shop.dbo.tbl_event_off E"
        sqlStr = sqlStr& " 	on A.evt_code=E.evt_code"
        sqlStr = sqlStr& " 	and A.AssignShopid=E.shopid"
        sqlStr = sqlStr& " 	Left Join db_shop.dbo.tbl_Shop_User U"
        sqlStr = sqlStr& " 	on A.AssignShopid=U.userid"
        sqlStr = sqlStr& " where A.evt_code="&frectevt_code
        sqlStr = sqlStr& " and E.shopid Is NULL"
        
        rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount
        
        
        if  not rsget.EOF  then
			redim preserve FItemList(FResultCount)
			
			do until rsget.EOF
				set FItemList(i) = new cevent_AddShop

				FItemList(i).fevt_Code  = rsget("evt_Code")
				FItemList(i).fshopid    = rsget("AssignShopid")
				FItemList(i).fshopName  = rsget("shopName")

				rsget.movenext
				i=i+1
			loop
		end if
        rsget.Close
    end sub

	'///admin/event_off/inc_eventprize.asp
	public sub fnGetPrize_off()
		dim sqlStr,i , strSearch

		IF frectevt_code <> "" THEN
			strSearch  = strSearch & " and evt_code = "&frectevt_code&""
		END IF

		'총 갯수 구하기
		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " count(evtprize_code) as cnt" + vbcrlf
		sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_event_prize_off]  " + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & strSearch

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " evtprize_code, evt_ranking,evt_rankname, a.itemid, evt_giftname,evt_winner,evt_winner_name" + vbcrlf
		sqlStr = sqlStr & " ,evt_regdate ,evtprize_startdate, evtprize_enddate, evtprize_status, a.giftkind_code" + vbcrlf
		sqlStr = sqlStr & " , b.giftkind_name, b.giftkind_img, b.itemid, evtprize_type,give_evtprizecode" + vbcrlf
		sqlStr = sqlStr & " FROM  [db_shop].[dbo].[tbl_event_prize_off] a " + vbcrlf
		sqlStr = sqlStr & " left outer join  [db_shop].[dbo].[tbl_giftkind_off] b " + vbcrlf
		sqlStr = sqlStr & " on a.giftkind_code = b.giftkind_code " + vbcrlf
		sqlStr = sqlStr & " where 1=1  "&strSearch&"" + vbcrlf
		sqlStr = sqlStr & " ORDER BY evt_ranking, evtprize_code desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new cevent_item

				FItemList(i).fevtprize_code = rsget("evtprize_code")
				FItemList(i).fevt_ranking = rsget("evt_ranking")
				FItemList(i).fevt_rankname = rsget("evt_rankname")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fevt_giftname = db2html(rsget("evt_giftname"))
				FItemList(i).fevt_winner = rsget("evt_winner")
				FItemList(i).fevt_winner_name = rsget("evt_winner_name")
				FItemList(i).fevt_regdate = rsget("evt_regdate")
				FItemList(i).fevtprize_startdate = rsget("evtprize_startdate")
				FItemList(i).fevtprize_enddate = rsget("evtprize_enddate")
				FItemList(i).fevtprize_status = rsget("evtprize_status")
				FItemList(i).fgiftkind_code = rsget("giftkind_code")
				FItemList(i).fgiftkind_name = db2html(rsget("giftkind_name"))
				FItemList(i).fevtprize_type = rsget("evtprize_type")
				FItemList(i).fgive_evtprizecode = rsget("give_evtprizecode")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/admin/offshop/gift/giftReg.asp '//admin/offshop/sale/saleReg.asp
	public sub fnGetEventConts()
		dim sqlStr,i , strSearch

		if frectevt_code <> "" then
			strSearch = strSearch + " and A.evt_code = '"&frectevt_code&"' " + vbcrlf
		end if

		'데이터 리스트
		sqlStr = "SELECT  top 1 " + vbcrlf
		sqlStr = sqlStr + " evt_name, evt_startdate, evt_enddate, evt_state, opendate ,closedate ,a.shopid" + vbcrlf
		sqlStr = sqlStr + " ,(select code_desc FROM  [db_shop].[dbo].[tbl_event_off_commoncode] " + vbcrlf
		sqlStr = sqlStr + " 	WHERE code_type = 'evt_state' and code_value = A.evt_state) evt_statedesc" + vbcrlf
		sqlStr = sqlStr + " FROM [db_shop].[dbo].[tbl_event_off] as A " + vbcrlf
		sqlStr = sqlStr + " join [db_shop].[dbo].[tbl_event_off_display] as B " + vbcrlf
		sqlStr = sqlStr + " on A.evt_code = B.evt_code" + vbcrlf
		sqlStr = sqlStr + " WHERE 1=1 " & strSearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FtotalCount = rsget.RecordCount

        set FOneItem = new cevent_item

        if Not rsget.Eof then

			FOneItem.fevt_name = db2html(rsget("evt_name"))
			FOneItem.fshopid = rsget("shopid")
			FOneItem.fevt_startdate = rsget("evt_startdate")
			FOneItem.fevt_enddate = rsget("evt_enddate")
			FOneItem.fevt_state = rsget("evt_state")
			FOneItem.fopendate = rsget("opendate")
			FOneItem.fclosedate = rsget("closedate")
			FOneItem.fevt_statedesc = db2html(rsget("evt_statedesc"))
        end if
        rsget.Close
	end sub

	'/admin/offshop/event_off/eventitem_regist.asp '//common/offshop/pop_eventitem_addinfo_off.asp
	public sub fnGetEventItem()
		dim sqlStr,i , addSql

       '// 추가 쿼리
        if (frectevt_code <> "") then
            addSql = addSql & " and a.evt_code='" + frectevt_code + "'"
        end if

		if FRectItemId<>"" then
			addSql = addSql + " and s.shopitemid=" + CStr(FRectItemId)
		end if

		if FRectItemName<>"" then
			addSql = addSql + " and s.shopitemname like '%" + FRectItemName + "%'"
		end if

		if FRectDesigner<>"" then
			addSql = addSql + " and s.makerid='" + FRectDesigner + "'"
		end if

		if FRectItemgubun<>"" then
			addSql = addSql + " and s.itemgubun='" + FRectItemgubun + "'"
		end if

        if (FRectCDL<>"") then
            addSql = addSql + " and s.catecdl='" + FRectCDL + "'"
        end if

        if (FRectCDM<>"") then
            addSql = addSql + " and s.catecdm='" + FRectCDM + "'"
        end if

        if (FRectCDS<>"") then
            addSql = addSql + " and s.catecdn='" + FRectCDS + "'"
        end if

        if (FRectOnlineExpiredItem<>"") then
            addSql = addSql + " and i.sellyn='N'"
            addSql = addSql + " and i.danjongyn in ('Y','M')"
            addSql = addSql + " and datediff(d,i.regdate,getdate())>91"
        end if

		'총 갯수 구하기
		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " count(A.itemid) as cnt" + vbcrlf
		sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_eventitem_off] AS A " + vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s" + vbcrlf
		sqlStr = sqlStr + " on a.itemid = s.shopitemid" + vbcrlf
		sqlStr = sqlStr + " and a.itemoption = s.itemoption and a.itemgubun = s.itemgubun" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i " + vbcrlf
		sqlStr = sqlStr + " on (s.shopitemid=i.itemid) and s.itemgubun='10'"

		''옵션 추가금액
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on "
		sqlStr = sqlStr + "		s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr & " where s.isusing='Y'" &addSql

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr + " s.itemgubun, a.itemid, s.itemoption,"
		sqlStr = sqlStr + " s.makerid, s.shopitemname, s.shopitemoptionname, s.orgsellprice, "
		sqlStr = sqlStr + " s.shopitemprice, s.shopsuplycash, s.isusing, s.regdate, s.extbarcode, s.discountsellprice "
		sqlStr = sqlStr + " , s.shopbuyprice,s.centermwdiv, i.sellyn, i.danjongyn,"
		sqlStr = sqlStr + " IsNull(i.orgprice,0) as onlineitemorgprice, IsNull(i.sellcash,0) as onlineitemprice ,"
		sqlStr = sqlStr + " IsNULL(i.smallimage,'') as imgsmall, IsNULL(s.offimgsmall,'') as offimgsmall"

		''옵션 추가금액
		sqlStr = sqlStr + " ,IsNULL(o.optaddprice,0) as optaddprice, IsNULL(o.optaddbuyprice,0) as optaddbuyprice"
		sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_eventitem_off] AS A " + vbcrlf
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item s" + vbcrlf
		sqlStr = sqlStr + " on a.itemid = s.shopitemid" + vbcrlf
		sqlStr = sqlStr + " and a.itemoption = s.itemoption and a.itemgubun = s.itemgubun" + vbcrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i " + vbcrlf
		sqlStr = sqlStr + " on (s.shopitemid=i.itemid) and s.itemgubun='10'"

		''옵션 추가금액
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o on "
		sqlStr = sqlStr + "		s.itemgubun='10' and s.shopitemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr & " where s.isusing='Y'" &addSql
		sqlStr = sqlStr + " order by s.itemgubun desc, s.shopitemid desc"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new cevent_item

				FItemList(i).Fitemgubun         = rsget("itemgubun")
				FItemList(i).fshopitemid = rsget("itemid")
				FItemList(i).Fitemoption     	= rsget("itemoption")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Fshopitemname      = db2html(rsget("shopitemname"))
				FItemList(i).Fshopitemoptionname= db2html(rsget("shopitemoptionname"))
				FItemList(i).Fshopitemprice     = rsget("shopitemprice")
				FItemList(i).Fshopsuplycash     = rsget("shopsuplycash")
				FItemList(i).Fisusing           = rsget("isusing")
				FItemList(i).Fregdate           = rsget("regdate")
				FItemList(i).Fextbarcode 		= rsget("extbarcode")
				FItemList(i).FOnLineItemprice	= rsget("onlineitemprice")
                FItemList(i).FOnlineitemorgprice= rsget("onlineitemorgprice")
                ''옵션 추가금액
			    FItemList(i).FOnlineOptaddprice = rsget("optaddprice")
			    FItemList(i).FOnlineOptaddbuyprice = rsget("optaddbuyprice")
				FItemList(i).Fdiscountsellprice = rsget("discountsellprice")
				FItemList(i).Fshopbuyprice		= rsget("shopbuyprice")
                FItemList(i).FShopItemOrgprice  = rsget("orgsellprice")
				FItemList(i).FimageSmall     = rsget("imgsmall")
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
				if FItemList(i).FimageSmall<>"" then FItemList(i).FimageSmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FimageSmall
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = "http://webimage.10x10.co.kr/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).Fshopitemid) + "/" + FItemList(i).FOffimgSmall
                FItemList(i).Fcentermwdiv  = rsget("centermwdiv")
                FItemList(i).Fsellyn        = rsget("sellyn")
                FItemList(i).FOnlinedanjongyn     = rsget("danjongyn")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'///admin/event_off/index.asp
	public sub fnGetEventList_off()
		dim sqlStr,i , strSearch

		If frectevt_startdate <> ""  or frectevt_enddate <> "" THEN
			if CStr(FrectSfDate) = "S" THEN
				strSearch  = strSearch & " and  datediff(day, '"&frectevt_startdate&"', evt_startdate) >= 0 and  datediff(day,'"&frectevt_startdate&"', evt_startdate) <=0  "
			elseif CStr(FrectSfDate) = "E" THEN
				strSearch  = strSearch & " and  datediff(day,'"&frectevt_enddate&"',evt_enddate) >= 0 and  datediff(day,'"&frectevt_enddate&"',evt_enddate) <=0  "
			end if
		END IF
		If FrectSeTxt <> "" THEN
			IF Cstr(FrectSfEvt) = "evt_code" THEN
				strSearch  = strSearch &  " and A.evt_code = "&FrectSeTxt
			ELSE
				strSearch  = strSearch &  " and  evt_name like '%"&FrectSeTxt&"%'"
			END IF
		End If

		If frectevt_state <> "" THEN
			IF frectevt_state = 9 THEN	'종료
				strSearch  = strSearch & " and   (evt_state = "&frectevt_state & " or getdate() >= dateadd(day,+1,a.evt_enddate) )"
			ELSEIF frectevt_state = 6 THEN	'오픈예정
				strSearch  = strSearch & " and   evt_state = 7 and  datediff(day,getdate(),evt_startdate)<= 0 and datediff(day,getdate(),evt_enddate) >= 0  "
			ELSEIF frectevt_state = 7 THEN	'오픈진행중
				strSearch  = strSearch & " and A.evt_state = 7 and getdate() >= a.evt_startdate and getdate() < dateadd(day,+1,a.evt_enddate)"
			ELSE
				strSearch  = strSearch & " and  evt_state = "&frectevt_state & ""
			END IF
		End If

		If FrectScategory <> "" THEN
			strSearch  = strSearch &  " and  evt_category = "&FrectScategory
		END IF
		If FrectScateMid <> "" THEN
			strSearch  = strSearch &  " and  evt_cateMid = "&FrectScateMid
		END IF

		IF frectevt_kind <> "" THEN
			strSearch  = strSearch &  " and evt_kind in ("& frectevt_kind & ") "
		END IF

		IF frectpartMDid <> "" THEN
			strSearch  = strSearch &  " and partMDid = '"&frectpartMDid&"'"
		END IF

		IF frectbrand <> "" THEN
			strSearch  = strSearch & " and brand = '"&frectbrand&"'"
		END IF

		IF frectissale <> "" THEN strSearch  = strSearch & " and issale = '"&frectissale&"'"
		IF frectisgift <> "" THEN strSearch  = strSearch & " and isgift = '"&frectisgift&"'"
		IF frectisrack <> "" THEN strSearch  = strSearch & " and israck = '"&frectisrack&"'"
		IF frectisprize <> "" THEN strSearch  = strSearch & " and isprize = '"&frectisprize&"'"

		IF frectshopid <> "" THEN
			strSearch  = strSearch & " and shopid = '"&frectshopid&"'"
		END IF

		'총 갯수 구하기
		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " count(a.evt_code) as cnt" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_event_off as A " + vbcrlf
		sqlStr = sqlStr & " LEFT JOIN db_shop.dbo.tbl_event_off_display as B " + vbcrlf
		sqlStr = sqlStr & " ON A.evt_code = B.evt_code" + vbcrlf
		sqlStr = sqlStr & " where evt_using ='Y'" & strSearch
		
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " A.evt_code, A.evt_kind, A.evt_name, A.evt_startdate, A.evt_enddate" + vbcrlf
		sqlStr = sqlStr & " , A.evt_regdate, a.shopid ,b.img_basic" + vbcrlf
		sqlStr = sqlStr & " ,(case when A.evt_state = 0 then '0' " + vbcrlf
		sqlStr = sqlStr & " when A.evt_state = 5 then '5'" + vbcrlf
		sqlStr = sqlStr & " when A.evt_state = 7 and getdate() >= a.evt_startdate and getdate() < dateadd(day,+1,a.evt_enddate) then '7'" + vbcrlf
		sqlStr = sqlStr & " when A.evt_state = 9 or getdate() >= dateadd(day,+1,a.evt_enddate) then '9'" + vbcrlf
		sqlStr = sqlStr & " end) as evt_state" + vbcrlf
		sqlStr = sqlStr & " ,(SELECT code_nm from  [db_item].[dbo].tbl_Cate_large " + vbcrlf
		sqlStr = sqlStr & " 	WHERE code_large = B.evt_category) categoryname" + vbcrlf
		sqlStr = sqlStr & " , B.evt_category, A.evt_prizedate ,B.brand, b.issale,b.isgift ,b.israck ,b.isprize" + vbcrlf
		sqlStr = sqlStr & " ,(SELECT COUNT(gift_code) FROM [db_shop].[dbo].[tbl_gift_off] " + vbcrlf
		sqlStr = sqlStr & " 	WHERE evt_code = A.evt_code and gift_using ='y') as gift_count" + vbcrlf
		sqlStr = sqlStr & " ,(SELECT COUNT(itemid) FROM [db_shop].dbo.tbl_eventitem_off " + vbcrlf
		sqlStr = sqlStr & " 	WHERE evt_code = A.evt_code ) as item_count" + vbcrlf
		sqlStr = sqlStr & " ,(SELECT COUNT(evtprize_code) FROM [db_shop].dbo.tbl_event_prize_off " + vbcrlf
		sqlStr = sqlStr & " 	WHERE evt_code = A.evt_code ) as prize_count" + vbcrlf
		sqlStr = sqlStr & " , A.prizeyn , b.isracknum" + vbcrlf
		sqlStr = sqlStr & " ,(select top 1 code_nm from db_item.dbo.tbl_Cate_mid where" + vbcrlf
		sqlStr = sqlStr & " code_large=b.evt_category and code_mid=b.evt_cateMid) as code_nm" + vbcrlf
		sqlStr = sqlStr & " , (Case When isNull(B.partMDid,'')<>'' Then (SELECT username " + vbcrlf
		sqlStr = sqlStr & " 	from db_partner.[dbo].tbl_user_tenbyten WHERE userid = B.partMDid ) Else '' end) as mdname" + vbcrlf
		sqlStr = sqlStr & " ,(select top 1 shopname from [db_shop].[dbo].tbl_shop_user where userid = a.shopid) as shopname" + vbcrlf
		sqlStr = sqlStr & " ,(select count(sale_code) from db_shop.dbo.tbl_sale_off s where a.evt_code = s.evt_code) as sale_count" + vbcrlf
		sqlStr = sqlStr & " ,(select count(*) from db_shop.dbo.tbl_event_off_AssignedShop T where T.evt_code=A.evt_code and T.AssignShopid<>A.shopid) as addShopCnt" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_event_off as A " + vbcrlf
		sqlStr = sqlStr & " LEFT JOIN db_shop.dbo.tbl_event_off_display as B " + vbcrlf
		sqlStr = sqlStr & " ON A.evt_code = B.evt_code" + vbcrlf
		sqlStr = sqlStr & " where evt_using ='Y' "&strSearch&"" + vbcrlf
		sqlStr = sqlStr & " ORDER BY A.evt_code DESC"

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
				set FItemList(i) = new cevent_item

				FItemList(i).fsale_count = rsget("sale_count")
				FItemList(i).fisracknum = rsget("isracknum")
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fevt_kind = rsget("evt_kind")
				FItemList(i).fevt_name = db2html(rsget("evt_name"))
				FItemList(i).fevt_startdate = rsget("evt_startdate")
				FItemList(i).fevt_enddate = rsget("evt_enddate")
				FItemList(i).fevt_regdate = rsget("evt_regdate")
				FItemList(i).fevt_state = rsget("evt_state")
				FItemList(i).fcategoryname = db2html(rsget("categoryname"))
				FItemList(i).fevt_category = rsget("evt_category")
				FItemList(i).fevt_prizedate = rsget("evt_prizedate")
				FItemList(i).fbrand = db2html(rsget("brand"))
				FItemList(i).fissale = rsget("issale")
				FItemList(i).fisgift = rsget("isgift")
				FItemList(i).fisrack = rsget("israck")
				FItemList(i).fisprize = rsget("isprize")
				FItemList(i).fgift_count = rsget("gift_count")
				FItemList(i).fprizeyn = rsget("prizeyn")
				FItemList(i).fcode_nm = rsget("code_nm")
				FItemList(i).fitem_count = rsget("item_count")
				FItemList(i).fprize_count = rsget("prize_count")
				FItemList(i).fmdname = db2html(rsget("mdname"))
				FItemList(i).fimg_basic = db2html(rsget("img_basic"))
                FItemList(i).faddShopCnt = rsget("addShopCnt")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/event_off/event_modify.asp
	public sub fnGetEventDisplay_off()
        dim sqlStr , strSearch

        '//검색
		If frectevt_code <> ""  THEN
			strSearch  = strSearch & " and evt_code ="&frectevt_code&""
		END IF

		'데이터 리스트
		sqlStr = "SELECT  " + vbcrlf
		sqlStr = sqlStr & " evt_category, evt_cateMid, isgift ,israck ,isprize, isNull(partMDid,'') as partMDid" + vbcrlf
		sqlStr = sqlStr & " , evt_forward, brand, evt_icon, evt_comment ,isracknum ,img_basic, issale" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_event_off_display" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " &strSearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FtotalCount = rsget.RecordCount

        set FOneItem = new cevent_item

        if Not rsget.Eof then

			FOneItem.FChkDisp = 1
			FOneItem.fisracknum = rsget("isracknum")
			FOneItem.fevt_category = rsget("evt_category")
			FOneItem.fevt_cateMid = rsget("evt_cateMid")
			FOneItem.fissale = rsget("issale")
			FOneItem.fisgift = rsget("isgift")
			FOneItem.fisrack = rsget("israck")
			FOneItem.fisprize = rsget("isprize")
			FOneItem.fpartMDid = rsget("partMDid")
			FOneItem.fevt_forward = db2html(rsget("evt_forward"))
			FOneItem.fbrand = rsget("brand")
			FOneItem.fevt_icon = rsget("evt_icon")
			FOneItem.fevt_comment = db2html(rsget("evt_comment"))
			FOneItem.fimg_basic = db2html(rsget("img_basic"))

        end if
        rsget.Close
    end Sub

	'//admin/event_off/event_modify.asp
	public sub fnGetEventCont_off()
        dim sqlStr , strSearch

        '//검색
		If frectevt_code <> ""  THEN
			strSearch  = strSearch & " and A.evt_code ="&frectevt_code&""
		END IF

		'데이터 리스트
		sqlStr = "SELECT  " + vbcrlf
		sqlStr = sqlStr & " evt_kind, evt_name, evt_startdate, evt_enddate" + vbcrlf
		sqlStr = sqlStr & " , evt_prizedate, evt_state, evt_regdate, evt_using, opendate" + vbcrlf
		sqlStr = sqlStr & " , closedate, a.prizeyn , a.shopid " + vbcrlf
		sqlStr = sqlStr & " ,(select code_desc FROM  db_shop.[dbo].[tbl_event_off_commoncode] " + vbcrlf
		sqlStr = sqlStr & " 	WHERE code_type = 'evt_kind' and code_value = a.evt_kind) evt_kinddesc" + vbcrlf
		sqlStr = sqlStr & " ,(select code_desc FROM  db_shop.[dbo].[tbl_event_off_commoncode] " + vbcrlf
		sqlStr = sqlStr & " 	WHERE code_type = 'evt_state' and code_value = a.evt_state) evt_statedesc" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_event_off a" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " &strSearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FtotalCount = rsget.RecordCount

        set FOneItem = new cevent_item

        if Not rsget.Eof then

			FOneItem.fevt_name = db2html(rsget("evt_name"))
			FOneItem.fevt_kind = rsget("evt_kind")
			FOneItem.fevt_startdate = rsget("evt_startdate")
			FOneItem.fevt_enddate = rsget("evt_enddate")
			FOneItem.fevt_prizedate = rsget("evt_prizedate")
			FOneItem.fevt_state = rsget("evt_state")
			FOneItem.fevt_regdate = rsget("evt_regdate")
			FOneItem.fevt_using = rsget("evt_using")
			FOneItem.fopendate = rsget("opendate")
			FOneItem.fclosedate = rsget("closedate")
			FOneItem.fprizeyn = rsget("prizeyn")
			FOneItem.fevt_kinddesc = db2html(rsget("evt_kinddesc"))
			FOneItem.fevt_statedesc = db2html(rsget("evt_statedesc"))
			FOneItem.fshopid = rsget("shopid")

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
end Class

%>

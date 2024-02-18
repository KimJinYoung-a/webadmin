<%
'####################################################
' Description : 판매대기상품LIST 클래스
' History : 이상구 생성
'			2023.10.4 한용민 수정(수정로그추가)
'####################################################

class CSellWaitBrandSum
    public FMakerID
    public Fpurchasetype
    public FCount
	public fpurchasetypename

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
    public function getPurchaseTypeName
        select CASE Fpurchasetype
            CASE "1"
                : getPurchaseTypeName="일반유통"
            CASE "4"
                : getPurchaseTypeName="사입"
            CASE "5"
                : getPurchaseTypeName="OFF사입"
            CASE "6"
                : getPurchaseTypeName="수입"
            CASE "7"
                : getPurchaseTypeName="브랜드수입"
            CASE "8"
                : getPurchaseTypeName="제작"
            CASE "9"
                : getPurchaseTypeName="해외직구"
            CASE "10"
                : getPurchaseTypeName="B2B"
            CASE ELSE
                : getPurchaseTypeName=""
        end select
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CItem
	public FItemID
	public FItemName
	public FMakerID
	public FSellPrice
	public FSellVat
	public FMarginrate
	public FBuyPrice
	public FBuyvat
	public FVatInclude
	public FSellYn
	public FBaesongGB
	public FMarginDiv
	public FPojangYn
	public FLimitYn
	public FLimitDiv
	public FLimitNo
	public FLimitSold
	public fisusing
	public FItemOptionName
	public FMwDiv

	public FImageList
	public FImageSmall
	public Fregdate
	public Fipgono
	public Fpreorderno
	public FSellno
	public Fchulno
	public Fcurrno

	public function getMwDivName()
		if FmwDiv="M" then
			getMwDivName = "매입"
		elseif FmwDiv="W" then
			getMwDivName = "위탁"
		elseif FmwDiv="U" then
			getMwDivName = "업체"
		end if
	end function

	public function getMwDivColor()
		if FmwDiv="M" then
			getMwDivColor = "#CC2222"
		elseif FmwDiv="W" then
			getMwDivColor = "#2222CC"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CItemList
	public FItemList()

	public FSearchItemid
	public FSearchItemName
	public FSearchDesigner
	public FSearchSellYn
	public FSearchLimitYn
	public FSearchBaedalDiv

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemId
	public FRectMakerid
	public FRectIpgoGubun
	public FRectDeliverType

	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectDispCate

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

    public Sub getSellWaitItemListByBrand()
        dim sqlStr, i, addSql
        dim regStDT : regStDT=LEFT(dateadd("d",-90,now()),10)

        if FRectItemid<>"" then
			addSql = addSql + " where i.itemid=" + FRectItemid + "" + VbCrlf
		else
			addSql = addSql + " where datediff(d,i.regdate,getdate())<90" + VbCrlf		'90일간 쿼리
			''addSql = addSql + " where i.regdate>='"&regStDT&"'" + VbCrlf		'90일간 쿼리
			addSql = addSql + " and s.sellno=0"

			if FRectMakerid<>"" then
				addSql = addSql + " and i.makerid='" + FRectMakerid + "'" + VbCrlf
			end if

			if FRectCate_Large<>"" then
	            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
	        end if

	        if FRectCate_Mid<>"" then
	            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
	        end if

	        if FRectCate_Small<>"" then
	            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
	        end if

			if FRectDeliverType="U" then
				addSql = addSql + " and i.deliverytype in ('2','5','9','7')" + VbCrlf
				''addSql = addSql + " and i.mwdiv in ('U')"
			else
				addSql = addSql + " and i.deliverytype in ('1','3','4')" + VbCrlf
				''addSql = addSql + " and i.mwdiv in ('M','W')"
			end if

			'if FRectMakerid="" or isNull(FRectMakerid) then
				if FRectIpgoGubun="BY" then
					addSql = addSql + " and IsNULL(s.totipgono,0)>0"
				else
					addSql = addSql + " and IsNULL(s.totipgono,0)=0"
				end if
			'end if

			'if FRectMakerid="" or isNull(FRectMakerid) then
				addSql = addSql + " and i.sellyn='N'" + VbCrlf
			'end if
		end if

		addSql = addSql + " and i.isusing='Y'" + VbCrlf
      ''  addSql = addSql + " and i.makerid not in ('Tilly','BONknot','BONlynns','onefineday1010')" ''일단제외. 2016/12/27
        IF FRectDispCate<>"" THEN ''1 depth 인경우 여길 타게.
            if (LEN(FRectDispCate)=3) then
		        addSql = addSql + " and i.dispcate1='"&FRectDispCate&"'"
		    end if
        end if

		'// 결과 카운트
		sqlStr = "select i.makerid, p.purchasetype, pc.pcomm_name as purchasetypename, count(*) as CNT "& VbCrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i with (nolock)" + VbCrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v with (nolock) on i.itemid=v.itemid and v.isusing='Y'" + VbCrlf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
		sqlStr = sqlStr + " on i.itemid=s.itemid and IsNULL(v.itemoption,'0000')=s.itemoption" + VbCrlf
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_partner p with (nolock)"
		sqlStr = sqlStr + " on i.makerid=p.id"
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and p.purchasetype=pc.pcomm_cd"

		IF FRectDispCate<>"" THEN	'2014-08-07 김진영 전시카테고리 검색 추가
		    if (LEN(FRectDispCate)>3) then
			    sqlStr = sqlStr + "  JOIN db_item.dbo.tbl_display_cate_item as dc "
			    sqlStr = sqlStr + " on i.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		    end if
		END IF
		sqlStr = sqlStr + addSql
        sqlStr = sqlStr + " group by i.makerid, p.purchasetype, pc.pcomm_name"
        sqlStr = sqlStr + " order by i.makerid"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount
		Ftotalcount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
    		do until rsget.eof
    			set FItemList(i) = new CSellWaitBrandSum
				FItemList(i).fpurchasetypename = rsget("purchasetypename")
    			FItemList(i).FMakerid = rsget("makerid")
    			FItemList(i).Fpurchasetype= rsget("purchasetype")
    			FItemList(i).Fcount = rsget("CNT")
    		    rsget.movenext
    			i=i+1
    		loop
		end if
		rsget.close
    end Sub

	' /admin/shopmaster/item_new_list.asp
	public Sub getSellWaitItemList()
		dim sqlStr, i, addSql
        dim regStDT : regStDT=LEFT(dateadd("d",-90,now()),10)
		'// 조건절 쿼리
		if FRectItemid<>"" then
			addSql = addSql + " where i.itemid=" + FRectItemid + "" + VbCrlf
		else
			addSql = addSql + " where datediff(d,i.regdate,getdate())<90" + VbCrlf		'90일간 쿼리
			''addSql = addSql + " where i.regdate>='"&regStDT&"'" + VbCrlf		'90일간 쿼리
			addSql = addSql + " and sellno=0"

			if FRectMakerid<>"" then
				addSql = addSql + " and i.makerid='" + FRectMakerid + "'" + VbCrlf
			end if

			if FRectCate_Large<>"" then
	            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
	        end if

	        if FRectCate_Mid<>"" then
	            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
	        end if

	        if FRectCate_Small<>"" then
	            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
	        end if

			if FRectDeliverType="U" then
				addSql = addSql + " and i.deliverytype in ('2','5','9','7')" + VbCrlf
				''addSql = addSql + " and i.mwdiv in ('U')"
			else
				addSql = addSql + " and i.deliverytype in ('1','3','4')" + VbCrlf
				''addSql = addSql + " and i.mwdiv in ('M','W')"
			end if

			'if FRectMakerid="" or isNull(FRectMakerid) then
				if FRectIpgoGubun="Y" then
					addSql = addSql + " and IsNULL(s.totipgono,0)>0"
				else
					addSql = addSql + " and IsNULL(s.totipgono,0)=0"
				end if
			'end if

			'if FRectMakerid="" or isNull(FRectMakerid) then
				addSql = addSql + " and i.sellyn='N'" + VbCrlf
			'end if
		end if

		addSql = addSql + " and i.isusing='Y'" + VbCrlf
     ''   addSql = addSql + " and i.makerid not in ('Tilly','BONknot','BONlynns','onefineday1010')" ''일단제외. 2016/12/27
        IF FRectDispCate<>"" THEN
            if (LEN(FRectDispCate)=3) then
		        addSql = addSql + " and i.dispcate1='"&FRectDispCate&"'"
		    end if
        end if

		'// 결과 카운트
		sqlStr = "select count(i.itemid), CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ")"& VbCrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i with (nolock)" + VbCrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v with (nolock) on i.itemid=v.itemid and v.isusing='Y'" + VbCrlf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
		sqlStr = sqlStr + " on s.itemgubun = '10' and i.itemid=s.itemid and IsNULL(v.itemoption,'0000')=s.itemoption" + VbCrlf
		IF FRectDispCate<>"" THEN	'2014-08-07 김진영 전시카테고리 검색 추가
		    if (LEN(FRectDispCate)>3) then
			    sqlStr = sqlStr + "  JOIN db_item.dbo.tbl_display_cate_item as dc "
			    sqlStr = sqlStr + " on i.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
			end if
		END IF
		sqlStr = sqlStr + addSql

		''response.write sqlStr

		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FTotalPage = rsget(1)
		rsget.Close

		'// 본문 쿼리
		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & " i.itemid,i.makerid,i.itemname," + VbCrlf
		sqlStr = sqlStr + " IsNULL(v.itemoption,'0000') as itemoption, IsNULL(v.optionname,'') as itemoptionname," + VbCrlf
		sqlStr = sqlStr + " sellcash, buycash, mwdiv, sellyn, deliverytype, limityn, limitno, limitsold," + VbCrlf
		sqlStr = sqlStr + " smallimage, listimage, i.regdate, IsNULL(s.totipgono,0) as ipgono, IsNULL(s.preorderno,0) as preorderno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(s.totsellno,0) as sellno, IsNULL(s.totchulgono,0) as chulno , IsNULL(s.realstock,0) as currno" + VbCrlf
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i with (nolock)" + VbCrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v with (nolock) on i.itemid=v.itemid and v.isusing='Y'" + VbCrlf
		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)"
		sqlStr = sqlStr + " on s.itemgubun = '10' and i.itemid=s.itemid and IsNULL(v.itemoption,'0000')=s.itemoption" + VbCrlf
		IF FRectDispCate<>"" THEN	'2014-08-07 김진영 전시카테고리 검색 추가
		    if (LEN(FRectDispCate)>3) then
			    sqlStr = sqlStr + "  JOIN db_item.dbo.tbl_display_cate_item as dc "
			    sqlStr = sqlStr + " on i.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
			end if
		END IF
		sqlStr = sqlStr + addSql
		sqlStr = sqlStr + " order by i.itemid desc"

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		rsget.absolutepage = FCurrPage
		do until rsget.eof
			set FItemList(i) = new CItem
			FItemList(i).FItemID    = rsget("itemid")
			FItemList(i).FItemName  = db2html(rsget("itemname"))
			FItemList(i).FItemOptionName  = db2html(rsget("itemoptionname"))
			FItemList(i).FMakerID   = rsget("makerid")
			FItemList(i).FSellPrice = rsget("sellcash")
			FItemList(i).FBuyPrice  = rsget("buycash")
			FItemList(i).FMwDiv= rsget("mwdiv")
			FItemList(i).FSellYn    = rsget("sellyn")
			FItemList(i).FBaesongGB = rsget("deliverytype")

			FItemList(i).FLimitYn = rsget("limityn")
			FItemList(i).FLimitNo = rsget("limitno")
			FItemList(i).FLimitSold = rsget("limitsold")
			FItemList(i).FRegdate = rsget("regdate")

			FItemList(i).Fipgono = rsget("ipgono")
			FItemList(i).Fpreorderno = rsget("preorderno")
			FItemList(i).FSellno = rsget("sellno")
			FItemList(i).Fchulno = rsget("chulno")
			FItemList(i).Fcurrno = rsget("currno")
			FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("listimage")
			FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("smallimage")

			rsget.movenext
			i=i+1
		loop
		end if
		rsget.close
	end sub

	' /admin/shopmaster/item_new_list.asp
	public Sub GetgiftItemIpgo()
		dim sqlStr, i, addSql
        dim regStDT : regStDT=LEFT(dateadd("d",-90,now()),10)

		if FRectItemid<>"" then
			addSql = addSql + " and si.shopitemid=" + FRectItemid + "" + VbCrlf
		end if

		'addSql = addSql + " and datediff(d,si.regdate,getdate())<90" + VbCrlf		'90일간 쿼리
		'addSql = addSql + " and sellno=0"

		if FRectMakerid<>"" then
			addSql = addSql + " and si.makerid='" + FRectMakerid + "'" + VbCrlf
		end if

		if FRectCate_Large<>"" then
            addSql = addSql + " and si.catecdl='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and si.catecdm='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and si.catecdn='" + FRectCate_Small + "'"
        end if

		if FRectIpgoGubun="GIY" then
			addSql = addSql + " and IsNULL(s.totipgono,0)>0"
		else
			addSql = addSql + " and IsNULL(s.totipgono,0)=0"
		end if

     ''   addSql = addSql + " and i.makerid not in ('Tilly','BONknot','BONlynns','onefineday1010')" ''일단제외. 2016/12/27
        IF FRectDispCate<>"" THEN
            if (LEN(FRectDispCate)=3) then
		        addSql = addSql + " and i.dispcate1='"&FRectDispCate&"'"
		    end if
        end if

		sqlStr = "select count(si.shopitemid), CEILING(CAST(Count(si.shopitemid) AS FLOAT)/" & FPageSize & ")"& VbCrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item si with (nolock)" & vbcrlf

		IF FRectDispCate<>"" THEN	'2014-08-07 김진영 전시카테고리 검색 추가
		    if (LEN(FRectDispCate)>3) then
			    sqlStr = sqlStr + " JOIN db_item.dbo.tbl_display_cate_item as dc with (nolock)"
			    sqlStr = sqlStr + " 	on i.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
				sqlStr = sqlStr & " 	and si.itemgubun='10'" & vbcrlf
			end if
		END IF

		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on si.shopitemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and si.itemgubun='10'" & vbcrlf
		'sqlStr = sqlStr & " 	and i.sellyn='N'" & vbcrlf
		sqlStr = sqlStr & " 	and i.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option v with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on si.shopitemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and si.itemoption=v.itemoption" & vbcrlf
		sqlStr = sqlStr & " 	and si.itemgubun='10'" & vbcrlf
		sqlStr = sqlStr & " 	and v.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on si.itemgubun = s.itemgubun and si.shopitemid=s.itemid and si.itemoption=s.itemoption" & vbcrlf
		sqlStr = sqlStr & " where si.isusing='Y' and s.itemgubun in ('80','85') " & addSql

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget(0)
			FTotalPage = rsget(1)
		rsget.Close

		if FTotalCount < 1 then exit Sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		'// 본문 쿼리
		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & " si.shopitemid as itemid ,si.makerid ,si.shopitemname as itemname" + VbCrlf
		sqlStr = sqlStr + " , si.itemoption, si.shopitemoptionname as itemoptionname" + VbCrlf
		sqlStr = sqlStr + " , si.shopitemprice, si.shopsuplycash, si.isusing, si.regdate, IsNULL(s.totipgono,0) as ipgono, IsNULL(s.preorderno,0) as preorderno," + VbCrlf
		sqlStr = sqlStr + " IsNULL(s.totsellno,0) as sellno, IsNULL(s.totchulgono,0) as chulno , IsNULL(s.realstock,0) as currno" + VbCrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item si with (nolock)" & vbcrlf

		IF FRectDispCate<>"" THEN	'2014-08-07 김진영 전시카테고리 검색 추가
		    if (LEN(FRectDispCate)>3) then
			    sqlStr = sqlStr + " JOIN db_item.dbo.tbl_display_cate_item as dc with (nolock)"
			    sqlStr = sqlStr + " 	on i.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
				sqlStr = sqlStr & " 	and si.itemgubun='10'" & vbcrlf
			end if
		END IF

		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on si.shopitemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and si.itemgubun='10'" & vbcrlf
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option v with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on si.shopitemid = i.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and si.itemoption=v.itemoption" & vbcrlf
		sqlStr = sqlStr & " 	and si.itemgubun='10'" & vbcrlf
		sqlStr = sqlStr & " 	and v.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join [db_summary].[dbo].tbl_current_logisstock_summary s with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on si.itemgubun = s.itemgubun and si.shopitemid=s.itemid and si.itemoption=s.itemoption" & vbcrlf
		sqlStr = sqlStr & " where si.isusing='Y' and s.itemgubun in ('80','85') " & addSql
		sqlStr = sqlStr & " order by si.shopitemid desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		rsget.absolutepage = FCurrPage
		do until rsget.eof
			set FItemList(i) = new CItem
			FItemList(i).FItemID    = rsget("itemid")
			FItemList(i).FItemName  = db2html(rsget("itemname"))
			FItemList(i).FItemOptionName  = db2html(rsget("itemoptionname"))
			FItemList(i).FMakerID   = rsget("makerid")
			FItemList(i).FSellPrice = rsget("shopitemprice")
			FItemList(i).FBuyPrice  = rsget("shopsuplycash")
			FItemList(i).fisusing    = rsget("isusing")
			FItemList(i).FRegdate = rsget("regdate")
			FItemList(i).Fipgono = rsget("ipgono")
			FItemList(i).Fpreorderno = rsget("preorderno")
			FItemList(i).FSellno = rsget("sellno")
			FItemList(i).Fchulno = rsget("chulno")
			FItemList(i).Fcurrno = rsget("currno")

			rsget.movenext
			i=i+1
		loop
		end if
		rsget.close
	end sub

	public sub getItemList()
		dim sqlStr
		dim sqlrect
		dim i

		sqlStr = "select count(i.itemid) as cnt from tbl_item i , tbl_item_image m"
		sqlStr = sqlStr + " where i.itemid=m.itemid"


		if (FSearchItemid<>"") then
			sqlrect = sqlrect + " and i.itemid=" + CStr(FSearchItemid)
		end if

		if (FSearchItemName<>"") then
			sqlrect = sqlrect + " and i.itemname like '%" + CStr(FSearchItemname) + "%'"
		end if

		if (FSearchDesigner<>"") then
			sqlrect = sqlrect + " and i.makerid = '" + CStr(FSearchDesigner) + "'"
		end if


		if (FSearchSellYn<>"") then
			sqlrect = sqlrect + " and i.sellyn = '" + CStr(FSearchSellYn) + "'"
		end if

		if (FSearchLimitYn<>"") then
			sqlrect = sqlrect + " and i.limityn = '" + CStr(FSearchLimitYn) + "'"
		end if

		if (FSearchBaedalDiv<>"") then
			sqlrect = sqlrect + " and i.deliverytype = '" + CStr(FSearchBaedalDiv) + "'"
		end if



		rsget.Open sqlStr + sqlrect,dbget,1
		FTotalCount = rsget("cnt")
		rsget.close

		sqlrect = sqlrect + " order by i.itemid desc"

		sqlStr = "select top " + CStr(FPageSize)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.makerid, i.buycash, i.buyvat, i.sellcash, i.sellvat, i.margin, i.sellyn, i.deliverytype, i.vatinclude, i.pojangok, i.limityn, i.limitdiv, i.limitno, i.limitsold, "
		sqlStr = sqlStr + " m.imglist, m.imgsmall"
		sqlStr = sqlStr + " from tbl_item i, tbl_item_image m"
		sqlStr = sqlStr + " where i.itemid=m.itemid"
		sqlStr = sqlStr + " and i.itemid not in ("
		sqlStr = sqlStr + " select top " + CStr((FCurrPage-1)*FPageSize)  + " i.itemid from tbl_item i, tbl_item_image m "
		sqlStr = sqlStr + " where i.itemid=m.itemid"
		sqlStr = sqlStr + sqlrect
		sqlStr = sqlStr + " )"
		rsget.Open sqlStr + sqlrect,dbget,1

		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)
		i=0
		do until rsget.eof
			set FItemList(i) = new CItem
			FItemList(i).FItemID    = rsget("itemid")
			FItemList(i).FItemName  = rsget("itemname")
			FItemList(i).FMakerID   = rsget("makerid")
			FItemList(i).FSellPrice = rsget("sellcash")
			FItemList(i).FSellVat   = rsget("sellvat")
			FItemList(i).FMarginrate= rsget("margin")
			FItemList(i).FBuyPrice  = rsget("buycash")
			FItemList(i).FBuyvat    = rsget("buyvat")
			FItemList(i).FVatInclude= rsget("vatinclude")
			FItemList(i).FSellYn    = rsget("sellyn")
			FItemList(i).FBaesongGB = rsget("deliverytype")
			FItemList(i).FPojangYn = rsget("pojangok")

			FItemList(i).FLimitYn = rsget("limityn")
			FItemList(i).FLimitNo = rsget("limitno")
			FItemList(i).FLimitSold = rsget("limitsold")
			FItemList(i).FLimitDiv = rsget("limitdiv")

			FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imglist")
			FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemId) + "/" + rsget("imgsmall")

			rsget.movenext
			i=i+1
		loop
		rsget.close
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

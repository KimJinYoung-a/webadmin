<%

Sub drawSelectBoxByUserDiv(puserdivinclude, puserdivexclude, cuserdivinclude, cuserdivexclude, selectBoxName, selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
	query1 = " select c.userid,c.socname "
	query1 = query1 & " from [db_partner].[dbo].tbl_partner p "
	query1 = query1 & " left join db_user.dbo.tbl_user_c c "
	query1 = query1 & " on c.userid=p.id "
	query1 = query1 & " where 1=1 "
	query1 = query1 & " and c.isusing='Y' "

	if (puserdivinclude <> "") then
		query1 = query1 & " and p.userdiv in (" + CStr(puserdivinclude) + ") "
	end if

	if (puserdivexclude <> "") then
		query1 = query1 & " and p.userdiv not in (" + CStr(puserdivexclude) + ") "
	end if

	if (cuserdivinclude <> "") then
		query1 = query1 & " and c.userdiv in (" + CStr(cuserdivinclude) + ") "
	end if

	if (cuserdivexclude <> "") then
		query1 = query1 & " and c.userdiv not in (" + CStr(cuserdivexclude) + ") "
	end if

	query1 = query1 & " order by c.userid "
	rsget.Open query1,dbget,1
	''response.write query1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("socname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end sub

Function DrawShopDivBox(shopdiv)
    dim buf
    '' 수출(7), 도매(5), 내부거래(1)
    buf = "<select name='shopdiv' class='select'>"
    buf = buf & "<option value=''>선택"
    buf = buf & "<option value='3' "&CHKIIF(shopdiv="3","selected","")&">가맹"
    buf = buf & "<option value='5' "&CHKIIF(shopdiv="5","selected","")&">도매"
    buf = buf & "<option value='7' "&CHKIIF(shopdiv="7","selected","")&">수출"
    buf = buf & "<option value='1' "&CHKIIF(shopdiv="1","selected","")&">내부"
    buf = buf & "<option value='9' "&CHKIIF(shopdiv="9","selected","")&">영세"
    buf = buf & "<option value='11' "&CHKIIF(shopdiv="11","selected","")&">띵소"
    buf = buf & "<option value='13' "&CHKIIF(shopdiv="13","selected","")&">제휴"
    buf = buf & "<option value='15' "&CHKIIF(shopdiv="15","selected","")&">용역"
    buf = buf & "</select>"

    response.write buf
End function

function fnGetShopName(shopid, byREF shopdiv, byREF papertype)
	dim sqlStr

    sqlStr = " select shopname, shopdiv "
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shop_user "
    sqlStr = sqlStr + " where userid='"&shopid&"'"

	'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1
    If Not rsget.EOF then
        fnGetShopName = rsget("shopname")
        shopdiv       = rsget("shopdiv")
    end if
    rsget.close

    if (shopdiv="2") then shopdiv="1"
    if (shopdiv="4") then shopdiv="3"
    if (shopdiv="6") then shopdiv="5"
    if (shopdiv="8") then shopdiv="7"
    if (shopdiv="12") then shopdiv="11"

	if (shopdiv = "7") then
		'// 수출신고필증
		papertype = "200"
	else
		'// 세금계산서
		papertype = "100"
	end if

    ''iTs
    ''if (shopid="streetshop874") then shopdiv="1"
    ''if (shopid="streetshop884") then shopdiv="1"
    ''29cm
    ''if (shopid="streetshop878") then shopdiv="1"
    ''텐텐
    if (shopid="cafe003") then shopdiv="1"
end function

Class CWitakSellJungsanTargetItem
	public Fidx
	public Fyyyymm
	public Fshopid
	public Fmakerid
	public Fjungsanid
	public Ftotitemcnt
	public Ftotorgsum
	public Ftotsum
	public Fminuscharge
	public Fchargepercent
	public Frealjungsansum
	public Fcurrstate
	public Fchargediv
	public Ffranchargediv
	public Fgroupidx
	public Foffgubun
	public FCurrchargediv
	public Fdefaultmargin
	public Fdefaultsuplymargin
	public Fprecheckidx
	public fyyyymmdd
	public fbuyprice
	public Fshopname

    public Ftotdeliveritemcnt
    public Ftotdeliverorgsum
    public Ftotdeliversum
    public Fbuydeliverprice

    Public Fbizsection_cd
    Public Fbizsection_nm

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CFranChulgojungsanTargetItem
	public Fid
	public Fcode
	public Fsocid
	public Fdivcode
	public Fexecutedt
	public Fscheduledate
	public FjumunRegDate
	public Ftotalsellcash
	public Ftotalsuplycash
	public Ftotalbuycash
	public Fbaljuidx
	public Fjumunrealsellcash
	public Fjumunrealsuplycash
	public Fjumunrealbuycash
	public Fipgodate
	public Fbaljucode
	public Fprecheckmasteridx
	public Fprecheckidx
	public Fbaljusegumdate

	public Fworktitle
	public Fworkstate
	public Fworkidx
	public Fbaljudate

	public FOrderStateCD

	public Fshopname
    Public Fbizsection_cd
    Public Fbizsection_nm

	public function GetOrderStateName()
		if FOrderStateCD="0" then
			GetOrderStateName = "주문접수"
		elseif FOrderStateCD="1" then
			GetOrderStateName = "주문확인"
		elseif FOrderStateCD="2" then
			GetOrderStateName = "입금대기"
		elseif FOrderStateCD="5" then
			GetOrderStateName = "배송준비"
		elseif FOrderStateCD="6" then
			GetOrderStateName = "출고대기"
		elseif FOrderStateCD="7" then
			GetOrderStateName = "출고완료"
		elseif FOrderStateCD="8" then
			GetOrderStateName = "입고대기"
		elseif FOrderStateCD="9" then
			GetOrderStateName = "입고완료"
		elseif FOrderStateCD=" " then
			GetOrderStateName = "작성중"
		end if
	end function

	public function GetOrderStateColor()
		if FOrderStateCD="0" then
			GetOrderStateColor = "#00000"
		elseif FOrderStateCD="1" then
			GetOrderStateColor = "#00AA00"
		elseif FOrderStateCD="2" then
			GetOrderStateColor = "#0000AA"
		elseif FOrderStateCD="5" then
			GetOrderStateColor = "#AAAA00"
		elseif FOrderStateCD="6" then
			GetOrderStateColor = "#AA00AA"
		elseif FOrderStateCD="7" then
			GetOrderStateColor = "#AA0000"
		elseif FOrderStateCD="8" then
			GetOrderStateColor = "#33AAAA"
		elseif FOrderStateCD="9" then
			GetOrderStateColor = "#AA33AA"
		elseif FOrderStateCD=" " then
			GetOrderStateColor = "#AAAAAA"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEtcMeachulMasterItem
	public Fidx
	public Fshopid
	public Ftitle
	public Ftotalsum
	public Ftotalsellcash
	public Ftotalbuycash
	public Ftotalsuplycash
	public Fdivcode
	public Ftaxdate
	public Ftaxregdate
	public Fregdate
	public Fipkumdate
	public Fetcstr
	public FStateCD
	public Freguserid
	public Fregusername
	public Ffinishuserid
	public Ffinishusername
	Public FtaxNo
	Public FbizNo

    Public Fyyyymm
    Public FdiffKey
    Public Fshopdiv
    Public FbrandDiv

    Public Fworkidx
    Public Fdelivermethod
    Public Finvoiceidx

    Public Ftotmatchedipkumsum
    Public Fmaymatchedipkumsum

    Public Fbizsection_cd
    Public Fbizsection_nm
    Public Fpapertype
    Public Fpaperissuetype
    Public Fetcpaperidx
    Public Fselltype
    Public Fselltypenm

    Public Fissuestatecd
    Public Fipkumstatecd

    Public Feserotaxkey
    Public Ftaxlinkidx

	Public Fjungsan_acctname

	public function GetBrandDivName()
		if FBrandDiv="02" then
			GetBrandDivName = "매입처"
		elseif FBrandDiv="14" then
			GetBrandDivName = "강사"
		elseif FBrandDiv="21" then
			GetBrandDivName = "출고처"
	    elseif FBrandDiv="50" then
			GetBrandDivName = "제휴사(온라인)"
		elseif FBrandDiv="95" then
			GetBrandDivName = "사용안함"
		else
			GetBrandDivName = FBrandDiv
		end if
	end function

    public function getShopDivName()
        SELECT CASE Fshopdiv
            CASE "1" : getShopDivName = "내부"  ''직영
            CASE "3" : getShopDivName = "가맹"
            CASE "5" : getShopDivName = "도매"
            CASE "7" : getShopDivName = "수출"
            CASE "9" : getShopDivName = "영세"
            CASE "11" : getShopDivName = "띵소"
            CASE "13" : getShopDivName = "제휴"
            CASE "15" : getShopDivName = "용역"
            CASE ELSE : getShopDivName = Fshopdiv
        ENd SELECT
    end function

	public function GetDivCodeName()
		if Fdivcode="MC" then
			GetDivCodeName = "출고분정산"
		elseif Fdivcode="WS" then
			GetDivCodeName = "판매분정산(가맹점)"
		elseif Fdivcode="GC" then
			GetDivCodeName = "가맹비"
		elseif Fdivcode="ET" then
			GetDivCodeName = "기타매출"
		elseif Fdivcode="TC" then
			GetDivCodeName = "B2C매출" '// 사용않함
		elseif Fdivcode="AA" then
			GetDivCodeName = "판매분정산(오프 입점몰)"
		elseif Fdivcode="BB" then
			GetDivCodeName = "판매분정산(온 입점몰)"
		elseif Fdivcode="CC" then
			GetDivCodeName = "배송비정산(온 입점몰)"
		else
			GetDivCodeName = Fdivcode
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="MC" then
			GetDivCodeColor = "#3333FF"
		elseif Fdivcode="WS" then
			GetDivCodeColor = "#FF3333"
		else
			GetDivCodeColor = "#000000"
		end if
	end function

	public function GetStateName()
		if FStateCD="0" then
			GetStateName = "수정중"
		elseif FStateCD="1" then
			GetStateName = "업체확인중"
	    elseif FStateCD="3" or FStateCD="4" then
			GetStateName = "업체확인완료"
		elseif FStateCD="7" then
			GetStateName = "완료"
		end if
	end function

	public function GetIssueStateName()
		if FIssueStateCD="0" then
			GetIssueStateName = "발행신청"
		elseif FIssueStateCD="9" then
			GetIssueStateName = "발행완료"
		else
			GetIssueStateName = FIssueStateCD
		end if
	end function

	public function GetIpkumStateName()
		if FIpkumStateCD="0" then
			GetIpkumStateName = "입금전"
		elseif FIpkumStateCD="5" then
			GetIpkumStateName = "일부입금"
		elseif FIpkumStateCD="9" then
			GetIpkumStateName = "입금완료"
		else
			GetIpkumStateName = FIpkumStateCD
		end if
	end function

    public function GetPaperTypeName()
        SELECT CASE Fpapertype
            CASE "100" : GetPaperTypeName = "세금"
            CASE "101" : GetPaperTypeName = "면세"
            CASE "102" : GetPaperTypeName = "영세"
            CASE "200" : GetPaperTypeName = "수출"
            CASE "999" : GetPaperTypeName = "없음"
            CASE ELSE : GetPaperTypeName = Fpapertype
        ENd SELECT
    end function

    public function GetPaperTypeColor()
        if (Fpaperissuetype = "2") then
        	'// 역발행
        	GetPaperTypeColor = "red"
        else
	        SELECT CASE Fpapertype
	            CASE "100" : GetPaperTypeColor = "green"
	            CASE "101" : GetPaperTypeColor = "green"
	            CASE "102" : GetPaperTypeColor = "green"
	            CASE "200" : GetPaperTypeColor = "blue"
	            CASE "999" : GetPaperTypeColor = "gray"
	            CASE ELSE : GetPaperTypeColor = "black"
	        ENd SELECT
        end if
    end function

	public function GetStateColor()
		if FStateCD="0" then
			GetStateColor = "#000000"
		elseif FStateCD="1" then
			GetStateColor = "#448888"
		elseif FStateCD="3" or FStateCD="4" then
			GetStateColor = "#884488"
		elseif FStateCD="7" then
			GetStateColor = "#FF0000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEtcMeachulSubMasterItem
	public Fidx
	public Fmasteridx
	public Flinkidx
	public Fshopid
	public Fcode01
	public Fcode02
	public Fexecdate
	public Ftotalcount
	public Ftotalsellcash
	public Ftotalbuycash
	public Ftotalsuplycash
	public Ftotalorgsellcash

	public Fbaljudate

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEtcMeachulSubDetailItem
	public Fidx
	public Fmasteridx
	public Ftopmasteridx
	public Flinkbaljucode
	public Flinkmastercode
	public Flinkdetailidx
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fmakerid
	public Fitemno
	public Fsellcash
	public Fsuplycash
	public Fbuycash
	public Forgsellcash

	public function GetBarCode()
		GetBarCode = Fitemgubun + Format00(6,Fitemid) + Fitemoption
		if (FItemID >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,FItemId)) + CStr(Fitemoption)
    	end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEtcMeachulSumItem
	public Fuserdiv
	public FpcUserDiv
	public fsocname
	public Fsocname_kor
	public Fyyyymm
	public Fshopid
	public Fshopdiv
	public FCNT
	public Ftotalsum
	public Ftotalsellcash
	public Ftotalsuplycash
	public Ftotalbuycash
	public Ftotalorgsellcash
	public Fdivcode
	public Fbizsection_cd
	public Fbizsection_nm
	public Fselltype
	public Fselltypenm
	public Ftotmatchedipkumsum

    Public FdtlsellsumITS
    Public FdtlsuplysumITS
    Public FdtlbuysumITS

    public function gettotalsum_Tax()
		if Fshopdiv="3" and Fselltype="4010009" then
			'// 해외 기타는 세금없음(가맹비 등)
			gettotalsum_Tax = 0
		elseif (Fshopdiv="7") or (Fshopdiv="9") then
            gettotalsum_Tax = 0
        else
            gettotalsum_Tax = CLNG(Ftotalsum*1/11)
        end if
    end function

	public function GetDivCodeName()
		if Fdivcode="MC" then
			GetDivCodeName = "출고분정산"
		elseif Fdivcode="WS" then
			GetDivCodeName = "판매분정산(가맹점)"
		elseif Fdivcode="GC" then
			GetDivCodeName = "가맹비"
		elseif Fdivcode="ET" then
			GetDivCodeName = "기타매출"
		elseif Fdivcode="TC" then
			GetDivCodeName = "B2C매출" '// 사용않함
		elseif Fdivcode="AA" then
			GetDivCodeName = "판매분정산(오프 입점몰)"
		elseif Fdivcode="BB" then
			GetDivCodeName = "판매분정산(온 입점몰)"
		elseif Fdivcode="CC" then
			GetDivCodeName = "배송비정산(온 입점몰)"
		else
			GetDivCodeName = Fdivcode
		end if
	end function

    public function getShopDivName()
        SELECT CASE Fshopdiv
            CASE "1" : getShopDivName = "내부"  ''직영
            CASE "3" : getShopDivName = "가맹"
            CASE "5" : getShopDivName = "도매"
            CASE "7" : getShopDivName = "수출"
            CASE "9" : getShopDivName = "영세"
            CASE "11" : getShopDivName = "띵소"
            CASE "13" : getShopDivName = "제휴"
            CASE "15" : getShopDivName = "용역"
            CASE ELSE : getShopDivName = Fshopdiv
        ENd SELECT
    end function

	public function GetStateColor()
		if FStateCD="0" then
			GetStateColor = "#000000"
		elseif FStateCD="1" then
			GetStateColor = "#448888"
		elseif FStateCD="3" or FStateCD="4" then
			GetStateColor = "#884488"
		elseif FStateCD="7" then
			GetStateColor = "#FF0000"
		end if
	end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CEtcMeachul
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FRectshopid
	public FRectStartDate
	public FRectEndDate
	public FRectShopDiv
	public FRectidx
	public FRectonlymifinish
	public FRectExclude3pl
	public FRectStateUpcheView
	public FRectdivcode
    public FRectStateCD
    public FRectOldData
	public FRectDateType

    public FRectBankInOutIdx

    public FRectBeforeIssueOnly		'// 증빙서류가 세금계산서(정발행) 인 경우중 발행신청되지 않은 내역만 표시
    public FRectOnlyDtlITS
    public FRectRemoveDupp
    public FRectSelltype
    public FRectSellBizCd

	public FRectSearchType
	public FRectSearchString
    public FRectRemoveDlvPay
    public FRectGroupByBrand
    public FRectMakerid

	public FRectCType
	public FRectExcTPL
	public FtplGubun
    public FRectInc3pl
	public frectipkumstate

    '//admin/meachul/managementSupport/etc_meachulSum.asp
	public sub getEtcMeachulSumList()
		dim i,sqlStr, addSql

        sqlStr = " select top " + CStr(FPageSize) + " "
        sqlStr = sqlStr & " a.yyyymm, a.shopid, a.shopdiv, a.divcode, a.selltype, a.bizsection_cd"
        sqlStr = sqlStr & " , m.bizsection_nm , c.pcomm_name as selltypenm "
        sqlStr = sqlStr & " , uc.socname, uc.socname_kor, isNull(p.userdiv,'') as puserdiv, uc.userdiv"
        sqlStr = sqlStr & " , count(*) as CNT"
        sqlStr = sqlStr & " , sum(a.totalsum)  as totalsum"
        sqlStr = sqlStr & " , sum(a.totalsellcash)  as totalsellcash"
		sqlStr = sqlStr & " , sum(a.totalsuplycash)  as totalsuplycash"
        sqlStr = sqlStr & " , sum(a.totalbuycash)	as totalbuycash"
        sqlStr = sqlStr & " , sum(a.totalorgsellcash)  as totalorgsellcash"
        sqlStr = sqlStr & " ,sum(isNULL(ip.totmatchedipkumsum,0)) as totmatchedipkumsum "

        IF (FRectOnlyDtlITS<>"") then
            sqlStr = sqlStr & " ,sum(isNULL(Ts.dtlsellsumITS,0)) as dtlsellsumITS"
            sqlStr = sqlStr & " ,sum(isNULL(Ts.dtlsuplysumITS,0)) as dtlsuplysumITS"
            sqlStr = sqlStr & " ,sum(isNULL(Ts.dtlbuysumITS,0)) as dtlbuysumITS"
        ENd IF

        sqlStr = sqlStr & " from [db_shop].[dbo].tbl_fran_meachuljungsan_master a "
        sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user s"
        sqlStr = sqlStr & " 	on s.userID = a.shopID "
        sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
        sqlStr = sqlStr & " 	on a.shopID=p.id "
        sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c uc"
        sqlStr = sqlStr & " 	on a.shopID = uc.userid"
        sqlStr = sqlStr & " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION m"
        sqlStr = sqlStr & "  	on m.bizsection_cd = a.bizsection_cd "
        sqlStr = sqlStr & " left join [db_partner].[dbo].tbl_partner_comm_code c"
        sqlStr = sqlStr & " 	on a.selltype = c.pcomm_cd and c.pcomm_group = 'sellacccd'"
        sqlStr = sqlStr & " left join db_jungsan.dbo.tbl_ipkum_match_master ip"
        sqlStr = sqlStr & " 	on ip.jungsanidx = a.idx "

        IF (FRectOnlyDtlITS<>"") then
            sqlStr = sqlStr & " left join ("
            sqlStr = sqlStr & " 	select d.topmasteridx"
            sqlStr = sqlStr & " 	, sum(d.sellcash*d.itemno) as dtlsellsumITS"
            sqlStr = sqlStr & " 	, sum(d.suplycash*d.itemno)  as dtlsuplysumITS"
            sqlStr = sqlStr & " 	, sum(d.buycash*d.itemno) as dtlbuysumITS"
            sqlStr = sqlStr & " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_master M"
            sqlStr = sqlStr & " 	join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail D"
            sqlStr = sqlStr & " 		on M.idx=D.topmasteridx"
            sqlStr = sqlStr & " 	Join db_partner.dbo.tbl_partner p"
            sqlStr = sqlStr & " 		on d.makerid=p.id"
            sqlStr = sqlStr & " 		and p.groupid='G02799'"
            sqlStr = sqlStr & " 	where 1=1"

            if (FRectDateType <> "") then
    			if (FRectDateType = "yyyymm") then
    				sqlStr = sqlStr & "  	and m.yyyymm>='" + Left(CStr(FRectStartDate), 7) + "'"
    				sqlStr = sqlStr & "  	and m.yyyymm<='" + Left(CStr(FRectendDate), 7) + "'"
    			elseif (FRectDateType = "taxdt") then
    				sqlStr = sqlStr & "  	and m.taxdate>='" + CStr(FRectStartDate) + "'"
    				sqlStr = sqlStr & "  	and m.taxdate<='" + CStr(FRectendDate) + "'"
    			end if
    		end if

            sqlStr = sqlStr & " 	group by d.topmasteridx"
            sqlStr = sqlStr & " ) Ts"
            sqlStr = sqlStr & "		on a.idx=Ts.topmasteridx"
        END IF

        sqlStr = sqlStr & "  where 1=1"
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        sqlStr = sqlStr & " and isNULL(p.tplcompanyid,'')=''"
	    end if

        IF (FRectOnlyDtlITS<>"") then
            sqlStr = sqlStr & " and (a.shopid in ('etcithinkso') or (dtlsellsumITS<>0))"
        ENd IF

        if (FRectRemoveDupp<>"") then
            sqlStr = sqlStr & " and a.divcode not in ('BB') and a.shopid not in ('streetshop012')"
        end if

		if (FRectShopDiv <> "") then
			sqlStr = sqlStr & " and a.shopdiv = '" + CStr(FRectShopDiv) + "' "
		end if

		if (FRectshopid <> "") then
			sqlStr = sqlStr & " and a.shopid = '" + CStr(FRectshopid) + "' "
		end if

		if (FRectdivcode <> "") then
			sqlStr = sqlStr & " and a.divcode = '" + CStr(FRectdivcode) + "' "
		end if

		if (FRectStateCd <> "") then
			sqlStr = sqlStr & " and a.statecd = '" + CStr(FRectStateCd) + "' "
		end if

        if (FRectSelltype<>"") then
            sqlStr = sqlStr & " and a.selltype="&FRectSelltype&""
        end if

        if (FRectSellBizCd<>"") then
            sqlStr = sqlStr & " and a.bizsection_cd='"&FRectSellBizCd&"'"
        end if

		if (FRectDateType <> "") then
			if (FRectDateType = "yyyymm") then
				sqlStr = sqlStr & "  and a.yyyymm>='" + Left(CStr(FRectStartDate), 7) + "'"
				sqlStr = sqlStr & "  and a.yyyymm<='" + Left(CStr(FRectendDate), 7) + "'"
			elseif (FRectDateType = "taxdt") then
				sqlStr = sqlStr & "  and a.taxdate>='" + CStr(FRectStartDate) + "'"
				sqlStr = sqlStr & "  and a.taxdate<='" + CStr(FRectendDate) + "'"
			end if
		end if

        sqlStr = sqlStr & " group by a.shopid"
        sqlStr = sqlStr & " ,a.yyyymm"
        sqlStr = sqlStr & " ,a.shopdiv"
        sqlStr = sqlStr & " ,a.divcode"
        sqlStr = sqlStr & " ,a.bizsection_cd"
        sqlStr = sqlStr & " ,m.bizsection_nm "
        sqlStr = sqlStr & " ,a.selltype"
        sqlStr = sqlStr & " ,c.pcomm_name, uc.socname, uc.socname_kor, isNull(p.userdiv,''), uc.userdiv"
        sqlStr = sqlStr & " order by a.yyyymm desc, a.shopid"

		'response.write sqlStr & "<Br>"
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
				set FItemList(i) = new CEtcMeachulSumItem

				FItemList(i).Fuserdiv  = rsget("userdiv")
				FItemList(i).FpcUserDiv  = rsget("puserdiv") &"_" & FItemList(i).Fuserdiv
				FItemList(i).Fsocname  				= db2html(rsget("socname"))
				FItemList(i).Fsocname_kor  			= db2html(rsget("socname_kor"))
				FItemList(i).Fyyyymm				= rsget("yyyymm")
				FItemList(i).Fshopid				= rsget("shopid")
				FItemList(i).Fshopdiv				= rsget("shopdiv")
				FItemList(i).FCNT					= rsget("CNT")
				FItemList(i).Ftotalsum				= rsget("totalsum")
				FItemList(i).Ftotalsellcash			= rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash		= rsget("totalsuplycash")
				FItemList(i).Ftotalbuycash			= rsget("totalbuycash")
				FItemList(i).Ftotalorgsellcash		= rsget("totalorgsellcash")
				FItemList(i).Fdivcode				= rsget("divcode")
				FItemList(i).Fbizsection_cd			= rsget("bizsection_cd")
				FItemList(i).Fbizsection_nm			= rsget("bizsection_nm")
				FItemList(i).Fselltype   		    = rsget("selltype")
				FItemList(i).Fselltypenm			= rsget("selltypenm")
				FItemList(i).Ftotmatchedipkumsum	= rsget("totmatchedipkumsum")

                IF (FRectOnlyDtlITS<>"") then
                    FItemList(i).FdtlsellsumITS     = rsget("dtlsellsumITS")
                    FItemList(i).FdtlsuplysumITS    = rsget("dtlsuplysumITS")
                    FItemList(i).FdtlbuysumITS      = rsget("dtlbuysumITS")

                    if (FItemList(i).Fshopid="etcithinkso") then
                        FItemList(i).FdtlsellsumITS     = FItemList(i).Ftotalsellcash
                        FItemList(i).FdtlsuplysumITS    = FItemList(i).Ftotalsuplycash
                        FItemList(i).FdtlbuysumITS      = FItemList(i).Ftotalbuycash
                    end if
                END IF

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close

    end sub

	'//admin/offshop/offshop_meachul.asp
	public sub getEtcMeachulList()
		dim i,sqlStr, addSql
		dim tmpStartDate, tmpEndDate

		'// ===================================================================
		addSql = " from " + vbcrlf
		addSql = addSql + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master a " + vbcrlf
		addSql = addSql + " 	left join db_shop.dbo.tbl_shop_user s " + vbcrlf
		addSql = addSql + " 	on " + vbcrlf
		addSql = addSql + " 		s.userID = a.shopID " + vbcrlf
		addSql = addSql + " 	left join db_partner.dbo.tbl_partner p " + vbcrlf
		addSql = addSql + " 	on " + vbcrlf
		addSql = addSql + " 		a.shopID = p.id " + vbcrlf									'// ON입점몰 나오도록 변경
		addSql = addSql + " 	left join db_partner.dbo.tbl_TMS_BA_BIZSECTION m " + vbcrlf
		addSql = addSql + " 	on " + vbcrlf
		addSql = addSql + " 		m.bizsection_cd = a.bizsection_cd " + vbcrlf
		addSql = addSql + " 	left join [db_partner].[dbo].tbl_partner_comm_code c " + vbcrlf
		addSql = addSql + " 	on " + vbcrlf
		addSql = addSql + " 		a.selltype = c.pcomm_cd and c.pcomm_group = 'sellacccd' " + vbcrlf
		addSql = addSql + " where a.idx<>0"

		if frectipkumstate<>"" then
			if frectipkumstate="0" then
				addSql = addSql + " and isnull(a.ipkumstatecd,'')=''"
			else
				addSql = addSql + " and isnull(a.ipkumstatecd,'')='"& frectipkumstate &"'"
			end if
		end if
		if FRectStateUpcheView<>"" then
			addSql = addSql + " and a.statecd>0"
		end if

		if FRectDateType = "regdate" then
			if FRectStartDate<>"" then
				tmpStartDate = FRectStartDate + "-01"
				addSql = addSql + " and a.regdate >='" + CStr(tmpStartDate) + "'"
			end if
			if FRectendDate<>"" then
				tmpEndDate = DateSerial(Left(FRectendDate, 4), Right(FRectendDate, 2), 1)
				tmpEndDate = DateAdd("m", 1, tmpEndDate)
				tmpEndDate = Left(tmpEndDate, 10)
				addSql = addSql + " and a.regdate < '" + CStr(tmpEndDate) + "'"
			end if
		elseif FRectDateType = "issuedate" then
			if FRectStartDate<>"" then
				tmpStartDate = FRectStartDate + "-01"
				addSql = addSql + " and a.taxdate >='" + CStr(tmpStartDate) + "'"
			end if
			if FRectendDate<>"" then
				tmpEndDate = DateSerial(Left(FRectendDate, 4), Right(FRectendDate, 2), 1)
				tmpEndDate = DateAdd("m", 1, tmpEndDate)
				tmpEndDate = Left(tmpEndDate, 10)
				addSql = addSql + " and a.taxdate < '" + CStr(tmpEndDate) + "'"
			end if
		else
			if FRectStartDate<>"" then
				addSql = addSql + " and IsNULL(a.yyyymm,'"&FRectStartDate&"')>='" + CStr(FRectStartDate) + "'"
			end if
			if FRectendDate<>"" then
				addSql = addSql + " and IsNULL(a.yyyymm,'"&FRectendDate&"')<='" + CStr(FRectendDate) + "'"
			end if
		end if

        if FRectStateCD<>"" then
            addSql = addSql + " and a.statecd=" & FRectStateCD
        end if

		if FRectshopid<>"" then
			addSql = addSql + " and a.shopid='" + FRectshopid + "'"
		end if

		if FRectdivcode<>"" then
			addSql = addSql + " and a.divcode='" + FRectdivcode + "'"
		end if

        if (FRectShopDiv<>"") then
            addSql = addSql + " and a.shopdiv='" + CStr(FRectShopDiv) + "'"
        end if

        if (FRectSelltype<>"") then
            addSql = addSql & " and a.selltype="&FRectSelltype&""
        end if

        if (FRectSellBizCd<>"") then
            addSql = addSql & " and a.bizsection_cd='"&FRectSellBizCd&"'"
        end if

        if (FRectBeforeIssueOnly<>"") then
            addSql = addSql + " and a.papertype in ('100', '101', '102') "
            addSql = addSql + " and a.paperissuetype = '1' "
            addSql = addSql + " and a.issuestatecd is null "
        end if

        if (FRectBankInOutIdx <> "") then
            addSql = addSql + " and a.idx in ( "
			addSql = addSql + " 	select "
			addSql = addSql + " 		m.jungsanidx "
			addSql = addSql + " 	from "
			addSql = addSql + " 		db_jungsan.dbo.tbl_ipkum_match_master m "
			addSql = addSql + " 		join db_jungsan.dbo.tbl_ipkum_match_detail d "
			addSql = addSql + " 		on "
			addSql = addSql + " 			m.idx = d.masteridx "
			addSql = addSql + " 	where "
			addSql = addSql + " 		d.ipkummethod = 'BNK' and d.ipkumidx = " + CStr(FRectBankInOutIdx) + " and d.useyn = 'Y' "
			addSql = addSql + " ) "
		end if

		if (FRectSearchType <> "") and (FRectSearchString <> "") then
			if (FRectSearchType = "groupcode") then
				addSql = addSql + " and p.groupid = '" + CStr(FRectSearchString) + "' "
			elseif (FRectSearchType = "taxidx") then
				addSql = addSql + " and a.taxlinkidx = " + CStr(FRectSearchString) + " "
			end if
		end if

        if (FRectExcTPL = "Y") then
            addSql = addSql & " and IsNull(p.tplcompanyid, '') = '' "
        end if

		if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				addSql = addSql + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				addSql = addSql + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FtplGubun) + "' "
			end if
		end if


		'// ===================================================================
		sqlStr = " select count(a.idx) as cnt "
		sqlStr = sqlStr + addSql
'rw sqlStr
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		'// ===================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.* "
		sqlStr = sqlStr + " , m.bizsection_nm " + vbcrlf
		sqlStr = sqlStr + " , c.pcomm_name as selltypenm " + vbcrlf
		sqlStr = sqlStr + " , (SELECT TOP 1 replace(p.company_no,'-','') FROM db_shop.dbo.tbl_shop_user s Join db_partner.dbo.tbl_partner p on s.userid=p.id WHERE s.userID = a.shopID ) bizNo " + vbcrlf
		sqlStr = sqlStr + " , (SELECT TOP 1 c.userdiv FROM db_partner.dbo.tbl_partner p left join [db_user].[dbo].tbl_user_c c on c.userid=p.id WHERE p.id = a.shopID ) brandDiv " + vbcrlf
		sqlStr = sqlStr + " , ( " + vbcrlf
		sqlStr = sqlStr + " 	SELECT TOP 1 m.totmatchedipkumsum " + vbcrlf
		sqlStr = sqlStr + " 	FROM " + vbcrlf
		sqlStr = sqlStr + " 		db_jungsan.dbo.tbl_ipkum_match_master m " + vbcrlf
		sqlStr = sqlStr + " 	WHERE " + vbcrlf
		sqlStr = sqlStr + " 		m.jungsanidx = a.idx " + vbcrlf
		sqlStr = sqlStr + " ) as totmatchedipkumsum " + vbcrlf
		sqlStr = sqlStr + " , ( " + vbcrlf
		sqlStr = sqlStr + " 	SELECT TOP 1 m.inoutidx " + vbcrlf
		sqlStr = sqlStr + " 	FROM " + vbcrlf
		sqlStr = sqlStr + " 		[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT m " + vbcrlf
		sqlStr = sqlStr + " 	WHERE " + vbcrlf
		sqlStr = sqlStr + " 		m.tx_amt = a.totalsum and m.INOUT_GUBUN = 2 and IsNull(m.matchstate, 'N') <> 'Y' " + vbcrlf
		sqlStr = sqlStr + " ) as maymatchedipkumsum " + vbcrlf

		sqlStr = sqlStr + addSql

        sqlStr = sqlStr + " order by a.idx desc"
''rw sqlStr
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
				set FItemList(i) = new CEtcMeachulMasterItem

				FItemList(i).Fidx         = rsget("idx")
				FItemList(i).Fshopid      = rsget("shopid")
				FItemList(i).Ftitle       = db2html(rsget("title"))
				FItemList(i).Ftotalsum    = rsget("totalsum")
				FItemList(i).Ftotalsellcash    = rsget("totalsellcash")
				FItemList(i).Ftotalbuycash    = rsget("totalbuycash")
				FItemList(i).Ftotalsuplycash   = rsget("totalsuplycash")
				FItemList(i).Fdivcode     = rsget("divcode")
				FItemList(i).Ftaxdate     = rsget("taxdate")
				FItemList(i).Ftaxregdate  = rsget("taxregdate")
				FItemList(i).Fregdate     = rsget("regdate")
				FItemList(i).Fipkumdate   = rsget("ipkumdate")
				FItemList(i).Fetcstr      = db2html(rsget("etcstr"))
				FItemList(i).FStateCD	  = rsget("statecd")
				FItemList(i).Freguserid      = rsget("reguserid")
				FItemList(i).Fregusername    = db2html(rsget("regusername"))
				FItemList(i).Ffinishuserid   = rsget("finishuserid")
				FItemList(i).Ffinishusername = db2html(rsget("finishusername"))
				FItemList(i).FtaxNo	  = rsget("neoTaxNo")
				FItemList(i).FbizNo	  = rsget("bizNo")

                FItemList(i).Fyyyymm    = rsget("yyyymm")
                FItemList(i).FdiffKey   = rsget("diffKey")
                FItemList(i).Fshopdiv   = rsget("shopdiv")

                FItemList(i).FbrandDiv   	= rsget("brandDiv")

                FItemList(i).Fworkidx   	= rsget("workidx")
                FItemList(i).Finvoiceidx   	= rsget("invoiceidx")

                FItemList(i).Ftotmatchedipkumsum   	= rsget("totmatchedipkumsum")
                FItemList(i).Fmaymatchedipkumsum   	= rsget("maymatchedipkumsum")

                FItemList(i).Fbizsection_cd   	= rsget("bizsection_cd")
                FItemList(i).Fbizsection_nm   	= db2html(rsget("bizsection_nm"))
                FItemList(i).Fpapertype   		= rsget("papertype")
                FItemList(i).Fpaperissuetype	= rsget("paperissuetype")
                FItemList(i).Fetcpaperidx   	= rsget("etcpaperidx")

                FItemList(i).Fselltype   		= rsget("selltype")
                FItemList(i).Fselltypenm   		= db2html(rsget("selltypenm"))

                FItemList(i).Fissuestatecd   	= rsget("issuestatecd")
                FItemList(i).Fipkumstatecd   	= rsget("ipkumstatecd")

                FItemList(i).Feserotaxkey   	= rsget("eserotaxkey")
                FItemList(i).Ftaxlinkidx   	= rsget("taxlinkidx")


				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getChulgoJungsanTargetList()
		dim i,sqlStr

		sqlStr = " select count(m.id) as cnt from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " where m.executedt>='" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.executedt<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and socid='" + FRectshopid + "'"
		sqlStr = sqlStr + " and m.deldt is null"
        ''sqlStr = sqlStr + " and m.cwFlag<>1"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close
		''??

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, s.baljuname, m.code, m.socid,m.divcode,s.scheduledate, s.regdate as jumunregdate, m.executedt,"
		sqlStr = sqlStr + " m.totalsellcash,m.totalsuplycash,m.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(s.totalsellcash,0) as jumunrealsellcash, IsNULL(s.totalsuplycash,0) as jumunrealsuplycash,"
		sqlStr = sqlStr + " IsNULL(s.totalbuycash,0) as jumunrealbuycash,"
		sqlStr = sqlStr + " s.ipgodate, s.baljucode, s.idx as baljuidx, s.segumdate as baljusegumdate, f.masteridx as precheckmasteridx, f.linkidx as precheckidx, (case when m.executedt is null then '0' else '7' end) as orderstatecd, T.idx as workidx, T.baljudate "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_master s"
		sqlStr = sqlStr + " on s.baljuid='" + FRectshopid + "' and s.deldt is null and m.code=s.alinkcode  "
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select sf.masteridx, sf.linkidx from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master fm"
		sqlStr = sqlStr + " 	    Join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sf"
		sqlStr = sqlStr + " 	    on fm.idx=sf.masteridx"
		sqlStr = sqlStr + " 	where fm.divcode='MC'"
		sqlStr = sqlStr + " 	and fm.shopid='" + FRectshopid + "'"
		sqlStr = sqlStr + " ) F"
		sqlStr = sqlStr + " on m.id=f.linkidx"

		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select "
		sqlStr = sqlStr + " 		cm.idx, T.baljudate, T.baljucode "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		db_storage.dbo.tbl_cartoonbox_master cm "
		sqlStr = sqlStr + " 		join db_storage.dbo.tbl_cartoonbox_detail cd "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			cm.idx = cd.masteridx "
		sqlStr = sqlStr + " 		join ( "
		sqlStr = sqlStr + " 			select b.baljuid, convert(varchar(10),b.baljudate,21) as baljudate, b.baljucode, IsNull(od.packingstate, 0) as innerboxno "
		sqlStr = sqlStr + " 			from "
		sqlStr = sqlStr + " 				db_storage.dbo.tbl_shopbalju b "
		sqlStr = sqlStr + " 				join [db_storage].[dbo].tbl_ordersheet_master om "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					b.baljucode = om.baljucode "
		sqlStr = sqlStr + " 				join [db_storage].[dbo].tbl_ordersheet_detail od "
		sqlStr = sqlStr + " 				on "
		sqlStr = sqlStr + " 					om.idx = od.masteridx "
		sqlStr = sqlStr + " 			where "
		sqlStr = sqlStr + " 				1 = 1 "
		sqlStr = sqlStr + " 				and om.deldt is null "
		sqlStr = sqlStr + " 				and od.deldt is null "
		sqlStr = sqlStr + " 				and IsNull(od.packingstate, 0) <> 0 "
		sqlStr = sqlStr + " 				and om.ipgodate >= '" + FRectStartDate + "' "
		sqlStr = sqlStr + " 				and om.ipgodate <= '" + FRectEndDate + "' "
		sqlStr = sqlStr + " 				and b.baljudate >= '" + FRectStartDate + "' "
		sqlStr = sqlStr + " 				and b.baljudate <= '" + FRectEndDate + "' "
		sqlStr = sqlStr + " 				and b.baljuid = '" + FRectshopid + "' "
		sqlStr = sqlStr + " 			group by "
		sqlStr = sqlStr + " 				b.baljuid, convert(varchar(10),b.baljudate,21), b.baljucode, IsNull(od.packingstate, 0) "
		sqlStr = sqlStr + " 		) T "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			1 = 1 "
		sqlStr = sqlStr + " 			and cd.shopid = T.baljuid "
		sqlStr = sqlStr + " 			and convert(varchar(10),cd.baljudate,21) = T.baljudate "
		sqlStr = sqlStr + " 			and cd.innerboxno = T.innerboxno "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and cm.shopid = '" + FRectshopid + "' "
		sqlStr = sqlStr + " 	group by "
		sqlStr = sqlStr + " 		cm.idx, T.baljudate, T.baljucode "
		sqlStr = sqlStr + " ) T "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	s.baljucode = T.baljucode "

		sqlStr = sqlStr + " where m.executedt>='" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.executedt<'" + FRectEndDate + "'"
		sqlStr = sqlStr + " and socid='" + FRectshopid + "'"
		sqlStr = sqlStr + " and m.deldt is null"
        sqlStr = sqlStr + " and isNULL(s.cwFlag,0)<>1"        '' 1 출고위탁 제외..

		if FRectonlymifinish<>"" then
			sqlStr = sqlStr + " and f.linkidx is null"
		end if

		sqlStr = sqlStr + " order by m.id, m.executedt"
''response.write sqlStr
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
				set FItemList(i) = new CFranChulgojungsanTargetItem

				FItemList(i).Fid         	= rsget("id")
				FItemList(i).Fcode      	= rsget("code")
				FItemList(i).Fsocid       	= rsget("socid")
				FItemList(i).Fshopname    	= rsget("baljuname")
				FItemList(i).Fdivcode    	= rsget("divcode")
				FItemList(i).Fexecutedt     = rsget("executedt")
				FItemList(i).Fscheduledate  = rsget("scheduledate")
				FItemList(i).FjumunRegDate  = rsget("jumunregdate")
				FItemList(i).Ftotalsellcash     = rsget("totalsellcash")*-1
				FItemList(i).Ftotalsuplycash  	= rsget("totalsuplycash")*-1
				FItemList(i).Ftotalbuycash     	= rsget("totalbuycash")*-1
				FItemList(i).Fjumunrealsellcash   	= rsget("jumunrealsellcash")
				FItemList(i).Fjumunrealsuplycash   	= rsget("jumunrealsuplycash")
				FItemList(i).Fjumunrealbuycash   	= rsget("jumunrealbuycash")
				FItemList(i).Fipgodate   			= rsget("ipgodate")
				FItemList(i).Fbaljucode			= rsget("baljucode")
				FItemList(i).Fbaljuidx			= rsget("baljuidx")
				FItemList(i).Fprecheckmasteridx		= rsget("precheckmasteridx")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).Fbaljusegumdate	= rsget("baljusegumdate")

				FItemList(i).Fworkidx			= rsget("workidx")
				FItemList(i).Fbaljudate			= rsget("baljudate")

				FItemList(i).Forderstatecd		= rsget("orderstatecd")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getChulgoJungsanTargetListNotReg()
		dim i,sqlStr

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, s.baljuname, m.code, m.socid,m.divcode,s.scheduledate, s.regdate as jumunregdate, m.executedt,"
		sqlStr = sqlStr + " m.totalsellcash,m.totalsuplycash,m.totalbuycash,"
		sqlStr = sqlStr + " IsNULL(s.totalsellcash,0) as jumunrealsellcash, IsNULL(s.totalsuplycash,0) as jumunrealsuplycash,"
		sqlStr = sqlStr + " IsNULL(s.totalbuycash,0) as jumunrealbuycash,"
		sqlStr = sqlStr + " s.ipgodate, s.baljucode, s.idx as baljuidx, s.segumdate as baljusegumdate, f.masteridx as precheckmasteridx, f.linkidx as precheckidx, (case when m.executedt is null then '0' else '7' end) as orderstatecd, b.bizsection_cd, b.bizsection_nm "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_master s"
		sqlStr = sqlStr + " on s.deldt is null and m.code=s.alinkcode  "
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + " 	select sf.masteridx, sf.linkidx from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master fm"
		sqlStr = sqlStr + " 	    Join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sf"
		sqlStr = sqlStr + " 	    on fm.idx=sf.masteridx"
		sqlStr = sqlStr + " 	where fm.divcode='MC'"
		sqlStr = sqlStr + " ) F"
		sqlStr = sqlStr + " on m.id=f.linkidx"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	p.id = m.socid "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	p.sellbizcd = b.BIZSECTION_CD "
		sqlStr = sqlStr + " where m.executedt>='" + FRectStartDate + "'"
		sqlStr = sqlStr + " and m.executedt<'" + FRectEndDate + "'"

		if (FRectshopid <> "") then
			sqlStr = sqlStr + " and socid='" + FRectshopid + "'"
		end if

		sqlStr = sqlStr + " and socid in ( "
		sqlStr = sqlStr + " 	select c.userid "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		[db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " 		left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			c.userid=p.id "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and IsNull(p.selltype, '') not in ('20166', '20032', '20046') "			'// B2C 제외
		if (FRectCType = "M") then
			sqlStr = sqlStr + " 		and IsNull(p.etcjungsantype, 0) in (2, 3) "
			sqlStr = sqlStr + " 		and p.userdiv in ('503','501') "
			sqlStr = sqlStr + " 		and c.userdiv in ('21') "
		elseif (FRectCType = "M_ETC") then
			sqlStr = sqlStr + " 		and IsNull(p.etcjungsantype, 0) in (2, 3) "
			sqlStr = sqlStr + " 		and p.userdiv in ('900') "
			sqlStr = sqlStr + " 		and c.userdiv in ('21') "
		end if
		if FRectExclude3pl="on" then
			sqlStr = sqlStr + " 		and p.userdiv not in ('903') "
			sqlStr = sqlStr + " 		and isNull(p.tplcompanyid,'')='' "
		end if
		sqlStr = sqlStr + " ) "

		sqlStr = sqlStr + " and m.deldt is null"
        sqlStr = sqlStr + " and isNULL(s.cwFlag,0)<>1"        '' 1 출고위탁 제외..

		if FRectonlymifinish<>"" then
			sqlStr = sqlStr + " and f.linkidx is null"
		end if

		sqlStr = sqlStr + " order by m.id, m.executedt"

		'response.write sqlStr
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
				set FItemList(i) = new CFranChulgojungsanTargetItem

				FItemList(i).Fid         	= rsget("id")
				FItemList(i).Fcode      	= rsget("code")
				FItemList(i).Fsocid       	= rsget("socid")
				FItemList(i).Fshopname    	= rsget("baljuname")
				FItemList(i).Fdivcode    	= rsget("divcode")
				FItemList(i).Fexecutedt     = rsget("executedt")
				FItemList(i).Fscheduledate  = rsget("scheduledate")
				FItemList(i).FjumunRegDate  = rsget("jumunregdate")
				FItemList(i).Ftotalsellcash     = rsget("totalsellcash")*-1
				FItemList(i).Ftotalsuplycash  	= rsget("totalsuplycash")*-1
				FItemList(i).Ftotalbuycash     	= rsget("totalbuycash")*-1
				FItemList(i).Fjumunrealsellcash   	= rsget("jumunrealsellcash")
				FItemList(i).Fjumunrealsuplycash   	= rsget("jumunrealsuplycash")
				FItemList(i).Fjumunrealbuycash   	= rsget("jumunrealbuycash")
				FItemList(i).Fipgodate   			= rsget("ipgodate")
				FItemList(i).Fbaljucode			= rsget("baljucode")
				FItemList(i).Fbaljuidx			= rsget("baljuidx")
				FItemList(i).Fprecheckmasteridx		= rsget("precheckmasteridx")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).Fbaljusegumdate	= rsget("baljusegumdate")

				'FItemList(i).Fworkidx			= rsget("workidx")
				'FItemList(i).Fbaljudate			= rsget("baljudate")

				FItemList(i).Forderstatecd		= rsget("orderstatecd")

				FItemList(i).Fbizsection_cd		= rsget("bizsection_cd")
				FItemList(i).Fbizsection_nm		= rsget("bizsection_nm")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

    public sub getWitakSellJungsanTargetList()
		dim i,sqlStr

        sqlStr = " select T.* ,sm.idx as precheckidx"
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + " select  m.idx, m.yyyymm, d.shopid, s.shopname, m.makerid, "
        sqlStr = sqlStr + " sum(itemno) as totitemcnt, sum(sellprice*itemno) as totorgsum, sum(realsellprice*itemno) as totsum, sum(suplyprice*itemno) as realjungsansum "
        sqlStr = sqlStr + " from  "
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master m, "
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_detail d, "
        sqlStr = sqlStr + " db_shop.dbo.tbl_shop_user s "
        sqlStr = sqlStr + " where m.idx=d.masteridx and d.shopid = s.userid "
        sqlStr = sqlStr + " and m.yyyymm>='" + Left(FRectStartDate,7) + "'"
        sqlStr = sqlStr + " and m.yyyymm<'" + Left(FRectEndDate,7) + "'"
        sqlStr = sqlStr + " and d.gubuncd in ('B012','B013') "                          '''출고위탁(B013 추가)
        sqlStr = sqlStr + " and d.shopid='" + FRectshopid + "'"
        sqlStr = sqlStr + " group by m.idx, m.yyyymm, d.shopid, s.shopname, m.makerid "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm "
        sqlStr = sqlStr + " 	on T.shopid=sm.shopid and T.makerid=sm.code02 and T.idx=sm.linkidx "
        if FRectonlymifinish<>"" then
            sqlStr = sqlStr + " where sm.idx is null "
        end if
        sqlStr = sqlStr + " order by T.yyyymm desc, T.idx "

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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fyyyymm         = rsget("yyyymm")
				FItemList(i).Fshopid         = rsget("shopid")
				FItemList(i).Fshopname       = rsget("shopname")
				FItemList(i).Fjungsanid      = rsget("makerid")
				FItemList(i).Ftotitemcnt     = rsget("totitemcnt")
				FItemList(i).Ftotorgsum      = rsget("totorgsum")
				FItemList(i).Ftotsum         = rsget("totsum")
				FItemList(i).Frealjungsansum = rsget("realjungsansum")
				FItemList(i).Fprecheckidx	= rsget("precheckidx")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

    public sub getWitakSellJungsanTargetListNotReg()
		dim i,sqlStr

		sqlStr = " exec [db_shop].[dbo].[usp_Ten_EtcMeachul_GetWitakSellJungsanTargetListNotReg] '" + CStr(FRectshopid) + "', '" + CStr(Left(FRectStartDate,7)) + "', '" + CStr(Left(FRectEndDate,7)) + "', '" + CStr(FRectonlymifinish) + "' "
		''response.write sqlStr &"<Br>"

		rsget.CursorLocation = 3
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 3, 1

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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fyyyymm         = rsget("yyyymm")
				FItemList(i).Fshopid         = rsget("shopid")
				FItemList(i).Fshopname       = rsget("shopname")
				FItemList(i).Fjungsanid      = rsget("makerid")
				FItemList(i).Ftotitemcnt     = rsget("totitemcnt")
				FItemList(i).Ftotorgsum      = rsget("totorgsum")
				FItemList(i).Ftotsum         = rsget("totsum")
				FItemList(i).Frealjungsansum = rsget("realjungsansum")
				FItemList(i).Fprecheckidx	= rsget("precheckidx")

				FItemList(i).Fbizsection_cd	= rsget("bizsection_cd")
				FItemList(i).Fbizsection_nm	= rsget("bizsection_nm")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 판매분정산(오프 입점몰)
    public sub getOfflineIpjumshopMaechulList()
		dim i,sqlStr

        sqlStr = " select"
        sqlStr = sqlStr & " t.yyyymmdd ,t.shopid ,t.totitemcnt , t.totorgsum, t.totsum ,t.realjungsansum ,t.buyprice"
        sqlStr = sqlStr & " ,sm.idx as precheckidx, s.shopname "
        sqlStr = sqlStr & " from ("
        sqlStr = sqlStr & "		select"
        sqlStr = sqlStr & "		convert(varchar(10),m.shopregdate,121) as yyyymmdd, m.shopid"
        sqlStr = sqlStr & "		,sum(d.itemno) as totitemcnt"
		sqlStr = sqlStr & "		,isnull(sum((d.sellprice+isnull(d.addtaxcharge,0))*d.itemno),0) as totorgsum"
        sqlStr = sqlStr & "		,isnull(sum((d.realsellprice+isnull(d.addtaxcharge,0))*d.itemno),0) as totsum"
        sqlStr = sqlStr & "		,isnull(sum(d.suplyprice*d.itemno),0) as realjungsansum"
        sqlStr = sqlStr & "		,isnull(sum(d.shopbuyprice*d.itemno),0) as buyprice"
		sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m "
	    sqlStr = sqlStr & "		join [db_shop].[dbo].tbl_shopjumun_detail d "
		sqlStr = sqlStr & "			on m.idx=d.masteridx"
        sqlStr = sqlStr & "			and m.cancelyn='N' and d.cancelyn='N'"
		sqlStr = sqlStr & "		left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd "
		sqlStr = sqlStr & "		on d.idx = sd.linkdetailidx "
        sqlStr = sqlStr & "		where 1=1 "

        if FRectStartDate <> "" and FRectEndDate <> "" then
        	sqlStr = sqlStr & " 	and m.shopregdate>='" & Left(FRectStartDate,10) & "'"
        	sqlStr = sqlStr & " 	and m.shopregdate<'" & Left(FRectEndDate,10) & "'"
        end if

        if FRectshopid <> "" then
        	sqlStr = sqlStr & " 	and m.shopid='" & FRectshopid & "'"
        end if

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " 	and sd.linkdetailidx is null "
        end if

        sqlStr = sqlStr & "		group by convert(varchar(10),m.shopregdate,121) ,m.shopid"
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr & " 	on T.shopid=sm.shopid"
        sqlStr = sqlStr & " 	and T.yyyymmdd=sm.code01"
        sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_user s "
        sqlStr = sqlStr & " 	on T.shopid = s.userid "
        sqlStr = sqlStr & "	where 1=1 "

        '' if FRectonlymifinish<>"" then
        ''     sqlStr = sqlStr & " and sm.idx is null"
        '' end if

        sqlStr = sqlStr + " order by t.shopid asc ,T.yyyymmdd desc, T.totsum desc"

		'response.write sqlStr &"<Br>"
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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         	= rsget("yyyymmdd")
				FItemList(i).Fshopid         	= rsget("shopid")
				FItemList(i).Fshopname       	= rsget("shopname")
				FItemList(i).Ftotitemcnt     	= rsget("totitemcnt")
				FItemList(i).Ftotorgsum      	= rsget("totorgsum")
				FItemList(i).Ftotsum         	= rsget("totsum")
				FItemList(i).Frealjungsansum 	= rsget("realjungsansum")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).fbuyprice         	= rsget("buyprice")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 판매분정산(오프 입점몰)
    public sub getOfflineIpjumshopMaechulListNotReg()
		dim i,sqlStr

		sqlStr = " exec [db_shop].[dbo].[usp_Ten_EtcMeachul_GetOfflineIpjumshopMaechulListNotReg] '" + CStr(FRectshopid) + "', '" + CStr(Left(FRectStartDate,10)) + "', '" + CStr(Left(FRectEndDate,10)) + "', '" + CStr(FRectonlymifinish) + "' "
		'response.write sqlStr &"<Br>"

		rsget.CursorLocation = 3
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 3, 1

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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         	= rsget("yyyymmdd")
				FItemList(i).Fshopid         	= rsget("shopid")
				FItemList(i).Fshopname       	= rsget("shopname")
				FItemList(i).Ftotitemcnt     	= rsget("totitemcnt")
				FItemList(i).Ftotorgsum      	= rsget("totorgsum")
				FItemList(i).Ftotsum         	= rsget("totsum")
				FItemList(i).Frealjungsansum 	= rsget("realjungsansum")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).fbuyprice         	= rsget("buyprice")

				FItemList(i).Fbizsection_cd		= rsget("bizsection_cd")
				FItemList(i).Fbizsection_nm		= rsget("bizsection_nm")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 판매분정산(온 입점몰)
    public sub getOnlineIpjumshopMaechulList()
		dim i,sqlStr

        sqlStr = " select"
        sqlStr = sqlStr & " t.yyyymmdd ,t.shopid ,t.makerid, t.totitemcnt , t.totorgsum, t.totsum ,t.totchulgosum ,t.buyprice"
        sqlStr = sqlStr & " ,sm.idx as precheckidx, s.socname as shopname "
        sqlStr = sqlStr & " , t.totdeliveritemcnt , t.totdeliverorgsum , t.totdeliversum , t.buydeliverprice "
        sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) as yyyymmdd "
		sqlStr = sqlStr & "			, m.sitename as shopid "
		if (FRectGroupByBrand="on") then
		    sqlStr = sqlStr & "			, d.makerid"
		else
		    sqlStr = sqlStr & "			, '' as makerid"
		end if
		sqlStr = sqlStr & "			, sum(case when d.itemid <> 0 then d.itemno else 0 end) as totitemcnt "
		sqlStr = sqlStr & "			, sum(case when d.itemid = 0 then d.itemno else 0 end) as totdeliveritemcnt "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliverorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliversum "
		sqlStr = sqlStr & "			, 0 as totchulgosum "
		sqlStr = sqlStr & "			, 0 as totdeliverchulgosum "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as buyprice "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as buydeliverprice "
		sqlStr = sqlStr & "		from "
		sqlStr = sqlStr & "			db_order.dbo.tbl_order_master m "
		sqlStr = sqlStr & "			join db_order.dbo.tbl_order_detail d "
		sqlStr = sqlStr & "			on "
		sqlStr = sqlStr & "				m.orderserial = d.orderserial "
		sqlStr = sqlStr & "		where "
		sqlStr = sqlStr & "			1 = 1 "
		if (FRectMakerid<>"") then
		    sqlStr = sqlStr & "			and d.makerid='"&FRectMakerid&"'"
	    end if

		if (FRectRemoveDlvPay<>"") then
    		sqlStr = sqlStr & "			and d.itemid <> 0 "		'// 배송비 제외
        end if

		if FRectStartDate <> "" and FRectEndDate <> "" then
			sqlStr = sqlStr & "			and d.beasongdate >= '" & Left(FRectStartDate,10) & "' "
			sqlStr = sqlStr & "			and d.beasongdate < '" & Left(FRectEndDate,10) & "' "
		end if

		sqlStr = sqlStr & "			and ((d.itemid=0 and d.beasongdate is Not NULL) or (d.currstate >= '7')) "

		if FRectshopid <> "" then
			sqlStr = sqlStr & "			and m.sitename = '" & FRectshopid & "' "
		end if

		sqlStr = sqlStr & "			and m.cancelyn='N' "
		sqlStr = sqlStr & "			and d.cancelyn<>'Y' "
		sqlStr = sqlStr & "			and d.currstate = '7' "
		sqlStr = sqlStr & "		group by "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) "
		sqlStr = sqlStr & "			, m.sitename "
		if (FRectGroupByBrand="on") then
		    sqlStr = sqlStr & "			, d.makerid"
		end if
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr & " 	on T.shopid=sm.shopid"
        sqlStr = sqlStr & " 	and T.yyyymmdd=sm.code01"
		sqlStr = sqlStr & " 	and sm.code02 <> 'beasongpay' " + vbcrlf
		if (FRectGroupByBrand="on") then
        	sqlStr = sqlStr & " 	and sm.code02 = T.makerid " + vbcrlf
		end if
        sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c s "
        sqlStr = sqlStr & " 	on T.shopid = s.userid "
        sqlStr = sqlStr & "	where 1=1 "

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " and sm.idx is null"
        end if

        sqlStr = sqlStr + " order by t.shopid asc ,T.yyyymmdd desc, T.totsum desc"
''rw sqlStr
		''response.write sqlStr &"<Br>"
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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         	= rsget("yyyymmdd")
				FItemList(i).Fshopid         	= rsget("shopid")
				FItemList(i).Fshopname       	= rsget("shopname")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Ftotitemcnt     	= rsget("totitemcnt")
				FItemList(i).Ftotorgsum      	= rsget("totorgsum")
				FItemList(i).Ftotsum         	= rsget("totsum")
				''FItemList(i).Frealjungsansum 	= rsget("totchulgosum")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).fbuyprice         	= rsget("buyprice")

                FItemList(i).Ftotdeliveritemcnt	= rsget("totdeliveritemcnt")
				FItemList(i).Ftotdeliverorgsum	= rsget("totdeliverorgsum")
				FItemList(i).Ftotdeliversum		= rsget("totdeliversum")
				FItemList(i).Fbuydeliverprice	= rsget("buydeliverprice")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 판매분정산(온 입점몰)
    public sub getOnlineIpjumshopMaechulListNotReg()
		dim i,sqlStr

        sqlStr = " select"
        sqlStr = sqlStr & " t.yyyymmdd ,t.shopid ,t.makerid, t.totitemcnt , t.totorgsum, t.totsum ,t.totchulgosum ,t.buyprice"
        sqlStr = sqlStr & " ,sm.idx as precheckidx, s.socname as shopname "
        sqlStr = sqlStr & " , t.totdeliveritemcnt , t.totdeliverorgsum , t.totdeliversum , t.buydeliverprice, b.bizsection_cd, b.bizsection_nm "
        sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) as yyyymmdd "
		sqlStr = sqlStr & "			, m.sitename as shopid "
		if (FRectGroupByBrand="on") then
		    sqlStr = sqlStr & "			, d.makerid"
		else
		    sqlStr = sqlStr & "			, '' as makerid"
		end if
		sqlStr = sqlStr & "			, sum(case when d.itemid <> 0 then d.itemno else 0 end) as totitemcnt "
		sqlStr = sqlStr & "			, sum(case when d.itemid = 0 then d.itemno else 0 end) as totdeliveritemcnt "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliverorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliversum "
		sqlStr = sqlStr & "			, 0 as totchulgosum "
		sqlStr = sqlStr & "			, 0 as totdeliverchulgosum "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as buyprice "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as buydeliverprice "
		sqlStr = sqlStr & "		from "
		sqlStr = sqlStr & "			db_order.dbo.tbl_order_master m "
		sqlStr = sqlStr & "			join db_order.dbo.tbl_order_detail d "
		sqlStr = sqlStr & "			on "
		sqlStr = sqlStr & "				m.orderserial = d.orderserial "
		sqlStr = sqlStr & "		where "
		sqlStr = sqlStr & "			1 = 1 "
		if (FRectMakerid<>"") then
		    sqlStr = sqlStr & "			and d.makerid='"&FRectMakerid&"'"
	    end if

		if (FRectRemoveDlvPay<>"") then
    		sqlStr = sqlStr & "			and d.itemid <> 0 "		'// 배송비 제외
        end if

		if FRectStartDate <> "" and FRectEndDate <> "" then
			sqlStr = sqlStr & "			and d.beasongdate >= '" & Left(FRectStartDate,10) & "' "
			sqlStr = sqlStr & "			and d.beasongdate < '" & Left(FRectEndDate,10) & "' "
		end if

		sqlStr = sqlStr & "			and ((d.itemid=0 and d.beasongdate is Not NULL) or (d.currstate >= '7')) "

		if FRectshopid <> "" then
			sqlStr = sqlStr & "			and m.sitename = '" & FRectshopid & "' "
		else
			sqlStr = sqlStr & "			and m.sitename <> '10x10' "
		end if

		sqlStr = sqlStr + " and m.sitename in ( "
		sqlStr = sqlStr + " 	select c.userid "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		[db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " 		left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			c.userid=p.id "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and IsNull(p.etcjungsantype, 0) = 1 "
		sqlStr = sqlStr + " 		and IsNull(p.selltype, '') not in ('20166', '20032', '20046') "			'// B2C 제외
		sqlStr = sqlStr + " 		and p.userdiv in ('999') "
		sqlStr = sqlStr + " 		and c.userdiv in ('50') "
		sqlStr = sqlStr + " ) "

		sqlStr = sqlStr & "			and m.cancelyn='N' "
		sqlStr = sqlStr & "			and d.cancelyn<>'Y' "
		sqlStr = sqlStr & "			and d.currstate = '7' "
		sqlStr = sqlStr & "		group by "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) "
		sqlStr = sqlStr & "			, m.sitename "
		if (FRectGroupByBrand="on") then
		    sqlStr = sqlStr & "			, d.makerid"
		end if
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr & " 	on T.shopid=sm.shopid"
        sqlStr = sqlStr & " 	and T.yyyymmdd=sm.code01"
		sqlStr = sqlStr & " 	and sm.code02 <> 'beasongpay' " + vbcrlf
		if (FRectGroupByBrand="on") then
        	sqlStr = sqlStr & " 	and sm.code02 = T.makerid " + vbcrlf
		end if
        sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c s "
        sqlStr = sqlStr & " 	on T.shopid = s.userid "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	p.id = T.shopid "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	p.sellbizcd = b.BIZSECTION_CD "
        sqlStr = sqlStr & "	where 1=1 "

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " and sm.idx is null"
        end if

        sqlStr = sqlStr + " order by t.shopid asc ,T.yyyymmdd desc, T.totsum desc"
''rw sqlStr
		''response.write sqlStr &"<Br>"
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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         	= rsget("yyyymmdd")
				FItemList(i).Fshopid         	= rsget("shopid")
				FItemList(i).Fshopname       	= rsget("shopname")
				FItemList(i).Fmakerid           = rsget("makerid")
				FItemList(i).Ftotitemcnt     	= rsget("totitemcnt")
				FItemList(i).Ftotorgsum      	= rsget("totorgsum")
				FItemList(i).Ftotsum         	= rsget("totsum")
				''FItemList(i).Frealjungsansum 	= rsget("totchulgosum")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).fbuyprice         	= rsget("buyprice")

                FItemList(i).Ftotdeliveritemcnt	= rsget("totdeliveritemcnt")
				FItemList(i).Ftotdeliverorgsum	= rsget("totdeliverorgsum")
				FItemList(i).Ftotdeliversum		= rsget("totdeliversum")
				FItemList(i).Fbuydeliverprice	= rsget("buydeliverprice")

				FItemList(i).Fbizsection_cd	= rsget("bizsection_cd")
				FItemList(i).Fbizsection_nm	= rsget("bizsection_nm")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 배송비정산(온 입점몰)
    public sub getOnlineIpjumshopBeasongPayMaechulList()
		dim i,sqlStr

        sqlStr = " select"
        sqlStr = sqlStr & " t.yyyymmdd ,t.shopid ,t.totitemcnt , t.totorgsum, t.totsum ,t.totchulgosum ,t.buyprice"
        sqlStr = sqlStr & " ,sm.idx as precheckidx, s.socname as shopname "
        sqlStr = sqlStr & " , t.totdeliveritemcnt , t.totdeliverorgsum , t.totdeliversum , t.buydeliverprice "
        sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) as yyyymmdd "
		sqlStr = sqlStr & "			, m.sitename as shopid "
		sqlStr = sqlStr & "			, sum(case when d.itemid <> 0 then d.itemno else 0 end) as totitemcnt "
		sqlStr = sqlStr & "			, sum(case when d.itemid = 0 then d.itemno else 0 end) as totdeliveritemcnt "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliverorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliversum "
		sqlStr = sqlStr & "			, 0 as totchulgosum "
		sqlStr = sqlStr & "			, 0 as totdeliverchulgosum "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as buyprice "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as buydeliverprice "
		sqlStr = sqlStr & "		from "
		sqlStr = sqlStr & "			db_order.dbo.tbl_order_master m "
		sqlStr = sqlStr & "			join db_order.dbo.tbl_order_detail d "
		sqlStr = sqlStr & "			on "
		sqlStr = sqlStr & "				m.orderserial = d.orderserial "
		sqlStr = sqlStr & "			left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd "
		sqlStr = sqlStr & "			on d.idx = sd.linkdetailidx "
		sqlStr = sqlStr & "		where "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and d.itemid = 0 "		'// 배송비만

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " 	and sd.linkdetailidx is null"
        end if

		if FRectStartDate <> "" and FRectEndDate <> "" then
			sqlStr = sqlStr & "			and d.beasongdate >= '" & Left(FRectStartDate,10) & "' "
			sqlStr = sqlStr & "			and d.beasongdate < '" & Left(FRectEndDate,10) & "' "
		end if

		sqlStr = sqlStr & "			and ((d.itemid=0 and d.beasongdate is Not NULL) or (d.currstate >= '7')) "

		if FRectshopid <> "" then
			sqlStr = sqlStr & "			and m.sitename = '" & FRectshopid & "' "
		end if

		sqlStr = sqlStr & "			and m.cancelyn='N' "
		sqlStr = sqlStr & "			and d.cancelyn<>'Y' "
		sqlStr = sqlStr & "			and d.currstate = '7' "
		sqlStr = sqlStr & "		group by "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) "
		sqlStr = sqlStr & "			, m.sitename "
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr & " 	on T.shopid=sm.shopid"
        sqlStr = sqlStr & " 	and T.yyyymmdd=sm.code01"
        sqlStr = sqlStr & " 	and sm.code02 = 'beasongpay' " + vbcrlf
        sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c s "
        sqlStr = sqlStr & " 	on T.shopid = s.userid "
        sqlStr = sqlStr & "	where 1=1 "

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " and sm.idx is null"
        end if

        sqlStr = sqlStr + " order by t.shopid asc ,T.yyyymmdd desc, T.totsum desc"
''rw sqlStr
		''response.write sqlStr &"<Br>"
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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         	= rsget("yyyymmdd")
				FItemList(i).Fshopid         	= rsget("shopid")
				FItemList(i).Fshopname       	= rsget("shopname")
				FItemList(i).Ftotitemcnt     	= rsget("totitemcnt")
				FItemList(i).Ftotorgsum      	= rsget("totorgsum")
				FItemList(i).Ftotsum         	= rsget("totsum")
				''FItemList(i).Frealjungsansum 	= rsget("totchulgosum")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).fbuyprice         	= rsget("buyprice")

				FItemList(i).Ftotdeliveritemcnt	= rsget("totdeliveritemcnt")
				FItemList(i).Ftotdeliverorgsum	= rsget("totdeliverorgsum")
				FItemList(i).Ftotdeliversum		= rsget("totdeliversum")
				FItemList(i).Fbuydeliverprice	= rsget("buydeliverprice")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'// 배송비정산(온 입점몰)
    public sub getOnlineIpjumshopBeasongPayMaechulListNotReg()
		dim i,sqlStr

        sqlStr = " select"
        sqlStr = sqlStr & " t.yyyymmdd ,t.shopid ,t.totitemcnt , t.totorgsum, t.totsum ,t.totchulgosum ,t.buyprice"
        sqlStr = sqlStr & " ,sm.idx as precheckidx, s.socname as shopname "
        sqlStr = sqlStr & " , t.totdeliveritemcnt , t.totdeliverorgsum , t.totdeliversum , t.buydeliverprice, b.bizsection_cd, b.bizsection_nm "
        sqlStr = sqlStr & " from ("
		sqlStr = sqlStr & "		select "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) as yyyymmdd "
		sqlStr = sqlStr & "			, m.sitename as shopid "
		sqlStr = sqlStr & "			, sum(case when d.itemid <> 0 then d.itemno else 0 end) as totitemcnt "
		sqlStr = sqlStr & "			, sum(case when d.itemid = 0 then d.itemno else 0 end) as totdeliveritemcnt "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.orgitemcost,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliverorgsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as totsum "
		sqlStr = sqlStr & "			, sum(isnull(d.reducedPrice,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as totdeliversum "
		sqlStr = sqlStr & "			, 0 as totchulgosum "
		sqlStr = sqlStr & "			, 0 as totdeliverchulgosum "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid <> 0 then d.itemno else 0 end)) as buyprice "
		sqlStr = sqlStr & "			, sum(isnull(d.buycash,0)*(case when d.itemid = 0 then d.itemno else 0 end)) as buydeliverprice "
		sqlStr = sqlStr & "		from "
		sqlStr = sqlStr & "			db_order.dbo.tbl_order_master m "
		sqlStr = sqlStr & "			join db_order.dbo.tbl_order_detail d "
		sqlStr = sqlStr & "			on "
		sqlStr = sqlStr & "				m.orderserial = d.orderserial "
		sqlStr = sqlStr & "			left join [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail sd "
		sqlStr = sqlStr & "			on d.idx = sd.linkdetailidx "
		sqlStr = sqlStr & "		where "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and d.itemid = 0 "		'// 배송비만

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " 	and sd.linkdetailidx is null"
        end if

		if FRectStartDate <> "" and FRectEndDate <> "" then
			sqlStr = sqlStr & "			and d.beasongdate >= '" & Left(FRectStartDate,10) & "' "
			sqlStr = sqlStr & "			and d.beasongdate < '" & Left(FRectEndDate,10) & "' "
		end if

		sqlStr = sqlStr & "			and ((d.itemid=0 and d.beasongdate is Not NULL) or (d.currstate >= '7')) "

		if FRectshopid <> "" then
			sqlStr = sqlStr & "			and m.sitename = '" & FRectshopid & "' "
		else
			sqlStr = sqlStr & "			and m.sitename <> '10x10' "
		end if

		sqlStr = sqlStr + " and m.sitename in ( "
		sqlStr = sqlStr + " 	select c.userid "
		sqlStr = sqlStr + " 	from "
		sqlStr = sqlStr + " 		[db_user].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " 		left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " 		on "
		sqlStr = sqlStr + " 			c.userid=p.id "
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and IsNull(p.etcjungsantype, 0) = 1 "
		sqlStr = sqlStr + " 		and p.userdiv in ('999') "
		sqlStr = sqlStr + " 		and c.userdiv in ('50') "
		sqlStr = sqlStr + " 		and IsNull(p.selltype, '') not in ('20166', '20032', '20046') "			'// B2C 제외
		sqlStr = sqlStr + " ) "

		sqlStr = sqlStr & "			and m.cancelyn='N' "
		sqlStr = sqlStr & "			and d.cancelyn<>'Y' "
		sqlStr = sqlStr & "			and d.currstate = '7' "
		sqlStr = sqlStr & "		group by "
		sqlStr = sqlStr & "			convert(varchar(10),d.beasongdate,121) "
		sqlStr = sqlStr & "			, m.sitename "
        sqlStr = sqlStr & " ) T"
        sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_fran_meachuljungsan_submaster sm"
        sqlStr = sqlStr & " 	on T.shopid=sm.shopid"
        sqlStr = sqlStr & " 	and T.yyyymmdd=sm.code01"
        sqlStr = sqlStr & " 	and sm.code02 = 'beasongpay' " + vbcrlf
        sqlStr = sqlStr & " left join db_user.dbo.tbl_user_c s "
        sqlStr = sqlStr & " 	on T.shopid = s.userid "
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	p.id = T.shopid "
		sqlStr = sqlStr + " left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	p.sellbizcd = b.BIZSECTION_CD "
        sqlStr = sqlStr & "	where 1=1 "

        if FRectonlymifinish<>"" then
            sqlStr = sqlStr & " and sm.idx is null"
        end if

        sqlStr = sqlStr + " order by t.shopid asc ,T.yyyymmdd desc, T.totsum desc"
''rw sqlStr
		''response.write sqlStr &"<Br>"
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
				set FItemList(i) = new CWitakSellJungsanTargetItem

				FItemList(i).Fyyyymmdd         	= rsget("yyyymmdd")
				FItemList(i).Fshopid         	= rsget("shopid")
				FItemList(i).Fshopname       	= rsget("shopname")
				FItemList(i).Ftotitemcnt     	= rsget("totitemcnt")
				FItemList(i).Ftotorgsum      	= rsget("totorgsum")
				FItemList(i).Ftotsum         	= rsget("totsum")
				''FItemList(i).Frealjungsansum 	= rsget("totchulgosum")
				FItemList(i).Fprecheckidx		= rsget("precheckidx")
				FItemList(i).fbuyprice         	= rsget("buyprice")

				FItemList(i).Ftotdeliveritemcnt	= rsget("totdeliveritemcnt")
				FItemList(i).Ftotdeliverorgsum	= rsget("totdeliverorgsum")
				FItemList(i).Ftotdeliversum		= rsget("totdeliversum")
				FItemList(i).Fbuydeliverprice	= rsget("buydeliverprice")

				FItemList(i).Fbizsection_cd	= rsget("bizsection_cd")
				FItemList(i).Fbizsection_nm	= rsget("bizsection_nm")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub getOneEtcMeachul()
		dim i,sqlStr

		sqlStr = " select top 1 j.*, IsNull(j.totalsum,0) as totalsum, IsNull(j.totalsellcash,0) as totalsellcash, IsNull(j.totalsuplycash,0) as totalsuplycash, IsNull(j.totalbuycash,0) as totalbuycash "
		sqlStr = sqlStr + " 	, c.delivermethod, replace(isnull(p.company_no,'') ,'-','') as bizno "
		sqlStr = sqlStr + " 	, isnull(g.jungsan_acctname, '') as jungsan_acctname "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_fran_meachuljungsan_master j "
		sqlStr = sqlStr + " 	left join [db_storage].[dbo].tbl_cartoonbox_master c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		j.workidx = c.idx "
		sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		j.shopid = p.id "
		sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner_group g "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		p.groupid = g.groupid "
		sqlStr = sqlStr + " where j.idx=" + CStr(FRectidx) + " "
		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		set FOneItem = new CEtcMeachulMasterItem
		if  not rsget.EOF  then

			FOneItem.Fidx         = rsget("idx")
			FOneItem.Fshopid      = rsget("shopid")
			FOneItem.Ftitle       = db2html(rsget("title"))
			FOneItem.Ftotalsum    		= rsget("totalsum")
			FOneItem.Ftotalsellcash    	= rsget("totalsellcash")
			FOneItem.Ftotalsuplycash    = rsget("totalsuplycash")
			FOneItem.Ftotalbuycash    	= rsget("totalbuycash")

			FOneItem.Fdivcode     = rsget("divcode")
			FOneItem.Ftaxdate     = rsget("taxdate")
			FOneItem.Ftaxregdate  = rsget("taxregdate")
			FOneItem.Fregdate     = rsget("regdate")
			FOneItem.Fipkumdate   = rsget("ipkumdate")
			FOneItem.Fetcstr      = db2html(rsget("etcstr"))
			FOneItem.FStateCD	  = rsget("statecd")

			FOneItem.Freguserid      = rsget("reguserid")
			FOneItem.Fregusername    = db2html(rsget("regusername"))
			FOneItem.Ffinishuserid   = rsget("finishuserid")
			FOneItem.Ffinishusername = db2html(rsget("finishusername"))

            FOneItem.Fyyyymm    	= rsget("yyyymm")
            FOneItem.FdiffKey   	= rsget("diffKey")
            FOneItem.Fshopdiv   	= rsget("shopdiv")

            FOneItem.Fworkidx   	= rsget("workidx")
            FOneItem.Finvoiceidx   	= rsget("invoiceidx")
            FOneItem.Fdelivermethod = rsget("delivermethod")

            FOneItem.Fbizsection_cd 	= rsget("bizsection_cd")
            FOneItem.Fpapertype 		= rsget("papertype")
            FOneItem.Fpaperissuetype 	= rsget("paperissuetype")
            FOneItem.Fselltype 			= rsget("selltype")

            FOneItem.Fissuestatecd   	= rsget("issuestatecd")
            FOneItem.Fipkumstatecd   	= rsget("ipkumstatecd")

            FOneItem.FtaxNo			   	= rsget("neoTaxNo")
            FOneItem.Feserotaxkey   	= rsget("eserotaxkey")
			FOneItem.FbizNo			   	= rsget("bizno")

			FOneItem.Fjungsan_acctname 	= rsget("jungsan_acctname")			'// 예금주명
		end if
		rsget.close
	end sub

	public sub getEtcMeachulSubmasterList()
		dim i,sqlStr
		sqlStr = " select count(idx) as cnt "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
		sqlStr = sqlStr + " where masteridx=" + CStr(FRectidx)

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " j.*, b.baljudate "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster j "
		sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_shopbalju b "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	j.code02 = b.baljucode "
		sqlStr = sqlStr + " where j.masteridx=" + CStr(FRectidx)
		''sqlStr = sqlStr + " order by j.idx desc"
		sqlStr = sqlStr + " order by j.execdate"

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
				set FItemList(i) = new CEtcMeachulSubMasterItem

				FItemList(i).Fidx               = rsget("idx")
				FItemList(i).Fmasteridx         = rsget("masteridx")
				FItemList(i).Flinkidx           = rsget("linkidx")
				FItemList(i).Fshopid            = rsget("shopid")
				FItemList(i).Fcode01            = rsget("code01")
				FItemList(i).Fcode02            = rsget("code02")
				FItemList(i).Fexecdate          = rsget("execdate")
				FItemList(i).Ftotalcount        = rsget("totalcount")
				FItemList(i).Ftotalsellcash     = rsget("totalsellcash")
				FItemList(i).Ftotalbuycash      = rsget("totalbuycash")
				FItemList(i).Ftotalsuplycash    = rsget("totalsuplycash")
				FItemList(i).Ftotalorgsellcash  = rsget("totalorgsellcash")

				FItemList(i).Fbaljudate  		= rsget("baljudate")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

' 입금상태		' 2018.08.21 한용민
function drawSelectBoxIpkumState(selectBoxName, selectedId, chgval)
%>
	<select name="<%= selectBoxName %>"  <%= chgval %>>
		<option value="">전체</option>
		<option value="0" <% if selectedId="0" then response.write " selected" %>>입금전</option>
		<option value="5" <% if selectedId="5" then response.write " selected" %>>일부입금</option>
		<option value="9" <% if selectedId="9" then response.write " selected" %>>입금완료</option>
	</select>
<%
End function
%>

<%
function getJungsanXsiteComboHTML(selBoxName, selVal, onclickstr)
	Dim retStr
	retStr = "<select class='select' name='"&selBoxName&"' "&onclickstr&" >"
	retStr = retStr & " <option value=''>- 선택 -</option>"

	retStr = retStr & " <option value='interpark' "& CHKIIF(selVal="interpark","selected","") &" >인터파크</option>"
	retStr = retStr & " <option value='lotteimall' "& CHKIIF(selVal="lotteimall","selected","") &" >롯데아이몰</option>"
	retStr = retStr & " <option value='lotteCom' "& CHKIIF(selVal="lotteCom","selected","") &" >롯데닷컴</option>"
	retStr = retStr & " <option value='11st1010' "& CHKIIF(selVal="11st1010","selected","") &" >11번가</option>"
	retStr = retStr & " <option value='auction1010' "& CHKIIF(selVal="auction1010","selected","") &" >옥션</option>"
	retStr = retStr & " <option value='gmarket1010' "& CHKIIF(selVal="gmarket1010","selected","") &" >지마켓(NEW)</option>"
	retStr = retStr & " <option value='gseshop' "& CHKIIF(selVal="gseshop","selected","") &" >GS샵</option>"
	retStr = retStr & " <option value='cjmall' "& CHKIIF(selVal="cjmall","selected","") &" >CJ몰</option>"
	'retStr = retStr & " <option value='homeplus' "& CHKIIF(selVal="homeplus","selected","") &" >홈플러스</option>"
	retStr = retStr & " <option value='ssg' "& CHKIIF(selVal="ssg","selected","") &" >SSG</option>"
	'retStr = retStr & " <option value='ssg6006' "& CHKIIF(selVal="ssg6006","selected","") &" >SSG-이마트</option>"
	'retStr = retStr & " <option value='ssg6007' "& CHKIIF(selVal="ssg6007","selected","") &" >SSG-ssg</option>"
	retStr = retStr & " <option value='shintvshopping' "& CHKIIF(selVal="shintvshopping","selected","") &" >신세계TV쇼핑</option>"
	retStr = retStr & " <option value='skstoa' "& CHKIIF(selVal="skstoa","selected","") &" >SKSTOA</option>"
	retStr = retStr & " <option value='wetoo1300k' "& CHKIIF(selVal="wetoo1300k","selected","") &" >1300k</option>"
	retStr = retStr & " <option value='wconcept1010' "& CHKIIF(selVal="wconcept1010","selected","") &" >W컨셉</option>"
	retStr = retStr & " <option value='GS25' "& CHKIIF(selVal="GS25","selected","") &" >GS25카달로그</option>"
	retStr = retStr & " <option value='nvstorefarm' "& CHKIIF(selVal="nvstorefarm","selected","") &" >스토어팜</option>"
	retStr = retStr & " <option value='Mylittlewhoopee' "& CHKIIF(selVal="Mylittlewhoopee","selected","") &" >스토어팜 캣앤독</option>"
'	retStr = retStr & " <option value='nvstorefarmclass' "& CHKIIF(selVal="nvstorefarmclass","selected","") &" >스토어팜-클래스</option>"
'	retStr = retStr & " <option value='nvstoremoonbangu' "& CHKIIF(selVal="nvstoremoonbangu","selected","") &" >스토어팜 문방구</option>"
	retStr = retStr & " <option value='nvstoregift' "& CHKIIF(selVal="nvstoregift","selected","") &" >스토어팜 선물하기</option>"
	retStr = retStr & " <option value='wadsmartstore' "& CHKIIF(selVal="wadsmartstore","selected","") &" >와드스마트스토어</option>"
	retStr = retStr & " <option value='ezwel' "& CHKIIF(selVal="ezwel","selected","") &" >이지웰페어</option>"
	retStr = retStr & " <option value='kakaogift' "& CHKIIF(selVal="kakaogift","selected","") &" >카카오기프트</option>"
	retStr = retStr & " <option value='kakaostore' "& CHKIIF(selVal="kakaostore","selected","") &" >카카오톡스토어</option>"
	retStr = retStr & " <option value='boribori1010' "& CHKIIF(selVal="boribori1010","selected","") &" >보리보리</option>"
	retStr = retStr & " <option value='coupang' "& CHKIIF(selVal="coupang","selected","") &" >쿠팡</option>"
	retStr = retStr & " <option value='halfclub' "& CHKIIF(selVal="halfclub","selected","") &" >하프클럽</option>"
	retStr = retStr & " <option value='hmall1010' "& CHKIIF(selVal="hmall1010","selected","") &" >Hmall</option>"
	retStr = retStr & " <option value='WMP' "& CHKIIF(selVal="WMP","selected","") &" >WMP</option>"
	retStr = retStr & " <option value='wmpfashion' "& CHKIIF(selVal="wmpfashion","selected","") &" >WMPW패션</option>"
	retStr = retStr & " <option value='LFmall' "& CHKIIF(selVal="LFmall","selected","") &" >LFmall</option>"
	retStr = retStr & " <option value='lotteon' "& CHKIIF(selVal="lotteon","selected","") &" >롯데On</option>"
	retStr = retStr & " <option value='yes24' "& CHKIIF(selVal="yes24","selected","") &" >yes24</option>"
	retStr = retStr & " <option value='alphamall' "& CHKIIF(selVal="alphamall","selected","") &" >알파몰</option>"
	retStr = retStr & " <option value='ohou1010' "& CHKIIF(selVal="ohou1010","selected","") &" >오늘의집</option>"
	retStr = retStr & " <option value='casamia_good_com' "& CHKIIF(selVal="casamia_good_com","selected","") &" >까사미아</option>"
	retStr = retStr & " <option value='cookatmall' "& CHKIIF(selVal="cookatmall","selected","") &" >쿠캣</option>"
	retStr = retStr & " <option value='aboutpet' "& CHKIIF(selVal="aboutpet","selected","") &" >어바웃펫</option>"
	retStr = retStr & " <option value='goodshop1010' "& CHKIIF(selVal="goodshop1010","selected","") &" >굿샵</option>"
	retStr = retStr & " <option value='withnature1010' "& CHKIIF(selVal="withnature1010","selected","") &" >자연이랑</option>"
	retStr = retStr & " </select>"

	getJungsanXsiteComboHTML =retStr
end function

function getExtsongjangInputNOTIStr()
	Dim ret
	ret = 		"* interpark : 주문관리>전체주문내역 샹태조회가능(반품있는경우 붉은색), 주문관리>배송중/배송완료에서 송장변경가능 (구매확정예정일검토)"
	ret = ret & "<br> * 11st1010 : 주문관리>배송관리> 주문상태(배송중,배송완료) 로 검색후 송장수정, (배송중) 검색시 구매확정요청 버튼 활성화됨."
	ret = ret & "<br> * nvstorefarm : 판매관리>배송현황 관리 에서 송장수정가능, 구매확정 요청 버튼 있음. (상품주문번호 클릭시 자세히 볼 수 있음)"
	ret = ret & "<br> * auction/gmarket : 주문관리>배송중/배송완료 배송정보수정 에서 송장수정가능, 우리쪽 배송완료내역이 있는데 정산이 안된것은 문제가 있는것임."
	ret = ret & "<br> * LFmall : 주문(배송)현황> 배송지연, 상품준비중 에서 송장수정 가능"
	ret = ret & "<br> * ssg : 주문/배송 현황에서 검색, 미배송 관리에서 송장수정 가능, 배송완료처리시 정산반영됨."
	ret = ret & "<br> * coupang : 주문/배송관리 > 배송관리, 배송상태 상품준비중 인경우 송장수정가능"
	ret = ret & "<br> * hmall1010 : MENU > 출고회수 > 협력사배송 > 출고/회수 LIST 에서 검색후 상태클릭(출고완료) , 송장수정및 배송완료진행 가능(배송완료=정산확정일)"
	ret = ret & "<br>  (송장 수정한다고 배송완료로 전환되지 않음. 배송완료처리 해야)"
	ret = ret & "<br> * WMP : 주문/클레임관리 > 배송현황 배송번호로 검색, 송장번호 누르면 추적 추적완료되면 배송완료로 바뀜."
	ret = ret & "<br> * cjmall : 협력사배송관리 > 협력사출고/고객인수등록 에서 인수등록처리(선택,실인수자(.)저장) 주요택배사는 자동인수처리."
	ret = ret & "<br> * ezwel : 주문관리 > 전체주문리스트 검색후 송장수정 or 수취완료진행, (날짜선택이 전일 이후가능한듯)"
	ret = ret & "<br>  (출고준비중 검색 - 송장입력 배송중 처리, 배송중검색 인수완료처리)"
	getExtsongjangInputNOTIStr = ret
end function

Class CExtOrderTmpItem
	public FOutMallOrderSeq
	public FOrderSerial
	public FOrgDetailKey
	public FSellSite
	public FOutMallOrderSerial
	public FSellDate
	public FPayDate
	public FmatchItemID
	public Fmatchitemoption
	public Fsellprice
	public Frealsellprice
	public FItemOrderCount
	public ForderDlvPay
	public Fsendstate
	public FoutMallGoodsNo
	public Fref_outmallorderserial
	public FbeasongNum11st

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CExtJungsanCheckCSItem
	public Fcsid
	public Fdivcd
	public FdivName
	public Fwriteuser
	public Ffinishuser
	public Ftitle
	public Fcurrstate
	public Fregdate
	public Ffinishdate
	public Fconfirmdate
	public Fdeletedate
	public Fdeleteyn
	public Frequireupche
	public Fmakerid
	public Fsongjangdiv
	public Fsongjangno
	public Fextsitename

	public Frefasid
	public Frefminusorderserial
	public Frefchangeorderserial

	public function getRefOrderSerial()
		if isNULL(Frefminusorderserial) and isNULL(Frefchangeorderserial) then Exit function

		if isNULL(Frefchangeorderserial) then
			getRefOrderSerial = Frefminusorderserial
		else
			getRefOrderSerial = Frefchangeorderserial
		end if
	end function

	public function getCsStateName()
		dim istate : istate = Fcurrstate
		if IsNULL(istate) then Exit Function

		getCsStateName = istate

		select CASE istate
			CASE "B007"
				: getCsStateName = "완료"
			CASE "B001"
				: getCsStateName = "접수"
			CASE "B004"
				: getCsStateName = "운송장입력"
			CASE "B005"
				: getCsStateName = "업체확인요청"
			CASE "B006"
				: getCsStateName = "업체처리완료"
			CASE ELSE
				:
		end Select

	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CExtJungsanCheckOrderItem
	public Forderserial
	public Fbuyname
	public Freqname
	public FreqZipAddr
	public Fipkumdiv
	public Fcancelyn
	public Fdcancelyn
	public Fregdate
	public Fipkumdate
	public Fbaljudate
	public Fbeadaldiv
	public Fsitename
	public Fjumundiv
	public Fidx
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fmakerid
	public Fupcheconfirmdate
	public FitemcostcouponnotApplied
	public Fitemcost
	public Freducedprice
	public Fitemno
	public Fodlvfixday
	public Fsongjangdiv
	public Fsongjangno
	public Fbeasongdate
	public Fdlvfinishdt
	public Fjungsanfixdate
	public Fbuycash
	public Fomwdiv

	public Fcomment
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CExtOrderJungsanCheckItem
	public Fsitename
	public FordCnt
	public FChgOrdCNT
	public FretOrdCNT
	public ForgOrderserial
	public Fauthcode
	public Fitemid
	public Fitemoption
	public Fitemno
	public FitemcostSum
	public FreducedpriceSum
	public FbeasongMonth
	public Forgsongjangdiv
	public Forgsongjangno
	public Forgdlvfinishdt
	public Forgjungsanfixdate

	public FMinus_itemno
	public FMinus_itemcostSum
	public FMinus_reducedpriceSum
	public FMinus_beasongmonth

	public FextItemNoSum
	public FextitemcostSum
	public FextReducedPriceSum
	public FextMeachulMonth

	public Fcomment
	public Fdiffitemno
	public FdiffSum
	public Fjorgorderserial


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanItem
	public Fsellsite
	public FextOrderserial
	public FextOrderserSeq
	public FextOrgOrderserial
	public FextItemNo
	public FextItemCost
	public FextReducedPrice
	public FextOwnCouponPrice
	public FextTenCouponPrice
	public FextJungsanType
	public FextCommPrice
	public FextTenMeachulPrice
	public FextTenJungsanPrice
	public FextMeachulDate
	public FextJungsanDate
	public FOrgOrderserial
	public Fitemid
	public Fitemoption
	public FsiteNo
	public FMinusOrderserial

	public FextTenJungsanPrice_ETC


	public Forderitemcost
	public Forderreducedprice
	public Forderitemno
	public Forderbeasongdate
	public Fdlvfinishdt
	public Fjungsanfixdate

	public Fdtlcancelyn
	public Fmastercancelyn


	public FMakerid
	public Fmwdiv
	public Ftenbuycash
	public Fjungsangain

	public FExtitemid
	public Fref_Slice_extOrderserSeq

	Public Fbeasongdate

	Public FRowNo
	Public FSumItemNo
	Public FSumitemcost
	Public FSumMeachulPrice
	Public FSumReducedPrice
	Public FSumOwnCouponPrice
	Public FSumTenCouponPrice
	Public FSumCommPrice
	Public FSumJungsanPrice
	Public FMiMapTTLCnt
	Public FSumJungsanPrice_ETC
	Public FSumtenReducedPrice
	Public FBigoSum
	Public FBigoBeasongSum
	Public FextMeachulMonth
	Public FjungsanfixMonth

	public Fsitename
	public FReducedPrice

	public function isCpnValEditAvailRow()
		isCpnValEditAvailRow = FALSE

		if isNULL(FOrgOrderserial) then  ''매칭된 내역이 없으면 RETURN
			Exit function
		end if

		if (Fmastercancelyn<>"N" or Fdtlcancelyn="Y") then
			Exit function
		end if

		if (getJOdiffitemno<>0) then
			Exit function
		end if

		if (getJOdiffReducedprice=-1) then
			Exit function
		end if

		if (ABS(Null2Zero(Forderitemno))>ABS(getJOdiffReducedprice)) then
			Exit function
		end if

		''if isNULL(Forderbeasongdate) then '' 출고가 아직 안된거는 수정 가능하다. but 정산이 올라올 가능성이 거의 없다.
		if isNULL(Fjungsanfixdate) then
			isCpnValEditAvailRow =TRUE
			Exit function
		end if

		'' 출고월이 이번달이면 수정가능
		''if LEFT(Forderbeasongdate,7)=LEFT(Date(),7) then
		if LEFT(Fjungsanfixdate,7)=LEFT(Date(),7) then
			isCpnValEditAvailRow =TRUE
			Exit function
		end if

		'' 출고일 기준 익월 3일 정도 까지 수정 가능
		if (Day(Date())>4) then
			isCpnValEditAvailRow =FALSE
			Exit function
		else
			''if (CDate(Forderbeasongdate)>=CDate(LEFT(dateadd("m",-4,Date()),7)+"-01") ) then
			if (CDate(Fjungsanfixdate)>=CDate(LEFT(dateadd("d",-3,Date()),7)+"-01") ) then
				isCpnValEditAvailRow =TRUE
			end if
		end if

	end function

	public function getBigoStr()
		dim ret
		if isNULL(FOrgOrderserial) then
			ret = "미매칭"
		end if

		if (Fmastercancelyn<>"N") then
			ret = ret & "주문취소"
		elseif (Fdtlcancelyn="Y") then
			ret = ret & "상품취소"
		elseif (Fdtlcancelyn="A") then
			ret = ret & "상품추가"
		end if

		getBigoStr = ret

	end function

	public function getJOdiffItemCost()
		if isNULL(Forderitemcost) then
			getJOdiffItemCost = FextItemCost
		else
			getJOdiffItemCost = FextItemCost-Forderitemcost
		end if
	end function

	public function getJOdiffReducedprice()
		if isNULL(Forderreducedprice) then
			getJOdiffReducedprice = FextReducedPrice
		else
			getJOdiffReducedprice = FextReducedPrice-Forderreducedprice
		end if
	end function

	public function getJOdiffitemno()
		if isNULL(Forderitemno) then
			getJOdiffitemno = FextItemNo
		else
			getJOdiffitemno = FextItemNo-Forderitemno
		end if
	end function

	'' 수수료율
	public function GetSusumargin()
		if (FextItemCost<>0) then
			GetSusumargin = FormatNumber((FextCommPrice/FextItemCost)*100,2)
		end if
	end function

	public function GetDiffReducedPrice()
		GetDiffReducedPrice = FextTenMeachulPrice-(FextReducedPrice)
	end function

	public function GetDiffMeachulPrice()
		GetDiffMeachulPrice = FextItemCost-FextTenMeachulPrice-(FextOwnCouponPrice+FextTenCouponPrice)
	end function

	public function GetDiffJungsanPrice()
		GetDiffJungsanPrice = FextTenJungsanPrice-(FextTenMeachulPrice-FextCommPrice)-FextTenJungsanPrice_ETC
	end function

	Public function GetSellSiteName()
		Select Case Fsellsite
			Case "interpark"
				GetSellSiteName = "인터파크"
			Case Else
				GetSellSiteName = Fsellsite
				if FsiteNo ="6006" then
					GetSellSiteName= GetSellSiteName&"-이마트"
				elseif 	FsiteNo ="6007" then
					GetSellSiteName= GetSellSiteName&"-ssg"
				end if
		End Select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanDiffItem
	public Fsitename
	public Fyyyymm
	public FTMeachulItem
	public FTMeachulDLV
	public FTbuycashItem
	public FTbuycashDLV

	public FXMeachulItem
	public FXMeachulDLV
	public FXJungsanItem
	public FXJungsanDLV
	public FregDt
	public FupdDt
	public FmonthItemDiff
	public FmonthdlvDiff
	public FdiffITEMsum
	public FdiffDlvsum

	public FmonthItemDiffMapErr
	public FmonthdlvDiffTMapErr

	public FmonthItemDiffNoExists
	public FmonthdlvDiffTNoExists

	public FMonthDiffSum
	public FMonthErrAsignSum
	public FMonthErrAsignItemSum
	public FMonthErrAsignItemSumReqCheck
	public FMonthnotAssignErr

	public function getSumVsDtlDiffSum()
		getSumVsDtlDiffSum = (FTMeachulItem+FTMeachulDLV)-(FXMeachulItem+FXMeachulDLV)-FMonthDiffSum
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanDiffItem_OLD
	public Fyyyymm
	public Fsellsite
	public Forderserial
	public FMeachulPriceSUM
	public FextMeachulPriceSUM
	public FMeachulPriceSUM1
	public FextMeachulPriceSUM1
	public FMeachulPriceSUM2
	public FextMeachulPriceSUM2
	public FMeachulPriceSUM3
	public FextMeachulPriceSUM3

	Public function GetSellSiteName()
		Select Case Fsellsite
			Case "interpark"
				GetSellSiteName = "인터파크"
			Case "lotteimall"
				GetSellSiteName = "롯데아이몰"
			Case Else
				GetSellSiteName = Fsellsite
		End Select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanFixedErrItem
	public Fyyyymm
	public Fsellsite
	public Forderserial
	public Fitemid
	public Fitemoption
	public Fitemnosum
	public Freducedsum
	public Fbuycashsum
	public FextItemNoSum
	public FextreducedpriceSum
	public FextTenJungsanPriceSum
	public Fupddt
	public FErrAsignMonth
	public FErrAsignSum
	public Fdiffthis
	public FaccErrNoSum
	public FacctErrsum
	public FaccAsgnErrSum
	public FaccTTLErrSum

	public Foutmallorderserial
	public Fcomment

	public function getDiffNo()
		getDiffNo = Fitemnosum-FextItemNoSum
	end function

	public function getDiffSum()
		getDiffSum = Freducedsum-FextreducedpriceSum
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanErrItem
	public Fsitename
	public Fyyyymmdd
	public Fauthcode
	public Foorderserial
	public Fitemid
	public Fitemoption
	public Fdiffno
	public Fdiffsum

	public Fjumundiv
	public Flinkorderserial
	public FregDt
	public FupdDt

	public Fcomment
	public Ferrortype

	public function getErrorTypeName()
		if isNULL(Ferrortype) then Exit function
		if (Ferrortype=0) then Exit function

		if (Ferrortype=1) then
			getErrorTypeName = "매핑오차"
		else
			getErrorTypeName=CStr(Ferrortype)
		end if
	end function

	public function getJumundivName()
		if isNULL(Fjumundiv) then Exit function

		if (Fjumundiv="9") then
			getJumundivName = "반품"
		elseif (Fjumundiv="6") then
			getJumundivName = "교환"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanStatisticItem
	public Fsellsite
	public FextMeachulDate

	public FtotExtTenMeachulPriceProduct
	public FtotExtCommPriceProduct
	public FtotExtTenJungsanPriceProduct

	public FtotExtTenMeachulPriceDeliver
	public FtotExtCommPriceDeliver
	public FtotExtTenJungsanPriceDeliver

	public FtotExtTenMeachulPriceEtc
	public FtotExtCommPriceEtc
	public FtotExtTenJungsanPriceEtc

	public FtotExtTenMeachulPrice
	public FtotExtCommPrice
	public FtotExtTenJungsanPrice
	public FextMiMapping
	public FextRowCount
    public FextMiMapping_C
	public FextRowCount_C
	public FMiMappOrder
	public FMiMappOrder_C

	public  FtotExtitemCostProduct
	public  FtotExtReducedPriceProduct
    public  FtotExtOwnCouponPriceProduct
    public  FtotExtTenCouponPriceProduct

	public  FtotExtitemCostDeliver
    public  FtotExtReducedPriceDeliver
    public  FtotExtOwnCouponPriceDeliver
    public  FtotExtTenCouponPriceDeliver

	public  FtotExtitemCost
	public  FtotExtReducedPrice
	public  FtotExtOwnCouponPrice
	public  FtotExtTenCouponPrice

	public function GetDiffMeachulPrice()
		GetDiffMeachulPrice = FtotExtitemCost-FtotExtTenMeachulPrice-(FtotExtOwnCouponPrice+FtotExtTenCouponPrice)
	end function

	public function GetDiffJungsanPrice()
		GetDiffJungsanPrice = FtotExtTenJungsanPrice-(FtotExtTenMeachulPrice-FtotExtCommPrice) - FtotExtTenJungsanPriceEtc
	end function

	Public function GetSellSiteName()
'2019-04-01 이문재 이사님 요청..전체 다 브랜드ID로 보여달라는..
'		Select Case Fsellsite
'			Case "interpark"
'				GetSellSiteName = "인터파크"
'			Case Else
				GetSellSiteName = Fsellsite
'		End Select
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsanCommentItem
	public Frowidx
	public Forderserial
	public Fitemid
	public Fitemoption
	public Freguserid
	public Fcomment
	public Fregdate
	public Fdeldate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CExtJungsan
    public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectSellSite
	public FRectJungsanType
	public FRectGroupGubun

	public FRectStartDate
	public FRectEndDate
	public FRectYYYYMM
	public FRectDiffType

	public FRectSearchField
	public FRectSearchText

	public FRowNo
	public FSumItemNo
	public FSumitemcost
	public FSumMeachulPrice
	public FSumReducedPrice
	public FSumOwnCouponPrice
	public FSumTenCouponPrice
	public FSumCommPrice
	public FSumJungsanPrice
	public FMiMapTTLCnt
	public FSumJungsanPrice_ETC

	public FRectMiMap
	public FRectJungsanfixdate
	public FRectVatYn
	public FRectReturnOnly
	public FRectErrexists
	public FRectExistsBigo
	public FRectExceptItemCostZero
	public FRectDlvMonth
	public FRectMiMapMinus

	public FRectMakerid
	public FRectItemid
	public FRectReturnExcept
	public FRectMinusGainOnly
	public FRectDiffType2

	public FRectOrderserial
	public FRectItemOption

	public FRectCheckBySum
	public FonlyErrNoExists
	public FRectErrorType
	public FRectAccerrtype

	public FdiffnoSum
	public FdiffsumSum
	public FErrAsignSum

	public FRectIpkumdateChk
	public FRectAStartdate
	public FRectAEndDate


	public function GetDiffMeachulPrice()
		GetDiffMeachulPrice = FSumitemcost-FSumMeachulPrice-(FSumOwnCouponPrice+FSumTenCouponPrice)
	end function

	public function GetDiffJungsanPrice()
		GetDiffJungsanPrice = FSumJungsanPrice-(FSumMeachulPrice-FSumCommPrice) - FSumJungsanPrice_ETC
	end function

	public function getExtjungsanCommentList()
		Dim sqlStr
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_List] '"&FRectOrderserial&"'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanCommentItem

				FItemList(i).Frowidx		= db3_rsget("rowidx")
				FItemList(i).Forderserial	= db3_rsget("orderserial")
				FItemList(i).Fitemid		= db3_rsget("itemid")
				FItemList(i).Fitemoption	= db3_rsget("itemoption")
				FItemList(i).Freguserid		= db3_rsget("reguserid")
				FItemList(i).Fcomment		= db3_rsget("comment")
				FItemList(i).Fregdate		= db3_rsget("regdate")
				FItemList(i).Fdeldate		= db3_rsget("deldate")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	end function

	public function getOutJungsanCheckCSInfo()
		Dim sqlStr
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTMALL_Jungsan_CheckCSInfo] '"&FRectOrderserial&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtJungsanCheckCSItem
				FItemList(i).Fcsid          = rsget("id")
				FItemList(i).Fdivcd         = rsget("divcd")
				FItemList(i).FdivName       = rsget("divName")
				FItemList(i).Fwriteuser     = rsget("writeuser")
				FItemList(i).Ffinishuser    = rsget("finishuser")
				FItemList(i).Ftitle         = rsget("title")
				FItemList(i).Fcurrstate     = rsget("currstate")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Ffinishdate    = rsget("finishdate")
				FItemList(i).Fconfirmdate   = rsget("confirmdate")
				FItemList(i).Fdeletedate    = rsget("deletedate")
				FItemList(i).Fdeleteyn      = rsget("deleteyn")
				FItemList(i).Frequireupche  = rsget("requireupche")
				FItemList(i).Fmakerid  		= rsget("makerid")
				FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
				FItemList(i).Fsongjangno    = rsget("songjangno")
				FItemList(i).Fextsitename   = rsget("extsitename")

				FItemList(i).Frefasid				= rsget("refasid")
				FItemList(i).Frefminusorderserial	= rsget("refminusorderserial")
				FItemList(i).Frefchangeorderserial	= rsget("refchangeorderserial")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Function

	public function getOutJungsanCheckOrderInfo()
		Dim sqlStr
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTMALL_Jungsan_CheckOrderInfo] '"&FRectOrderserial&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtJungsanCheckOrderItem
				FItemList(i).Forderserial	= rsget("orderserial")
				FItemList(i).Fbuyname		= rsget("buyname")
				FItemList(i).Freqname		= rsget("reqname")
				FItemList(i).FreqZipAddr	= rsget("reqZipAddr")
				FItemList(i).Fipkumdiv		= rsget("ipkumdiv")
				FItemList(i).Fcancelyn		= rsget("cancelyn")
				FItemList(i).Fdcancelyn		= rsget("dcancelyn")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fipkumdate		= rsget("ipkumdate")
				FItemList(i).Fbaljudate		= rsget("baljudate")
				FItemList(i).Fbeadaldiv		= rsget("beadaldiv")
				FItemList(i).Fsitename		= rsget("sitename")
				FItemList(i).Fjumundiv		= rsget("jumundiv")
				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Fitemname		= rsget("itemname")
				FItemList(i).Fitemoptionname	= rsget("itemoptionname")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fupcheconfirmdate	= rsget("upcheconfirmdate")
				FItemList(i).FitemcostcouponnotApplied = rsget("itemcostcouponnotApplied")
				FItemList(i).Fitemcost		= rsget("itemcost")
				FItemList(i).Freducedprice	= rsget("reducedprice")
				FItemList(i).Fitemno		= rsget("itemno")
				FItemList(i).Fodlvfixday	= rsget("odlvfixday")
				FItemList(i).Fsongjangdiv	= rsget("songjangdiv")
				FItemList(i).Fsongjangno	= rsget("songjangno")
				FItemList(i).Fbeasongdate	= rsget("beasongdate")
				FItemList(i).Fdlvfinishdt	= rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate	= rsget("jungsanfixdate")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Fomwdiv		= rsget("omwdiv")

				FItemList(i).Fcomment		= rsget("comment")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Function

	public function getExtJungsanOrderDiffList()
		Dim sqlStr
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_OrderDiffCheck] '" & FRectSellSite & "','" & FRectStartDate & "','" & FRectEndDate & "', '" & FRectDiffType & "', '" & FPageSize & "','"&FRectDlvMonth&"',"&CHKIIF(FRectCheckBySum<>"",1,"NULL")&""

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				'FItemList(i).FextJungsanType		= rsget("extJungsanType")
				'FItemList(i).FextCommPrice			= rsget("extCommPrice")
				'FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				'FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				'FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				'FItemList(i).FsiteNo				= rsget("siteNo")

				FItemList(i).Forderitemcost			= rsget("itemcost")
				FItemList(i).Forderreducedprice		= rsget("reducedprice")
				FItemList(i).Forderitemno			= rsget("itemno")
				FItemList(i).Forderbeasongdate		= rsget("beasongdate")
				if NOT isNULL(FItemList(i).Forderbeasongdate) THEN
					FItemList(i).Forderbeasongdate = LEFT(FItemList(i).Forderbeasongdate,10)
				end if

				FItemList(i).Fdtlcancelyn		= rsget("dtlcancelyn")
				FItemList(i).Fmastercancelyn	= rsget("mastercancelyn")

				FItemList(i).Fdlvfinishdt		= rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate	= rsget("jungsanfixdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function getExtJungsanOrderDiffList_replica()
		Dim sqlStr
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_OrderDiffCheck_replica] '" & FRectSellSite & "','" & FRectStartDate & "','" & FRectEndDate & "', '" & FRectDiffType & "', '" & FPageSize & "','"&FRectDlvMonth&"',"&CHKIIF(FRectCheckBySum<>"",1,"NULL")&""

        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= db3_rsget("sellsite")
				FItemList(i).FextOrderserial		= db3_rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= db3_rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= db3_rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= db3_rsget("extItemNo")
				FItemList(i).FextItemCost			= db3_rsget("extItemCost")
				FItemList(i).FextReducedPrice		= db3_rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= db3_rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= db3_rsget("extTenCouponPrice")
				'FItemList(i).FextJungsanType		= db3_rsget("extJungsanType")
				'FItemList(i).FextCommPrice			= db3_rsget("extCommPrice")
				'FItemList(i).FextTenMeachulPrice	= db3_rsget("extTenMeachulPrice")
				'FItemList(i).FextTenJungsanPrice	= db3_rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= db3_rsget("extMeachulDate")
				'FItemList(i).FextJungsanDate		= db3_rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= db3_rsget("OrgOrderserial")
				FItemList(i).Fitemid				= db3_rsget("itemid")
				FItemList(i).Fitemoption			= db3_rsget("itemoption")
				'FItemList(i).FsiteNo				= db3_rsget("siteNo")

				FItemList(i).Forderitemcost			= db3_rsget("itemcost")
				FItemList(i).Forderreducedprice		= db3_rsget("reducedprice")
				FItemList(i).Forderitemno			= db3_rsget("itemno")
				FItemList(i).Forderbeasongdate		= db3_rsget("beasongdate")
				if NOT isNULL(FItemList(i).Forderbeasongdate) THEN
					FItemList(i).Forderbeasongdate = LEFT(FItemList(i).Forderbeasongdate,10)
				end if

				FItemList(i).Fdtlcancelyn		= db3_rsget("dtlcancelyn")
				FItemList(i).Fmastercancelyn	= db3_rsget("mastercancelyn")

				FItemList(i).Fdlvfinishdt		= db3_rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate	= db3_rsget("jungsanfixdate")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function getExtOrderJungsanDiffList()
		Dim sqlStr
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_Check_Outmall_OrderVsOutJungsan] '"&FRectDlvMonth&"','" & FRectSellSite & "',"&CHKIIF(FRectDiffType="","NULL",FRectDiffType)&","&CHKIIF(FRectDiffType2="","NULL",FRectDiffType2)&""

        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof

				set FItemList(i) = new CExtOrderJungsanCheckItem

				FItemList(i).Fsitename 			= db3_rsget("sitename")
				FItemList(i).FordCnt			= db3_rsget("ordCnt")
				FItemList(i).FChgOrdCNT			= db3_rsget("ChgOrdCNT")
				FItemList(i).FretOrdCNT			= db3_rsget("retOrdCNT")
				FItemList(i).ForgOrderserial	= db3_rsget("orgOrderserial")
				FItemList(i).Fauthcode			= db3_rsget("authcode")
				FItemList(i).Fitemid			= db3_rsget("itemid")
				FItemList(i).Fitemoption		= db3_rsget("itemoption")
				FItemList(i).Fitemno			= db3_rsget("itemno")
				FItemList(i).FitemcostSum		= db3_rsget("itemcostSum")
				FItemList(i).FreducedpriceSum	= db3_rsget("reducedpriceSum")
				FItemList(i).FbeasongMonth		= db3_rsget("beasongMonth")
				FItemList(i).Forgsongjangdiv	= db3_rsget("orgsongjangdiv")
				FItemList(i).Forgsongjangno		= db3_rsget("orgsongjangno")
				FItemList(i).Forgdlvfinishdt	= db3_rsget("orgdlvfinishdt")
				FItemList(i).Forgjungsanfixdate	= db3_rsget("orgjungsanfixdate")

				FItemList(i).FMinus_itemno			= db3_rsget("Minus_itemno")
				FItemList(i).FMinus_itemcostSum		= db3_rsget("Minus_itemcostSum")
				FItemList(i).FMinus_reducedpriceSum	= db3_rsget("Minus_reducedpriceSum")
				FItemList(i).FMinus_beasongmonth		= db3_rsget("Minus_beasongmonth")

				FItemList(i).FextItemNoSum			= db3_rsget("extItemNoSum")
				FItemList(i).FextitemcostSum		= db3_rsget("extitemcostSum")
				FItemList(i).FextReducedPriceSum	= db3_rsget("extReducedPriceSum")
				FItemList(i).FextMeachulMonth		= db3_rsget("extMeachulMonth")

				FItemList(i).Fcomment				= db3_rsget("comment")
				FItemList(i).Fdiffitemno			= db3_rsget("diffitemno")
				FItemList(i).FdiffSum				= db3_rsget("diffSum")

				FItemList(i).Fjorgorderserial		= db3_rsget("jorgorderserial")  ''정산내역이 있는지 여부판단.

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	Public Function getExtOrderJungsanFixdate()
		Dim sqlStr
		sqlStr = " exec  [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_Fixdate] '" & FRectSellSite & "' "
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem
				FItemList(i).Fsitename				= rsget("sitename")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if
				FItemList(i).Fdlvfinishdt			= rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate		= rsget("jungsanfixdate")
				FItemList(i).Fbeasongdate			= rsget("beasongdate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	End Function

	public function GetExtJungsanByItemDW()
		Dim sqlStr, i
        '' @styyyymmdd ,@edyyyymmdd ,@sellsite varchar(32) = NULL,@makerid varchar(32) = NULL
        '',@itemid int = NULL,@returnExcept int = 0 -- 반품정산건 제외	,@minuscOnly int = 0 -- 마이너스 상품이 있는경우만.
		sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_GainSumDtlByItemCNT] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "'," & FRectItemid & ", " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ""
''rw sqlStr
        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly
		    FTotalCount = rsSTSget("CNT")
        rsSTSget.Close

        if (FTotalCount<1) then
            FResultCount = 0
            Exit function
        end if

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

        sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_GainSumDtlByItem] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "'," & FRectItemid & ", " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ","&FPageSize&","&FCurrPage
        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsSTSget.RecordCount
		if FResultCount<0 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsSTSget("sellsite")
				FItemList(i).FextOrderserial		= rsSTSget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsSTSget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsSTSget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsSTSget("extItemNo")

				FItemList(i).FextItemCost			= rsSTSget("extItemCost")
				FItemList(i).FextReducedPrice		= rsSTSget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsSTSget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsSTSget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsSTSget("extJungsanType")
				FItemList(i).FextCommPrice			= rsSTSget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsSTSget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsSTSget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsSTSget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsSTSget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsSTSget("OrgOrderserial")
				FItemList(i).Fitemid				= rsSTSget("itemid")
				FItemList(i).Fitemoption			= rsSTSget("itemoption")
				'FItemList(i).FsiteNo				= rsSTSget("siteNo")
				'FItemList(i).FMinusOrderserial		= rsSTSget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if

				FItemList(i).Fmakerid			= rsSTSget("makerid")
				FItemList(i).Fmwdiv				= rsSTSget("omwdiv")
				FItemList(i).Ftenbuycash		= rsSTSget("tenbuycash")
				FItemList(i).Fjungsangain		= rsSTSget("jungsangain")

				i=i+1
				rsSTSget.moveNext
			loop
		end if
		rsSTSget.Close
	end Function

	public function GetExtJungsanCheckTargetList()
		dim i, sqlStr

        sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_CheckRequireList] '" & FRectSellSite & "','"&FRectStartdate&"','"&FRectEndDate&"','"&FRectDiffType&"'"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if
				FItemList(i).Fdlvfinishdt			= rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate		= rsget("jungsanfixdate")
				FItemList(i).Fbeasongdate			= rsget("beasongdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetExtJungsanMapCheckListTmpOrder()
		dim i, sqlStr
		dim iextOrderserial, iOrgOrderserial, iextitemid

		if (FRectSearchField="extOrderserial") then
			iextOrderserial = FRectSearchText
		elseif (FRectSearchField="OrgOrderserial") then
			iOrgOrderserial = FRectSearchText
		elseif (FRectSearchField="extitemid") then
			iextitemid = FRectSearchText
		end if

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_MiMapCheckList_TmpOrder] '" & FRectSellSite & "', '" & iextOrderserial & "', '" & iOrgOrderserial & "','"&iextitemid&"' "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtOrderTmpItem
				FItemList(i).FOutMallOrderSeq	= rsget("OutMallOrderSeq")
				FItemList(i).FOrderSerial		= rsget("OrderSerial")
				FItemList(i).FOrgDetailKey		= rsget("OrgDetailKey")
				FItemList(i).FSellSite			= rsget("SellSite")
				FItemList(i).FOutMallOrderSerial	= rsget("OutMallOrderSerial")
				FItemList(i).FSellDate			= rsget("SellDate")
				FItemList(i).FPayDate			= rsget("PayDate")
				FItemList(i).FmatchItemID		= rsget("matchItemID")
				FItemList(i).Fmatchitemoption	= rsget("matchitemoption")
				FItemList(i).Fsellprice			= rsget("sellprice")
				FItemList(i).Frealsellprice		= rsget("realsellprice")
				FItemList(i).FItemOrderCount	= rsget("ItemOrderCount")
				FItemList(i).ForderDlvPay		= rsget("orderDlvPay")
				FItemList(i).Fsendstate			= rsget("sendstate")
				FItemList(i).FoutMallGoodsNo	= rsget("outMallGoodsNo")
				FItemList(i).Fref_outmallorderserial	= rsget("ref_outmallorderserial")
				FItemList(i).FbeasongNum11st	= rsget("beasongNum11st")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetExtJungsanMapCheckList()
		dim i, sqlStr
		dim iextOrderserial, iOrgOrderserial, iextitemid

		if (FRectSearchField="extOrderserial") then
			iextOrderserial = FRectSearchText
		elseif (FRectSearchField="OrgOrderserial") then
			iOrgOrderserial = FRectSearchText
		elseif (FRectSearchField="extitemid") then
			iextitemid = FRectSearchText
		end if

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_MiMapCheckList] '" & FRectSellSite & "', '" & FRectJungsanType & "', '" & iextOrderserial & "', '" & iOrgOrderserial & "','"&iextitemid&"' "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")
				FItemList(i).Fref_Slice_extOrderserSeq = rsget("ref_Slice_extOrderserSeq")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if

				FItemList(i).FExtitemid = rsget("Extitemid")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetExtJungsanSum
		dim i, sqlStr, addSqlStr
		'// ====================================================================
		addSqlStr = ""

		 if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		        addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		 elseif (FRectSellSite<>"") then
			    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		 end if

		if (FRectJungsanType <> "") then
			addSqlStr = addSqlStr + " and j.extJungsanType = '" + CStr(FRectJungsanType) + "' "
		end if

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField="extOrderserial") then
				addSqlStr = addSqlStr + " and (LEFT(j.extOrderserial,"&LEN(FRectSearchText)&")='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " 	or j.extOrgOrderserial='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " )"
			elseif (FRectSearchField="matchitemid") then
				addSqlStr = addSqlStr + " and j.itemid='"&FRectSearchText&"'"&vbCRLF
			else
				addSqlStr = addSqlStr + " and j." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectMiMap<>"") then
			addSqlStr = addSqlStr + " and isNULL(j.OrgOrderserial,'')=''"
		end if

		if (FRectMiMapMinus<>"") then
			addSqlStr = addSqlStr + " and j.extitemno<0 and isNULL(j.MinusOrderserial,'')=''"
		end if

		if (FRectVatYn<>"") then
			addSqlStr = addSqlStr + " and extvatyn='"&FRectVatYn&"'"
		end if

		if (FRectReturnOnly<>"") then
			addSqlStr = addSqlStr + " and extitemno<1"
		end if

		if (FRectErrexists<>"") then
			addSqlStr = addSqlStr + " and extTenMeachulPrice - extReducedPrice <> 0"
		end if

		if (FRectExistsBigo <> "") Then
			addSqlStr = addSqlStr + " and ("
			addSqlStr = addSqlStr + "	(extItemCost-extReducedPrice-extOwnCouponPrice-extTenCouponPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extTenJungsanPrice-extReducedPrice+extCommPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extReducedPrice<>extTenMeachulPrice)"
			addSqlStr = addSqlStr + " )"
		End If

		if (FRectExceptItemCostZero<>"") then
			addSqlStr = addSqlStr + " and extItemCost<>0"
		end if

		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""

		dim sqlSum
		sqlSum = " select LEFT(extMeachulDate, 7) as extMeachulMonth, isNull(convert(varchar(7),d.jungsanfixdate,121), '') as jungsanfixMonth "
		sqlSum = sqlSum & ",count(*) cnt, sum(j.extitemNO) as itemno , sum(j.extItemCost*j.extItemno) as itemcost "
		sqlSum = sqlSum & " , sum(j.extTenMeachulPrice*j.extItemNo) as MeachulPrice "
		sqlSum = sqlSum & ", sum(j.extReducedPrice*j.extItemNo) as  ReducedPrice "
		sqlSum = sqlSum & ", sum(j.extOwnCouponPrice*j.extItemNo) as  OwnCouponPrice "
		sqlSum = sqlSum & ", sum(j.extTenCouponPrice*j.extItemNo) as TenCouponPrice "
		sqlSum = sqlSum & ", sum(j.extCommPrice*j.extItemNo) as CommPrice "
		sqlSum = sqlSum & ", sum(j.extTenJungsanPrice*j.extItemNo) as JungsanPrice "
		sqlSum = sqlSum & ", sum((CASE WHEN j.extJungsanType not in ('C','D') then j.extTenJungsanPrice*j.extItemNo else 0 END)) as JungsanPrice_ETC "
		sqlSum = sqlSum & ", sum(CASE WHEN isNULL(j.OrgOrderserial,'')='' THEN 1 ELSE 0 END) as MiMapTTLCnt"
'		sqlSum = sqlSum & ", sum((extTenMeachulPrice - extReducedPrice) * extitemno) as bigoSum "
'		sqlSum = sqlSum & ", isnull(sum((extTenMeachulPrice - reducedPrice) * extitemno), 0) as bigoBeasongSum "
		sqlSum = sqlSum & ", isnull(sum(d.reducedPrice*d.itemno), 0) as tenReducedPrice "
		sqlSum = sqlSum & " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
		sqlSum = sqlSum + " LEFT JOIN db_order.dbo.tbl_order_detail d WITH(NOLOCK) on j.OrgOrderserial = d.orderserial and j.itemid = d.itemid and j.itemoption = d.itemoption "
		sqlSum = sqlSum + " where 1=1"
		sqlSum = sqlSum + addSqlStr
		sqlSum = sqlSum + " GROUP BY LEFT(extMeachulDate, 7), isNull(convert(varchar(7),d.jungsanfixdate,121), '') "
		sqlSum = sqlSum + " ORDER BY 1, 2 DESC "

	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlSum,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem
					FItemList(i).FextMeachulMonth	= rsget("extMeachulMonth")
					FItemList(i).FjungsanfixMonth	= rsget("jungsanfixMonth")
					FItemList(i).FRowNo 			= rsget("cnt")
					FItemList(i).FSumItemNo 		= rsget("itemno")
					FItemList(i).FSumitemcost 		= rsget("itemcost")
					FItemList(i).FSumMeachulPrice 	= rsget("MeachulPrice")
					FItemList(i).FSumReducedPrice	= rsget("ReducedPrice")
					FItemList(i).FSumOwnCouponPrice = rsget("OwnCouponPrice")
					FItemList(i).FSumTenCouponPrice = rsget("TenCouponPrice")
					FItemList(i).FSumCommPrice 		= rsget("CommPrice")
					FItemList(i).FSumJungsanPrice 	= rsget("JungsanPrice")
					FItemList(i).FMiMapTTLCnt		= rsget("MiMapTTLCnt")
					FItemList(i).FSumJungsanPrice_ETC = rsget("JungsanPrice_ETC")
					FItemList(i).FSumtenReducedPrice = rsget("tenReducedPrice")
'					FItemList(i).FBigoSum 			= rsget("bigoSum")
'					FItemList(i).FBigoBeasongSum 	= rsget("bigoBeasongSum")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end function

	public function GetExtJungsan()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

		 if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		        addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		 elseif (FRectSellSite<>"") then
			    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		 end if

		if (FRectJungsanType <> "") then
			addSqlStr = addSqlStr + " and j.extJungsanType = '" + CStr(FRectJungsanType) + "' "
		end if

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField="extOrderserial") then
				addSqlStr = addSqlStr + " and (LEFT(j.extOrderserial,"&LEN(FRectSearchText)&")='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " 	or j.extOrgOrderserial='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " )"
			elseif (FRectSearchField="matchitemid") then
				addSqlStr = addSqlStr + " and j.itemid='"&FRectSearchText&"'"&vbCRLF
			else
				addSqlStr = addSqlStr + " and j." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectMiMap<>"") then
			addSqlStr = addSqlStr + " and isNULL(j.OrgOrderserial,'')=''"
		end if

		if (FRectMiMapMinus<>"") then
			addSqlStr = addSqlStr + " and j.extitemno<0 and isNULL(j.MinusOrderserial,'')=''"
		end if

		if (FRectVatYn<>"") then
			addSqlStr = addSqlStr + " and extvatyn='"&FRectVatYn&"'"
		end if

		if (FRectReturnOnly<>"") then
			addSqlStr = addSqlStr + " and extitemno<1"
		end if

		if (FRectErrexists<>"") then
			''FSumitemcost-FSumMeachulPrice-(FSumOwnCouponPrice+FSumTenCouponPrice)
			''FSumJungsanPrice-(FSumMeachulPrice-FSumCommPrice+FSumOwnCouponPrice)

			addSqlStr = addSqlStr + " and ("
			addSqlStr = addSqlStr + "	(extItemCost-extReducedPrice-extOwnCouponPrice-extTenCouponPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extTenJungsanPrice-extReducedPrice+extCommPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extReducedPrice<>extTenMeachulPrice)"
			addSqlStr = addSqlStr + " )"

		end if

		if (FRectExistsBigo<>"") then
			addSqlStr = addSqlStr + " and extTenMeachulPrice - extReducedPrice <> 0"
		end if

		if (FRectExceptItemCostZero<>"") then
			addSqlStr = addSqlStr + " and extItemCost<>0"
		end if

		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""

		'// ====================================================================
		If FRectJungsanfixdate <> "" Then
			sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
			sqlStr = sqlStr + " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
			sqlStr = sqlStr + " LEFT JOIN db_order.dbo.tbl_order_detail d WITH(NOLOCK) on j.OrgOrderserial = d.orderserial and j.itemid = d.itemid and j.itemoption = d.itemoption "
			sqlStr = sqlStr + " where 1=1"
			Select Case FRectJungsanfixdate
				Case "A"
					sqlStr = sqlStr + " and d.jungsanfixdate between '" & CStr(FRectStartDate) & "' AND '"& CStr(FRectEndDate) &"' "
				Case "B"
					sqlStr = sqlStr + " and d.jungsanfixdate < '" & CStr(FRectStartDate) & "' "
				Case "C"
					sqlStr = sqlStr + " and d.jungsanfixdate is null "
			End Select
			sqlStr = sqlStr + addSqlStr
		Else
			sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
			sqlStr = sqlStr + " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
			sqlStr = sqlStr + " where 1=1"
			sqlStr = sqlStr + addSqlStr
		End If

		' response.write sqlstr & "<Br>"
    	rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		Select Case FRectJungsanfixdate
			Case "A"
				addSqlStr = addSqlStr + " and d.jungsanfixdate between '" & CStr(FRectStartDate) & "' AND '"& CStr(FRectEndDate) &"' "
			Case "B"
				addSqlStr = addSqlStr + " and d.jungsanfixdate < '" & CStr(FRectStartDate) & "' "
			Case "C"
				addSqlStr = addSqlStr + " and d.jungsanfixdate is null "
		End Select

		'// ====================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " j.* "
		sqlStr = sqlStr + " , convert(varchar(10),d.beasongdate,121) as beasongdate, convert(varchar(10),d.[dlvfinishdt],121) as dlvfinishdt, d.jungsanfixdate, isnull(d.reducedPrice, 0) as reducedPrice "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
		sqlStr = sqlStr + " LEFT JOIN db_order.dbo.tbl_order_detail d WITH(NOLOCK) on j.OrgOrderserial = d.orderserial and j.itemid = d.itemid and j.itemoption = d.itemoption "
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by j.extMeachulDate desc, j.sellsite, j.extOrderserial, j.extOrderserSeq "

		' response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if
				FItemList(i).Fbeasongdate = rsget("beasongdate")
				FItemList(i).Fdlvfinishdt = rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate = rsget("jungsanfixdate")
				FItemList(i).FReducedPrice = rsget("reducedPrice")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function GetExtJungsanExcelDown()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

		if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		      addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		elseif (FRectSellSite<>"") then
		    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		end if

		if (FRectJungsanType <> "") then
			addSqlStr = addSqlStr + " and j.extJungsanType = '" + CStr(FRectJungsanType) + "' "
		end if

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField="extOrderserial") then
				addSqlStr = addSqlStr + " and (LEFT(j.extOrderserial,"&LEN(FRectSearchText)&")='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " 	or j.extOrgOrderserial='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " )"
			else
				addSqlStr = addSqlStr + " and j." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectMiMap<>"") then
			addSqlStr = addSqlStr + " and isNULL(j.OrgOrderserial,'')=''"
		end if

		if (FRectVatYn<>"") then
			addSqlStr = addSqlStr + " and extvatyn='"&FRectVatYn&"'"
		end if

		if (FRectReturnOnly<>"") then
			addSqlStr = addSqlStr + " and extitemno<1"
		end if

		if (FRectErrexists<>"") then
			''FSumitemcost-FSumMeachulPrice-(FSumOwnCouponPrice+FSumTenCouponPrice)
			''FSumJungsanPrice-(FSumMeachulPrice-FSumCommPrice+FSumOwnCouponPrice)

			addSqlStr = addSqlStr + " and ("
			addSqlStr = addSqlStr + "	(extItemCost-extReducedPrice-extOwnCouponPrice-extTenCouponPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extTenJungsanPrice-extReducedPrice+extCommPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extReducedPrice<>extTenMeachulPrice)"
			addSqlStr = addSqlStr + " )"

		end if

		if (FRectExceptItemCostZero<>"") then
			addSqlStr = addSqlStr + " and extItemCost<>0"
		end if

		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""

		'// ====================================================================
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
	    sqlStr = sqlStr + " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

		''response.write sqlstr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
			FResultCount = FTotalCount
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if FTotalCount < 1 then
			FResultCount = 0
			exit function
		end if


		'// ====================================================================
		'' 최대 2만건만 하자.
		sqlStr = "select top 20000 j.sellsite,extMeachulDate,extOrderserial,extOrderserSeq,extOrgOrderserial "
		sqlStr = sqlStr + " ,extItemNo,extItemCost,extOwnCouponPrice"
		sqlStr = sqlStr + " ,extTenCouponPrice,extReducedPrice,extTenMeachulPrice,extCommPrice"
		sqlStr = sqlStr + " ,extTenJungsanPrice,OrgOrderserial"
		sqlStr = sqlStr + " ,itemid,itemoption,siteNo,extJungsanDate,extJungsanType"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by j.extMeachulDate desc, j.sellsite, j.extOrderserial, j.extOrderserSeq "

		''response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		if  not rsget.EOF  then
			GetExtJungsanExcelDown = rsget.getRows
		end if

		rsget.close()

		' redim preserve FItemList(FResultCount)
		' i=0
		' if  not rsget.EOF  then
		' 	rsget.absolutepage = 1
		' 	do until rsget.eof
		' 		set FItemList(i) = new CExtJungsanItem

		' 		FItemList(i).Fsellsite				= rsget("sellsite")
		' 		FItemList(i).FextOrderserial		= rsget("extOrderserial")
		' 		FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
		' 		FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
		' 		FItemList(i).FextItemNo				= rsget("extItemNo")
		' 		FItemList(i).FextItemCost			= rsget("extItemCost")
		' 		FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
		' 		FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
		' 		FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
		' 		FItemList(i).FextJungsanType		= rsget("extJungsanType")
		' 		FItemList(i).FextCommPrice			= rsget("extCommPrice")
		' 		FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
		' 		FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
		' 		FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
		' 		FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
		' 		FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
		' 		FItemList(i).Fitemid				= rsget("itemid")
		' 		FItemList(i).Fitemoption			= rsget("itemoption")
		' 		FItemList(i).FsiteNo				= rsget("siteNo")

		' 		i=i+1
		' 		rsget.moveNext
		' 	loop
		' end if
    end function

	public function GetExtJungsanFixedErrDetailListByOrder()
		dim i, sqlStr

		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_FIXED_DIFF_GETLIST_ByOrder] '"&FRectOrderserial&"',"&CHKIIF(FRectItemID="","NULL",FRectItemID)&",'"&FRectItemOption&"'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly


		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanFixedErrItem

				FItemList(i).Fyyyymm					= db3_rsget("yyyymm")
				FItemList(i).Fsellsite					= db3_rsget("sellsite")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).Fitemid					= db3_rsget("itemid")
				FItemList(i).Fitemoption				= db3_rsget("itemoption")
				FItemList(i).Fitemnosum					= db3_rsget("itemnosum")
				FItemList(i).Freducedsum				= db3_rsget("reducedsum")
				FItemList(i).Fbuycashsum				= db3_rsget("buycashsum")
				FItemList(i).FextItemNoSum				= db3_rsget("extItemNoSum")
				FItemList(i).FextreducedpriceSum		= db3_rsget("extreducedpriceSum")
				FItemList(i).FextTenJungsanPriceSum		= db3_rsget("extTenJungsanPriceSum")
				FItemList(i).Fupddt				= db3_rsget("upddt")
				FItemList(i).FErrAsignMonth		= db3_rsget("ErrAsignMonth")
				FItemList(i).FErrAsignSum		= db3_rsget("ErrAsignSum")
				FItemList(i).Fdiffthis			= db3_rsget("diffthis")

				' FItemList(i).FacctErrsum		= db3_rsget("acctErrsum")
				' FItemList(i).FaccAsgnErrSum		= db3_rsget("accAsgnErrSum")
				' FItemList(i).FaccTTLErrSum		= db3_rsget("accTTLErrSum")

				FItemList(i).FOutMallOrderSerial = db3_rsget("OutMallOrderSerial")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanFixedErrDetailList()
		dim i, sqlStr

		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_FIXED_DIFF_GETCNT] '"&FRectSellSite&"','"&FRectYYYYMM&"','"&FRectJungsanType&"',"&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&","&CHKIIF(FRectAccerrtype<>"",FRectAccerrtype,"NULL")

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
		if NOT db3_rsget.Eof then
			FTotalCount = db3_rsget("cnt")
			FdiffnoSum  = db3_rsget("diffnoSum")
			FdiffsumSum = db3_rsget("diffsumSum")
			FErrAsignSum = db3_rsget("ErrAsignSum")
		end if
		db3_rsget.Close

		if FTotalCount < 1 then
			FResultCount = 0
			exit function
		end if


		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_FIXED_DIFF_GETLIST] "&FCurrPage&","&FPageSize&",'"&FRectSellSite&"','"&FRectYYYYMM&"','"&FRectJungsanType&"',"&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&","&CHKIIF(FRectAccerrtype<>"",FRectAccerrtype,"NULL")
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanFixedErrItem

				FItemList(i).Fyyyymm					= db3_rsget("yyyymm")
				FItemList(i).Fsellsite					= db3_rsget("sellsite")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).Fitemid					= db3_rsget("itemid")
				FItemList(i).Fitemoption				= db3_rsget("itemoption")
				FItemList(i).Fitemnosum					= db3_rsget("itemnosum")
				FItemList(i).Freducedsum				= db3_rsget("reducedsum")
				FItemList(i).Fbuycashsum				= db3_rsget("buycashsum")
				FItemList(i).FextItemNoSum				= db3_rsget("extItemNoSum")
				FItemList(i).FextreducedpriceSum		= db3_rsget("extreducedpriceSum")
				FItemList(i).FextTenJungsanPriceSum		= db3_rsget("extTenJungsanPriceSum")
				FItemList(i).Fupddt				= db3_rsget("upddt")
				FItemList(i).FErrAsignMonth		= db3_rsget("ErrAsignMonth")
				FItemList(i).FErrAsignSum		= db3_rsget("ErrAsignSum")
				FItemList(i).Fdiffthis			= db3_rsget("diffthis")

				FItemList(i).FaccErrNoSum		= db3_rsget("accErrNoSum")
				FItemList(i).FacctErrsum		= db3_rsget("acctErrsum")
				FItemList(i).FaccAsgnErrSum		= db3_rsget("accAsgnErrSum")
				FItemList(i).FaccTTLErrSum		= db3_rsget("accTTLErrSum")

				FItemList(i).FOutMallOrderSerial= db3_rsget("OutMallOrderSerial")
				FItemList(i).Fcomment			= db3_rsget("comment")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanErrDetailList()
		dim i, sqlStr

		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_GETCNT] '"&FRectSellSite&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectJungsanType&"',"&CHKIIF(FonlyErrNoExists<>"",1,0)&","&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&""
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
		if NOT db3_rsget.Eof then
			FTotalCount = db3_rsget("cnt")
			FdiffnoSum  = db3_rsget("diffnoSum")
			FdiffsumSum = db3_rsget("diffsumSum")
		end if
		db3_rsget.Close

		if FTotalCount < 1 then
			FResultCount = 0
			exit function
		end if


		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_GETLIST] "&FCurrPage&","&FPageSize&",'"&FRectSellSite&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectJungsanType&"',"&CHKIIF(FonlyErrNoExists<>"",1,0)&","&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&""
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanErrItem

				FItemList(i).Fsitename		= db3_rsget("sitename")
				FItemList(i).Fyyyymmdd		= db3_rsget("yyyymmdd")
				FItemList(i).Fauthcode		= db3_rsget("authcode")
				FItemList(i).Foorderserial	= db3_rsget("oorderserial")
				FItemList(i).Fitemid		= db3_rsget("itemid")
				FItemList(i).Fitemoption	= db3_rsget("itemoption")
				FItemList(i).Fdiffno		= db3_rsget("diffno")
				FItemList(i).Fdiffsum		= db3_rsget("diffsum")

				FItemList(i).Fjumundiv			= db3_rsget("jumundiv")
				FItemList(i).Flinkorderserial	= db3_rsget("linkorderserial")
				FItemList(i).FregDt				= db3_rsget("regDt")
				FItemList(i).FupdDt				= db3_rsget("updDt")

				FItemList(i).Fcomment			= db3_rsget("comment")
				FItemList(i).Ferrortype			= db3_rsget("errortype")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanDiff()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_MonthList] '" & FRectSellSite & "', '" & FRectDiffType & "' "

        db3_rsget.CursorLocation = adUseClient
    	db3_rsget.Open  sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount  = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanDiffItem
				FItemList(i).Fsitename       = db3_rsget("sitename")
				FItemList(i).Fyyyymm         = db3_rsget("yyyymm")
				FItemList(i).FTMeachulItem   = db3_rsget("TMeachulItem")
				FItemList(i).FTMeachulDLV    = db3_rsget("TMeachulDLV")
				FItemList(i).FTbuycashItem   = db3_rsget("TbuycashItem")
				FItemList(i).FTbuycashDLV    = db3_rsget("TbuycashDLV")

				FItemList(i).FXMeachulItem   = db3_rsget("XMeachulItem")
				FItemList(i).FXMeachulDLV    = db3_rsget("XMeachulDLV")
				FItemList(i).FXJungsanItem   = db3_rsget("XJungsanItem")
				FItemList(i).FXJungsanDLV    = db3_rsget("XJungsanDLV")
				FItemList(i).FregDt          = db3_rsget("regDt")
				FItemList(i).FupdDt      	 = db3_rsget("updDt")

				FItemList(i).FmonthItemDiff  = db3_rsget("monthItemDiff")
				FItemList(i).FmonthdlvDiff   = db3_rsget("monthdlvDiff")
				FItemList(i).FdiffITEMsum    = db3_rsget("diffITEMsum")
				FItemList(i).FdiffDlvsum     = db3_rsget("diffDlvsum")

				FItemList(i).FmonthItemDiffMapErr	= db3_rsget("monthItemDiffMapErr")
				FItemList(i).FmonthdlvDiffTMapErr	= db3_rsget("monthdlvDiffTMapErr")

				FItemList(i).FmonthItemDiffNoExists	= db3_rsget("monthItemDiffNoExists")
				FItemList(i).FmonthdlvDiffTNoExists	= db3_rsget("monthdlvDiffTNoExists")

				FItemList(i).FMonthDiffSum = db3_rsget("MonthDiffSum")
				FItemList(i).FMonthErrAsignSum  = db3_rsget("MonthErrAsignSum")
				FItemList(i).FMonthnotAssignErr = db3_rsget("MonthnotAssignErr")
				FItemList(i).FMonthErrAsignItemSum = db3_rsget("MonthErrAsignItemSum")

				FItemList(i).FMonthErrAsignItemSumReqCheck = db3_rsget("MonthErrAsignItemSumReqCheck")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function


	public function GetExtJungsanDiff_OLD()
	    dim i, sqlStr, addSqlStr

		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_GetExtSiteMeachulDiff_Count] '" & FRectYYYYMM & "', '" & FRectSellSite & "', '" & FRectDiffType & "' "

		''response.write sqlstr & "<Br>"
    	db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

		if FTotalCount<1 then exit function

		'// ====================================================================
		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_GetExtSiteMeachulDiff_List] '" & FRectYYYYMM & "', '" & FRectSellSite & "', '" & FRectDiffType & "', '" & FPageSize & "', '" & FCurrPage & "' "

        db3_rsget.CursorLocation = adUseClient
    	db3_rsget.CursorType = adOpenStatic
    	db3_rsget.LockType = adLockOptimistic

		db3_rsget.Open sqlStr,db3_dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if
		'
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		'
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			''rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanDiffItem_OLD

				FItemList(i).Fyyyymm				= db3_rsget("yyyymm")
				FItemList(i).Fsellsite				= db3_rsget("sitename")

				FItemList(i).Forderserial			= db3_rsget("orderserial")
				FItemList(i).FMeachulPriceSUM		= db3_rsget("MeachulPriceSUM")
				FItemList(i).FextMeachulPriceSUM	= db3_rsget("extMeachulPriceSUM")
				FItemList(i).FMeachulPriceSUM1		= db3_rsget("MeachulPriceSUM1")
				FItemList(i).FextMeachulPriceSUM1	= db3_rsget("extMeachulPriceSUM1")
				FItemList(i).FMeachulPriceSUM2		= db3_rsget("MeachulPriceSUM2")
				FItemList(i).FextMeachulPriceSUM2	= db3_rsget("extMeachulPriceSUM2")
				FItemList(i).FMeachulPriceSUM3		= db3_rsget("MeachulPriceSUM3")
				FItemList(i).FextMeachulPriceSUM3	= db3_rsget("extMeachulPriceSUM3")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	public function GetExtJungsanStatistic()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		If FRectIpkumdateChk = "Y" Then
			if (FRectAStartDate <> "") then
				addSqlStr = addSqlStr + " and m.ipkumdate >= '" + CStr(FRectAStartDate) + "' "
			end if

			if (FRectAEndDate <> "") then
				addSqlStr = addSqlStr + " and m.ipkumdate < '" + CStr(FRectAEndDate) + "' "
			end if
		End If

		if(FRectSellSite <> "") then
		    if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		        addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		    else
			    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		    end if
		end if

		'// ====================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "

		if (FRectGroupGubun = "sellsite") then
			sqlStr = sqlStr + " 	j.sellsite "
		else
			sqlStr = sqlStr + " 	j.extMeachulDate "
		end if

		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extTenMeachulPrice*j.extItemNo else 0 end) as totExtTenMeachulPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extitemCost*j.extItemNo else 0 end) as totExtitemCostProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extReducedPrice*j.extItemNo else 0 end) as totExtReducedPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extOwnCouponPrice*j.extItemNo else 0 end) as totExtOwnCouponPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extTenCouponPrice*j.extItemNo else 0 end) as totExtTenCouponPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extCommPrice*j.extItemNo else 0 end) as totExtCommPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extTenJungsanPrice*j.extItemNo else 0 end) as totExtTenJungsanPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extTenMeachulPrice*j.extItemNo else 0 end) as totExtTenMeachulPriceDeliver "

		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extitemCost*j.extItemNo else 0 end) as totExtitemCostDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extReducedPrice*j.extItemNo else 0 end) as totExtReducedPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extOwnCouponPrice*j.extItemNo else 0 end) as totExtOwnCouponPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extTenCouponPrice*j.extItemNo else 0 end) as totExtTenCouponPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extCommPrice*j.extItemNo else 0 end) as totExtCommPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extTenJungsanPrice*j.extItemNo else 0 end) as totExtTenJungsanPriceDeliver "

		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') not in ('C', 'D') then j.extTenMeachulPrice*j.extItemNo else 0 end) as totExtTenMeachulPriceEtc "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') not in ('C', 'D') then j.extCommPrice*j.extItemNo else 0 end) as totExtCommPriceEtc "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') not in ('C', 'D') then j.extTenJungsanPrice*j.extItemNo else 0 end) as totExtTenJungsanPriceEtc "
		sqlStr = sqlStr + " 	, sum(j.extTenMeachulPrice*j.extItemNo) as totExtTenMeachulPrice "
		sqlStr = sqlStr + " 	, sum(j.extCommPrice*j.extItemNo) as totExtCommPrice "
		sqlStr = sqlStr + " 	, sum(j.extTenJungsanPrice*j.extItemNo) as totExtTenJungsanPrice "

		sqlStr = sqlStr + " 	, sum(j.extitemCost*j.extItemNo) as totExtitemCost "
		sqlStr = sqlStr + " 	, sum(j.extReducedPrice*j.extItemNo) as totExtReducedPrice "
		sqlStr = sqlStr + " 	, sum(j.extOwnCouponPrice*j.extItemNo) as totExtOwnCouponPrice "
		sqlStr = sqlStr + " 	, sum(j.extTenCouponPrice*j.extItemNo) as totExtTenCouponPrice "

		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(extJungsanDate,'')<>'' THEN 0 ELSE 1 END) as MiMapp, count(*) as RowCnt "
		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(extJungsanDate,'')='' and IsNull(j.extJungsanType, 'C') = 'C'  THEN 1 ELSE 0 END) as MiMapp_C"
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then 1 else 0 END) as RowCnt_C "

		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(OrgOrderserial,'')='' THEN 1 ELSE 0 END) as MiMappOrder"
		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(OrgOrderserial,'')='' and IsNull(j.extJungsanType, 'C') = 'C'  THEN 1 ELSE 0 END) as MiMappOrder_C"

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
		If FRectIpkumdateChk = "Y" Then
			sqlStr = sqlStr + " LEFT JOIN db_order.dbo.tbl_order_master m WITH(NOLOCK) on j.OrgOrderserial = m.orderserial "
		End If
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "

		sqlStr = sqlStr + addSqlStr

		if (FRectGroupGubun = "sellsite") then
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	j.sellsite "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	j.sellsite "
		else
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	j.extMeachulDate "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	j.extMeachulDate desc "
		end if

		' response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanStatisticItem

				if (FRectGroupGubun = "sellsite") then
					FItemList(i).Fsellsite				= rsget("sellsite")
				else
					FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				end if

				FItemList(i).FtotExtTenMeachulPriceProduct	= rsget("totExtTenMeachulPriceProduct")
				FItemList(i).FtotExtCommPriceProduct			= rsget("totExtCommPriceProduct")
				FItemList(i).FtotExtTenJungsanPriceProduct	= rsget("totExtTenJungsanPriceProduct")

				FItemList(i).FtotExtTenMeachulPriceDeliver	= rsget("totExtTenMeachulPriceDeliver")
				FItemList(i).FtotExtCommPriceDeliver			= rsget("totExtCommPriceDeliver")
				FItemList(i).FtotExtTenJungsanPriceDeliver	= rsget("totExtTenJungsanPriceDeliver")

				FItemList(i).FtotExtTenMeachulPriceEtc			= rsget("totExtTenMeachulPriceEtc")
				FItemList(i).FtotExtCommPriceEtc				= rsget("totExtCommPriceEtc")
				FItemList(i).FtotExtTenJungsanPriceEtc			= rsget("totExtTenJungsanPriceEtc")

				FItemList(i).FtotExtTenMeachulPrice			= rsget("totExtTenMeachulPrice")
				FItemList(i).FtotExtCommPrice					= rsget("totExtCommPrice")
				FItemList(i).FtotExtTenJungsanPrice			= rsget("totExtTenJungsanPrice")
				FItemList(i).FextMiMapping						= rsget("MiMapp")
				FItemList(i).FextRowCount						= rsget("RowCnt")

                FItemList(i).FextMiMapping_C					= rsget("MiMapp_C")
				FItemList(i).FextRowCount_C						= rsget("RowCnt_C")

				FItemList(i).FMiMappOrder					= rsget("MiMappOrder")
				FItemList(i).FMiMappOrder_C						= rsget("MiMappOrder_C")

				''2018/06/26
				FItemList(i).FtotExtitemCostProduct	    = rsget("totExtitemCostProduct")
				FItemList(i).FtotExtReducedPriceProduct	    = rsget("totExtReducedPriceProduct")
				FItemList(i).FtotExtOwnCouponPriceProduct	= rsget("totExtOwnCouponPriceProduct")
				FItemList(i).FtotExtTenCouponPriceProduct	= rsget("totExtTenCouponPriceProduct")

				FItemList(i).FtotExtitemCostDeliver	    = rsget("totExtitemCostDeliver")
				FItemList(i).FtotExtReducedPriceDeliver	    = rsget("totExtReducedPriceDeliver")
				FItemList(i).FtotExtOwnCouponPriceDeliver	= rsget("totExtOwnCouponPriceDeliver")
				FItemList(i).FtotExtTenCouponPriceDeliver	= rsget("totExtTenCouponPriceDeliver")

				FItemList(i).FtotExtitemCost            = rsget("totExtitemCost")
				FItemList(i).FtotExtReducedPrice            = rsget("totExtReducedPrice")
                FItemList(i).FtotExtOwnCouponPrice          = rsget("totExtOwnCouponPrice")
                FItemList(i).FtotExtTenCouponPrice          = rsget("totExtTenCouponPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0

		FdiffnoSum = 0
		FdiffsumSum = 0
		FErrAsignSum = 0
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


Function getCsOrgOrderserila(v)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 OutMallOrderSerial "
	strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPCS "
	strSql = strSql & " WHERE orgOutMallOrderSerial = '"& v &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getCsOrgOrderserila = rsget("OutMallOrderSerial")
	Else
		getCsOrgOrderserila = ""
	End If
	rsget.Close
End Function
%>

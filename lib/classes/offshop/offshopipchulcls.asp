<%
'####################################################
' Description :  오프라인 개별 입출고 클래스
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################

class CBarCodeItem
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fbrand

	public Forgsellprice
	public Fitemprice
	public Fitemtype
	public Fitemno

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CBrandShopInfoItem
	public Fchargeid
	public FChargeName
	public FSocNo
	public FCeoName
	public FAddress
	public FSocName
	public FManagerName
	public FManagerHp
	public FUptae
	public FUpjong
	public FRectChargeId

	'//common/offshop/pop_ipgosheet.asp
	public function GetBrandShopInFo()
		dim sqlStr
		sqlStr = "select top 1 p.company_name, "
		sqlStr = sqlStr + " (p.address + ' ' + p.manager_address) as socaddr, p.manager_name, p.company_no,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, p.ceoname, p.manager_hp"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
		''sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		''sqlStr = sqlStr + " on p.id=c.userid"
		sqlStr = sqlStr + " where id='" + FRectChargeId + "'"

		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1

		if Not rsget.Eof then
			Fchargeid  = FRectChargeId
			FChargeName = db2html(rsget("company_name"))
			FSocNo     = db2html(rsget("company_no"))
			FCeoName   = db2html(rsget("ceoname"))
			FManagerName = db2html(rsget("manager_name"))
			FAddress   = db2html(rsget("socaddr"))
			FSocName   = db2html(rsget("company_name"))
			FUptae = db2html(rsget("company_uptae"))
			FUpjong = db2html(rsget("company_upjong"))
			FManagerHp = db2html(rsget("manager_hp"))
		end if

		rsget.Close
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CIpChulItem
	public Fidx
	public Fchargeid
	public Fshopid
	public Fdivcode
	public Ftotalsellcash
	public Ftotalsuplycash
	public Ftotalshopbuyprice
	public fcomment
	public Fvatcode
	public Fexecdt
	public Fregdate
	public FRegUserid
	public Fstatecd
	public Fdeleteyn
	public Flinkidx
	public FScheduleDt
	public fsendsms
	public Fshopname
	public Fsongjangdiv
	public Fsongjangname
	public Fsongjangno
	public Ffindurl
	public fshopdiv
	public Fshopconfirmdate
	public Fupcheconfirmdate
	public Fshopconfirmuserid
	public Fupcheconfirmuserid
    public Fbaljuconfirmdate
    public fipchulmoveidx
    public FComm_cd
    public FComm_cd_jungsan
    public FisbaljuExists

	public function IsPriceEditEnabled()
		IsPriceEditEnabled = (session("ssBctDiv")<=9)
	end function

    public function IsDispReqNo()
        IsDispReqNo = (FisbaljuExists="Y")
    end function


    '' 입고요청 상태인지
    public function IsRequireConfirm()
        IsRequireConfirm = (Fstatecd = -2) and (FisbaljuExists="Y")
    end function

    '' 업체 확인 진행
    public function UpcheConfirmProcess()
        dim sqlStr, resultRow
        sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" & VbCrlf
        sqlStr = sqlStr + " set baljuconfirmdate=getdate()" & VbCrlf
        sqlStr = sqlStr + " ,statecd=-1" & VbCrlf
        sqlStr = sqlStr + " where idx=" & CStr(Fidx) & VbCrlf
        sqlStr = sqlStr + " and statecd=-2"

        dbget.Execute sqlStr, resultRow

        if resultRow>0 then
            Fstatecd = -1
        end if
    end function

	public function IsWaitState()
	    ''' 입고 대기.
		IsWaitState = ((Fstatecd = 0) or (Fstatecd = -1))
	end function

	public function getInputAreaStr()
		if (IsUpcheInput) then
			getInputAreaStr = "업체"
		else
			getInputAreaStr = "매장"
		end if
	end function

	public function getInputAreaColor()
		if (IsUpcheInput) then
			getInputAreaColor = "#000000"
		else
			getInputAreaColor = "#3333CC"
		end if
	end function

    public function IsEditEnabled()
		IsEditEnabled = false

		''입고대기/입고요청 상태가 아니면 삭제불가
		if (Fstatecd>0) then exit function

		'' 업체인경우
		if (session("ssBctDiv") = "9999") then
		    '' statecd==-1 : 입고요청 확인, statecd==0 : 입고대기(발송)
			if (IsUpcheInput) or ((Not IsUpcheInput) and (Fstatecd<0)) then IsEditEnabled=true
		'' 매장인경우
		elseif ((session("ssBctDiv") = "501") or (session("ssBctDiv") = "101") or (session("ssBctDiv") = "111") or (session("ssBctDiv") = "112") or (session("ssBctDiv") = "502") or (session("ssBctDiv") = "503")) then
			if Not (IsUpcheInput) then IsEditEnabled=true

		    if (FisbaljuExists="Y") and (Fstatecd=0) then IsEditEnabled=false
		else
			IsEditEnabled = true
		end if
	end function

	public function IsAvailDelete()
		IsAvailDelete = false

		''입고대기/입고요청 상태가 아니면 삭제불가
		if (Fstatecd>0) then exit function

		'' 업체인경우
		if (session("ssBctDiv") = "9999") then
			if (LCase(FRegUserid)=LCase(session("ssBctId"))) then IsAvailDelete=true
		'' 매장인경우
		elseif ((session("ssBctDiv") = "501") or (session("ssBctDiv") = "101") or (session("ssBctDiv") = "111") or (session("ssBctDiv") = "112") or (session("ssBctDiv") = "502") or (session("ssBctDiv") = "503")) then
			if (LCase(FRegUserid)=LCase(session("ssBctId"))) then IsAvailDelete=true

		    if (FisbaljuExists="Y") and (Fstatecd=0) then IsAvailDelete=false
		else
			IsAvailDelete = false
		end if


	end function

	public function IsUpcheInput()
		'' 기존 FRegUserid 가 널인 경우도 업체에서 입력한값으로 봄
		IsUpcheInput = (Lcase(Fchargeid)=LCase(FRegUserid)) or IsNULL(FRegUserid)
	end function

	public function IsUpcheStateChangeEnabled()
		IsUpcheStateChangeEnabled = false

		if (IsUpcheInput) then
			if (Fstatecd=7) then
				IsUpcheStateChangeEnabled = true
			end if
		else
			if (Fstatecd=0) and (Not FisbaljuExists="Y") then
				IsUpcheStateChangeEnabled = true
			end if
		end if
	end function


	public function IsShopStateChangeEnabled()
		IsShopStateChangeEnabled = false

		if (IsUpcheInput) then
			if (Fstatecd<1) then
				IsShopStateChangeEnabled = true
			end if
		else
			if (Fstatecd=7) or (Fstatecd=-1) or (Fstatecd=0) then
				IsShopStateChangeEnabled = true
			end if
		end if
	end function


	public function getMinusColor(value)
		getMinusColor = "#000000"
		if IsNull(value) then Exit Function

		if (value<0) then
			getMinusColor = "#FF3333"
		else
			getMinusColor = "#000000"
		end if
	end function

	public function getStateName()
		getStateName = ""
		if IsNull(Fstatecd) then Exit Function

		if Fstatecd=0 then
			getStateName = "입고대기"
		elseif Fstatecd=7 then
			if IsUpcheInput then
				getStateName = "매장 입고확인"
			else
				getStateName = "업체 입고확인"
			end if
		elseif Fstatecd=8 then
			getStateName = "입고확정"
	    elseif Fstatecd=-1 then
			getStateName = "입고요청확인"
		elseif Fstatecd=-2 then
			getStateName = "입고요청"
		elseif Fstatecd=-5 then
			getStateName = "임시저장"
		end if

	end function

	public function getStateColor()
		getStateColor = "#000000"
		if IsNull(Fstatecd) then Exit Function

		if Fstatecd=0 then
			getStateColor = "#000000"
		elseif Fstatecd=7 then
			getStateColor = "#FF3333"
		elseif Fstatecd=8 then
			getStateColor = "#3333FF"
		elseif Fstatecd=-1 then
			getStateColor = "#33AA33"
		elseif Fstatecd=-2 then
			getStateColor = "#FF33FF"
		end if
	end function

    public function GetContractColor()
        if (FComm_cd="B011") then
        	GetContractColor = "#000000"
        elseif (FComm_cd="B012") then
        	GetContractColor = "#0000FF"
        elseif (FComm_cd="B021") then
        	GetContractColor = "#FF0000"
        elseif (FComm_cd="B022") then
        	GetContractColor = "#FF0000"
        elseif (FComm_cd="B031") then
        	GetContractColor = "#000000"
        elseif (FComm_cd="B032") then
        	GetContractColor = "#000000"
        elseif (FComm_cd="B999") then
        	GetContractColor = "#000000"
        else
            GetContractColor = "#000000"
        end if

    end function

	'//현재 정산구분
    public function GetContractName()
        if (FComm_cd="B011") then
        	GetContractName = "위탁판매"
        elseif (FComm_cd="B012") then
        	GetContractName = "업체위탁"
        elseif (FComm_cd="B021") then
        	GetContractName = "오프매입"
        elseif (FComm_cd="B022") then
        	GetContractName = "매장매입"
        elseif (FComm_cd="B031") then
        	GetContractName = "출고매입"
        elseif (FComm_cd="B032") then
        	GetContractName = "센터매입"
        elseif (FComm_cd="B999") then
        	GetContractName = "기타보정"
        else
            GetContractName = FComm_cd
        end if
    end function

	'//주문당시 정산구분
    public function GetContractName_jungsan()
        if (FComm_cd_jungsan="B011") then
        	GetContractName_jungsan = "위탁판매"
        elseif (FComm_cd_jungsan="B012") then
        	GetContractName_jungsan = "업체위탁"
        elseif (FComm_cd_jungsan="B021") then
        	GetContractName_jungsan = "오프매입"
        elseif (FComm_cd_jungsan="B022") then
        	GetContractName_jungsan = "매장매입"
        elseif (FComm_cd_jungsan="B031") then
        	GetContractName_jungsan = "출고매입"
        elseif (FComm_cd_jungsan="B032") then
        	GetContractName_jungsan = "센터매입"
        elseif (FComm_cd_jungsan="B999") then
        	GetContractName_jungsan = "기타보정"
        else
            GetContractName_jungsan = FComm_cd_jungsan
        end if
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CIpChulDetailItem
	public Fidx
	public Fmasteridx
	public Fitemgubun
	public Fshopitemid
	public Fitemoption
	public Fdesignerid
	public Fsellcash
	public Fsuplycash
	public Fshopbuyprice
	public Fitemno
	public Freqno

	public Fdeleteyn
	public Flinkidx

	public FItemName
	public FItemOptionName
    public FCurrMakerid
    public FrackcodeByOption

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fshopitemid) & Fitemoption
		if (Fshopitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fshopitemid)) + CStr(Fitemoption)
    	end if
	end function

	public function getStateName()
		getStateName = ""
		if IsNull(Fstatecd) then Exit Function

		if Fstatecd="0" then
			getStateName = "입고대기"
		elseif Fstatecd="7" then
			getStateName = "입고완료"
		end if

	end function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CIpChulDetailByShopByItemItem
    public Fidx
	public Fshopid
	public Fitemgubun
	public Fitemid
	public Fitemoption
	public Fitemno
	public Fexecdt
	public Fdesignerid
	public Fsellcash
	public Fsuplycash
	public Fitemname
	public Fitemoptionname

    public Fchargeid
    public Fcomm_cd

	public function GetBarCode()
		GetBarCode = Fitemgubun & Format00(6,Fitemid) & Fitemoption
		if (Fitemid >= 1000000) then
    		GetBarCode = CStr(Fitemgubun) + CStr(Format00(8,Fitemid)) + CStr(Fitemoption)
    	end if
	end function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CShopMaechulItem
	public Fshopid
	public Fshopname

	public Ftenout
	public Ftenreturn
	public Fupcheout
	public Fupchereturn
	public Fshopsell

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CShopIpChul
	public FItemList()
	public FOneItem
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectmoveipchulyn
	public FRectShopId
	public FRectChargeId
	public FRectStartDay
	public FRectEndDay
	public FRectItemGubun
	public FRectItemId
	public FRectItemOption
	public FRectMakerId

	public FRectIdx
	public FRectIdxArr
	public FRectNotIpgo
	public FRectDatesearchtype
	public frect_IS_Maker_Upche

	public function GetShopMaechulList
		dim sqlStr,i

        sqlStr = " select userid as shopid,shopname, isnull(A.ttlout,0) as tenout, isnull(A.ttlreturn,0) as tenreturn, isnull(B.ttlout,0) as upcheout, isnull(B.ttlreturn,0) as upchereturn, (isnull(C.totalsum,0) + isnull(D.totalsum,0)) as shopsell "
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user u "
        sqlStr = sqlStr + " left join "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + "     select socid, "
        sqlStr = sqlStr + "     sum(case  when totalsellcash<0 then totalsellcash*-1 else 0 end) as ttlout, "
        sqlStr = sqlStr + "     sum(case  when totalsellcash>0 then totalsellcash*-1 else 0 end) as ttlreturn "
        sqlStr = sqlStr + "     from [db_storage].[dbo].tbl_acount_storage_master m "
        sqlStr = sqlStr + "     where m.ipchulflag = 'S' "
        sqlStr = sqlStr + "     and m.deldt is null "
        sqlStr = sqlStr + "     and executedt >='" + CStr(FRectStartDay) + "' "
        sqlStr = sqlStr + "     and executedt < '" + CStr(FRectEndDay) + "' "
        sqlStr = sqlStr + "     and executedt is not null "
        sqlStr = sqlStr + "     group by socid "
        sqlStr = sqlStr + " ) A on A.socid = u.userid "
        sqlStr = sqlStr + " left join "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + "     select shopid, "
        sqlStr = sqlStr + "     sum(case  when totalsellcash>0 then totalsellcash else 0 end) as ttlout, "
        sqlStr = sqlStr + "     sum(case  when totalsellcash<0 then totalsellcash else 0 end) as ttlreturn "
        sqlStr = sqlStr + "     from [db_shop].[dbo].tbl_shop_ipchul_master "
        sqlStr = sqlStr + "     where chargeid<>'10x10' "
        sqlStr = sqlStr + "     and deleteyn='N' "
        sqlStr = sqlStr + "     and execdt >='" + CStr(FRectStartDay) + "' "
        sqlStr = sqlStr + "     and execdt < '" + CStr(FRectEndDay) + "' "
        sqlStr = sqlStr + "     and execdt is not null "
        sqlStr = sqlStr + "     group by shopid "
        sqlStr = sqlStr + " ) B on B.shopid = u.userid "
        sqlStr = sqlStr + " left join "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + "     select shopid, sum(totalsum) as totalsum "
        sqlStr = sqlStr + "     from [db_shop].[dbo].tbl_shopjumun_master "
        sqlStr = sqlStr + "     where shopregdate >='" + CStr(FRectStartDay) + "' "
        sqlStr = sqlStr + "     and shopregdate < '" + CStr(FRectEndDay) + "' "
        sqlStr = sqlStr + "     and cancelyn='N' "
        sqlStr = sqlStr + "     group by shopid "
        sqlStr = sqlStr + " ) C on C.shopid = u.userid "
        sqlStr = sqlStr + " left join "
        sqlStr = sqlStr + " ( "
        sqlStr = sqlStr + "     select shopid, sum(totalsum) as totalsum "
        sqlStr = sqlStr + "     from [db_shoplog].[dbo].tbl_old_shopjumun_master "
        sqlStr = sqlStr + "     where shopregdate >='" + CStr(FRectStartDay) + "' "
        sqlStr = sqlStr + "     and shopregdate < '" + CStr(FRectEndDay) + "' "
        sqlStr = sqlStr + "     and cancelyn='N' "
        sqlStr = sqlStr + "     group by shopid "
        sqlStr = sqlStr + " ) D on D.shopid = u.userid "
        sqlStr = sqlStr + " where isusing='Y' "
        sqlStr = sqlStr + " and userid<>'streetshop000' "
        sqlStr = sqlStr + " and userid<>'streetshop800' "
        sqlStr = sqlStr + " and ((A.ttlout is not null) or (A.ttlreturn is not null) or (B.ttlout is not null) or (B.ttlreturn is not null) or (C.totalsum is not null) or (D.totalsum is not null)) "
        sqlStr = sqlStr + " order by (isnull(A.ttlout,0)+isnull(A.ttlreturn,0)) desc, userid "

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CShopMaechulItem

				FItemList(i).Fshopid        = rsget("shopid")
				FItemList(i).Fshopname      = rsget("shopname")

				FItemList(i).Ftenout        = rsget("tenout")
				FItemList(i).Ftenreturn     = rsget("tenreturn")
				FItemList(i).Fupcheout      = rsget("upcheout")
				FItemList(i).Fupchereturn   = rsget("upchereturn")
				FItemList(i).Fshopsell      = rsget("shopsell")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetShopMaechulListBySuplyCash
		dim sqlStr,i

                sqlStr = " select userid as shopid,shopname, isnull(A.ttlout,0) as tenout, isnull(A.ttlreturn,0) as tenreturn, isnull(B.ttlout,0) as upcheout, isnull(B.ttlreturn,0) as upchereturn "
                sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user u "
                sqlStr = sqlStr + " left join "
                sqlStr = sqlStr + " ( "
                sqlStr = sqlStr + "     select socid, "
                sqlStr = sqlStr + "     sum(case  when totalsuplycash<0 then totalsuplycash*-1 else 0 end) as ttlout, "
                sqlStr = sqlStr + "     sum(case  when totalsuplycash>0 then totalsuplycash*-1 else 0 end) as ttlreturn "
                sqlStr = sqlStr + "     from [db_storage].[dbo].tbl_acount_storage_master m "
                sqlStr = sqlStr + "     where m.ipchulflag = 'S' "
                sqlStr = sqlStr + "     and m.deldt is null "
                sqlStr = sqlStr + "     and executedt >='" + CStr(FRectStartDay) + "' "
                sqlStr = sqlStr + "     and executedt < '" + CStr(FRectEndDay) + "' "
                sqlStr = sqlStr + "     and executedt is not null "
                sqlStr = sqlStr + "     group by socid "
                sqlStr = sqlStr + " ) A on A.socid = u.userid "
                sqlStr = sqlStr + " left join "
                sqlStr = sqlStr + " ( "
                sqlStr = sqlStr + "     select shopid, "
                sqlStr = sqlStr + "     sum(case  when totalsuplycash>0 then totalsuplycash else 0 end) as ttlout, "
                sqlStr = sqlStr + "     sum(case  when totalsuplycash<0 then totalsuplycash else 0 end) as ttlreturn "
                sqlStr = sqlStr + "     from [db_shop].[dbo].tbl_shop_ipchul_master "
                sqlStr = sqlStr + "     where chargeid<>'10x10' "
                sqlStr = sqlStr + "     and deleteyn='N' "
                sqlStr = sqlStr + "     and execdt >='" + CStr(FRectStartDay) + "' "
                sqlStr = sqlStr + "     and execdt < '" + CStr(FRectEndDay) + "' "
                sqlStr = sqlStr + "     and execdt is not null "
                sqlStr = sqlStr + "     group by shopid "
                sqlStr = sqlStr + " ) B on B.shopid = u.userid "
                sqlStr = sqlStr + " where isusing='Y' "
                sqlStr = sqlStr + " and userid<>'streetshop000' "
                sqlStr = sqlStr + " and userid<>'streetshop800' "
                sqlStr = sqlStr + " and ((A.ttlout is not null) or (A.ttlreturn is not null) or (B.ttlout is not null) or (B.ttlreturn is not null)) "
                sqlStr = sqlStr + " order by (isnull(A.ttlout,0)+isnull(A.ttlreturn,0)) desc, userid "

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CShopMaechulItem

				FItemList(i).Fshopid        = rsget("shopid")
				FItemList(i).Fshopname      = rsget("shopname")

				FItemList(i).Ftenout        = rsget("tenout")
				FItemList(i).Ftenreturn     = rsget("tenreturn")
				FItemList(i).Fupcheout      = rsget("upcheout")
				FItemList(i).Fupchereturn   = rsget("upchereturn")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function GetBaCodeListByIdxList
		dim sqlStr,i
		sqlStr = " select d.shopitemid, d.itemoption, s.shopitemname,"
		sqlStr = sqlStr + " s.shopitemoptionname, c.socname, s.orgsellprice, s.shopitemprice, d.itemgubun, d.itemno"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item s,"
		sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " where d.masteridx in (" + FRectIdxArr + ")"
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " and d.shopitemid=s.shopitemid"
		sqlStr = sqlStr + " and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " and s.makerid=c.userid"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CBarCodeItem
				FItemList(i).Fitemid        = rsget("shopitemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2Html(rsget("shopitemname"))
				FItemList(i).Fitemoptionname= db2Html(rsget("shopitemoptionname"))
				FItemList(i).Fbrand         = db2Html(rsget("socname"))
				FItemList(i).Fitemprice     = rsget("shopitemprice")
				FItemList(i).Fitemtype      = rsget("itemgubun")
				FItemList(i).Fitemno     	= rsget("itemno")

                ''소비자가 로 출력.. X 오프라인은 실판매가로 출력
                FItemList(i).FOrgSellPrice  = rsget("orgsellprice")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	'//common/offshop/pop_ipgosheet.asp
	public function GetIpChulDetail()
		dim sqlStr,i
		sqlStr = " select top 750 d.*, s.shopitemname, s.shopitemoptionname, s.makerid "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_detail d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " where d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " and d.shopitemid=s.shopitemid"
		sqlStr = sqlStr + " and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and d.masteridx=" + Cstr(FRectIdx)
		sqlStr = sqlStr + " order by  d.itemgubun desc, s.shopitemid asc, s.itemoption"
		''sqlStr = " select top 750 d.*, s.shopitemname, s.shopitemoptionname, s.makerid, os.rackcodeByOption "
		''sqlStr = sqlStr + " from "
		''sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail d "
		''sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shop_item s "
		''sqlStr = sqlStr + " 	on "
		''sqlStr = sqlStr + " 		1 = 1 "
		''sqlStr = sqlStr + " 		and d.itemgubun = s.itemgubun "
		''sqlStr = sqlStr + " 		and d.shopitemid = s.shopitemid "
		''sqlStr = sqlStr + " 		and d.itemoption = s.itemoption "
		''sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
		''sqlStr = sqlStr + " 	on "
		''sqlStr = sqlStr + " 		1 = 1 "
		''sqlStr = sqlStr + " 		and d.itemgubun = os.itemgubun "
		''sqlStr = sqlStr + " 		and d.shopitemid = os.itemid "
		''sqlStr = sqlStr + " 		and d.itemoption = os.itemoption "
		''sqlStr = sqlStr + " where "
		''sqlStr = sqlStr + " 	1 = 1 "
		''sqlStr = sqlStr + " and d.masteridx=" + Cstr(FRectIdx)
        ''sqlStr = sqlStr + " order by  d.itemgubun desc, s.shopitemid asc, s.itemoption"


		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CIpChulDetailItem
				FItemList(i).Fidx        = rsget("idx")
				FItemList(i).Fmasteridx  = rsget("masteridx")
				FItemList(i).Fitemgubun = rsget("itemgubun")
				FItemList(i).Fshopitemid = rsget("shopitemid")
				FItemList(i).Fitemoption = rsget("itemoption")
				FItemList(i).Fdesignerid = rsget("designerid")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")
				FItemList(i).Fshopbuyprice  = rsget("shopbuyprice")
				FItemList(i).Fitemno     = rsget("itemno")
				FItemList(i).Fdeleteyn   = rsget("deleteyn")
				FItemList(i).Flinkidx    = rsget("linkidx")
				FItemList(i).FItemName	 = db2html(rsget("shopitemname"))
				FItemList(i).FItemOptionName = db2html(rsget("shopitemoptionname"))

				FItemList(i).Freqno      = rsget("reqno")
				FItemList(i).FCurrMakerid = rsget("makerid")
                ''FItemList(i).FrackcodeByOption = rsget("rackcodeByOption")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	public function GetIpChulDetailByShopByItem()
		dim sqlStr,i
        'response.write "관리자 문의 요망"
        'dbget.close()	:	response.End

        sqlStr = " select top 1000 m.idx, m.shopid, d.itemgubun, d.shopitemid as itemid, d.itemoption, d.itemno, "
        sqlStr = sqlStr + " m.execdt, d.designerid, d.sellcash, d.suplycash, d.itemname, d.itemoptionname,m.chargeid,m.comm_cd "
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m, "
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "
        sqlStr = sqlStr + " and m.chargeid <> '10x10' "
        ''sqlStr = sqlStr + " and m.divcode = '006' "         '''2012/06/04 서동석 수정
        sqlStr = sqlStr + " and m.statecd >= '7' "
        sqlStr = sqlStr + " and m.deleteyn = 'N' "
        sqlStr = sqlStr + " and d.deleteyn = 'N' "

		if (FRectShopId <> "") then
		        sqlStr = sqlStr + " and m.shopid='" + Cstr(FRectShopId) + "' "
		end if
		if (FRectStartDay <> "") then
		        sqlStr = sqlStr + " and m.execdt>='" + Cstr(FRectStartDay) + "' "
		end if
		if (FRectEndDay <> "") then
		        sqlStr = sqlStr + " and m.execdt<'" + Cstr(FRectEndDay) + "' "
		end if

		if (FRectItemGubun <> "") then
		    sqlStr = sqlStr + " and d.itemgubun = '" + FRectItemGubun + "'"
        end if


		if (FRectItemId <> "") then
		        sqlStr = sqlStr + " and d.shopitemid=" + Cstr(FRectItemId)
		end if

		if (FRectItemOption <> "") then
		        sqlStr = sqlStr + " and d.itemoption='" + FRectItemOption + "'"
		end if

		if (FRectChargeId <> "") then
		        sqlStr = sqlStr + " and m.chargeid='" + Cstr(FRectChargeId) + "' "
		end if
		
		if (FRectMakerId <> "") then
		        sqlStr = sqlStr + " and d.designerid='" + Cstr(FRectMakerId) + "' "
		end if

        sqlStr = sqlStr + " order by m.execdt, m.shopid, d.itemgubun, d.shopitemid, d.itemoption "
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CIpChulDetailByShopByItemItem
				FItemList(i).Fidx        = rsget("idx")
				FItemList(i).Fshopid     = rsget("shopid")
				FItemList(i).Fitemgubun  = rsget("itemgubun")
				FItemList(i).Fitemid     = rsget("itemid")
				FItemList(i).Fitemoption = rsget("itemoption")
				FItemList(i).Fitemno     = rsget("itemno")
				FItemList(i).Fexecdt     = rsget("execdt")
				FItemList(i).Fdesignerid = rsget("designerid")
				FItemList(i).Fsellcash   = rsget("sellcash")
				FItemList(i).Fsuplycash  = rsget("suplycash")
				FItemList(i).FItemname	 = db2html(rsget("itemname"))
				FItemList(i).FItemoptionname = db2html(rsget("itemoptionname"))

                FItemList(i).Fchargeid  = rsget("chargeid")
                FItemList(i).Fcomm_cd   = rsget("comm_cd")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	'//admin/offshop/popupchejumunsms_off.asp
	public Sub GetOneIpChulMaster()
		dim sqlStr

		sqlStr = " select top 1"
		sqlStr = sqlStr + " idx,chargeid,m.shopid,divcode,totalsellcash,totalsuplycash,totalshopbuyprice"
		sqlStr = sqlStr + " ,vatcode,execdt,m.regdate,reguserid,statecd,deleteyn,linkidx,scheduledate"
		sqlStr = sqlStr + " ,lastupdate,shopconfirmdate,upcheconfirmdate,shopconfirmuserid,upcheconfirmuserid"
		sqlStr = sqlStr + " ,songjangdiv,songjangname,songjangno,isbaljuExists,baljuconfirmdate,comment,sendsms"
		sqlStr = sqlStr + " ,s.shopname ,j.divname ,j.findurl ,d.comm_cd"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s"
		sqlStr = sqlStr + " 	on m.shopid=s.userid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " 	on m.shopid=d.shopid and m.chargeid=d.makerid"
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div j"
		sqlStr = sqlStr + " 	on m.songjangdiv=j.divcd"
		sqlStr = sqlStr + " where idx=" + CStr(FRectIdx) + ""

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then

		set FOneItem = new CIpChulItem

			FOneItem.Fidx            = rsget("idx")
			FOneItem.Fchargeid       = rsget("chargeid")
			FOneItem.Fshopid         = rsget("shopid")
			FOneItem.Fshopname       = db2html(rsget("shopname"))
			FOneItem.Fdivcode        = rsget("divcode")
			FOneItem.Ftotalsellcash  = rsget("totalsellcash")
			FOneItem.Ftotalsuplycash = rsget("totalsuplycash")
			FOneItem.Ftotalshopbuyprice = rsget("totalshopbuyprice")
			FOneItem.Fvatcode        = rsget("vatcode")
			FOneItem.Fexecdt         = rsget("execdt")
			FOneItem.Fregdate        = rsget("regdate")
			FOneItem.Fstatecd        = rsget("statecd")
			FOneItem.Fdeleteyn       = rsget("deleteyn")
			FOneItem.Flinkidx        = rsget("linkidx")
			FOneItem.FScheduleDt	 = rsget("scheduledate")
			FOneItem.fcomment	 = db2html(rsget("comment"))
			FOneItem.Fsongjangdiv	 = rsget("songjangdiv")
			FOneItem.Fsongjangname	 = db2html(rsget("divname"))
			FOneItem.Fsongjangno	 = db2html(rsget("songjangno"))
            FOneItem.Ffindurl        = db2html(rsget("findurl"))
            FOneItem.fsendsms	 = rsget("sendsms")
			FOneItem.Fsongjangdiv = Trim(FOneItem.Fsongjangdiv)
			FOneItem.Fshopconfirmdate = rsget("shopconfirmdate")
			FOneItem.Fupcheconfirmdate = rsget("upcheconfirmdate")
			FOneItem.Fshopconfirmuserid = rsget("shopconfirmuserid")
			FOneItem.Fupcheconfirmuserid = rsget("upcheconfirmuserid")
			FOneItem.FRegUserid     = rsget("reguserid")
			FOneItem.FComm_cd       = rsget("comm_cd")
			FOneItem.FisbaljuExists = rsget("isbaljuExists")
			FOneItem.Fbaljuconfirmdate = rsget("baljuconfirmdate")

		end if
		rsget.Close

	end Sub

	'//common/offshop/shop_ipchullist.asp		'//common/offshop/pop_ipgosheet.asp
	public function GetIpChulMasterList()
		dim sqlStr,i , sqlsearch

		If frect_IS_Maker_Upche Then
			sqlsearch = sqlsearch + " and m.statecd <> -5 "
		End IF

		if FRectmoveipchulyn ="Y" then
			sqlsearch = sqlsearch + " and m.ipchulmoveidx is not null"
		elseif FRectmoveipchulyn ="N" then
			sqlsearch = sqlsearch + " and m.ipchulmoveidx is null"
		end if
		if FRectIdx<>"" then
			sqlsearch = sqlsearch + " and m.idx=" + CStr(FRectIdx)
		end if

		if FRectNotIpgo<>"" then
			sqlsearch = sqlsearch + " and m.statecd<7"

		else
			if (FRectDatesearchtype="execdt") then
				if FRectStartDay<>"" then
					sqlsearch = sqlsearch + " and m.execdt>='" + FRectStartDay + "'"
				end if

				if FRectEndDay<>"" then
					sqlsearch = sqlsearch + " and m.execdt<'" + FRectEndDay + "'"
				end if
			else
				if FRectStartDay<>"" then
					sqlsearch = sqlsearch + " and m.scheduledate>='" + FRectStartDay + "'"
					sqlsearch = sqlsearch + " and m.scheduledate>='" + FRectStartDay + "'"
				end if

				if FRectEndDay<>"" then
					sqlsearch = sqlsearch + " and m.scheduledate<'" + FRectEndDay + "'"
				end if
			end if
		end if

		if FRectShopId<>"" then
			'//재고 이동의 경우 연관재고이동내역도 같이 보여줌(XXXX -> skyer9, 2014-08-21)
			sqlsearch = sqlsearch + " and m.shopid='"&FRectShopId&"' "
			''sqlsearch = sqlsearch + " and ("
			''sqlsearch = sqlsearch + " 	m.shopid='"&FRectShopId&"'"
			''sqlsearch = sqlsearch + " 	or ("
			''sqlsearch = sqlsearch + " 		m.idx in ("
			''sqlsearch = sqlsearch + " 			select tm.ipchulmoveidx"
			''sqlsearch = sqlsearch + " 			from [db_shop].[dbo].tbl_shop_ipchul_master tm"
			''sqlsearch = sqlsearch + " 			where tm.deleteyn='N'"
			''sqlsearch = sqlsearch + " 			and tm.shopid='"&FRectShopId&"'"
			''sqlsearch = sqlsearch + " 			and tm.scheduledate>='" + FRectStartDay + "' and tm.scheduledate<'" + FRectEndDay + "'"
			''sqlsearch = sqlsearch + " 			and isnull(tm.ipchulmoveidx,'') <> ''"
			''sqlsearch = sqlsearch + " 		)"
			''sqlsearch = sqlsearch + " 	)"
			''sqlsearch = sqlsearch + " )"
		end if

		if FRectChargeId<>"" then
			sqlsearch = sqlsearch + " and m.chargeid='" + FRectChargeId + "'"
		end if

		sqlStr = " select count(*) as cnt"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s"
		sqlStr = sqlStr + " 	on m.shopid=s.userid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " 	on m.shopid=d.shopid and m.chargeid=d.makerid"
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div j"
		sqlStr = sqlStr + " 	on m.songjangdiv=j.divcd"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_ipchul_master tm"
		sqlStr = sqlStr + "		on m.ipchulmoveidx = tm.ipchulmoveidx"
		sqlStr = sqlStr + "		and tm.deleteyn='N'"
		sqlStr = sqlStr + " where m.deleteyn='N' " & sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr + " m.idx ,m.chargeid ,m.shopid ,m.divcode ,m.totalsellcash ,m.totalsuplycash ,m.totalshopbuyprice"
		sqlStr = sqlStr + " ,m.vatcode ,m.execdt ,m.regdate ,m.reguserid ,m.statecd ,m.deleteyn ,m.linkidx ,m.scheduledate"
		sqlStr = sqlStr + " ,m.lastupdate ,m.shopconfirmdate ,m.upcheconfirmdate ,m.shopconfirmuserid ,m.upcheconfirmuserid"
		sqlStr = sqlStr + " ,m.songjangdiv ,m.songjangname ,m.songjangno ,m.isbaljuExists ,m.baljuconfirmdate ,m.comment"
		sqlStr = sqlStr + " ,m.sendsms ,m.ipchulmoveidx ,m.comm_cd as comm_cd_jungsan"
		sqlStr = sqlStr + " ,s.shopname, s.shopdiv, d.comm_cd, j.divname, j.findurl "
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_ipchul_master m"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_user s"
		sqlStr = sqlStr + " 	on m.shopid=s.userid"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d"
		sqlStr = sqlStr + " 	on m.shopid=d.shopid and m.chargeid=d.makerid"
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div j"
		sqlStr = sqlStr + " 	on m.songjangdiv=j.divcd"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_ipchul_master tm"
		sqlStr = sqlStr + "		on m.ipchulmoveidx = tm.ipchulmoveidx"
		sqlStr = sqlStr + "		and tm.deleteyn='N'"
		sqlStr = sqlStr + " where m.deleteyn='N' " & sqlsearch
		sqlStr = sqlStr + " order by m.idx desc"

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CIpChulItem

				FItemList(i).fipchulmoveidx            = rsget("ipchulmoveidx")
				FItemList(i).fshopdiv            = rsget("shopdiv")
				FItemList(i).Fidx            = rsget("idx")
				FItemList(i).Fchargeid       = rsget("chargeid")
				FItemList(i).Fshopid         = rsget("shopid")
				FItemList(i).Fshopname       = db2html(rsget("shopname"))
				FItemList(i).Fdivcode        = rsget("divcode")
				FItemList(i).Ftotalsellcash  = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash = rsget("totalsuplycash")
				FItemList(i).Ftotalshopbuyprice = rsget("totalshopbuyprice")
				FItemList(i).Fvatcode        = rsget("vatcode")
				FItemList(i).Fexecdt         = rsget("execdt")
				FItemList(i).Fregdate        = rsget("regdate")
				FItemList(i).Fstatecd        = rsget("statecd")
				FItemList(i).Fdeleteyn       = rsget("deleteyn")
				FItemList(i).Flinkidx        = rsget("linkidx")
				FItemList(i).FScheduleDt	 = rsget("scheduledate")
				FItemList(i).fcomment	 = db2html(rsget("comment"))
				FItemList(i).Fsongjangdiv	 = rsget("songjangdiv")
				FItemList(i).Fsongjangname	 = db2html(rsget("divname"))
				FItemList(i).Fsongjangno	 = db2html(rsget("songjangno"))
                FItemList(i).Ffindurl        = db2html(rsget("findurl"))
                FItemList(i).fsendsms	 = rsget("sendsms")
				FItemList(i).Fsongjangdiv = Trim(FItemList(i).Fsongjangdiv)
				FItemList(i).Fshopconfirmdate = rsget("shopconfirmdate")
				FItemList(i).Fupcheconfirmdate = rsget("upcheconfirmdate")
				FItemList(i).Fshopconfirmuserid = rsget("shopconfirmuserid")
				FItemList(i).Fupcheconfirmuserid = rsget("upcheconfirmuserid")
				FItemList(i).FRegUserid     = rsget("reguserid")
				FItemList(i).FComm_cd       = rsget("comm_cd")
				FItemList(i).fcomm_cd_jungsan       = rsget("comm_cd_jungsan")
				FItemList(i).FisbaljuExists = rsget("isbaljuExists")
				FItemList(i).Fbaljuconfirmdate = rsget("baljuconfirmdate")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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

%>

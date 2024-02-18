<%
'####################################################
' Description :  물류 입고리스트
' History : 2007.01.01 이상구 생성
'			2017.01.06 한용민 수정
'####################################################

'// mode = "mwonly" or "etcchulgo" or "etclosschulgo" or "offline"
Sub drawSelectBoxIpChulDivcode(mode, selectBoxName, selectedId)
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<% if (mode = "mwonly") then %>
			<option value="002" <% if (selectedId="002") then response.write "selected" %> >위탁
			<option value="001" <% if (selectedId="001") then response.write "selected" %> >매입
		<% end if %>
		<% if (mode = "etcchulgo") then %>
			<option value="003" <% if (selectedId="003") then response.write "selected" %> >판촉
			<option value="004" <% if (selectedId="004") then response.write "selected" %> >외부
			<option value="005" <% if (selectedId="005") then response.write "selected" %> >협찬
			<option value="006" <% if (selectedId="006") then response.write "selected" %> >B2B
			<option value="007" <% if (selectedId="007") then response.write "selected" %> >기타
			<option value="999" <% if (selectedId="999") then response.write "selected" %> >기타(정산않함)
			<option value="101" <% if (selectedId="101") then response.write "selected" %> >위탁출고
		<% end if %>
		<% if (mode = "etclosschulgo") then %>
			<option value="007" <% if (selectedId="007") then response.write "selected" %> >기타
			<option value="999" <% if (selectedId="999" or selectedId="") then response.write "selected" %> >기타(정산않함)
		<% end if %>
		<% if (mode = "offline") then %>
			<option value="801" <% if (selectedId="801") then response.write "selected" %> >Off매입
			<option value="802" <% if (selectedId="802") then response.write "selected" %> >Off위탁
		<% end if %>
	</select>
<%
End Sub

Function fnGetAGVCheckBalju(baljucode)
	dim sqlStr
	sqlStr = "select top 1 idx from [db_aLogistics].[dbo].tbl_agv_scheduleditems where requestMaster='STOCKIN("&baljucode&")' and isusing='Y'"
	rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	if not rsget_Logistics.eof then
		fnGetAGVCheckBalju = True
	else
		fnGetAGVCheckBalju = False
	end if
	rsget_Logistics.close
end Function

class CIpgo2AgvDiff
    public Fskucd
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fmakerid
    public Fitemname
    public Fitemoptionname
    public Fbaljuitemno
    public Fcheckitemno
    public Frealitemno
    public Fagvipgoitemno
    public FlocationCd1
    public FlocationCd2
    public FlocationCdCnt

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

Class CIpCulmasterItem
	public Fid
	public Fcode
	public Fsocid
	public Fdivcode
	public Fexecutedt
	public Fscheduledt
	public Ftotalsellcash
	public Ftotalsuplycash
	public ftotalitemno
	public Fvatcode
	public Fchargeid
	public Fcomment
	public Findt
	public Fupdt
	public Fdeldt
	public Ftotalbuycash
	public Fsocname
	public Fchargename
	public Frackipgoyn
	public fpurchasetypename
	public FBrandMaeipdiv
	public FppMasterIdx
	public Falinkcode
	public Fblinkcode

	public FpurchaseType

	public Fstatecd
	public Freportidx
	public Freportstate
	public pcuserdiv
	public FtplGubun

	public Ffinishid
	public Ffinishname
	public Fprizecnt
	public fipchulflag
	public fcheckusersn
	public frackipgousersn
	public fbigo
	public flogregdate
	public flogadminid
	public flogidx

	' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
	public function GetPurchaseTypeName()
		Select Case FpurchaseType
			Case "1"
				GetPurchaseTypeName = "일반유통"
			Case "4"
				GetPurchaseTypeName = "사입"
			Case "5"
				GetPurchaseTypeName = "OFF사입"
			Case "6"
				GetPurchaseTypeName = "수입"
			Case "7"
				GetPurchaseTypeName = "브랜드수입"
			Case "8"
				GetPurchaseTypeName = "제작"
			Case "9"
				GetPurchaseTypeName = "해외직구"
			Case "10"
				GetPurchaseTypeName = "B2B"
			Case Else
				GetPurchaseTypeName = FpurchaseType
		End Select
	end function

	public function GetMinusColor(icash)
		if (icash<0) then
			GetMinusColor = "#EE3333"
		else
			GetMinusColor = "#000000"
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="002" then
			GetDivCodeColor = "#000000"
		elseif Fdivcode="001" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="801" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="802" then
			GetDivCodeColor = "#5555DD"
		end if
	end function

	public function GetDivCodeName()
		if Fdivcode="002" then
			GetDivCodeName = "위탁"
		elseif Fdivcode="001" then
			GetDivCodeName = "매입"
		elseif Fdivcode="003" then
			GetDivCodeName = "판촉"
		elseif Fdivcode="004" then
			GetDivCodeName = "외부"
		elseif Fdivcode="005" then
			GetDivCodeName = "협찬"
		elseif Fdivcode="006" then
			GetDivCodeName = "B2B"
		elseif Fdivcode="007" then
			GetDivCodeName = "기타"
		elseif Fdivcode="101" then
			GetDivCodeName = "위탁출고"
		elseif Fdivcode="801" then
			GetDivCodeName = "Off매입"
		elseif Fdivcode="802" then
			GetDivCodeName = "Off위탁"
		elseif Fdivcode="999" then
			GetDivCodeName = "기타(정산않함)"
		end if
	end function

	public function GetBrandMaeipDivCodeName()
		if FBrandMaeipdiv="W" then
			GetBrandMaeipDivCodeName = "위탁"
		elseif FBrandMaeipdiv="M" then
			GetBrandMaeipDivCodeName = "매입"
		elseif FBrandMaeipdiv="U" then
			GetBrandMaeipDivCodeName = "업체"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CIpCulDetailItem
	public Fid
	public Fmastercode
	public fitemgubun
	public Fitemid
	public Fitemoption
	public Forgprice
	public Forgsuplycash
	public Fsellcash
	public Fsuplycash
	public Fitemno
	public Findt
	public Fupdt
	public Fdeldt
	public Fbuycash
	public Fmwgubun '매입/위탁
	public Fiitemgubun  '온라인(10)/오프전용(90)
	public Fiitemname
	public Fiitemoptionname
	public Fimakerid
	public FDtComment
    public FrackcodeByOption

	public FSellYn
	public FLimitYn
	public FLimitNo
	public FLimitSold
	public FOptionLimitNo
	public FOptionLimitSold
	public FOptUsing
	public FOptionCount
	public FIsNewItem

	public Flastrealno
	public Fcurrno
	public FOnlineMwdiv
	public FCenterMwdiv

	public FSmallimage
	public FOffimgMain
	public FOffimgList
	public FOffimgSmall

	public FMaystockno

	public Fdanjongyn

	public FPublicBarcode
    public Fbaljuitemno
    public FUpcheManageCode

    public Flastmwdiv

	public function IsSoldOut
		IsSoldOut = ((FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa=0)))
	end function

	public function GetIsSlodOutText
		if IsSoldOut then
			GetIsSlodOutText = "품절"
		else
			GetIsSlodOutText = ""
		end if
	end function

	public function GetIsSlodOutColor
		if IsSoldOut then
			GetIsSlodOutColor = "#FF2222"
		else
			GetIsSlodOutColor = "#000000"
		end if
	end function

	public function GetLimitEa
		GetLimitEa = FLimitNo-FLimitSold
		if GetLimitEa<1 then GetLimitEa=0
	end function

	public function getOptionLimitEa()
		getOptionLimitEa =0
		if (Flimityn="Y") then
			getOptionLimitEa = FOptionLimitNo-FOptionLimitSold
		end if

		if getOptionLimitEa<1 then getOptionLimitEa=0

		if (Foptioncount < 1) then
		    getOptionLimitEa = GetLimitEa
		end if
	end function

	public function GetMayCheckColor()
		if (Fitemno<1) and (FSellYn="Y") then
			GetMayCheckColor = "#88CC88"
		elseif (Fitemno>0) and (FSellYn<>"Y")  then
			GetMayCheckColor = "#8888CC"
		else
			GetMayCheckColor = "#FFFFFF"
		end if
	end function

	public function GetSellYnColor()
		if FSellYn="N" then
			GetSellYnColor = "#FF2222"
		elseif FSellYn="S" then
			GetSellYnColor = "#FF2222"
		else
			GetSellYnColor = "#000000"
		end if
	end function

	public function GetLimitYnColor()
		if FLimitYn="Y" then
			GetLimitYnColor = "#FF2222"
		else
			GetLimitYnColor = "#000000"
		end if
	end function



	public function getMwDivColor()
		if Fmwgubun="M" then
			getMwDivColor = "#CC2222"
		elseif Fmwgubun="W" then
			getMwDivColor = "#000000"
		else
		    getMwDivColor = "#000000"
		end if
	end function

	public function getOnlineMwdivColor()
		if FOnlineMwdiv="M" then
			getOnlineMwdivColor = "#CC2222"
		elseif FOnlineMwdiv="W" then
			getOnlineMwdivColor = "#2222CC"
		else
			getOnlineMwdivColor = "#000000"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CIpChulStorage
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectCodeGubun
	public FRectExecuteDtStart
	public FRectExecuteDtEnd
	public FRectScheduleDtStart
	public FRectScheduleDtEnd

	public FRectSocID
	public FRectId
	public FRectStoragecode
	public FRectDivCode
	public FRectMaeipDiv
	public FRectCode
	public FRectALinkCode
	public FRectBLinkCode
	public FRectChulgoState
	public FRectOnOffGubun
	public FRectRackipgoyn
	public FRectReturnyn

	public FRectItemID
    public FRectBrandPurchaseType

	public FRectSearchType
	public FRectSearchText
	public FRectMinusOnly
    public FRectPCuserDiv
    public FRectChargename
    public FtplGubun
    public FRectReportState
	Public FRectPrcGbn
	public FRectNotalinkcode
    public FRectComment
    public FRectYYYYMMDD
    Public FRectYYYYMM
    public FRectDiffOnly

    public function GetIdxFromMasterCode()
		dim sqlStr,i

		sqlStr = "select top 1 m.id "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " where m.code='" + FRectStoragecode + "'"
		rsget.Open sqlStr, dbget, 1
		if  not rsget.EOF  then
			i = rsget("id")
		end if
		rsget.close

                GetIdxFromMasterCode = i
        end function

	public sub GetIpgoToAgvDiffList
		dim sqlStr,i

		sqlStr = " exec [db_storage].[dbo].[usp_Logics_Ipgo2AgvDiffList_Get] '" & FRectYYYYMMDD & "', '" & FRectDiffOnly & "' "

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CIpgo2AgvDiff

				FItemList(i).Fskucd          = rsget("skucd")
                FItemList(i).Fitemgubun      = rsget("itemgubun")
                FItemList(i).Fitemid         = rsget("itemid")
                FItemList(i).Fitemoption     = rsget("itemoption")
                FItemList(i).Fmakerid        = rsget("makerid")
                FItemList(i).Fitemname       = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))
                FItemList(i).Fbaljuitemno    = rsget("baljuitemno")
                FItemList(i).Fcheckitemno    = rsget("checkitemno")
                FItemList(i).Frealitemno     = rsget("realitemno")
                FItemList(i).Fagvipgoitemno  = rsget("itemno")
                FItemList(i).FlocationCd1    = rsget("locationCd1")
                FItemList(i).FlocationCd2    = rsget("locationCd2")
                FItemList(i).FlocationCdCnt  = rsget("locationCdCnt")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetMwDiffMasterList
		dim sqlStr,i
		sqlStr = "select top 300 m.id,m.code,m.divcode, m.socid,m.totalsellcash,m.totalsuplycash,"
		sqlStr = sqlStr + " m.scheduledt,m.executedt,m.chargeid,m.socname,m.chargename,"
		sqlStr = sqlStr + " c.maeipdiv from "
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
		sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " on m.socid=c.userid"
		sqlStr = sqlStr + " where m.executedt>='" + FRectExecuteDtStart + "'"
		sqlStr = sqlStr + " and m.executedt<'" + FRectExecuteDtEnd + "'"
		sqlStr = sqlStr + " and m.divcode='" + FRectDivCode + "'"
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and c.maeipdiv<>'" + FRectMaeipDiv + "'"

		if FRectSocID<>"" then
			sqlStr = sqlStr + " and m.socid='" + FRectSocID + "'"
		end if

		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CIpCulmasterItem

				FItemList(i).Fid             = rsget("id")
				FItemList(i).Fcode           = rsget("code")
				FItemList(i).Fsocid          = rsget("socid")
				FItemList(i).Fdivcode		 = rsget("divcode")
				FItemList(i).Fexecutedt      = rsget("executedt")
				FItemList(i).Fscheduledt     = rsget("scheduledt")
				FItemList(i).Ftotalsellcash  = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash = rsget("totalsuplycash")
				FItemList(i).Fchargeid       = rsget("chargeid")
				FItemList(i).Fsocname        = db2html(rsget("socname"))
				FItemList(i).Fchargename     = db2html(rsget("chargename"))
				FItemList(i).FBrandMaeipdiv  = rsget("maeipdiv")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetIpChulMaster
		dim sqlStr
		sqlStr = " select top 1 m.id, "
		sqlStr = sqlStr + " m.code, m.socid, m.divcode, m.executedt, m.scheduledt, "
		sqlStr = sqlStr + " IsNULL(totalsellcash,0) as totalsellcash, "
		sqlStr = sqlStr + " IsNULL(totalsuplycash,0) as totalsuplycash, "
		sqlStr = sqlStr + " vatcode, chargeid, comment, indt, updt, deldt,"
		sqlStr = sqlStr + " IsNULL(totalbuycash,0) as totalbuycash, "
		sqlStr = sqlStr + " m.socname, chargename, rackipgoyn , statecd , ep.reportidx, ep.reportstate "
		sqlStr = sqlSTr + " , isNull(p.userdiv,'') as pUserdiv, isNull(c.userdiv,'') as cUserdiv ,IsNull(p.tplcompanyid, '') as tplgubun"
		sqlStr = sqlStr + " ,( select count(id) from db_storage.dbo.tbl_acount_storage_detail as d where m.code = d.mastercode and sellcash>50000 and m.socid ='itemgift') as ipcnt"
		sqlStr = sqlSTr + ", (select top 1 jj.baljucode from [db_storage].dbo.tbl_ordersheet_master jj where alinkcode = m.code) as alinkcode"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master as m"
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep on m.id = ep.scmlinkNo and ep.isUsing =1   and ep.edmsidx = 58 "
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_partner as p on m.socid=p.id "
		sqlStr = sqlStr + " left outer join [db_user].[dbo].tbl_user_c as c on c.userid=p.id"
		sqlStr = sqlStr + " where m.id=" + CStr(FRectId)
		''response.write sqlSTr
		rsget.Open sqlStr, dbget, 1

		ftotalcount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CIpCulmasterItem

		if  not rsget.EOF  then

			FOneItem.Fid             = rsget("id")
			FOneItem.Fcode           = rsget("code")
			FOneItem.Fsocid          = rsget("socid")
			FOneItem.Fdivcode        = rsget("divcode")
			FOneItem.Fexecutedt      = rsget("executedt")
			FOneItem.Fscheduledt     = rsget("scheduledt")
			FOneItem.Ftotalsellcash  = rsget("totalsellcash")
			FOneItem.Ftotalsuplycash = rsget("totalsuplycash")
			FOneItem.Fvatcode        = rsget("vatcode")
			FOneItem.Fchargeid       = rsget("chargeid")
			FOneItem.Fcomment        = db2html(rsget("comment"))
			FOneItem.Findt           = rsget("indt")
			FOneItem.Fupdt           = rsget("updt")
			FOneItem.Fdeldt          = rsget("deldt")
			FOneItem.Ftotalbuycash   = rsget("totalbuycash")
			FOneItem.Fsocname        = db2html(rsget("socname"))
			FOneItem.Fchargename     = db2html(rsget("chargename"))
			FOneItem.Frackipgoyn     = rsget("rackipgoyn")
			FOneItem.Fstatecd		 = rsget("statecd")
			FOneItem.Freportidx      = rsget("reportidx")
			FOneItem.Freportstate    = rsget("reportstate")
			FOneItem.pcuserdiv       = Cstr(rsget("pUserdiv"))&"_"&Cstr(rsget("cUserdiv"))
			FOneITem.FtplGubun       = rsget("tplgubun")
			FoneItem.Fprizecnt			 = rsget("ipcnt")
			FoneItem.Falinkcode			 = rsget("alinkcode")
		end if
		rsget.close
	end sub

	public sub GetIpChulDetail_edit_log
		dim sqlStr,i, sqlsearch

		if FRectStoragecode="" or isnull(FRectStoragecode) then exit sub

		if FRectStoragecode<>"" then
			sqlsearch = sqlsearch & " and ml.code='" & CStr(FRectStoragecode) & "'"
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " logidx, id, code, socid, divcode, executedt, scheduledt, totalsellcash, totalsuplycash, vatcode"
		sqlStr = sqlStr & " , chargeid, comment, indt, updt, deldt, totalbuycash, socname, chargename, ipchulflag"
		sqlStr = sqlStr & " , rackipgoyn, checkusersn, rackipgousersn, statecd, finishid, finishname, bigo, logregdate, logadminid"
		sqlStr = sqlStr & " from db_storage.dbo.tbl_acount_storage_master_edit_log ml with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by logidx desc"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CIpCulmasterItem

				FItemList(i).flogidx              = rsget("logidx")
				FItemList(i).Fid              = rsget("id")
				FItemList(i).fcode              = rsget("code")
				FItemList(i).fsocid              = rsget("socid")
				FItemList(i).fdivcode              = rsget("divcode")
				FItemList(i).fexecutedt              = rsget("executedt")
				FItemList(i).fscheduledt              = rsget("scheduledt")
				FItemList(i).ftotalsellcash              = rsget("totalsellcash")
				FItemList(i).ftotalsuplycash              = rsget("totalsuplycash")
				FItemList(i).fvatcode              = rsget("vatcode")
				FItemList(i).fchargeid              = rsget("chargeid")
				FItemList(i).fcomment              = rsget("comment")
				FItemList(i).findt              = rsget("indt")
				FItemList(i).fupdt              = rsget("updt")
				FItemList(i).fdeldt              = rsget("deldt")
				FItemList(i).ftotalbuycash              = rsget("totalbuycash")
				FItemList(i).fsocname              = rsget("socname")
				FItemList(i).fchargename              = rsget("chargename")
				FItemList(i).fipchulflag              = rsget("ipchulflag")
				FItemList(i).frackipgoyn              = rsget("rackipgoyn")
				FItemList(i).fcheckusersn              = rsget("checkusersn")
				FItemList(i).frackipgousersn              = rsget("rackipgousersn")
				FItemList(i).fstatecd              = rsget("statecd")
				FItemList(i).ffinishid              = rsget("finishid")
				FItemList(i).fbigo              = rsget("bigo")
				FItemList(i).flogregdate              = rsget("logregdate")
				FItemList(i).flogadminid              = rsget("logadminid")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetIpChulDetail
		dim sqlStr,i

		sqlStr = "select top 2000"
		sqlStr = sqlStr + " d.*, i.mwdiv, i.smallimage, si.offimgmain, si.offimglist, si.offimgsmall, i.orgprice, i.orgsuplycash, "
		sqlStr = sqlStr + " s.barcode,d.comment, isNull(s.upchemanagecode,'') AS upchemanagecode, si.centermwdiv, s.rackcodeByOption"
        if FRectYYYYMM <> "" then
            sqlStr = sqlStr + "	, a.lastmwdiv "
        else
            sqlStr = sqlStr + "	, NULL as lastmwdiv "
        end if
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i with (nolock)"
		sqlStr = sqlStr + " 	on d.iitemgubun='10' and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s with (nolock)"
		sqlStr = sqlStr + "		on d.iitemgubun=s.itemgubun and d.itemid=s.itemid and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_item si with (nolock)"
		sqlStr = sqlStr + "		on d.iitemgubun=si.itemgubun and d.itemid=si.shopitemid and d.itemoption=si.itemoption "

		if FRectYYYYMM <> "" then
		    sqlStr = sqlStr + "	left join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] a with (nolock) "
		    sqlStr = sqlStr + "	on "
		    sqlStr = sqlStr + "		1 = 1 "
		    sqlStr = sqlStr + "		and d.iitemgubun = a.itemgubun "
		    sqlStr = sqlStr + "		AND d.itemid = a.itemid "
		    sqlStr = sqlStr + "		AND d.itemoption = a.itemoption "
		    sqlStr = sqlStr + "		and a.yyyymm = '" & FRectYYYYMM & "' "
        end if

		sqlStr = sqlStr + " where d.mastercode='" + CStr(FRectStoragecode) + "'"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " order by d.imakerid,d.iitemgubun,d.itemid,d.itemoption"		' 다른매뉴와 동일하게 맞춤

		''response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CIpCulDetailItem

				FItemList(i).Fid              = rsget("id")
				FItemList(i).Fmastercode      = rsget("mastercode")
				FItemList(i).Fitemgubun      = rsget("iitemgubun")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Forgprice        = rsget("orgprice")
				FItemList(i).Forgsuplycash    = rsget("orgsuplycash")
				FItemList(i).Fsellcash        = rsget("sellcash")
				FItemList(i).Fsuplycash       = rsget("suplycash")
				FItemList(i).Fitemno          = rsget("itemno")
				FItemList(i).Findt            = rsget("indt")
				FItemList(i).Fupdt            = rsget("updt")
				FItemList(i).Fdeldt           = rsget("deldt")
				FItemList(i).Fbuycash         = rsget("buycash")
				FItemList(i).Fmwgubun         = rsget("mwgubun")
				FItemList(i).Fiitemgubun      = rsget("iitemgubun")
				FItemList(i).Fiitemname       = db2html(rsget("iitemname"))
				FItemList(i).Fiitemoptionname = db2html(rsget("iitemoptionname"))
				FItemList(i).Fimakerid		  = rsget("imakerid")
				FItemList(i).FOnlineMwdiv		= rsget("mwdiv")
				FItemList(i).FCenterMwdiv		= rsget("centermwdiv")
				FItemList(i).Fsmallimage	= rsget("smallimage")
				FItemList(i).FPublicBarcode = rsget("barcode")
				FItemList(i).FUpcheManageCode = rsget("upchemanagecode")
				FItemList(i).FDtComment = rsget("comment")
                FItemList(i).FrackcodeByOption = rsget("rackcodeByOption")
                FItemList(i).Flastmwdiv          = rsget("lastmwdiv")

				if Not IsNull(FItemList(i).Fsmallimage) then
					FItemList(i).Fsmallimage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).Fsmallimage
				end if

				FItemList(i).FOffimgMain	= rsget("offimgmain")
					if isnull(FItemList(i).FOffimgMain) then FItemList(i).FOffimgMain=""
				FItemList(i).FOffimgList	= rsget("offimglist")
					if isnull(FItemList(i).FOffimgList) then FItemList(i).FOffimgList=""
				FItemList(i).FOffimgSmall	= rsget("offimgsmall")
					if isnull(FItemList(i).FOffimgSmall) then FItemList(i).FOffimgSmall=""

				if FItemList(i).FOffimgMain<>"" then FItemList(i).FOffimgMain = webImgUrl + "/offimage/offmain/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgMain
				if FItemList(i).FOffimgList<>"" then FItemList(i).FOffimgList = webImgUrl + "/offimage/offlist/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgList
				if FItemList(i).FOffimgSmall<>"" then FItemList(i).FOffimgSmall = webImgUrl + "/offimage/offsmall/i" + FItemList(i).Fitemgubun + "/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + FItemList(i).FOffimgSmall

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetIpChulDetailCheck
		dim sqlStr,i
		dim lasteBasedate

        ''get ordersheet_master's idx
        dim ordersheetidx

        sqlStr = " select top 1 idx from [db_storage].[dbo].tbl_ordersheet_master"
        sqlStr = sqlStr + " where blinkcode='" + CStr(FRectStoragecode) + "'"
        sqlStr = sqlStr + " and deldt is null"
        sqlStr = sqlStr + " order by idx desc"

        rsget.Open sqlStr, dbget, 1

        ordersheetidx = 0
        if Not rsget.Eof then
            ordersheetidx = rsget("idx")
        end if
        rsget.close

		'sqlStr = "select convert(varchar(10),dateadd(d,-14,getdate()),21) as lasteBasedate " + VbCrlf
		'rsget.Open sqlStr, dbget, 1
		'	lasteBasedate = rsget("lasteBasedate")
		'rsget.close

		sqlStr = " select d.*, IsNULL(o.baljuitemno,0) as baljuitemno, "
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, isnull(i.optioncnt,0) as optioncnt, i.danjongyn, v.isusing as optusing, v.optlimitno, v.optlimitsold, "+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.realstock,0) as realstock, IsNULL(s.ipkumdiv5,0) as ipkumdiv5,"+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.offconfirmno,0) as offconfirmno, "
		sqlStr = sqlStr + " IsNULL(s.offjupno,0) as offjupno, "+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.ipkumdiv2,0) as ipkumdiv2, IsNULL(s.ipkumdiv4,0) as ipkumdiv4, (case when IsNULL(s.itemid,-1) = -1 then 'Y' else 'N' end) as isnewitem "+ VbCrlf
		sqlStr = sqlStr + " ,d.comment " +vbcrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail d"

		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlStr + " on d.iitemgubun='10'"
		sqlStr = sqlStr + " and d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + " and d.itemid=s.itemid"
		sqlStr = sqlStr + " and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " on d.iitemgubun='10'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlStr + " on d.iitemgubun='10'"
		sqlStr = sqlStr + " and d.itemid=v.itemid "
		sqlStr = sqlStr + " and d.itemoption=v.itemoption "

        sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_detail o"
		sqlStr = sqlStr + " on o.masteridx=" + CStr(ordersheetidx)
		sqlStr = sqlStr + " and d.iitemgubun=o.itemgubun"
		sqlStr = sqlStr + " and d.itemid=o.itemid"
		sqlStr = sqlStr + " and d.itemoption=o.itemoption"



		'sqlStr = sqlStr + " left join ( " + VbCrlf
		'sqlStr = sqlStr + " 	select d.itemid,d.itemoption, " + VbCrlf
		'sqlStr = sqlStr + " 	sum(case when m.ipkumdiv='2' then d.itemno*-1 else 0 end) as ipkumdiv2, " + VbCrlf
		'sqlStr = sqlStr + " 	sum(case when m.ipkumdiv='4' then d.itemno*-1 else 0 end) as ipkumdiv4 " + VbCrlf
		'sqlStr = sqlStr + "  	from " + VbCrlf
		'sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m, " + VbCrlf
		'sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d " + VbCrlf
		'sqlStr = sqlStr + " 	where m.orderserial=d.orderserial " + VbCrlf
		'sqlStr = sqlStr + " 	and m.regdate>'" + lasteBasedate + "' " + VbCrlf
		'sqlStr = sqlStr + " 	and m.ipkumdiv>1 " + VbCrlf
		'sqlStr = sqlStr + " 	and m.ipkumdiv<5 " + VbCrlf
		'sqlStr = sqlStr + " 	and m.cancelyn='N' " + VbCrlf
		'sqlStr = sqlStr + " 	and d.cancelyn<>'Y' " + VbCrlf
		'sqlStr = sqlStr + " 	group by d.itemid,d.itemoption " + VbCrlf
		'sqlStr = sqlStr + " ) T " + VbCrlf
		'sqlStr = sqlStr + " on d.mastercode='" + CStr(FRectStoragecode) + "' and d.iitemgubun='10' and d.itemid=T.itemid and d.itemoption=T.itemoption" + VbCrlf

		sqlStr = sqlStr + " where d.mastercode='" + CStr(FRectStoragecode) + "'"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " order by d.itemid,d.itemoption"


		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CIpCulDetailItem

				FItemList(i).Fid              = rsget("id")
				FItemList(i).Fmastercode      = rsget("mastercode")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fsellcash        = rsget("sellcash")
				FItemList(i).Fsuplycash       = rsget("suplycash")
				FItemList(i).Fbaljuitemno     = rsget("baljuitemno")
				FItemList(i).Fitemno          = rsget("itemno")
				FItemList(i).Findt            = rsget("indt")
				FItemList(i).Fupdt            = rsget("updt")
				FItemList(i).Fdeldt           = rsget("deldt")
				FItemList(i).Fbuycash         = rsget("buycash")
				FItemList(i).Fmwgubun         = rsget("mwgubun")
				FItemList(i).Fiitemgubun      = rsget("iitemgubun")
				FItemList(i).Fiitemname       = db2html(rsget("iitemname"))
				FItemList(i).Fiitemoptionname = db2html(rsget("iitemoptionname"))
				FItemList(i).Fimakerid		  = rsget("imakerid")

				FItemList(i).FSellYn          = rsget("sellyn")
				FItemList(i).FLimitYn         = rsget("limityn")
				FItemList(i).FLimitNo         = rsget("limitno")
				FItemList(i).FLimitSold       = rsget("limitsold")
				FItemList(i).FOptionLimitNo   = rsget("optlimitno")
				FItemList(i).FOptionLimitSold = rsget("optlimitsold")
				FItemList(i).FOptUsing        = rsget("optusing")
				FItemList(i).FOptionCount     = rsget("optioncnt")

				FItemList(i).Fdanjongyn				= rsget("danjongyn")

				''재고파악재고
				FItemList(i).Fcurrno	= rsget("realstock") + rsget("ipkumdiv5") + rsget("offconfirmno")
				''한정비교재고 2011-06-23 수정 : 오프 준비중 수량 제외. => 2011-06-24 다시 원상복귀
				FItemList(i).FMaystockno = FItemList(i).Fcurrno + rsget("ipkumdiv4") + rsget("ipkumdiv2") ''' + rsget("offjupno") (오프라인주문접수건 뺌)
                ''FItemList(i).FMaystockno = rsget("realstock") + rsget("ipkumdiv5") + rsget("ipkumdiv4") + rsget("ipkumdiv2")

                FItemList(i).FIsNewItem     = rsget("isnewitem")
				FItemList(i).FDtComment		= rsget("comment")



				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	public sub GetshopIpChulDetailCheck
		dim sqlStr,i
		dim lasteBasedate

        ''get ordersheet_master's idx
        dim ordersheetidx

        sqlStr = " select top 1 idx from [db_storage].[dbo].tbl_ordersheet_master"
        sqlStr = sqlStr + " where alinkcode='" + CStr(FRectStoragecode) + "'"
        sqlStr = sqlStr + " and deldt is null"
        sqlStr = sqlStr + " order by idx desc"

        rsget.Open sqlStr, dbget, 1

        ordersheetidx = 0
        if Not rsget.Eof then
            ordersheetidx = rsget("idx")
        end if
        rsget.close

		'sqlStr = "select convert(varchar(10),dateadd(d,-14,getdate()),21) as lasteBasedate " + VbCrlf
		'rsget.Open sqlStr, dbget, 1
		'	lasteBasedate = rsget("lasteBasedate")
		'rsget.close

		sqlStr = " select d.*, IsNULL(o.baljuitemno,0) as baljuitemno, "
		sqlStr = sqlStr + " i.sellyn, i.limityn, i.limitno, i.limitsold, isnull(i.optioncnt,0) as optioncnt, i.danjongyn, v.isusing as optusing, v.optlimitno, v.optlimitsold, "+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.realstock,0) as realstock, IsNULL(s.ipkumdiv5,0) as ipkumdiv5,"+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.offconfirmno,0) as offconfirmno, "
		sqlStr = sqlStr + " IsNULL(s.offjupno,0) as offjupno, "+ VbCrlf
		sqlStr = sqlStr + " IsNULL(s.ipkumdiv2,0) as ipkumdiv2, IsNULL(s.ipkumdiv4,0) as ipkumdiv4, (case when IsNULL(s.itemid,-1) = -1 then 'Y' else 'N' end) as isnewitem "+ VbCrlf
		sqlStr = sqlStr + " ,d.comment " +vbcrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail d"

		sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s"
		sqlStr = sqlStr + " on d.iitemgubun='10'"
		sqlStr = sqlStr + " and d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + " and d.itemid=s.itemid"
		sqlStr = sqlStr + " and d.itemoption=s.itemoption"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + " on d.iitemgubun='10'"
		sqlStr = sqlStr + " and d.itemid=i.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
		sqlStr = sqlStr + " on d.iitemgubun='10'"
		sqlStr = sqlStr + " and d.itemid=v.itemid "
		sqlStr = sqlStr + " and d.itemoption=v.itemoption "

        sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_detail o"
		sqlStr = sqlStr + " on o.masteridx=" + CStr(ordersheetidx)
		sqlStr = sqlStr + " and d.iitemgubun=o.itemgubun"
		sqlStr = sqlStr + " and d.itemid=o.itemid"
		sqlStr = sqlStr + " and d.itemoption=o.itemoption"

		sqlStr = sqlStr + " where d.mastercode='" + CStr(FRectStoragecode) + "'"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " order by d.itemid,d.itemoption"


		'response.write sqlStr
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new CIpCulDetailItem

				FItemList(i).Fid              = rsget("id")
				FItemList(i).Fmastercode      = rsget("mastercode")
				FItemList(i).Fitemid          = rsget("itemid")
				FItemList(i).Fitemoption      = rsget("itemoption")
				FItemList(i).Fsellcash        = rsget("sellcash")
				FItemList(i).Fsuplycash       = rsget("suplycash")
				FItemList(i).Fbaljuitemno     = rsget("baljuitemno")
				FItemList(i).Fitemno          = rsget("itemno")
				FItemList(i).Findt            = rsget("indt")
				FItemList(i).Fupdt            = rsget("updt")
				FItemList(i).Fdeldt           = rsget("deldt")
				FItemList(i).Fbuycash         = rsget("buycash")
				FItemList(i).Fmwgubun         = rsget("mwgubun")
				FItemList(i).Fiitemgubun      = rsget("iitemgubun")
				FItemList(i).Fiitemname       = db2html(rsget("iitemname"))
				FItemList(i).Fiitemoptionname = db2html(rsget("iitemoptionname"))
				FItemList(i).Fimakerid		  = rsget("imakerid")

				FItemList(i).FSellYn          = rsget("sellyn")
				FItemList(i).FLimitYn         = rsget("limityn")
				FItemList(i).FLimitNo         = rsget("limitno")
				FItemList(i).FLimitSold       = rsget("limitsold")
				FItemList(i).FOptionLimitNo   = rsget("optlimitno")
				FItemList(i).FOptionLimitSold = rsget("optlimitsold")
				FItemList(i).FOptUsing        = rsget("optusing")
				FItemList(i).FOptionCount     = rsget("optioncnt")

				FItemList(i).Fdanjongyn				= rsget("danjongyn")

				''재고파악재고
				FItemList(i).Fcurrno	= rsget("realstock") + rsget("ipkumdiv5") + rsget("offconfirmno")
				''한정비교재고 2011-06-23 수정 : 오프 준비중 수량 제외. => 2011-06-24 다시 원상복귀
				FItemList(i).FMaystockno = FItemList(i).Fcurrno + rsget("ipkumdiv4") + rsget("ipkumdiv2") ''' + rsget("offjupno") (오프라인주문접수건 뺌)
                ''FItemList(i).FMaystockno = rsget("realstock") + rsget("ipkumdiv5") + rsget("ipkumdiv4") + rsget("ipkumdiv2")

                FItemList(i).FIsNewItem     = rsget("isnewitem")
				FItemList(i).FDtComment		= rsget("comment")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end sub

	'//admin/newstorage/culgolist.asp
	public Sub GetIpChulgoByItemID()
		dim i,sqlStr, sqlsearch, tmpStr

		if FRectCode<>"" then
			if (InStr(FRectCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and m.code in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and m.code = '" + FRectCode + "'"
			end if
		end if

		if FRectCodeGubun<>"" then
			sqlsearch = sqlsearch + " and Left(m.code,2)='" + FRectCodeGubun + "'"
		end if

		if FRectSocID<>"" then
			sqlsearch = sqlsearch + " and m.socid='" + trim(FRectSocID) + "'"
		end if

		if FRectExecuteDtStart<>"" then
			sqlsearch = sqlsearch + " and m.executedt>='" + FRectExecuteDtStart + "'"
		end if

		if FRectExecuteDtEnd<>"" then
			sqlsearch = sqlsearch + " and m.executedt<'" + FRectExecuteDtEnd + "'"
		end if

		IF (FRectChargename <> "") THEN
		     sqlsearch = sqlsearch + " and m.chargename='"& trim(FRectChargename) &"'"
	    END IF

         IF (FRectPCuserDiv<>"") then
		    sqlsearch = sqlsearch + " and p.userdiv='"&splitValue(FRectPCuserDiv,"_",0)&"'"
		    sqlsearch = sqlsearch + " and c.userdiv='"&splitValue(FRectPCuserDiv,"_",1)&"'"
		end if

	    if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlsearch = sqlsearch + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				sqlsearch = sqlsearch + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FtplGubun) + "' "
			end if
		end If

		if (FRectReportstate = "0" ) then
		    sqlsearch = sqlsearch + " and ( ep.reportstate is null or ep.reportstate ='' ) "
	    elseif (FRectReportstate <> "" ) then
	        sqlsearch = sqlsearch + " and   ep.reportstate  = "+ CStr(FRectReportstate)
	    end if
		if FRectNotalinkcode<>"" then
		    sqlsearch = sqlsearch + " and isnull(nj.alinkcode,'')=''"
		end if

		sqlStr = " select count(m.id) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m "
		sqlStr = sqlStr + " inner join [db_storage].[dbo].tbl_acount_storage_detail d on m.code=d.mastercode "
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep on m.id = ep.scmlinkNo and ep.isUsing =1   and ep.edmsidx = 58 "

		if ( FRectPCuserDiv<>"" or FtplGubun <> "" ) then
		    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p"
		    sqlStr = sqlStr + " on m.socid=p.id"
		    sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		    sqlStr = sqlStr + " on c.userid=p.id"
		end if

		if (FRectALinkCode<>"") then
		    sqlStr = sqlStr + " Join [db_storage].dbo.tbl_ordersheet_master j"
		    sqlStr = sqlStr + " on m.code = j.alinkcode "
			if (InStr(FRectALinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectALinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectALinkCode + "'"
			end if
		elseif (FRectBLinkCode<>"") then
		    sqlStr = sqlStr + " Join [db_storage].dbo.tbl_ordersheet_master j"
		    sqlStr = sqlStr + " on m.code = j.blinkcode "
			if (InStr(FRectBLinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectBLinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectBLinkCode + "'"
			end if
		end If

		if FRectNotalinkcode<>"" then
		    sqlStr = sqlStr + " left Join [db_storage].dbo.tbl_ordersheet_master nj"
		    sqlStr = sqlStr + " 	on m.code = nj.alinkcode "
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemID)

		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, "
		sqlStr = sqlStr + " m.code,m.socid,m.divcode,m.executedt,m.scheduledt,"
		sqlStr = sqlStr + " IsNULL(m.totalsellcash,0) as totalsellcash,"
		sqlStr = sqlStr + " IsNULL(m.totalsuplycash,0) as totalsuplycash,"
		sqlStr = sqlStr + " m.vatcode,m.chargeid,m.comment,m.indt,m.updt,m.deldt,"
		sqlStr = sqlStr + " IsNULL(m.totalbuycash,0) as totalbuycash,"
		sqlStr = sqlStr + " m.socname,m.chargename"
		sqlStr = sqlStr + " ,m.statecd , ep.reportidx, ep.reportstate "
		sqlStr = sqlStr & " , isnull((select sum(itemno)" & vbcrlf
		sqlStr = sqlStr & " 	from [db_storage].[dbo].tbl_acount_storage_detail dd" & vbcrlf
		sqlStr = sqlStr & " 	where dd.mastercode = m.code" & vbcrlf
		sqlStr = sqlStr & " 	and dd.deldt is null),0) as totalitemno" & vbcrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m "
		sqlStr = sqlStr + " inner join [db_storage].[dbo].tbl_acount_storage_detail d on m.code=d.mastercode "
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep on m.id = ep.scmlinkNo and ep.isUsing =1   and ep.edmsidx = 58 "

		if (FRectALinkCode<>"") then
		    sqlStr = sqlStr + " Join [db_storage].dbo.tbl_ordersheet_master j"
		    sqlStr = sqlStr + " on m.code = j.alinkcode "
			if (InStr(FRectALinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectALinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectALinkCode + "'"
			end if
		elseif (FRectBLinkCode<>"") then
		    sqlStr = sqlStr + " Join [db_storage].dbo.tbl_ordersheet_master j"
		    sqlStr = sqlStr + " on m.code = j.blinkcode "
			if (InStr(FRectBLinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectBLinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectBLinkCode + "'"
			end if
		end If

		if FRectNotalinkcode<>"" then
		    sqlStr = sqlStr + " left Join [db_storage].dbo.tbl_ordersheet_master nj"
		    sqlStr = sqlStr + " 	on m.code = nj.alinkcode "
		end if

		if ( FRectPCuserDiv<>"" or FtplGubun<>"") then
		    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p"
		    sqlStr = sqlStr + " on m.socid=p.id"
		    sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c"
		    sqlStr = sqlStr + " on c.userid=p.id"
		end if

		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " and d.itemid=" + CStr(FRectItemID)
		sqlStr = sqlStr + " order by m.id desc"

		'response.write sqlStr & "<br>"
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
				set FItemList(i) = new CIpCulmasterItem

				FItemList(i).Fid             = rsget("id")
				FItemList(i).Fcode           = rsget("code")
				FItemList(i).Fsocid          = rsget("socid")
				FItemList(i).Fdivcode        = rsget("divcode")
				FItemList(i).Fexecutedt      = rsget("executedt")
				FItemList(i).Fscheduledt     = rsget("scheduledt")
				FItemList(i).Ftotalsellcash  = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash = rsget("totalsuplycash")
				FItemList(i).Fvatcode        = rsget("vatcode")
				FItemList(i).Fchargeid       = rsget("chargeid")
				FItemList(i).Fcomment        = db2html(rsget("comment"))
				FItemList(i).Findt           = rsget("indt")
				FItemList(i).Fupdt           = rsget("updt")
				FItemList(i).Fdeldt          = rsget("deldt")
				FItemList(i).Ftotalbuycash   = rsget("totalbuycash")
				FItemList(i).Fsocname        = db2html(rsget("socname"))
				FItemList(i).Fchargename     = db2html(rsget("chargename"))
				FItemList(i).Fstatecd  		= rsget("statecd")
 				FItemList(i).Freportidx		= rsget("reportidx")
 				FItemList(i).Freportstate	= rsget("reportstate")
 				FItemList(i).ftotalitemno	= rsget("totalitemno")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

	'//admin/newstorage/culgolist.asp
	public Sub GetIpChulgoList()
		dim i, sqlStr, sqlsearch, tmpStr

		if FRectMinusOnly<>"" then
			sqlsearch = sqlsearch + " and IsNULL(m.totalsellcash,0)<0"
		end if

		if (FRectBrandPurchaseType<>"" and  FRectPCuserDiv = "" and FtplGubun = "" ) then
			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sqlsearch = sqlsearch + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sqlsearch = sqlsearch & " 	and p.purchasetype in ('3','5','6')"
			else
				sqlsearch = sqlsearch + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
		end if

		if FRectCode<>"" then
			if (InStr(FRectCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and m.code in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and m.code = '" + FRectCode + "'"
			end if
		end if

		if FRectSocID<>"" then
			sqlsearch = sqlsearch + " and m.socid='" + trim(FRectSocID) + "'"
		end if

		if FRectCodeGubun<>"" then
			sqlsearch = sqlsearch + " and Left(m.code,2)='" + FRectCodeGubun + "'"
		end if

		if FRectDivCode<>"" then
			sqlsearch = sqlsearch + " and m.divcode='" + FRectDivCode + "'"
		end if

		if FRectChulgoState="0" then
			sqlsearch = sqlsearch + " and  m.statecd = 0 "
		elseif FRectChulgoState="1" then
		    sqlsearch = sqlsearch + " and  m.statecd = 1 "
		elseif FRectChulgoState="7" then
			sqlsearch = sqlsearch + " and ( m.executedt is not NULL or  m.statecd = 7 ) "
		end if

		if FRectExecuteDtStart<>"" then
			sqlsearch = sqlsearch + " and m.executedt>='" + FRectExecuteDtStart + "'"
		end if

		if FRectExecuteDtEnd<>"" then
			sqlsearch = sqlsearch + " and m.executedt<'" + FRectExecuteDtEnd + "'"
		end if

		if FRectScheduleDtStart<>"" then
			sqlsearch = sqlsearch + " and m.scheduledt>='" + FRectScheduleDtStart + "'"
		end if

		if FRectScheduleDtEnd<>"" then
			sqlsearch = sqlsearch + " and m.scheduledt<'" + FRectScheduleDtEnd + "'"
		end if

		if FRectRackipgoyn<>"" then
			sqlsearch = sqlsearch + " and m.rackipgoyn='" + FRectRackipgoyn + "'"
		end if

		if FRectReturnyn<>"" then
		        if (FRectReturnyn = "N") then
		                sqlsearch = sqlsearch + " and m.totalsellcash>=0 "
		        else
		                sqlsearch = sqlsearch + " and m.totalsellcash<0 "
		        end if
		end if

		if FRectOnOffGubun="on" then
			sqlsearch = sqlsearch + " and m.divcode<>'801'"
			sqlsearch = sqlsearch + " and m.divcode<>'802'"
		elseif FRectOnOffGubun="off" then
			sqlsearch = sqlsearch + " and m.divcode in ('801','802')"
		end if

         IF (FRectPCuserDiv<>"") then
		    sqlsearch = sqlsearch + " and p.userdiv='"&splitValue(FRectPCuserDiv,"_",0)&"'"
		    sqlsearch = sqlsearch + " and c.userdiv='"&splitValue(FRectPCuserDiv,"_",1)&"'"
		end if

		IF (FRectChargename <> "") THEN
		     sqlsearch = sqlsearch + " and m.chargename='"& trim(FRectChargename) &"'"
	    END IF

	    if (FtplGubun <> "") then
			if (FtplGubun = "3X") then
				sqlsearch = sqlsearch + " 	and IsNull(p.tplcompanyid, '') = '' "
			else
				sqlsearch = sqlsearch + " 	and IsNull(p.tplcompanyid, '') = '" + CStr(FtplGubun) + "' "
			end if
		end If

		if (FRectReportstate = "0" ) then
		    sqlsearch = sqlsearch + " and ( ep.reportstate is null or ep.reportstate ='' ) "
	    elseif (FRectReportstate <> "" ) then
	        sqlsearch = sqlsearch + " and   ep.reportstate  = "+ CStr(FRectReportstate)
	    end if
		if (trim(FRectSearchType) <> "") and (trim(FRectSearchText) <> "") then
			Select Case trim(FRectSearchType)
				Case "socname"
					sqlsearch = sqlsearch + " and m.socid in ( "
					sqlsearch = sqlsearch + " 	select distinct p1.id "
					sqlsearch = sqlsearch + " 	from "
					sqlsearch = sqlsearch + " 		db_partner.dbo.tbl_partner p1 with (nolock)"
					sqlsearch = sqlsearch + " 		join [db_partner].[dbo].tbl_partner_group g1 with (nolock)"
					sqlsearch = sqlsearch + " 		on "
					sqlsearch = sqlsearch + " 			p1.groupid = g1.groupid "
					sqlsearch = sqlsearch + " 	where "
					sqlsearch = sqlsearch + " 		1 = 1 "
					sqlsearch = sqlsearch + " 		and g1.company_name like '%" & trim(FRectSearchText) & "%' "
					sqlsearch = sqlsearch + " 		and p1.isusing = 'Y' "
					sqlsearch = sqlsearch + " ) "
				Case "socno"
					sqlsearch = sqlsearch + " and m.socid in ( "
					sqlsearch = sqlsearch + " 	select distinct p1.id "
					sqlsearch = sqlsearch + " 	from "
					sqlsearch = sqlsearch + " 		db_partner.dbo.tbl_partner p1 with (nolock)"
					sqlsearch = sqlsearch + " 		join [db_partner].[dbo].tbl_partner_group g1 with (nolock)"
					sqlsearch = sqlsearch + " 		on "
					sqlsearch = sqlsearch + " 			p1.groupid = g1.groupid "
					sqlsearch = sqlsearch + " 	where "
					sqlsearch = sqlsearch + " 		1 = 1 "
					sqlsearch = sqlsearch + " 		and replace(g1.company_no,'-','') = '" & trim(replace(FRectSearchText,"-","")) & "' "
					sqlsearch = sqlsearch + " 		and p1.isusing = 'Y' "
					sqlsearch = sqlsearch + " ) "
				Case Else
					''
			End Select
		end if
		if FRectNotalinkcode<>"" then
		    sqlsearch = sqlsearch + " and isnull(nj.alinkcode,'')=''"
		end if

		if FRectComment<>"" then
		    sqlsearch = sqlsearch + " and m.comment like '%" & Replace(FRectComment, "'", "") & "%' "
		end if

		sqlStr = " select count(m.id) as cnt "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.id = ep.scmlinkNo and ep.isUsing =1   and ep.edmsidx = 58 "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].[tbl_partner] as pp with (nolock) on m.socid = pp.id "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and pp.purchasetype=pc.pcomm_cd"
		sqlStr = sqlStr + " left Join [db_storage].dbo.tbl_ordersheet_master j with (nolock)"
		sqlStr = sqlStr + " 	on m.code = j.blinkcode"
		sqlStr = sqlStr + " 	and j.deldt is Null"

		if (FRectBrandPurchaseType<>"" and  FRectPCuserDiv = "" and FtplGubun = "" ) then
		    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
		    sqlStr = sqlStr + " 	on m.socid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
'			if FRectBrandPurchaseType = "101" then
'				sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
'			elseif FRectBrandPurchaseType = "102" then
'				sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
'			else
'				sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
'			end if
		elseif (FRectPCuserDiv<>"" or FtplGubun <> "") then
		    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
		    sqlStr = sqlStr + " on m.socid=p.id"

		    if FRectBrandPurchaseType<>"" then
				'/일반유통(101)제외. 일반유통 코드값(1)
				if FRectBrandPurchaseType = "101" then
					sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
				' 전략상품만(3 PB / 5 ODM / 6 수입)
				elseif FRectBrandPurchaseType = "102" then
					sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
				else
					sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
				end if
		    end if

		    sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c with (nolock)"
		    sqlStr = sqlStr + " on c.userid=p.id"
		end if

		if (FRectALinkCode<>"") then
		    if (InStr(FRectALinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectALinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectALinkCode + "'"
			end if
		elseif (FRectBLinkCode<>"") then
			if (InStr(FRectBLinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectBLinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectBLinkCode + "'"
			end if
		end If

		if FRectNotalinkcode<>"" then
		    sqlStr = sqlStr + " left Join [db_storage].dbo.tbl_ordersheet_master nj with (nolock)"
		    sqlStr = sqlStr + " 	on m.code = nj.alinkcode "
		end if

		If (FRectPrcGbn <> "") Then
			If (FRectSocID = "itemgift") And (FRectPrcGbn = "50000") Then
				'// 출고처 itemgift 일 경우만
				sqlStr = sqlStr + " join ( "
				sqlStr = sqlStr + "	 	select m.code, Max(d.sellcash) as sellcash "
				sqlStr = sqlStr + " 	from "
				sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
				sqlStr = sqlStr + " 		join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
				sqlStr = sqlStr + " 		on "
				sqlStr = sqlStr + " 			m.code = d.mastercode "
				sqlStr = sqlStr + " 	where "
				sqlStr = sqlStr + " 		m.socid = 'itemgift' "
				sqlStr = sqlStr + " 	group by "
				sqlStr = sqlStr + " 		m.code "
				sqlStr = sqlStr + " 	having Max(d.sellcash) > 50000 "
				sqlStr = sqlStr + " ) T "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	m.code = T.code "
			End If
		End If

		sqlStr = sqlStr + " where m.deldt is NULL " & sqlsearch

		'response.write sqlStr & "<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " m.id, "
		sqlStr = sqlStr + " m.code,m.socid,m.divcode,m.executedt,m.scheduledt,m.rackipgoyn,"
		sqlStr = sqlStr + " IsNULL(m.totalsellcash,0) as totalsellcash,"
		sqlStr = sqlStr + " IsNULL(m.totalsuplycash,0) as totalsuplycash,"
		sqlStr = sqlStr + " m.vatcode,m.chargeid,m.comment,m.indt,m.updt,m.deldt,"
		sqlStr = sqlStr + " IsNULL(m.totalbuycash,0) as totalbuycash,"
		sqlStr = sqlStr + " m.socname,m.chargename "
		sqlStr = sqlStr + " , (select top 1 jj.baljucode from [db_storage].dbo.tbl_ordersheet_master jj with (nolock) where alinkcode = m.code) as alinkcode "
		sqlStr = sqlStr + " , (select top 1 jj.baljucode from [db_storage].dbo.tbl_ordersheet_master jj with (nolock) where blinkcode = m.code and jj.idx >= 270000) as blinkcode "
		'sqlStr = sqlStr + " , (select top 1 pp.purchaseType from db_partner.dbo.tbl_partner pp where pp.id = m.socid) as purchaseType "
		sqlStr = sqlStr + " ,pp.purchaseType , pc.pcomm_name as purchasetypename"
		sqlStr = sqlStr + " , m.statecd , ep.reportidx, ep.reportstate "
		sqlStr = sqlStr + " , m.finishid , m.finishname "
		sqlStr = sqlStr & " , isnull((select sum(itemno) from [db_storage].[dbo].tbl_acount_storage_detail with (nolock) where isnull(deldt,'')='' and m.code = mastercode),0) as totalitemno" & vbcrlf
		sqlStr = sqlStr & " , isnull((select sum(itemno)" & vbcrlf
		sqlStr = sqlStr & " 	from [db_storage].[dbo].tbl_acount_storage_detail dd" & vbcrlf
		sqlStr = sqlStr & " 	where dd.mastercode = m.code" & vbcrlf
		sqlStr = sqlStr & " 	and dd.deldt is null),0) as totalitemno" & vbcrlf
		sqlStr = sqlStr + " , ("
		sqlStr = sqlStr + " 	select top 1"
		sqlStr = sqlStr + " 	pl.ppmasteridx"
		sqlStr = sqlStr + " 	from [db_storage].[dbo].[tbl_pp_product_master] pm with (nolock)"
		sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_pp_product_link] pl with (nolock)"
		sqlStr = sqlStr + " 		on pm.idx=pl.ppMasterIdx"
		sqlStr = sqlStr + " 	where pm.deldt is null"
		sqlStr = sqlStr + " 	and pl.deldt is null"
		sqlStr = sqlStr + " 	and j.idx=pl.linkIdx"
		sqlStr = sqlStr + " 	order by pl.ppmasteridx desc"
		sqlStr = sqlStr + " 	) as ppmasteridx"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
		sqlStr = sqlStr + " left outer join db_partner.dbo.tbl_eappreport as ep with (nolock) on m.id = ep.scmlinkNo and ep.isUsing =1   and ep.edmsidx = 58 "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].[tbl_partner] as pp with (nolock) on m.socid = pp.id "
		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].tbl_partner_comm_code as pc with (nolock)"
		sqlStr = sqlStr & " 	on pc.pcomm_group='purchasetype' and pc.pcomm_isusing='Y' and pp.purchasetype=pc.pcomm_cd"
		sqlStr = sqlStr + " left Join [db_storage].dbo.tbl_ordersheet_master j with (nolock)"
		sqlStr = sqlStr + " 	on m.code = j.blinkcode"
		sqlStr = sqlStr + " 	and j.deldt is Null"

		if (FRectBrandPurchaseType<>"" and FRectPCuserDiv = "" and FtplGubun = "" ) then
		    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
		    sqlStr = sqlStr + " 	on m.socid=p.id"

			'/일반유통(101)제외. 일반유통 코드값(1)
			if FRectBrandPurchaseType = "101" then
				sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
			' 전략상품만(3 PB / 5 ODM / 6 수입)
			elseif FRectBrandPurchaseType = "102" then
				sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
			else
				sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
			end if
	    elseif (FRectPCuserDiv<>"" or FtplGubun <> "") then
		    sqlStr = sqlStr + " left Join db_partner.dbo.tbl_partner p with (nolock)"
		    sqlStr = sqlStr + " 	on m.socid=p.id"

		    if FRectBrandPurchaseType<>"" then
				'/일반유통(101)제외. 일반유통 코드값(1)
				if FRectBrandPurchaseType = "101" then
					sqlStr = sqlStr + " 	and p.purchasetype <> '1' "
				' 전략상품만(3 PB / 5 ODM / 6 수입)
				elseif FRectBrandPurchaseType = "102" then
					sqlStr = sqlStr & " 	and p.purchasetype in ('3','5','6')"
				else
					sqlStr = sqlStr + " 	and p.purchasetype = '" & FRectBrandPurchaseType & "' "
				end if
		    end if

		    sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c with (nolock)"
		    sqlStr = sqlStr + " on c.userid=p.id"
		end if

		if (FRectALinkCode<>"") then
		    if (InStr(FRectALinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectALinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectALinkCode + "'"
			end if
		elseif (FRectBLinkCode<>"") then
			if (InStr(FRectBLinkCode, vbCrLf) > 0) then
				tmpStr = Replace(FRectBLinkCode, vbCrLf, "','")
				sqlsearch = sqlsearch + " and j.baljucode in ('" + tmpStr + "') "
			else
				sqlsearch = sqlsearch + " and j.baljucode = '" + FRectBLinkCode + "'"
			end if
		end If

		if FRectNotalinkcode<>"" then
		    sqlStr = sqlStr + " left Join [db_storage].dbo.tbl_ordersheet_master nj with (nolock)"
		    sqlStr = sqlStr + " 	on m.code = nj.alinkcode "
		end if

		If (FRectPrcGbn <> "") Then
			If (FRectSocID = "itemgift") And (FRectPrcGbn = "50000") Then
				'// 출고처 itemgift 일 경우만
				sqlStr = sqlStr + " join ( "
				sqlStr = sqlStr + "	 	select m.code, Max(d.sellcash) as sellcash "
				sqlStr = sqlStr + " 	from "
				sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_acount_storage_master m with (nolock)"
				sqlStr = sqlStr + " 		join [db_storage].[dbo].tbl_acount_storage_detail d with (nolock)"
				sqlStr = sqlStr + " 		on "
				sqlStr = sqlStr + " 			m.code = d.mastercode "
				sqlStr = sqlStr + " 	where "
				sqlStr = sqlStr + " 		m.socid = 'itemgift' "
				sqlStr = sqlStr + " 	group by "
				sqlStr = sqlStr + " 		m.code "
				sqlStr = sqlStr + " 	having Max(d.sellcash) > 50000 "
				sqlStr = sqlStr + " ) T "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	m.code = T.code "
			End If
		End If

		sqlStr = sqlStr + " where m.deldt is NULL " & sqlsearch
		sqlStr = sqlStr + " order by m.id desc"

	 	'response.write sqlStr &"<Br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
				set FItemList(i) = new CIpCulmasterItem
				FItemList(i).fpurchasetypename             = rsget("purchasetypename")
				FItemList(i).Fid             = rsget("id")
				FItemList(i).Fcode           = rsget("code")
				FItemList(i).Fsocid          = rsget("socid")
				FItemList(i).Fdivcode        = rsget("divcode")
				FItemList(i).Fexecutedt      = rsget("executedt")
				FItemList(i).Fscheduledt     = rsget("scheduledt")
				FItemList(i).Ftotalsellcash  = rsget("totalsellcash")
				FItemList(i).Ftotalsuplycash = rsget("totalsuplycash")
				FItemList(i).ftotalitemno             = rsget("totalitemno")
				FItemList(i).Fvatcode        = rsget("vatcode")
				FItemList(i).Fchargeid       = rsget("chargeid")
				FItemList(i).Fcomment        = db2html(rsget("comment"))
				FItemList(i).Findt           = rsget("indt")
				FItemList(i).Fupdt           = rsget("updt")
				FItemList(i).Fdeldt          = rsget("deldt")
				FItemList(i).Ftotalbuycash   = rsget("totalbuycash")
				FItemList(i).Fsocname        = db2html(rsget("socname"))
				FItemList(i).Fchargename     = db2html(rsget("chargename"))
				FItemList(i).Frackipgoyn     = rsget("rackipgoyn")
				FItemList(i).Falinkcode     = rsget("alinkcode")
				FItemList(i).Fblinkcode     = rsget("blinkcode")
				FItemList(i).FpurchaseType  = rsget("purchaseType")
 				FItemList(i).Fstatecd  		= rsget("statecd")
 				FItemList(i).Freportidx		= rsget("reportidx")
 				FItemList(i).Freportstate	= rsget("reportstate")
				FItemList(i).Ffinishid		= rsget("finishid")
				FItemList(i).Ffinishname	= rsget("finishname")
				FItemList(i).ftotalitemno	= rsget("totalitemno")
                FItemList(i).FppMasterIdx	= rsget("ppMasterIdx")
				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
	end Sub

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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

Class COffChulgoJungsanItem
	public Fmakerid
	public FItemCnt

	public Ftotalsellcash
	public Ftotalbuycash
	public Ftotalsuplycash
	public Frealjungsansum

	public Fchargediv
	public Fjungsanchargediv
	public Ffranchargediv
	public Fcurrstate

	public function GetCurrStateName()
		if IsNull(Fcurrstate) or (Fcurrstate="") then
			GetCurrStateName = "미정산"
		elseif Fcurrstate="0" then
			GetCurrStateName = "수정중"
		elseif Fcurrstate="1" then
			GetCurrStateName = "업체확인중"
		elseif Fcurrstate="2" then
			GetCurrStateName = "업체확인완료"
		elseif Fcurrstate="3" then
			GetCurrStateName = "정산확정"
		elseif Fcurrstate="7" then
			GetCurrStateName = "입금완료"
		elseif Fcurrstate="8" then
			GetCurrStateName = "정산안함"
		elseif Fcurrstate="9" then
			GetCurrStateName = "통합정산"
		end if
	end function

	public function GetStateColor()
		if IsNull(Fcurrstate) or (Fcurrstate="") then
			GetStateColor = "#000000"
		elseif Fcurrstate="0" then
			GetStateColor = "#000000"
		elseif Fcurrstate="1" then
			GetStateColor = "#448888"
		elseif Fcurrstate="2" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="3" then
			GetStateColor = "#0000FF"
		elseif Fcurrstate="7" then
			GetStateColor = "#FF0000"
		elseif Fcurrstate=" " then
			GetStateColor = "#AAAAAA"
		else

		end if
	end function

	public function getChargeDivName()
		if FChargeDiv="2" then
			getChargeDivName = "10x10 위탁"
		elseif FChargeDiv="4" then
			getChargeDivName = "10x10 매입"
		elseif FChargeDiv="6" then
			getChargeDivName = "업체 위탁"
		elseif FChargeDiv="8" then
			getChargeDivName = "업체 매입"
		elseif FChargeDiv="9" then
			getChargeDivName = "가맹점"
		elseif FChargeDiv="0" then
			getChargeDivName = "통합"
		else
			getChargeDivName = FChargeDiv
		end if
	end function

	public function getJungSanChargeDivName()
		if FJungsanChargediv="2" then
			getJungSanChargeDivName = "10x10 위탁"
		elseif FJungsanChargediv="4" then
			getJungSanChargeDivName = "10x10 매입"
		elseif FJungsanChargediv="6" then
			getJungSanChargeDivName = "업체 위탁"
		elseif FJungsanChargediv="8" then
			getJungSanChargeDivName = "업체 매입"
		elseif FJungsanChargediv="9" then
			getJungSanChargeDivName = "가맹점"
		elseif FJungsanChargediv="0" then
			getJungSanChargeDivName = "통합"
		else
			getJungSanChargeDivName = FJungsanChargediv
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class  COffChulgoJungsan
	public FItemList()
	public FOneItem

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectShopid
	public FRectJungsanYYYY
	public FRectJungsanMM


	''-------------------
	public FRectCodeGubun
	public FRectExecuteDtStart
	public FRectExecuteDtEnd
	public FRectScheduleDtStart
	public FRectScheduleDtEnd
	public FRectSocID
	public FRectId
	public FRectStoragecode
	public FRectDivCode
	public FRectMaeipDiv
	public FRectCode
	public FRectChulgoState
	public FRectOnOffGubun
	public FRectItemID

	public sub GetChulgoJungsanList()
		dim sqlStr,i

		sqlStr = sqlStr + " select d.imakerid, sum(d.itemno*-1) as ccnt,"
		sqlStr = sqlStr + " sum(d.itemno*d.sellcash*-1) as totalsellcash,"
		sqlStr = sqlStr + " sum(d.itemno*d.buycash*-1) as totalbuycash,"
		sqlStr = sqlStr + " sum(d.itemno*d.suplycash*-1) as totalsuplycash,"
		sqlStr = sqlStr + " IsNull(j.realjungsansum,0) as realjungsansum, "
		sqlStr = sqlStr + " j.chargediv as jungsanchargediv, j.franchargediv, j.currstate, s.chargediv"

		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d "

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_jungsanmaster j"
		sqlStr = sqlStr + " on j.yyyymm='" + FRectJungsanYYYY + "-" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and d.imakerid=j.jungsanid and j.shopid='" + FRectShopid + "'"

		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer s"
		sqlStr = sqlStr + " on s.shopid='" + FRectShopid + "'"
		sqlStr = sqlStr + " and d.imakerid=s.makerid"

		sqlStr = sqlStr + " where m.code=d.mastercode"
		sqlStr = sqlStr + " and year(m.executedt)='" + FRectJungsanYYYY + "'"
		sqlStr = sqlStr + " and month(m.executedt)='" + FRectJungsanMM + "'"
		sqlStr = sqlStr + " and Left(m.code,2)='SO'"

		if FRectShopid<>"" then
			sqlStr = sqlStr + " and m.socid='" + FRectShopid + "'"
		end if

		sqlStr = sqlStr + " and d.mwgubun='H'"
		sqlStr = sqlStr + " and m.deldt is null"
		sqlStr = sqlStr + " and d.deldt is null"
		sqlStr = sqlStr + " group by d.imakerid, IsNull(j.realjungsansum,0), j.chargediv, j.franchargediv, j.currstate, s.chargediv"

		rsget.Open sqlStr, dbget, 1

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			do until rsget.eof
				set FItemList(i) = new COffChulgoJungsanItem

				FItemList(i).Fmakerid        = rsget("imakerid")
				FItemList(i).FItemCnt        = rsget("ccnt")

				FItemList(i).Ftotalsellcash  = rsget("totalsellcash")
				FItemList(i).Ftotalbuycash   = rsget("totalbuycash")
				FItemList(i).Ftotalsuplycash   = rsget("totalsuplycash")
				FItemList(i).Frealjungsansum      = rsget("realjungsansum")
				FItemList(i).Fjungsanchargediv		= rsget("jungsanchargediv")
				FItemList(i).Ffranchargediv	= rsget("franchargediv")
				FItemList(i).Fcurrstate	= rsget("currstate")

				FItemList(i).Fchargediv	= rsget("chargediv")


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
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

public Function GetDanJongStat(byval v)
	if v="Y" then
		GetDanJongStat="단종"
	elseif v="S" then
		GetDanJongStat="일시품절"
	elseif v="M" then
		GetDanJongStat="MD품절"
	elseif v="N" then
		GetDanJongStat="생산중"
	else
	    GetDanJongStat=v
	end if
End Function

' 수정로그저장		' 2021.03.09 한용민
public Function chulgo_edit_log(masterid,finishid,bigo)
	dim sqlStr

	if masterid="" or isnull(masterid) then exit Function

	sqlStr = "insert into db_storage.dbo.tbl_acount_storage_master_edit_log ("
	sqlStr = sqlStr & " id, code, socid, divcode, executedt, scheduledt, totalsellcash, totalsuplycash, vatcode"
	sqlStr = sqlStr & " , chargeid, comment, indt, updt, deldt, totalbuycash, socname, chargename, ipchulflag"
	sqlStr = sqlStr & " , rackipgoyn, checkusersn, rackipgousersn, statecd, finishid, finishname, bigo, logregdate, logadminid)"
	sqlStr = sqlStr & " 	select"
	sqlStr = sqlStr & " 	id, code, socid, divcode, executedt, scheduledt, totalsellcash, totalsuplycash, vatcode"
	sqlStr = sqlStr & " 	, chargeid, comment, indt, updt, deldt, totalbuycash, socname, chargename, ipchulflag"
	sqlStr = sqlStr & " 	, rackipgoyn, checkusersn, rackipgousersn, statecd, finishid, finishname, '"& bigo &"', getdate(), '"& finishid &"'"
	sqlStr = sqlStr & " 	from [db_storage].[dbo].tbl_acount_storage_master ml with (nolock)"
	sqlStr = sqlStr & " 	where id='"& masterid &"'"

	'response.write sqlStr & "<br>"
	dbget.Execute sqlStr
End Function
%>

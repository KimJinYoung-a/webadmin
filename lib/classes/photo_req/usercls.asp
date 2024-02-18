<%

class COffContractInfo
	public FItemList()
	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectDesignerID

	public function GetSpecialChargeDivName(byval ishopid)
		dim i

		for i=LBound(FItemList) to UBound(FItemList)
			if IsObject(FItemList(i)) then
				if FItemList(i).Fshopid=ishopid then
					GetSpecialChargeDivName = FItemList(i).GetChargeDivName
				end if
			end if
		next

	end function

	public function GetSpecialDefaultMargin(byval ishopid)
		dim i

		for i=LBound(FItemList) to UBound(FItemList)
			if IsObject (FItemList(i)) then
				if FItemList(i).Fshopid=ishopid then
					GetSpecialDefaultMargin = FItemList(i).Fdefaultmargin
				end if
			end if
		next

	end function

	public Sub GetPartnerOffContractInfo()
		dim sqlStr, i
		sqlStr = "select d.shopid,d.chargediv,d.defaultmargin,u.shopname,u.shopdiv"
		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer d,"
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_user u"
		sqlStr = sqlStr + " where d.shopid=u.userid"
		sqlStr = sqlStr + " and u.isusing='Y'"
		sqlStr = sqlStr + " and makerid='" + FRectDesignerID + "'"

		rsget.Open sqlStr,dbget,1


		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)
		i=0
		do until rsget.eof
			set FItemList(i) = new COffContractInfoItem
			FItemList(i).Fshopid = rsget("shopid")
			FItemList(i).Fchargediv = rsget("chargediv")
			FItemList(i).Fdefaultmargin = rsget("defaultmargin")
			FItemList(i).Fshopname = db2html(rsget("shopname"))
			FItemList(i).FShopdiv = rsget("shopdiv")
			i=i+1
			rsget.movenext
		loop
		rsget.Close
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CPartnerGroupItem
	public Fgroupid
	public Fcompany_name
	public Fcompany_no
	public Fceoname
	public Fcompany_uptae
	public Fcompany_upjong
	public Fcompany_zipcode
	public Fcompany_address
	public Fcompany_address2
	public Fcompany_tel
	public Fcompany_fax
	public Freturn_zipcode				'업체 사무실주소로 전용
	public Freturn_address				'업체 사무실주소로 전용
	public Freturn_address2				'업체 사무실주소로 전용
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_date_off     '' 추가.
	public Fjungsan_acctname
	public Fjungsan_acctno
	public Fmanager_name
	public Fmanager_phone
	public Fmanager_hp
	public Fmanager_email
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fregdate
	public Flastupdate

	public FBrandList

	public Fdefaultsongjangdiv

	public Fpopularid
    public FpartnerCnt

    public function getPartnerIdInfoStr()
        getPartnerIdInfoStr = ""
        if IsNULL(Fpopularid) or (Fpopularid="") or (FpartnerCnt<1) then
            exit function
        end if

        getPartnerIdInfoStr = Fpopularid

        if (FpartnerCnt>1) then
            getPartnerIdInfoStr = getPartnerIdInfoStr & " (외 " & CStr(FpartnerCnt-1) & "건)"
        end if
    end function

	public function getBrandList()
		if Right(FBrandList,1)="," then
			getBrandList = Left(FBrandList,Len(FBrandList)-1)
		else
			getBrandList = FBrandList
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CPartnerGroup
	public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectGroupid
	public FrectDesigner
	public Frectconame
    public FRectSocno

	public Sub GetGroupInfoListByBrand
		dim sqlStr,i
		sqlStr = "select count(g.groupid) as cnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group g ,"
		sqlStr = sqlStr + " [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + " where g.groupid=p.groupid"
		if FrectDesigner<>"" then
			sqlStr = sqlStr + " and p.id like '%" + FrectDesigner + "%'"
		end if
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close


		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " g.*, T.popularid, T.cnt as partnerCnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p,"
		sqlStr = sqlStr + " [db_partner].[dbo].tbl_partner_group g"
		sqlStr = sqlStr + "     left join ("
		sqlStr = sqlStr + "     select Max(p.id) as popularid, p.groupid, count(id) cnt  "
		sqlStr = sqlStr + "     from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + "     where p.isusing='Y'"
		sqlStr = sqlStr + "     group by p.groupid"
		sqlStr = sqlStr + "     ) T on g.groupid=T.groupid"
		sqlStr = sqlStr + " where g.groupid=p.groupid"
		if FrectDesigner<>"" then
			sqlStr = sqlStr + " and p.id like '%" + FrectDesigner + "%'"
		end if
		sqlStr = sqlStr + " order by g.groupid"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		''올림.
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerGroupItem
				FItemList(i).Fgroupid         = rsget("groupid")
				FItemList(i).Fcompany_name    = db2html(rsget("company_name"))
				FItemList(i).Fcompany_no      = rsget("company_no")
				FItemList(i).Fceoname         = db2html(rsget("ceoname"))
				FItemList(i).Fcompany_uptae   = db2html(rsget("company_uptae"))
				FItemList(i).Fcompany_upjong  = db2html(rsget("company_upjong"))
				FItemList(i).Fcompany_zipcode = rsget("company_zipcode")
				FItemList(i).Fcompany_address = db2html(rsget("company_address"))
				FItemList(i).Fcompany_address2= db2html(rsget("company_address2"))
				FItemList(i).Fcompany_tel     = rsget("company_tel")
				FItemList(i).Fcompany_fax     = rsget("company_fax")
				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = db2html(rsget("return_address"))
				FItemList(i).Freturn_address2 = db2html(rsget("return_address2"))
				FItemList(i).Fjungsan_gubun   = rsget("jungsan_gubun")
				FItemList(i).Fjungsan_bank    = rsget("jungsan_bank")
				FItemList(i).Fjungsan_date    = rsget("jungsan_date")
				FItemList(i).Fjungsan_date_off= rsget("jungsan_date_off")
				FItemList(i).Fjungsan_acctname= db2html(rsget("jungsan_acctname"))
				FItemList(i).Fjungsan_acctno  = rsget("jungsan_acctno")
				FItemList(i).Fmanager_name    = db2html(rsget("manager_name"))
				FItemList(i).Fmanager_phone   = rsget("manager_phone")
				FItemList(i).Fmanager_hp      = rsget("manager_hp")
				FItemList(i).Fmanager_email   = db2html(rsget("manager_email"))
				FItemList(i).Fdeliver_name    = db2html(rsget("deliver_name"))
				FItemList(i).Fdeliver_phone   = rsget("deliver_phone")
				FItemList(i).Fdeliver_hp      = rsget("deliver_hp")
				FItemList(i).Fdeliver_email   = db2html(rsget("deliver_email"))
				FItemList(i).Fjungsan_name    = db2html(rsget("jungsan_name"))
				FItemList(i).Fjungsan_phone   = rsget("jungsan_phone")
				FItemList(i).Fjungsan_hp      = rsget("jungsan_hp")
				FItemList(i).Fjungsan_email   = db2html(rsget("jungsan_email"))
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")

                FItemList(i).Fpopularid       = rsget("popularid")
                FItemList(i).FpartnerCnt      = rsget("partnerCnt")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public Sub GetGroupInfoList
		dim sqlStr,i
		sqlStr = "select count(groupid) as cnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group"
		sqlStr = sqlStr + " where 1=1"

		if Frectconame<>"" then
			sqlStr = sqlStr + " and company_name like '%" + Frectconame + "%'"
		end if

		if FRectSocno<>"" then
		    sqlStr = sqlStr + " and Replace(company_no,'-','')='" + Replace(FRectSocno,"-","") + "'"
		end if
        
        if FRectGroupid<>"" then
		    sqlStr = sqlStr + " and groupid ='" + FRectGroupid + "'"
		end if
        
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " g.*, T.popularid, T.cnt as partnerCnt "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner_group g"
		sqlStr = sqlStr + " left join ("
		sqlStr = sqlStr + "     select Max(p.id) as popularid, p.groupid, count(id) cnt  "
		sqlStr = sqlStr + "     from [db_partner].[dbo].tbl_partner p"
		sqlStr = sqlStr + "     where p.isusing='Y'"
		sqlStr = sqlStr + "     group by p.groupid"
		sqlStr = sqlStr + "     ) T on g.groupid=T.groupid"
		sqlStr = sqlStr + " where 1=1"

		if Frectconame<>"" then
			sqlStr = sqlStr + " and g.company_name like '%" + Frectconame + "%'"
		end if

		if FRectSocno<>"" then
		    sqlStr = sqlStr + " and Replace(g.company_no,'-','')='" + Replace(FRectSocno,"-","") + "'"
		end if
        
        if FRectGroupid<>"" then
		    sqlStr = sqlStr + " and g.groupid ='" + FRectGroupid + "'"
		end if
		
		sqlStr = sqlStr + " order by g.groupid"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		''올림.
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerGroupItem
				FItemList(i).Fgroupid         = rsget("groupid")
				FItemList(i).Fcompany_name    = db2html(rsget("company_name"))
				FItemList(i).Fcompany_no      = rsget("company_no")
				FItemList(i).Fceoname         = db2html(rsget("ceoname"))
				FItemList(i).Fcompany_uptae   = db2html(rsget("company_uptae"))
				FItemList(i).Fcompany_upjong  = db2html(rsget("company_upjong"))
				FItemList(i).Fcompany_zipcode = rsget("company_zipcode")
				FItemList(i).Fcompany_address = db2html(rsget("company_address"))
				FItemList(i).Fcompany_address2= db2html(rsget("company_address2"))
				FItemList(i).Fcompany_tel     = rsget("company_tel")
				FItemList(i).Fcompany_fax     = rsget("company_fax")
				FItemList(i).Freturn_zipcode  = rsget("return_zipcode")
				FItemList(i).Freturn_address  = db2html(rsget("return_address"))
				FItemList(i).Freturn_address2 = db2html(rsget("return_address2"))
				FItemList(i).Fjungsan_gubun   = rsget("jungsan_gubun")
				FItemList(i).Fjungsan_bank    = rsget("jungsan_bank")
				FItemList(i).Fjungsan_date    = rsget("jungsan_date")
				FItemList(i).Fjungsan_date_off= rsget("jungsan_date_off")
				FItemList(i).Fjungsan_acctname= db2html(rsget("jungsan_acctname"))
				FItemList(i).Fjungsan_acctno  = rsget("jungsan_acctno")
				FItemList(i).Fmanager_name    = db2html(rsget("manager_name"))
				FItemList(i).Fmanager_phone   = rsget("manager_phone")
				FItemList(i).Fmanager_hp      = rsget("manager_hp")
				FItemList(i).Fmanager_email   = db2html(rsget("manager_email"))
				FItemList(i).Fdeliver_name    = db2html(rsget("deliver_name"))
				FItemList(i).Fdeliver_phone   = rsget("deliver_phone")
				FItemList(i).Fdeliver_hp      = rsget("deliver_hp")
				FItemList(i).Fdeliver_email   = db2html(rsget("deliver_email"))
				FItemList(i).Fjungsan_name    = db2html(rsget("jungsan_name"))
				FItemList(i).Fjungsan_phone   = rsget("jungsan_phone")
				FItemList(i).Fjungsan_hp      = rsget("jungsan_hp")
				FItemList(i).Fjungsan_email   = db2html(rsget("jungsan_email"))
				FItemList(i).Fregdate         = rsget("regdate")
				FItemList(i).Flastupdate      = rsget("lastupdate")

                FItemList(i).Fpopularid       = rsget("popularid")
                FItemList(i).FpartnerCnt      = rsget("partnerCnt")
				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public Sub GetOneGroupInfo
		dim sqlStr
		sqlStr = "select top 1 * from [db_partner].[dbo].tbl_partner_group"
		sqlStr = sqlStr + " where groupid='" + html2db(FRectGroupid) + "'"

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new CPartnerGroupItem
		if Not rsget.Eof then
			FOneItem.Fgroupid         = db2html(rsget("groupid"))
			FOneItem.Fcompany_name    = db2html(rsget("company_name"))
			FOneItem.Fcompany_no      = rsget("company_no")
			FOneItem.Fceoname         = db2html(rsget("ceoname"))
			FOneItem.Fcompany_uptae   = db2html(rsget("company_uptae"))
			FOneItem.Fcompany_upjong  = db2html(rsget("company_upjong"))
			FOneItem.Fcompany_zipcode = rsget("company_zipcode")
			FOneItem.Fcompany_address = db2html(rsget("company_address"))
			FOneItem.Fcompany_address2= db2html(rsget("company_address2"))
			FOneItem.Fcompany_tel     = rsget("company_tel")
			FOneItem.Fcompany_fax     = rsget("company_fax")
			FOneItem.Freturn_zipcode  = rsget("return_zipcode")					'업체 사무실 주소로 전용
			FOneItem.Freturn_address  = db2html(rsget("return_address"))
			FOneItem.Freturn_address2 = db2html(rsget("return_address2"))
			FOneItem.Fjungsan_gubun   = rsget("jungsan_gubun")
			FOneItem.Fjungsan_bank    = rsget("jungsan_bank")
			FOneItem.Fjungsan_date    = rsget("jungsan_date")
			FOneItem.Fjungsan_date_off= rsget("jungsan_date_off")
			FOneItem.Fjungsan_acctname= db2html(rsget("jungsan_acctname"))
			FOneItem.Fjungsan_acctno  = rsget("jungsan_acctno")
			FOneItem.Fmanager_name    = db2html(rsget("manager_name"))
			FOneItem.Fmanager_phone   = rsget("manager_phone")
			FOneItem.Fmanager_hp      = rsget("manager_hp")
			FOneItem.Fmanager_email   = db2html(rsget("manager_email"))
			FOneItem.Fdeliver_name    = db2html(rsget("deliver_name"))
			FOneItem.Fdeliver_phone   = rsget("deliver_phone")
			FOneItem.Fdeliver_hp      = rsget("deliver_hp")
			FOneItem.Fdeliver_email   = db2html(rsget("deliver_email"))
			FOneItem.Fjungsan_name    = db2html(rsget("jungsan_name"))			'정산담당자
			FOneItem.Fjungsan_phone   = rsget("jungsan_phone")
			FOneItem.Fjungsan_hp      = rsget("jungsan_hp")
			FOneItem.Fjungsan_email   = db2html(rsget("jungsan_email"))
			FOneItem.Fregdate         = rsget("regdate")
			FOneItem.Flastupdate      = rsget("lastupdate")
		end if
		rsget.close

        dim bufStr
		if FOneItem.Fgroupid<>"" then
			sqlStr = "select id, isusing from [db_partner].[dbo].tbl_partner"
			sqlStr = sqlStr + " where groupid='" + FRectGroupid + "'"
			'sqlStr = sqlStr + " and isusing='Y'"
			rsget.Open sqlStr,dbget,1

				do until rsget.eof
				    if rsget("isusing")="Y" then
				        bufStr = rsget("id")
				    else
				        bufStr = "<font color='#BBBBBB'>" & rsget("id") & "</font>"
				    end if

					FOneItem.FBrandList = FOneItem.FBrandList + bufStr + ","
					rsget.movenext
				loop
			rsget.close
		end if
	end sub

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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
end class

class COutBrandItem
	public Fyyyymm
	public Fmakerid
	public Fmakername
	public Fmakerlevel
	public Fnewitemcount

	public Flastonjungsansum
	public Flastoffjungsansum
	public Flastminuscnt
	public Flastminussum

	public Fusingitemcount
	public Fregdate

	public Fbrandregdate
	public Fmaeipdiv
	public Fdefaultmargine
	public Femail

	public Fisusing
	public Fisextusing
	public Fstreetusing
	public Fextstreetusing
	public Fspecialbrand

	public Fcurrentusingitemcnt
    public Foffcurrentusingitemcnt
	public Fmduserid
    public Fpartnerusing

	public function GetMWUName()
		if Fmaeipdiv="M" then
			GetMWUName = "매입"
		elseif Fmaeipdiv="W" then
			GetMWUName = "위탁"
		elseif Fmaeipdiv="U" then
			GetMWUName = "업체"
		end if
	end function

	public function GetMwName()
		if Fmaeipdiv="M" then
			GetMwName = "매입"
		elseif Fmaeipdiv="W" then
			GetMwName = "위탁"
		end if
	end function

	public function GetMwColor()
		if Fmaeipdiv="M" then
			GetMwColor = "#FF4444"
		elseif Fmaeipdiv="W" then
			GetMwColor = "#4444FF"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CDesignerUserItem
	public FUserID
	public FSocNo
	public FSocName
	public FSocMail
	public FSocUrl
	public FSocPhone
	public FSocFax
	public FIsUsing
	public FIsB2B
	public FUserDiv
	public FUserDivName
	public FIsExtUsing


	public function Is10x10Using()
		Is10x10Using = false
		if IsNull(FIsUsing) or FIsUsing="N" then Exit function
		if FIsUsing="Y" then
			Is10x10Using = true
		end if
	end function

	public function IsExtUsing()
		IsExtUsing = false
		if IsNull(FIsExtUsing) or FIsExtUsing="N" then Exit function
		if FIsExtUsing="Y" then
			IsExtUsing = true
		end if
	end function

	public function IsB2B()
		IsB2B = false
		if IsNull(FIsB2B) or FIsB2B="N" then Exit function
		if FIsB2B="Y" then
			IsB2B = true
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CPartnerUserItem
	public Fid
	public Fpassword
	public Fdiscountrate
	public Fcompany_name
	public Faddress
	public Ftel
	public Ffax
	public Fbigo
	public Furl
	public Fmanager_name
	public Fmanager_address
	public Fcommission
	public Femail
	public Fuserdiv
	public Fcatecode
	public Fisusing
	public Fbuseo
	public Fpart
	public Fcposition
	public Fintro
	public Fmsn
	public Fbirthday
	public Fuserimg


	public Fonlyflg
	public Fartistflg
	public Fkdesignflg


	public FVatinclude
	public Fmaeipdiv		'기본계약구분
	public Fdefaultmargine	'기본마진
	public FM_margin		'별도 매입시마진
	public FW_margin		'별도 위탁시마진
	public FU_margin		'별도 업체배송시마진

	public Fpid

	public Fcompany_no
	public Fzipcode
	public Fceoname
	public Fmanager_phone
	public Fmanager_hp
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_date_off
	public Fjungsan_date_frn

	public Fjungsan_acctname
	public Fjungsan_acctno

	public Fcompany_upjong
	public Fcompany_uptae

	public FGroupId
	public FSubId
	public Fppass

	public Fsocname
	public Fsocname_kor

	public Fisextusing
	public Fspecialbrand
	public FPrtIdx

	public Fstreetusing
	public Fextstreetusing
	public FTotalitemcount
	Public Fsocicon
	public Fsoclog
	public Ftitleimgurl
	public Fdgncomment
	public Fsamebrand

	public Fmduserid
	public Fregdate

	public Fpartnerusing

    public FdefaultFreeBeasongLimit         ''기본무료배송비기준
    public FdefaultDeliverPay               ''기본배송비
    public FdefaultDeliveryType             ''기본배송정책

    public Fdefaultsongjangdiv

    public Ftakbae_name
    public Ftakbae_tel

    public Flec_yn
    public Fdiy_yn
    public Flec_margin
    public Fmat_margin
    public Fdiy_margin
    public Fitemid


	public function getSocIconUrl()
		IF application("Svr_Info") = "Dev" THEN
			getSocIconUrl = "http://testwebimage.10x10.co.kr/image/brandicon/" + Fsocicon
		Else
			getSocIconUrl = "http://webimage.10x10.co.kr/image/brandicon/" + Fsocicon
		End If
	end function

	public function getSocLogoUrl()
		IF application("Svr_Info") = "Dev" THEN
			getSocLogoUrl = "http://testwebimage.10x10.co.kr/image/brandlogo/" + Fsoclog
		Else
			getSocLogoUrl = "http://webimage.10x10.co.kr/image/brandlogo/" + Fsoclog
		End If
	end function

	public function getTitleImgUrl()
		IF application("Svr_Info") = "Dev" THEN
			getTitleImgUrl = "http://testwebimage.10x10.co.kr/image/brandlogo/" + Ftitleimgurl
		Else
			getTitleImgUrl = "http://webimage.10x10.co.kr/image/brandlogo/" + Ftitleimgurl
		End If
	end function

	public function getRackCode()
		getRackCode = format00(4,FPrtIdx)
	end function

	public function GetBrandDivName()
		if Fuserdiv="02" then
			GetBrandDivName = "디자인업체"
		elseif Fuserdiv="03" then
			GetBrandDivName = "플라워업체"
		elseif Fuserdiv="04" then
			GetBrandDivName = "패션업체"
		elseif Fuserdiv="05" then
			GetBrandDivName = "쥬얼리업체"
		elseif Fuserdiv="06" then
			GetBrandDivName = "케어업체"
		elseif Fuserdiv="07" then
			GetBrandDivName = "애견업체"
		elseif Fuserdiv="08" then
			GetBrandDivName = "보드게임"
		elseif Fuserdiv="13" then
			GetBrandDivName = "여행몰업체"
		elseif Fuserdiv="14" then
			GetBrandDivName = "강사"
		else
			GetBrandDivName = Fuserdiv
		end if
	end function

	public function GetUserDivName()
		if Fuserdiv="02" then
			GetUserDivName = "매입처"
		elseif Fuserdiv="14" then
			GetUserDivName = "강사"
		elseif Fuserdiv="21" then
			GetUserDivName = "출고처"
		elseif Fuserdiv="95" then
			GetUserDivName = "사용안함"
		else
			GetUserDivName = Fuserdiv
		end if
	end function

	public function GetCateCodeName()
		if Fcatecode="10" then
			GetCateCodeName = "문구,사무"
		elseif Fcatecode="15" then
			GetCateCodeName = "인테리어,리빙데코"
		elseif Fcatecode="20" then
			GetCateCodeName = "취미,여가,뷰티"
		elseif Fcatecode="25" then
			GetCateCodeName = "주방,욕실,생활"
		elseif Fcatecode="30" then
			GetCateCodeName = "패션,잡화"
		elseif Fcatecode="35" then
			GetCateCodeName = "보석,액세서리"
		elseif Fcatecode="40" then
			GetCateCodeName = "키덜트,얼리"
		elseif Fcatecode="45" then
			GetCateCodeName = "선물샵"
		elseif Fcatecode="50" then
			GetCateCodeName = "플라워샵"
		elseif Fcatecode="94" then
			GetCateCodeName = "아카데미 DIY"
		elseif Fcatecode="95" then
			GetCateCodeName = "아카데미 강좌"
		else
			GetCateCodeName = Fuserdiv
		end if
	end function

	public function GetMWUName()
		if Fmaeipdiv="M" then
			GetMWUName = "매입"
		elseif Fmaeipdiv="W" then
			GetMWUName = "위탁"
		elseif Fmaeipdiv="U" then
			GetMWUName = "업체배송"
		end if
	end function

	public function GetMaeipDivName()
		if Fmaeipdiv="M" then
			GetMaeipDivName = "매입"
		elseif Fmaeipdiv="W" then
			GetMaeipDivName = "위탁"
		end if
	end function

	''// 업체별 배송비 부과 상품(업체 조건 배송)
	public Function IsUpcheParticleDeliverItem()
	    IsUpcheParticleDeliverItem = (FDefaultFreeBeasongLimit>0) and (FDefaultDeliverPay>0) and (FdefaultDeliveryType="9")
	end function

	''// 업체착불 배송여부
	public Function IsUpcheReceivePayDeliverItem()
	    IsUpcheReceivePayDeliverItem = (FdefaultDeliveryType="7")
	end function

	'// 무료 배송 여부
	public Function IsFreeBeasong()

		if (FdefaultDeliveryType="2") or (FdefaultDeliveryType="4") or (FdefaultDeliveryType="5") then
			IsFreeBeasong = true
		end if

		''//착불 배송은 무료배송이 아님
		if (FdefaultDeliveryType="7") then
		    IsFreeBeasong = false
		end if
	end Function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CPartnerUser
	public FPartnerList()
	public FOneItem

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectDesignerID
	public FRectDesignerName
	public FRectDesignerDiv
	public FRectIsUsing
	public FRectIsB2BUsing
	public FRectIsExtUsing

	public FRectOrder

    public FPass_yn
	public Fpassword

	public Fcompany_name
	public Faddress
	public Ftel
	public Ffax
	public Fbigo
	public Furl
	public Fmanager_name
	public Fmanager_address
	public Fcommission
	public Femail
	public Fuserdiv
	public Fcatecode

	public Fisusing
	public Fisextusing
	public Fstreetusing
	public Fextstreetusing
	public Fspecialbrand

	public FVatinclude
	public Fmaeipdiv
	public Fdefaultmargine

	public Fpid

	public Fcompany_no
	public Fzipcode
	public Fceoname
	public Fmanager_phone
	public Fmanager_hp
	public Fdeliver_name
	public Fdeliver_phone
	public Fdeliver_hp
	public Fdeliver_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Fjungsan_gubun
	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_acctname
	public Fjungsan_acctno

	public Fcompany_upjong
	public Fcompany_uptae

	public FRectInitial
	public FRectUserDiv
	public FRectYYYYMM
	public FRectUserDivUnder
	public FRectMdUserID
	public FRectCatecode
	public FRectmakerlevel
	public FRectCompanyName
	public FRectManagerName
	
	public FRectCompanyNo
    public FRectSOCName
    public Fitemid
    
	Private Sub Class_Initialize()
		'redim preserve FPartnerList(0)
		redim  FPartnerList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub GetGroupList()
		dim sqlStr

	end sub

    ''201010 추가
    public sub GetAcademyPartnerList()
        dim sqlStr, i

        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " c.userid, "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " p.company_name,c.regdate,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.isusing as partnerusing,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass"
		sqlStr = sqlStr + " ,U.lec_yn, U.diy_yn, U.lec_margin, U.mat_margin, U.diy_margin, U.diy_dlv_gubun, U.defaultFreeBeasongLimit, U.defaultDeliveryPay "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p "
		sqlStr = sqlStr + "     on c.userid=p.id"
		sqlStr = sqlStr + "     left join [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U "
		sqlStr = sqlStr + "     on c.userid=U.lecturer_id"
		sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
		sqlStr = sqlStr + " and c.userdiv ='14'" + vbCrlf
		
        if FRectIsUsing="on" then
            sqlStr = sqlStr + " and c.isusing='Y'"
        end if

		if FRectMdUserID<>"" then
			sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
		end if

		if FRectInitial<>"" then
			sqlStr = sqlStr + " and (c.userid like '" + FRectInitial + "%')"
		end if

        if (FRectDesignerID<>"") then
            sqlStr = sqlStr + " and c.userid='"&FRectDesignerID&"'"
        end if

		sqlStr = sqlStr + " order by c.userid asc"

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.recordCount
		FTotalCount = FResultCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set FPartnerList(i) = new CPartnerUserItem
			FPartnerList(i).Fid    			= db2html(rsget("userid"))
			FPartnerList(i).Fcompany_name  	= db2html(rsget("company_name"))
			FPartnerList(i).Faddress        = db2html(rsget("address"))
			FPartnerList(i).Ftel            = rsget("tel")
			FPartnerList(i).Ffax            = rsget("fax")
			FPartnerList(i).Furl            = rsget("url")
			FPartnerList(i).Fmanager_name   = db2html(rsget("manager_name"))
			FPartnerList(i).Fmanager_address  = db2html(rsget("manager_address"))
			FPartnerList(i).Femail          = db2html(rsget("email"))

			FPartnerList(i).FVatinclude     = rsget("vatinclude")
			FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
			FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")
			FPartnerList(i).Fpid			= rsget("pid")
			'oneitem.Fisusing          = rsget("isusing")

			FPartnerList(i).Fcompany_no		= rsget("company_no")
			FPartnerList(i).Fzipcode		= rsget("zipcode")
			FPartnerList(i).Fceoname		= db2html(rsget("ceoname"))
			FPartnerList(i).Fmanager_phone	= rsget("manager_phone")
			FPartnerList(i).Fmanager_hp		= rsget("manager_hp")
			FPartnerList(i).Fdeliver_name	= rsget("deliver_name")
			FPartnerList(i).Fdeliver_phone	= rsget("deliver_phone")
			FPartnerList(i).Fdeliver_hp		= rsget("deliver_hp")
			FPartnerList(i).Fdeliver_email	= rsget("deliver_email")
			FPartnerList(i).Fjungsan_name	= db2html(rsget("jungsan_name"))
			FPartnerList(i).Fjungsan_phone	= rsget("jungsan_phone")
			FPartnerList(i).Fjungsan_hp		= rsget("jungsan_hp")
			FPartnerList(i).Fjungsan_email	= rsget("jungsan_email")
			FPartnerList(i).Fjungsan_gubun	= rsget("jungsan_gubun")
			FPartnerList(i).Fjungsan_bank	= rsget("jungsan_bank")
			FPartnerList(i).Fjungsan_date	= rsget("jungsan_date")
			FPartnerList(i).Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FPartnerList(i).Fjungsan_acctno		= rsget("jungsan_acctno")

			FPartnerList(i).Fcompany_upjong = rsget("company_upjong")
			FPartnerList(i).Fcompany_uptae  = rsget("company_uptae")

			FPartnerList(i).FGroupId  = rsget("groupid")
			FPartnerList(i).FSubId  = rsget("subid")
			FPartnerList(i).Fppass  = rsget("ppass")

			FPartnerList(i).Fsocname  = db2html(rsget("socname"))
			FPartnerList(i).Fsocname_kor  = db2html(rsget("socname_kor"))

			FPartnerList(i).Fisusing	 = rsget("isusing")
			FPartnerList(i).Fisextusing	 = rsget("isextusing")
			FPartnerList(i).Fspecialbrand	= rsget("specialbrand")
			FPartnerList(i).FPrtIdx	= rsget("prtidx")

			FPartnerList(i).Fstreetusing  = rsget("streetusing")
			FPartnerList(i).Fextstreetusing  = rsget("extstreetusing")
			FPartnerList(i).Fuserdiv = rsget("userdiv")
			FPartnerList(i).Fcatecode= rsget("catecode")

			FPartnerList(i).Fregdate = rsget("regdate")

            FPartnerList(i).Fpartnerusing = rsget("partnerusing")

            FPartnerList(i).Flec_yn         = rsget("lec_yn")
            FPartnerList(i).Fdiy_yn         = rsget("diy_yn")
            FPartnerList(i).Flec_margin     = rsget("lec_margin")
            FPartnerList(i).Fmat_margin			= rsget("mat_margin")
            FPartnerList(i).Fdiy_margin			= rsget("diy_margin")
			FPartnerList(i).FdefaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
			FPartnerList(i).FdefaultDeliveryType = rsget("diy_dlv_gubun")
            FPartnerList(i).FdefaultDeliverPay	= rsget("defaultDeliveryPay")


			i=i+1
			rsget.movenext
		loop
		rsget.close

    end Sub

	public Sub GetPartnerQuickSearch()
		dim sqlStr


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " c.userid, "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " p.company_name,c.regdate,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.isusing as partnerusing,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
		sqlStr = sqlStr + " and c.userdiv < 22" + vbCrlf

		if FRectMdUserID<>"" then
			sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
		end if

		if FRectInitial<>"" then
			sqlStr = sqlStr + " and (c.userid like '" + FRectInitial + "%')"
		end if

		if FRectUserDiv<>"" then
			sqlStr = sqlStr + " and c.userdiv='" + FRectUserDiv + "'"
		end if

		sqlStr = sqlStr + " order by c.userid asc"

		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		FTotalCount = FResultCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set FPartnerList(i) = new CPartnerUserItem
			FPartnerList(i).Fid    			= db2html(rsget("userid"))
			FPartnerList(i).Fcompany_name  	= db2html(rsget("company_name"))
			FPartnerList(i).Faddress        = db2html(rsget("address"))
			FPartnerList(i).Ftel            = rsget("tel")
			FPartnerList(i).Ffax            = rsget("fax")
			FPartnerList(i).Furl            = rsget("url")
			FPartnerList(i).Fmanager_name   = db2html(rsget("manager_name"))
			FPartnerList(i).Fmanager_address  = db2html(rsget("manager_address"))
			FPartnerList(i).Femail          = db2html(rsget("email"))

			FPartnerList(i).FVatinclude     = rsget("vatinclude")
			FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
			FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")
			FPartnerList(i).Fpid			= rsget("pid")
			'oneitem.Fisusing          = rsget("isusing")

			FPartnerList(i).Fcompany_no		= rsget("company_no")
			FPartnerList(i).Fzipcode		= rsget("zipcode")
			FPartnerList(i).Fceoname		= db2html(rsget("ceoname"))
			FPartnerList(i).Fmanager_phone	= rsget("manager_phone")
			FPartnerList(i).Fmanager_hp		= rsget("manager_hp")
			FPartnerList(i).Fdeliver_name	= rsget("deliver_name")
			FPartnerList(i).Fdeliver_phone	= rsget("deliver_phone")
			FPartnerList(i).Fdeliver_hp		= rsget("deliver_hp")
			FPartnerList(i).Fdeliver_email	= rsget("deliver_email")
			FPartnerList(i).Fjungsan_name	= db2html(rsget("jungsan_name"))
			FPartnerList(i).Fjungsan_phone	= rsget("jungsan_phone")
			FPartnerList(i).Fjungsan_hp		= rsget("jungsan_hp")
			FPartnerList(i).Fjungsan_email	= rsget("jungsan_email")
			FPartnerList(i).Fjungsan_gubun	= rsget("jungsan_gubun")
			FPartnerList(i).Fjungsan_bank	= rsget("jungsan_bank")
			FPartnerList(i).Fjungsan_date	= rsget("jungsan_date")
			FPartnerList(i).Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FPartnerList(i).Fjungsan_acctno		= rsget("jungsan_acctno")

			FPartnerList(i).Fcompany_upjong = rsget("company_upjong")
			FPartnerList(i).Fcompany_uptae  = rsget("company_uptae")

			FPartnerList(i).FGroupId  = rsget("groupid")
			FPartnerList(i).FSubId  = rsget("subid")
			FPartnerList(i).Fppass  = rsget("ppass")

			FPartnerList(i).Fsocname  = db2html(rsget("socname"))
			FPartnerList(i).Fsocname_kor  = db2html(rsget("socname_kor"))

			FPartnerList(i).Fisusing	 = rsget("isusing")
			FPartnerList(i).Fisextusing	 = rsget("isextusing")
			FPartnerList(i).Fspecialbrand	= rsget("specialbrand")
			FPartnerList(i).FPrtIdx	= rsget("prtidx")

			FPartnerList(i).Fstreetusing  = rsget("streetusing")
			FPartnerList(i).Fextstreetusing  = rsget("extstreetusing")
			FPartnerList(i).Fuserdiv = rsget("userdiv")
			FPartnerList(i).Fcatecode= rsget("catecode")

			FPartnerList(i).Fregdate = rsget("regdate")

            FPartnerList(i).Fpartnerusing = rsget("partnerusing")
			i=i+1
			rsget.movenext
		loop
		rsget.close
	end sub


	public Sub GetOutBrandList()
		dim i,sqlstr

		sqlStr = "select top 1000 o.*, c.mduserid, c.regdate as brandregdate, c.maeipdiv, c.defaultmargine"
		sqlStr = sqlStr + " ,c.isusing, c.isextusing, c.streetusing, c.extstreetusing, c.specialbrand"
		sqlStr = sqlStr + " ,IsNULL(T.cnt,0) as currentusingitemcnt"
		sqlStr = sqlStr + " ,IsNULL(T2.cnt,0) as offcurrentusingitemcnt"
		sqlStr = sqlStr + " ,p.isusing as partnerusing"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_outbrand o," + vbCrlf
		sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select makerid, count(itemid) as cnt"
		sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " 	where isusing='Y'"
		sqlStr = sqlStr + "		group by makerid"
		sqlStr = sqlStr + " ) as T on T.makerid=c.userid"
		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " 	select makerid, count(shopitemid) as cnt"
		sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shop_item "
		sqlStr = sqlStr + " 	where isusing='Y'"
		sqlStr = sqlStr + "		group by makerid"
		sqlStr = sqlStr + " ) as T2 on T2.makerid=c.userid"

		sqlStr = sqlStr + " where yyyymm='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and o.makerid=c.userid"
		if FRectIsUsing="on" then
			sqlStr = sqlStr + " and c.isusing='Y'"
		end if

		if FRectCatecode<>"" then
			sqlStr = sqlStr + " and c.catecode='" + FRectCatecode + "'"
		end if

		if FRectMdUserID<>"" then
			sqlStr = sqlStr + " and c.mduserid='" + FRectMdUserID + "'"
		end if

		if FRectmakerlevel<>"" then
			sqlStr = sqlStr + " and o.makerlevel=" + FRectmakerlevel + ""
		end if

		sqlStr = sqlStr + " order by o.makerlevel desc, o.lastonjungsansum "
'response.write sqlStr
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set FPartnerList(i) = new COutBrandItem
			FPartnerList(i).Fyyyymm         = rsget("yyyymm")
			FPartnerList(i).Fmakerid        = rsget("makerid")
			FPartnerList(i).Fmakername      = db2html(rsget("makername"))
			FPartnerList(i).Fmakerlevel     = rsget("makerlevel")
			FPartnerList(i).Fnewitemcount   = rsget("newitemcount")
			FPartnerList(i).Flastonjungsansum = rsget("lastonjungsansum")

            FPartnerList(i).Flastoffjungsansum = rsget("lastoffjungsansum")
   			FPartnerList(i).Flastminuscnt = rsget("lastminuscnt")
   			FPartnerList(i).Flastminussum = rsget("lastminussum")

			FPartnerList(i).Fusingitemcount = rsget("usingitemcount")
			FPartnerList(i).Fregdate        = rsget("regdate")

			FPartnerList(i).Fbrandregdate	= rsget("brandregdate")
			FPartnerList(i).Fmaeipdiv       = rsget("maeipdiv")
			FPartnerList(i).Fdefaultmargine = rsget("defaultmargine")

			'FPartnerList(i).Femail		= rsget("email")
			FPartnerList(i).Fisusing	= rsget("isusing")
			FPartnerList(i).Fisextusing	= rsget("isextusing")

			FPartnerList(i).Fstreetusing 	= rsget("streetusing")
			FPartnerList(i).Fextstreetusing = rsget("extstreetusing")
			FPartnerList(i).Fspecialbrand	= rsget("specialbrand")

			FPartnerList(i).Fcurrentusingitemcnt = rsget("currentusingitemcnt")
            FPartnerList(i).Foffcurrentusingitemcnt = rsget("offcurrentusingitemcnt")
			FPartnerList(i).Fmduserid       = rsget("mduserid")
            FPartnerList(i).Fpartnerusing   = rsget("partnerusing")

			i=i+1
			rsget.movenext
		loop
		rsget.close
	end sub

	public Sub GetOnePartnerNUser()
		dim sqlStr

		sqlStr = "select top 1 c.userid, "
		sqlStr = sqlStr + " c.vatinclude, c.maeipdiv, c.defaultmargine, c.defaultFreeBeasongLimit, c.defaultDeliverPay, c.defaultDeliveryType, c.socname, c.socname_kor,"
		sqlStr = sqlStr + " c.isusing, c.isextusing, c.specialbrand, c.prtidx,c.streetusing,c.extstreetusing,c.userdiv,c.catecode,"
		sqlStr = sqlStr + " c.socicon, c.soclogo, c.titleimgurl,c.dgncomment,c.samebrand, c.mduserid, c.regdate, c.onlyflg, c.artistflg, c.kdesignflg, "
		sqlStr = sqlStr + " IsNull(p.M_margin,0) as M_margin, IsNull(p.W_margin,0) as W_margin, IsNull(p.U_margin,0) as U_margin, "
		sqlStr = sqlStr + " c.socicon, c.soclogo, c.titleimgurl,c.dgncomment, c.mduserid, c.regdate, "
		sqlStr = sqlStr + " p.company_name,"
		sqlStr = sqlStr + " p.email, p.address, p.manager_address,"
		sqlStr = sqlStr + " p.tel, p.fax, p.url, p.manager_name, p.id as pid,"
		sqlStr = sqlStr + " p.company_no, p.zipcode, p.ceoname, p.manager_phone,"
		sqlStr = sqlStr + " p.manager_hp, p.deliver_name, p.deliver_phone, "
		sqlStr = sqlStr + " p.deliver_hp, p.deliver_email, p.jungsan_name, "
		sqlStr = sqlStr + " p.jungsan_phone, p.jungsan_hp, p.jungsan_email,"
		sqlStr = sqlStr + " p.jungsan_gubun, p.jungsan_bank, p.jungsan_date,p.jungsan_date_off, p.jungsan_date_frn,"
		sqlStr = sqlStr + " p.jungsan_acctname, p.jungsan_acctno,"
		sqlStr = sqlStr + " p.company_upjong, p.company_uptae, IsNULL(p.groupid,'') as groupid, p.subid, p.password as ppass, p.isusing as partnerusing,"
		sqlStr = sqlStr + " IsNULL(T.cnt,0) as ttlitemcnt, p.defaultsongjangdiv,"
		sqlStr = sqlStr + " IsNULL(s.divname,'') as takbae_name, IsNULL(s.tel,'') as takbae_tel"
		sqlStr = sqlStr + " ,U.lec_yn, U.diy_yn, U.lec_margin, U.mat_margin, U.diy_margin, U.diy_dlv_gubun, U.defaultFreeBeasongLimit as defaultFreeBeasongLimitAcademy, U.defaultDeliveryPay as defaultDeliveryPayAcademy "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " left join [ACADEMYDB].[db_academy].[dbo].tbl_lec_user U on c.userid=U.lecturer_id"
		sqlStr = sqlStr + " left join ( "
		sqlStr = sqlStr + " select makerid, count(itemid) as cnt from [db_item].[dbo].tbl_item where makerid='" + FRectDesignerID + "'"
		sqlStr = sqlStr + " group by makerid "
		sqlStr = sqlStr + " ) as T on c.userid=T.makerid"
		''택배사 명,전화 추가.
		sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div s"
		sqlStr = sqlStr + "     on p.defaultsongjangdiv=s.divcd"
		sqlStr = sqlStr + " where c.userid='" + FRectDesignerID + "'"
		
		'response.write sqlstr &"<Br>"
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount

		if Not rsget.Eof then
			set FOneItem = new CPartnerUserItem
			FOneItem.Fid    			= db2html(rsget("userid"))
			FOneItem.Fcompany_name  	= db2html(rsget("company_name"))
			FOneItem.Faddress        	= db2html(rsget("address"))
			FOneItem.Ftel            	= rsget("tel")
			FOneItem.Ffax            	= rsget("fax")
			FOneItem.Furl            	= rsget("url")
			FOneItem.Fmanager_name   	= db2html(rsget("manager_name"))
			FOneItem.Fmanager_address  	= db2html(rsget("manager_address"))
			FOneItem.Femail          	= db2html(rsget("email"))

			FOneItem.FVatinclude     	= rsget("vatinclude")
			FOneItem.Fmaeipdiv       	= rsget("maeipdiv")
			FOneItem.Fdefaultmargine 	= rsget("defaultmargine")
			FOneItem.FM_margin 		 	= rsget("M_margin")
			FOneItem.FW_margin       	= rsget("W_margin")
			FOneItem.FU_margin       	= rsget("U_margin")

			FOneItem.Fpid				= rsget("pid")
			'oneitem.Fisusing          	= rsget("isusing")

			FOneItem.Fcompany_no		= rsget("company_no")
			FOneItem.Fzipcode			= rsget("zipcode")
			FOneItem.Fceoname			= db2html(rsget("ceoname"))
			FOneItem.Fmanager_phone		= rsget("manager_phone")
			FOneItem.Fmanager_hp		= rsget("manager_hp")
			FOneItem.Fdeliver_name		= rsget("deliver_name")
			FOneItem.Fdeliver_phone		= rsget("deliver_phone")
			FOneItem.Fdeliver_hp		= rsget("deliver_hp")
			FOneItem.Fdeliver_email		= rsget("deliver_email")
			FOneItem.Fjungsan_name		= db2html(rsget("jungsan_name"))
			FOneItem.Fjungsan_phone		= rsget("jungsan_phone")
			FOneItem.Fjungsan_hp		= rsget("jungsan_hp")
			FOneItem.Fjungsan_email		= rsget("jungsan_email")
			FOneItem.Fjungsan_gubun		= rsget("jungsan_gubun")
			FOneItem.Fjungsan_bank		= rsget("jungsan_bank")
			FOneItem.Fjungsan_date		= rsget("jungsan_date")
			FOneItem.Fjungsan_date_off	= rsget("jungsan_date_off")
			FOneItem.Fjungsan_date_frn	= rsget("jungsan_date_frn")

			FOneItem.Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FOneItem.Fjungsan_acctno	= rsget("jungsan_acctno")

			FOneItem.Fcompany_upjong 	= db2html(rsget("company_upjong"))
			FOneItem.Fcompany_uptae  	= db2html(rsget("company_uptae"))

			FOneItem.FGroupId  			= rsget("groupid")
			FOneItem.FSubId  			= rsget("subid")
			FOneItem.Fppass  			= rsget("ppass")

			FOneItem.Fsocname  			= db2html(rsget("socname"))
			FOneItem.Fsocname_kor  		= db2html(rsget("socname_kor"))

			FOneItem.Fisusing	 		= rsget("isusing")
			FOneItem.Fisextusing	 	= rsget("isextusing")
			FOneItem.Fspecialbrand		= rsget("specialbrand")
			FOneItem.FPrtIdx			= rsget("prtidx")

			FOneItem.Fstreetusing  		= rsget("streetusing")
			FOneItem.Fextstreetusing  	= rsget("extstreetusing")
			FOneItem.Fuserdiv 			= rsget("userdiv")
			FOneItem.Fcatecode			= rsget("catecode")

			FOneItem.FTotalitemcount 	= rsget("ttlitemcnt")
			FOneItem.Fsocicon 			= db2html(rsget("socicon"))
			FOneItem.Fsoclog 			= db2html(rsget("soclogo"))
			FOneItem.Ftitleimgurl 		= db2html(rsget("titleimgurl"))
			FOneItem.Fdgncomment 		= db2html(rsget("dgncomment"))
			FOneItem.Fsamebrand 		= db2html(rsget("samebrand"))
			
			FOneItem.Fmduserid 			= rsget("mduserid")
			FOneItem.Fregdate 			= rsget("regdate")

			FOneItem.Fonlyflg			= rsget("onlyflg")
			FOneItem.Fartistflg			= rsget("artistflg")
			FOneItem.Fkdesignflg		= rsget("kdesignflg")

			FOneItem.Fpartnerusing 		= rsget("partnerusing")

			FOneItem.FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
			FOneItem.FdefaultDeliverPay         = rsget("defaultDeliverPay")
			FOneItem.FdefaultDeliveryType       = rsget("defaultDeliveryType")

			FOneItem.Fdefaultsongjangdiv 		= rsget("defaultsongjangdiv")
			FOneItem.Ftakbae_name 				= db2html(rsget("takbae_name"))
			FOneItem.Ftakbae_tel  				= rsget("takbae_tel")

            FOneItem.Flec_yn         			= rsget("lec_yn")
            FOneItem.Fdiy_yn         			= rsget("diy_yn")
            FOneItem.Flec_margin     			= rsget("lec_margin")
            FOneItem.Fmat_margin				= rsget("mat_margin")
            FOneItem.Fdiy_margin				= rsget("diy_margin")

            if (CStr(FOneItem.Fuserdiv) = "14") then
            	'강사
				FOneItem.FdefaultFreeBeasongLimit 	= rsget("defaultFreeBeasongLimitAcademy")
				FOneItem.FdefaultDeliveryType 		= rsget("diy_dlv_gubun")
	            FOneItem.FdefaultDeliverPay			= rsget("defaultDeliveryPayAcademy")
            end if

		end if
		rsget.close

	end sub


	public Sub GetPartnerNUserCList()
		dim sqlStr
		''#################################################
		''총 갯수.
		''#################################################
		sqlStr = "select Count(userid) as cnt"
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_user_tenbyten where userid <> '' " 
		
		'response.write FRectSOCName

		if FRectInitial<>"" then
			sqlStr = sqlStr + " and userid like '%" + FRectInitial + "%'"
		end If
		
		If FRectSOCName <> "" Then
			sqlstr = sqlstr + " and username like '%" + FRectSOCName + "%'"
		End If 
			sqlstr = sqlstr  + " group by userid, username order by userid asc"

		'response.write  sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount =  rsget.recordCount 
		rsget.Close

		sqlStr = "select  top " + CStr(FPageSize*FCurrPage) + " userid, username "
		sqlStr = sqlStr + " from [db_partner].[dbo].tbl_user_tenbyten where userid <> '' " 
		
		if FRectInitial<>"" then
			sqlStr = sqlStr + " and userid like '%" + FRectInitial + "%' "
		end If
		If FRectSOCName <> "" Then
			sqlstr = sqlstr + " and username like '%" + FRectSOCName + "%'"
		End If 
			sqlstr = sqlstr  + "  group by userid, username order by userid asc"
		
		'response.write sqlstr

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FPartnerList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof

				set FPartnerList(i) = new CPartnerUserItem
				FPartnerList(i).Fid    						= db2html(rsget("userid"))
				FPartnerList(i).Fcompany_name  	= db2html(rsget("username"))

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
	end sub

	public function duplicateUserID(byval userid)
		dim sqlStr
		sqlStr = "select count(id) as cnt from [db_partner].[dbo].tbl_partner"
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		rsget.Open sqlStr,dbget,1
		duplicateUserID = rsget("cnt")>0
		rsget.close
	end function

	public Sub addNewPartner(byval userid,userpass,username,usermail,userdiv, discountrate,commission,bigo)
	    dim sqlStr

		sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + "(id,password,company_name,email,userdiv,discountrate,commission,bigo)" + vbCrlf
		sqlStr = sqlStr + " values('" + userid + "'," + vbCrlf
		sqlStr = sqlStr + " '" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " '" + username + "'," + vbCrlf
		sqlStr = sqlStr + " '" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " " + userdiv + "," + vbCrlf
		sqlStr = sqlStr + " " + discountrate + "," + vbCrlf
		sqlStr = sqlStr + " " + commission + "," + vbCrlf
		sqlStr = sqlStr + " '" + bigo + "'" + vbCrlf
		sqlStr = sqlStr + ")"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
        end sub

        public Sub editPartner(byval userid,userpass,username,usermail,userdiv, isusing, discountrate,commission,bigo)
	        dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set password='" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " company_name='" + username + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " userdiv=" + CStr(userdiv) + "," + vbCrlf
		sqlStr = sqlStr + " isusing='" + isusing + "'," + vbCrlf
		sqlStr = sqlStr + " discountrate=" + discountrate + "," + vbCrlf
		sqlStr = sqlStr + " commission=" + commission + "," + vbCrlf
		sqlStr = sqlStr + " bigo='" + bigo + "'" + vbCrlf

		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
        end sub

	public Sub addNewEmploy(byval userid,userpass,username,usermail,userdiv,bigo)
		dim sqlStr
		sqlStr = "insert into [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + "(id,password,company_name,email,bigo,userdiv)" + vbCrlf
		sqlStr = sqlStr + " values('" + userid + "'," + vbCrlf
		sqlStr = sqlStr + " '" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " '" + username + "'," + vbCrlf
		sqlStr = sqlStr + " '" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " '" + bigo + "'," + vbCrlf
		sqlStr = sqlStr + " " + userdiv + "" + vbCrlf
		sqlStr = sqlStr + ")"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
	end sub

	public Sub editEmploy(byval userid,userpass,username,usermail,userdiv,bigo,isusing)
		dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set password='" + userpass + "'," + vbCrlf
		sqlStr = sqlStr + " company_name='" + username + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + usermail + "'," + vbCrlf
		sqlStr = sqlStr + " userdiv=" + CStr(userdiv) + "," + vbCrlf
		sqlStr = sqlStr + " isusing='" + isusing + "'," + vbCrlf
		sqlStr = sqlStr + " bigo='" + bigo + "'" + vbCrlf
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
	end sub

	public Sub GetOnePartner(byval userid)
		dim sqlStr
		dim oneitem

		sqlStr = "select top 1 * from [db_partner].[dbo].tbl_partner"
		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		if Not rsget.Eof then
			set oneitem = new CPartnerUserItem
			oneitem.Fid               = rsget("id")
			oneitem.Fpassword         = rsget("password")
			oneitem.Fdiscountrate     = rsget("discountrate")
			oneitem.Fcompany_name     = rsget("company_name")
			oneitem.Faddress          = rsget("address")
			oneitem.Ftel              = rsget("tel")
			oneitem.Fmanager_hp    = rsget("manager_hp")
			oneitem.Ffax              = rsget("fax")
			oneitem.Fbigo             = rsget("bigo")
			oneitem.Furl              = rsget("url")
			oneitem.Fmanager_name     = rsget("manager_name")
			oneitem.Fmanager_address  = rsget("manager_address")
			oneitem.Fcommission       = rsget("commission")
			oneitem.Femail            = rsget("email")
			oneitem.Fbirthday     = rsget("birthday")
			oneitem.Fmsn     = rsget("msn")
			oneitem.Fzipcode     = rsget("zipcode")
			oneitem.Fbuseo     = rsget("buseo")
			oneitem.Fpart    = rsget("part")
			oneitem.Fcposition     = rsget("cposition")
			oneitem.Fintro     = rsget("intro")
			oneitem.Fuserimg          = rsget("userimg")
			oneitem.Fuserdiv          = rsget("userdiv")
			oneitem.Fisusing          = rsget("isusing")
			set FPartnerList(0) = oneitem
		end if
		rsget.close
	end Sub

	public Function UpdateDesignerSet(byval idesignerid, isusing, isextusing, isb2b)
		dim sqlStr
		sqlStr = "update [db_user].[dbo].tbl_user_c" + vbCrlf
		sqlStr = sqlStr + " set isusing='" + isusing + "'," + vbCrlf
		sqlStr = sqlStr + " isextusing='" + isextusing + "'," + vbCrlf
		sqlStr = sqlStr + " isb2b='" + isb2b + "'" + vbCrlf
		sqlStr = sqlStr + " where userid='" + idesignerid + "'"

		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
	end function

	public Sub editPartnerDesigner2(byval userid)
	    dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set " + vbCrlf
		if FPass_yn = "" then
		else
		sqlStr = sqlStr + " password='" + Fpassword + "'," + vbCrlf
		end if
		sqlStr = sqlStr + " company_name='" + Fcompany_name + "'," + vbCrlf
		sqlStr = sqlStr + " zipcode='" + Fzipcode + "'," + vbCrlf
		sqlStr = sqlStr + " address='" + Faddress + "'," + vbCrlf
		sqlStr = sqlStr + " manager_address='" + Fmanager_address + "'," + vbCrlf
		sqlStr = sqlStr + " ceoname='" + Fceoname + "'," + vbCrlf
		sqlStr = sqlStr + " company_upjong ='" + Fcompany_upjong + "'," + vbCrlf
		sqlStr = sqlStr + " company_uptae='" + Fcompany_uptae + "'," + vbCrlf
		sqlStr = sqlStr + " company_no='" + Fcompany_no + "'," + vbCrlf
		sqlStr = sqlStr + " url='" + Furl + "'," + vbCrlf
		sqlStr = sqlStr + " tel='" + Ftel + "'," + vbCrlf
		sqlStr = sqlStr + " fax='" + Ffax + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_bank='" + Fjungsan_bank + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_acctno='" + Fjungsan_acctno + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_acctname='" + Fjungsan_acctname + "'," + vbCrlf
		sqlStr = sqlStr + " manager_name='" + Fmanager_name + "'," + vbCrlf
		sqlStr = sqlStr + " manager_phone='" + Fmanager_phone + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + Femail + "'," + vbCrlf
		sqlStr = sqlStr + " manager_hp ='" + Fmanager_hp + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_name ='" + Fdeliver_name + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_phone='" + Fdeliver_phone + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_email='" + Fdeliver_email + "'," + vbCrlf
		sqlStr = sqlStr + " deliver_hp   ='" + Fdeliver_hp + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_name ='" + Fjungsan_name + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_phone='" + Fjungsan_phone + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_email='" + Fjungsan_email + "'," + vbCrlf
		sqlStr = sqlStr + " jungsan_hp	 ='" + Fjungsan_hp + "'" + vbCrlf

		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
    end sub

'==========================================태훈 추가
        public Sub editPartnerDesigner(byval userid)
	        dim sqlStr
		sqlStr = "update [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " set " + vbCrlf
		if FPass_yn = "" then
		else
		sqlStr = sqlStr + " password='" + Fpassword + "'," + vbCrlf
		end if
		sqlStr = sqlStr + " company_name='" + Fcompany_name + "'," + vbCrlf
		sqlStr = sqlStr + " address='" + Faddress + "'," + vbCrlf
		sqlStr = sqlStr + " tel='" + Ftel + "'," + vbCrlf
		sqlStr = sqlStr + " fax='" + Ffax + "'," + vbCrlf
		sqlStr = sqlStr + " url='" + Furl + "'," + vbCrlf
		sqlStr = sqlStr + " manager_name='" + Fmanager_name + "'," + vbCrlf
		sqlStr = sqlStr + " manager_address='" + Fmanager_address + "'," + vbCrlf
		sqlStr = sqlStr + " email='" + Femail + "'" + vbCrlf

		sqlStr = sqlStr + " where id='" + CStr(userid) + "'"

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1
        end sub
'====================================================

	public Sub GetDesignerList()
		dim i,sqlStr,wheredetail
		dim oneitem

		if FRectDesignerID<>"" then
			wheredetail =" and c.userid='" + FRectDesignerID + "'"
		end if

		if FRectDesignerName<>"" then
			wheredetail =" and c.socname like '%" + FRectDesignerName + "'%"
		end if

		if FRectDesignerDiv<>"" then
			wheredetail =" and c.userdiv = '" + FRectDesignerDiv + "'"
		end if

		if FRectIsUsing<>"" then
			wheredetail =" and IsNull(c.isusing,'N') = '" + FRectIsUsing + "'"
		end if

		if FRectIsB2BUsing<>"" then
			wheredetail =" and IsNull(c.isb2b,'N') = '" + FRectIsB2BUsing + "'"
		end if

		if FRectIsExtUsing<>"" then
			wheredetail =" and IsNull(c.isextusing,'N') = '" + FRectIsExtUsing + "'"
		end if

		''#################################################
		''총 갯수.
		''#################################################
		sqlStr = "select Count(userid) as cnt from [db_user].[dbo].tbl_user_c c, [db_user].[dbo].tbl_user_div d" + vbCrlf
		sqlStr = sqlStr + " where c.userdiv=d.divcode" + vbCrlf
		rsget.Open sqlStr + wheredetail,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FPageSize) + " c.*, d.divename from [db_user].[dbo].tbl_user_c c, [db_user].[dbo].tbl_user_div d" + vbCrlf
		sqlStr = sqlStr + " where c.userdiv=d.divcode" + vbCrlf
		sqlStr = sqlStr + " and c.userid not in (" + vbCrlf
		sqlStr = sqlStr + " select top " + CStr((FCurrPage-1)*FPageSize)  + " c.userid from [db_user].[dbo].tbl_user_c c, [db_user].[dbo].tbl_user_div d" + vbCrlf
		sqlStr = sqlStr + " where c.userdiv=d.divcode" + vbCrlf
		sqlStr = sqlStr + wheredetail + vbCrlf
		sqlStr = sqlStr + " )" + vbCrlf
		sqlStr = sqlStr + wheredetail + vbCrlf

		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FPartnerList(FResultCount)
		i=0
		do until rsget.eof
			set oneitem = new CDesignerUserItem
			oneitem.FUserID     = rsget("userid")
			oneitem.FSocNo      = rsget("socno")
			oneitem.FSocName    = rsget("socname")
			oneitem.FSocMail    = rsget("socmail")
			oneitem.FSocUrl     = rsget("socurl")
			oneitem.FSocPhone   = rsget("socphone")
			oneitem.FSocFax     = rsget("socfax")
			oneitem.FIsUsing    = rsget("isusing")
			oneitem.FIsB2B      = rsget("isb2b")
			oneitem.FUserDiv    = rsget("userdiv")
			oneitem.FIsExtUsing = rsget("isextusing")
			oneitem.FUserDivName= rsget("divename")

			set FPartnerList(i) = oneitem
			i=i+1
			rsget.movenext
		loop
		rsget.close
	end Sub

	public Sub GetPartnerList(byval ix)
		dim sqlStr, wheredetail
		dim oneitem,i

		if ix=1 then
			'' 파트너만.
			wheredetail = " and userdiv=999"
		elseif ix=2 then
			'' 직원만.
			wheredetail = " and userdiv<999"
		elseif ix=3 then
			'' 디자이너
			wheredetail = " and userdiv=9999"
			if FRectInitial="etc" then
				wheredetail = wheredetail + " and ((Left(id,1)<'a') or (Left(id,1)>'Z'))"
			elseif FRectInitial<>"" then
				wheredetail = wheredetail + " and (id like '" + FRectInitial + "%')"
			end if
		else
			wheredetail = ""
		end if

		''#################################################
		''총 갯수.
		''#################################################
		sqlStr = "select Count(id) as cnt from [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " where id<>''" + vbCrlf
		if FRectUserDiv<>"" then
			sqlStr = sqlStr + " and userdiv='" + FRectUserDiv + "'"
		end if
		rsget.Open sqlStr + wheredetail,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		''#################################################
		''현재 페이지 리스트.
		''#################################################
		sqlStr = "select top " + CStr(FCurrPage*FPageSize) + " * from [db_partner].[dbo].tbl_partner" + vbCrlf
		sqlStr = sqlStr + " where id<>''" + vbCrlf
		if FRectUserDiv<>"" then
			sqlStr = sqlStr + " and userdiv='" + FRectUserDiv + "'"
		end if
		sqlStr = sqlStr + wheredetail
'response.write sqlStr
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FPartnerList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set oneitem = new CPartnerUserItem
				oneitem.Fid               = rsget("id")
				oneitem.Fpassword         = rsget("password")
				oneitem.Fdiscountrate     = rsget("discountrate")
				oneitem.Fcompany_name     = db2html(rsget("company_name"))
				oneitem.Faddress          = db2html(rsget("address"))
				oneitem.Ftel              = rsget("tel")
				oneitem.Ffax              = rsget("fax")
				oneitem.Fbigo             = rsget("bigo")
				oneitem.Furl              = rsget("url")
				oneitem.Fmanager_name     = db2html(rsget("manager_name"))
				oneitem.Fmanager_address  = db2html(rsget("manager_address"))
				oneitem.Fcommission       = rsget("commission")
				oneitem.Femail            = db2html(rsget("email"))
				oneitem.Fuserdiv          = rsget("userdiv")
				oneitem.Fisusing          = rsget("isusing")
				set FPartnerList(i) = oneitem
				i=i+1
				rsget.movenext
			loop
		end if
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

end class
%>
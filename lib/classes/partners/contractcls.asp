<%
'###########################################################
' Description : 브랜드 계약 클래스
' Hieditor : 2009.04.07 서동석 생성
'			 2010.05.25 한용민 수정
'###########################################################

Class CPartnerContractDetailTypeItem
    public FcontractType
    public FdetailKey
    public FdetailDesc

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPartnerContractTypeItem
    public FContractType
    public FContractName
    public FContractContents
    public Fisusing
    public Fregdate
    public fonoffgubun

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPartnerContractDetailItem
    public FcontractID
    public FdetailKey
    public FdetailValue
    public FdetailDesc

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CPartnerContractItem
    public FcontractID
    public Fmakerid
    public FContractType
    public FContractState
    public FcontractName
    public FcontractNo
    public FContractContents
    public FContractEtcContetns
    public Freguserid
    public Fregdate
    public Fconfirmdate
    public Ffinishdate

    public Fusername
    public Fusermail
    public Finterphoneno
    public Fextension
    public Fdirect070

    public function GetContractStateColor()
        Select Case FContractState
            Case 0
                : GetContractStateColor = "#000000"
            Case 1
                : GetContractStateColor = "#77FF77"
            Case 3
                : GetContractStateColor = "#7777FF"
            Case 7
                : GetContractStateColor = "#FF7777"
            Case -1
                : GetContractStateColor = "#AAAAAA"
            Case else
                : GetContractStateColor = "#000000"
        end Select
    end function

    public function GetContractStateName()
        Select Case FContractState
            Case 0
                : GetContractStateName = "수정중"
            Case 1
                : GetContractStateName = "계약오픈"
            Case 3
                : GetContractStateName = "업체확인"
            Case 7
                : GetContractStateName = "계약완료"
            Case -1
                : GetContractStateName = "삭제"
            Case else
                : GetContractStateName = FContractState
        end Select
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CPartnerContract
    public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FRectCateCode
	public FRectMakerid
	public FRectCompanyName
	public FRectManagerName
	public FRectContractType
	public FRectdetailKey
	public FRectContractID
	public FRectContractno
	public FRectContractState
	public FRectOnOffGubun

	'//designer/company/popContract.asp
	public Sub GetOneContract()
	    dim sqlStr
	    sqlStr = " select C.*,t.username, t.usermail, t.interphoneno, t.extension,t.direct070 "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract C"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_user_tenbyten t"
	    sqlStr = sqlStr & "     on c.reguserid=t.userid"
	    sqlStr = sqlStr & " where C.ContractID=" & FRectContractID & ""
	    if FRectMakerid<>"" then
	        sqlStr = sqlStr & " and makerid='" & FRectMakerid & "'"
	    end if

	    'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem
			FOneItem.FcontractID          = rsget("contractID")
            FOneItem.Fmakerid             = rsget("makerid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FContractState       = rsget("ContractState")
            FOneItem.FcontractName        = db2html(rsget("contractName"))
            FOneItem.FcontractNo          = rsget("contractNo")
            FOneItem.FContractContents    = db2html(rsget("ContractContents"))
            FOneItem.FContractEtcContetns = db2html(rsget("ContractEtcContetns"))
            FOneItem.Freguserid           = rsget("reguserid")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

            FOneItem.Fusername            = rsget("username")
            FOneItem.Fusermail            = rsget("usermail")
            FOneItem.Finterphoneno        = rsget("interphoneno")
            FOneItem.Fextension           = rsget("extension")
            FOneItem.Fdirect070           = rsget("direct070")

		end if
		rsget.close

    end Sub

    '//admin/member/contractReg.asp
    public Sub GetContractDetailList()
	    dim sqlStr

	    sqlStr = " select A.*, t.detailDesc ,t.orderno from "
	    sqlStr = sqlStr & " ("
	    sqlStr = sqlStr & "     select c.ContractType, d.* from db_partner.dbo.tbl_partner_contract c,"
	    sqlStr = sqlStr & "     db_partner.dbo.tbl_partner_contractDetail d"
	    sqlStr = sqlStr & "     where d.ContractID=" & FRectContractID & ""
	    sqlStr = sqlStr & "     and d.ContractID=c.ContractID"
	    sqlStr = sqlStr & " ) A"
	    sqlStr = sqlStr & "     left join db_partner.dbo.tbl_partner_contractDetailType t"
	    sqlStr = sqlStr & "     on A.ContractType=t.ContractType"
	    sqlStr = sqlStr & "     and A.detailKey=t.detailKey"
	    sqlStr = sqlStr & "     order by t.orderno asc"

	    'response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
    			set FItemList(i) = new CPartnerContractDetailItem
    			FItemList(i).FcontractID            = rsget("contractID")
                FItemList(i).FdetailKey             = rsget("detailKey")
                FItemList(i).FdetailValue           = db2html(rsget("detailValue"))
                FItemList(i).FdetailDesc            = db2html(rsget("detailDesc"))
                i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

    end Sub

    public Sub GetRecentContractbyOnOff()
         dim sqlStr
	    sqlStr = " select C.* from db_partner.dbo.tbl_partner_contract C"
	    sqlStr = sqlStr & "     Join db_partner.dbo.tbl_partner_contractType T"
	    sqlStr = sqlStr & "     on C.contractType=T.contractType"
	    sqlStr = sqlStr & "     and T.onoffgubun='"&FRectOnOffGubun&"'"
	    sqlStr = sqlStr & " where C.makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and C.contractState>=0"
        sqlStr = sqlStr & " order by C.contractID desc"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem

			FOneItem.FcontractID          = rsget("contractID")
            FOneItem.Fmakerid             = rsget("makerid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FContractState       = rsget("ContractState")
            FOneItem.FcontractName        = db2html(rsget("contractName"))
            FOneItem.FcontractNo          = rsget("contractNo")
            FOneItem.FContractContents    = db2html(rsget("ContractContents"))
            FOneItem.FContractEtcContetns = db2html(rsget("ContractEtcContetns"))
            FOneItem.Freguserid           = rsget("reguserid")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

		end if
		rsget.close
    end Sub

	public Sub GetLastOneContract()
	    dim sqlStr
	    sqlStr = " select * from db_partner.dbo.tbl_partner_contract"
	    sqlStr = sqlStr & " where makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and contractState>=0"
	    sqlStr = sqlStr & " and contractState<7"
        sqlStr = sqlStr & " order by contractID desc"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractItem

			FOneItem.FcontractID          = rsget("contractID")
            FOneItem.Fmakerid             = rsget("makerid")
            FOneItem.FContractType        = rsget("ContractType")
            FOneItem.FContractState       = rsget("ContractState")
            FOneItem.FcontractName        = db2html(rsget("contractName"))
            FOneItem.FcontractNo          = rsget("contractNo")
            FOneItem.FContractContents    = db2html(rsget("ContractContents"))
            FOneItem.FContractEtcContetns = db2html(rsget("ContractEtcContetns"))
            FOneItem.Freguserid           = rsget("reguserid")
            FOneItem.Fregdate             = rsget("regdate")
            FOneItem.Fconfirmdate         = rsget("confirmdate")
            FOneItem.Ffinishdate          = rsget("finishdate")

		end if
		rsget.close


    end Sub

	'/admin/member/contractPrototypeReg.asp
	public sub getContractDetailProtoType()
	    dim sqlStr, i

	    sqlStr = " select * from db_partner.dbo.tbl_partner_contractDetailType"
	    if FRectContractType<>"" then
	        sqlStr = sqlStr & " where contractType=" & FRectContractType & ""
	    else
	        sqlStr = sqlStr & " where 1=0"
        end if

        sqlStr = sqlStr & " order by orderno asc"

        'response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount
	    redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			i=0
			do until rsget.eof
    			set FItemList(i) = new CPartnerContractDetailTypeItem

    			FItemList(i).FcontractType          = rsget("contractType")
                FItemList(i).FdetailKey             = rsget("detailKey")
                FItemList(i).FdetailDesc            = db2html(rsget("detailDesc"))

                i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

    end Sub

    public sub getOneContractDetailProtoType()
	    dim sqlStr, i

	    sqlStr = " select * from db_partner.dbo.tbl_partner_contractDetailType"
	    sqlStr = sqlStr & " where contractType=" & FRectContractType & ""
        sqlStr = sqlStr & " and detailKey='" & FRectdetailKey & "'"

        rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractDetailTypeItem
			FOneItem.FcontractType          = rsget("contractType")
            FOneItem.FdetailKey             = rsget("detailKey")
            FOneItem.FdetailDesc            = db2html(rsget("detailDesc"))
		end if
		rsget.close

    end Sub

    '//admin/member/contractPrototypeReg.asp
	public sub getOneContractProtoType()
	    dim sqlStr, i

	    sqlStr = " select * "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType"
	    if FRectContractType<>"" then
	        sqlStr = sqlStr & " where contractType=" & FRectContractType & ""
	    else
	        sqlStr = sqlStr & " where 1=0"
        end if
        sqlStr = sqlStr & " and subtype=-999"

	    'response.write sqlStr &"<br>"
	    rsget.Open sqlStr,dbget,1

	    FResultCount = rsget.RecordCount

	    if Not rsget.Eof then

			set FOneItem = new CPartnerContractTypeItem

			FOneItem.FContractType           = rsget("contractType")
            FOneItem.FContractName           = db2html(rsget("contractName"))
            FOneItem.FContractContents       = db2html(rsget("ContractContents"))
            FOneItem.Fregdate                = rsget("regdate")
            FOneItem.fonoffgubun           = rsget("onoffgubun")

		end if
		rsget.close

    end sub

	'//admin/member/contractPrototypeReg.asp
	public Sub getValidContractProtoTypeList()
	    dim sqlStr, i
	    sqlStr = "select count(contractType) as cnt "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType"
	    sqlStr = sqlStr & " where isusing='Y'"

	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & CStr(FPageSize*FCurrPage) & " contractType,contractName,regdate ,onoffgubun "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType"
	    sqlStr = sqlStr & " where isusing='Y'"

		'response.write sqlStr &"<Br>"
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof

				set FItemList(i) = new CPartnerContractTypeItem

				FItemList(i).fonoffgubun           = rsget("onoffgubun")
				FItemList(i).FcontractType           = rsget("contractType")
                FItemList(i).FcontractName           = db2html(rsget("contractName"))
                FItemList(i).Fregdate               = rsget("regdate")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close

    end Sub

    public Sub GetMakerNotConfirmContractList()
        dim sqlStr, i
        sqlStr = "select top " & CStr(FPageSize*FCurrpage) & " c.contractID "
		sqlStr = sqlStr & " , c.makerid, c.contractType, c.contractName"
		sqlStr = sqlStr & " , c.contractNo, c.contractState, c.reguserid, c.regdate, c.confirmdate, c.finishdate"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
	    sqlStr = sqlStr & " where makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and contractState>0"
	    sqlStr = sqlStr & " and contractState<3"
	    sqlStr = sqlStr & " order by c.contractID desc"

	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem

				FItemList(i).FcontractID           = rsget("contractID")
                FItemList(i).Fmakerid              = rsget("makerid")
                FItemList(i).FcontractType         = rsget("contractType")
                FItemList(i).FcontractNo           = rsget("contractNo")
                FItemList(i).FcontractName         = db2html(rsget("contractName"))
                FItemList(i).FcontractState        = rsget("contractState")
                FItemList(i).Freguserid            = rsget("reguserid")
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Fconfirmdate          = rsget("confirmdate")
                FItemList(i).Ffinishdate           = rsget("finishdate")
                'FItemList(i).FcontractContents     = rsget("contractContents")
                'FItemList(i).FcontractEtcContetns  = rsget("contractEtcContetns")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

    '//designer/company/popContract.asp
    public Sub GetMakerValidContractList()
        dim sqlStr, i
        sqlStr = "select top " & CStr(FPageSize*FCurrpage) & " c.contractID "
		sqlStr = sqlStr & " , c.makerid, c.contractType, c.contractName"
		sqlStr = sqlStr & " , c.contractNo, c.contractState, c.reguserid, c.regdate, c.confirmdate, c.finishdate"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
	    sqlStr = sqlStr & " where makerid='" & FRectMakerid & "'"
	    sqlStr = sqlStr & " and contractState>0"
	    sqlStr = sqlStr & " order by c.contractID desc"

	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem
				FItemList(i).FcontractID           = rsget("contractID")
                FItemList(i).Fmakerid              = rsget("makerid")
                FItemList(i).FcontractType         = rsget("contractType")
                FItemList(i).FcontractNo           = rsget("contractNo")
                FItemList(i).FcontractName         = db2html(rsget("contractName"))
                FItemList(i).FcontractState        = rsget("contractState")
                FItemList(i).Freguserid            = rsget("reguserid")
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Fconfirmdate          = rsget("confirmdate")
                FItemList(i).Ffinishdate           = rsget("finishdate")

                'FItemList(i).FcontractContents     = rsget("contractContents")
                'FItemList(i).FcontractEtcContetns  = rsget("contractEtcContetns")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

	public Sub GetContractList()
	    dim sqlStr, i
	    sqlStr = "select count(contractID) as cnt "
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract"
	    sqlStr = sqlStr & " where 1=1"

	    if (FRectContractState<>"") then
	        sqlStr = sqlStr & " and contractState=" & FRectContractState
	    else
	        sqlStr = sqlStr & " and (contractState>=0 or contractState=-2)"
	    end if

	    if FRectMakerid<>"" then
	        sqlStr = sqlStr & " and makerid='" & FRectMakerid & "'"
	    end if

	    if FRectContractno<>"" then
	        sqlStr = sqlStr & " and ContractNo='" & FRectContractno & "'"
	    end if

	    if FRectCateCode<>"" then
	        sqlStr = sqlStr & " and makerid in ("
	        sqlStr = sqlStr & "     select userid from db_user.[dbo].tbl_user_c"
	        sqlStr = sqlStr & "     where catecode='" & FRectCateCode & "'"
            sqlStr = sqlStr & " )"
	    end if

	    if FRectCompanyName<>"" then
	        sqlStr = sqlStr & " and makerid in ("
	        sqlStr = sqlStr & "     select id from db_partner.[dbo].tbl_partner"
	        sqlStr = sqlStr & "     where company_name like '%" & FRectCompanyName & "%'"
            sqlStr = sqlStr & " )"
	    end if

	    if FRectManagerName<>"" then
	        sqlStr = sqlStr & " and makerid in ("
	        sqlStr = sqlStr & "     select id from db_partner.[dbo].tbl_partner"
	        sqlStr = sqlStr & "     where manager_name like '%" & FRectManagerName & "%'"
            sqlStr = sqlStr & " )"
	    end if

	    rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & CStr(FPageSize*FCurrpage) & " c.contractID "
		sqlStr = sqlStr & " , c.makerid, c.contractType, c.contractName"
		sqlStr = sqlStr & " , c.contractNo, c.contractState, c.reguserid, c.regdate, c.confirmdate, c.finishdate"
	    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contract c"
	    'sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner_contractType t"
	    'sqlStr = sqlStr & " on c.contractType=t.contractType"
	    sqlStr = sqlStr & " where 1=1"
	    if (FRectContractState<>"") then
	        sqlStr = sqlStr & " and contractState=" & FRectContractState
	    else
	        sqlStr = sqlStr & " and (contractState>=0 or contractState=-2)"
	    end if
	    if FRectMakerid<>"" then
	        sqlStr = sqlStr & " and makerid='" & FRectMakerid & "'"
	    end if

	    if FRectContractno<>"" then
	        sqlStr = sqlStr & " and ContractNo='" & FRectContractno & "'"
	    end if

	    if FRectCateCode<>"" then
	        sqlStr = sqlStr & " and makerid in ("
	        sqlStr = sqlStr & "     select userid from db_user.[dbo].tbl_user_c"
	        sqlStr = sqlStr & "     where catecode='" & FRectCateCode & "'"
            sqlStr = sqlStr & " )"
	    end if

	    if FRectCompanyName<>"" then
	        sqlStr = sqlStr & " and makerid in ("
	        sqlStr = sqlStr & "     select id from db_partner.[dbo].tbl_partner"
	        sqlStr = sqlStr & "     where company_name like '%" & FRectCompanyName & "%'"
            sqlStr = sqlStr & " )"
	    end if

	    if FRectManagerName<>"" then
	        sqlStr = sqlStr & " and makerid in ("
	        sqlStr = sqlStr & "     select id from db_partner.[dbo].tbl_partner"
	        sqlStr = sqlStr & "     where manager_name like '%" & FRectManagerName & "%'"
            sqlStr = sqlStr & " )"
	    end if
		sqlStr = sqlStr & " order by c.regdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			i=0
			do until rsget.eof
				set FItemList(i) = new CPartnerContractItem

				FItemList(i).FcontractID           = rsget("contractID")
                FItemList(i).Fmakerid              = rsget("makerid")
                FItemList(i).FcontractType         = rsget("contractType")
                FItemList(i).FcontractNo           = rsget("contractNo")
                FItemList(i).FcontractName         = db2html(rsget("contractName"))
                FItemList(i).FcontractState        = rsget("contractState")
                FItemList(i).Freguserid            = rsget("reguserid")
                FItemList(i).Fregdate              = rsget("regdate")
                FItemList(i).Fconfirmdate          = rsget("confirmdate")
                FItemList(i).Ffinishdate           = rsget("finishdate")
                'FItemList(i).FcontractContents     = rsget("contractContents")
                'FItemList(i).FcontractEtcContetns  = rsget("contractEtcContetns")

				i=i+1
				rsget.movenext
			loop
		end if
		rsget.close
    end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage       = 1
		FPageSize       = 12
		FResultCount    = 0
		FScrollCount    = 10
		FTotalCount     = 0
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

Sub drawSelectBoxContractTypeWithChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%= selectBoxName %>" onchange="ChangeContractType(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select contractType,contractName from db_partner.dbo.tbl_partner_contractType"
   query1 = query1 & " where isusing='Y'"
   query1 = query1 & " and subtype=-999"
   query1 = query1 & " order by contractType"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("contractType")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("contractType")&"' "&tmp_str&">"&rsget("contractType")&" ["&db2html(rsget("contractName"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

function getMdUserName(mduserid)
    dim sqlStr
    sqlStr = "select company_name from [db_partner].[dbo].tbl_partner"
    sqlStr = sqlStr & " where id='" & mduserid & "'"
    rsget.Open sqlStr,dbget,1
    if  not rsget.EOF  then
        getMdUserName = db2html(rsget("company_name"))
    end if
    rsget.close
end function

function getDefaultContractValue(aKey, opartner)
    dim mdusername

    select case aKey
        ''case "$$CONTRACT_NO$$"          ''계약서 번호. -> 자동생성
        ''    : getDefaultContractValue = mdusername
        case "$$A_CHARGE$$"             ''계약담당자.
            mdusername = getMdUserName(opartner.FOneItem.Fmduserid)
            : getDefaultContractValue = mdusername
        case "$$A_UPCHENAME$$"
            : getDefaultContractValue = "(주)텐바이텐"
        case "$$A_CEONAME$$"
            : getDefaultContractValue = "최은희"
        case "$$A_COMPANY_NO$$"
            : getDefaultContractValue = "211-87-00620"
        case "$$A_COMPANY_ADDR$$"
            : getDefaultContractValue = "서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐"

        case "$$B_CHARGE$$"
            : getDefaultContractValue = opartner.FOneItem.FManager_Name
        case "$$B_UPCHENAME$$"
            : getDefaultContractValue = opartner.FOneItem.Fcompany_name
        case "$$B_CEONAME$$"
            : getDefaultContractValue = opartner.FOneItem.Fceoname
        case "$$B_COMPANY_NO$$"
            : getDefaultContractValue = opartner.FOneItem.Fcompany_no
        case "$$B_COMPANY_ADDR$$"
            : getDefaultContractValue = opartner.FOneItem.Faddress & " " & opartner.FOneItem.Fmanager_address
        case "$$B_BRANDNAME$$"
            : getDefaultContractValue = opartner.FOneItem.Fsocname_kor
        ''case "$$B_DELIVER_MANAGER$$"
        ''    : getDefaultContractValue = opartner.FOneItem.Fdeliver_name


        case "$$DEFAULT_ITEM_MARGIN$$"
            : getDefaultContractValue = opartner.FOneItem.Fdefaultmargine & "%"
        case "$$DEFAULT_SERVICE_MARGIN$$"
            : getDefaultContractValue = opartner.FOneItem.Fdefaultmargine & "%"


        case "$$DEFAULT_JUNGSANDATE$$"
            :
            if (opartner.FOneItem.Fjungsan_date="15일") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 15일"
            elseif (opartner.FOneItem.Fjungsan_date="말일") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 말일"
            elseif (opartner.FOneItem.Fjungsan_date="수시") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 5일"
            elseif (opartner.FOneItem.Fjungsan_date="5일") then
                getDefaultContractValue = "판매(제공)월의 " & "익월 5일"
            end if

        case "$$CONTRACT_DATE$$"
            : getDefaultContractValue = Left(Now(),10)

        case "$$INSURANCE_FEE$$"                        '' 보증보험
            : getDefaultContractValue = "0 만원"

        case Else
            : getDefaultContractValue = ""
    end select
end function

function drawonoffgubun(boxname ,stats)
%>
	<select name='<%=boxname%>'>
		<option value='' <% if stats = "" then response.write " selected" %>>선택</option>
		<option value='ON' <% if stats = "ON" then response.write " selected" %>>온라인</option>
		<option value='OFF' <% if stats = "OFF" then response.write " selected" %>>오프라인</option>
	</select>

<%
end function

'//해당 브랜드에 대한 샵의 마진을 반환한다
Sub drawSelectOffShopmargin(selectBoxName,selectedId)
dim tmp_str,query1
%>
   <select class="select" name="<%=selectBoxName%>" onchange="SelectOffShopmargin(this.value);">
   <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   'query1 = query1 & " and userid<>'streetshop000'"	'직영점(대표)
   'query1 = query1 & " and userid<>'streetshop800'"		'가맹점(대표)
   'query1 = query1 & " and userid<>'streetshop870'"		'도매(대표)
   'query1 = query1 & " and userid<>'streetshop700'"		'해외(대표)

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
   'response.write query1 &"<br>"
%>
	<script language='javascript'>
		function SelectOffShopmargin(tmp){
			frmReg.shopid.value = tmp;
			frmReg.target = 'view';
			frmReg.action = 'contractReg_selectshopmargin.asp';
			frmReg.submit();
		}
	</script>
	<input type="hidden" name="shopid">
	<iframe id="view" name="view" frameborder="0" width=0 height=0></iframe>
<%
end sub

function drawSelectshopuser(selectBoxName,selectedId,btcid,changeflg)
dim tmp_str,query1
%>
   <select class="select" name="<%=selectBoxName%>" <%=changeflg%>>
   <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	query1 = "select top 30" + VbCrlf
	query1 = query1 & " ut.userid , ps.shopid ,su.shopname, ps.firstisusing" + VbCrlf
	query1 = query1 & " ,(case when ps.firstisusing='Y' then '[대표]' end) as firstname" + VbCrlf
	query1 = query1 + " from db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf
	query1 = query1 + " join db_partner.dbo.tbl_partner_shopuser ps" + vbcrlf
	query1 = query1 + " 	on ps.empno = ut.empno" + vbcrlf
	query1 = query1 + " join db_shop.dbo.tbl_shop_user su" + vbcrlf
	query1 = query1 + " 	on ps.shopid = su.userid" + vbcrlf
	query1 = query1 + " where ut.isusing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (ut.statediv ='Y' or (ut.statediv ='N' and datediff(dd,ut.retireday,getdate())<=0))" & vbcrlf
	query1 = query1 & "	and ut.userid = '"&btcid&"'" & vbcrlf

	'response.write query1 &"<br>"
	rsget.Open query1,dbget,1

	if  not rsget.EOF  then
	   rsget.Movefirst

	   do until rsget.EOF
	       if Lcase(selectedId) = Lcase(rsget("shopid")) then
	           tmp_str = " selected"
	       end if
	       response.write("<option value='"&rsget("shopid")&"' "&tmp_str&">"&rsget("shopid")&"/"&rsget("shopname")&rsget("firstname")&"</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   loop
	end if
	rsget.close
	response.write("</select>")

end function
%>

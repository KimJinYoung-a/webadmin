<%

class CElecTaxHistory_F
	public F_idx
	public F_uniq_id
	public F_jungsangubun
	public F_makerid
	public F_jungsanname
	public F_biz_no
	public F_corp_nm
	public F_write_date
	public F_item_amt
	public F_item_price
	public F_item_vat

	public F_cur_dam_nm
	public F_tax_no
	public F_deleteyn
	public F_resultmsg
    public F_tax_type
    public F_regdate
    public F_jgubun

    public function getTaxTypeName()
        if (F_tax_type="01") then
            getTaxTypeName = "과세"
        elseif (F_tax_type="02") then
            getTaxTypeName = "<font color=red>면세</font>"
        elseif (F_tax_type="03") then
            getTaxTypeName = "<font color=blue>영세</font>"
        else
            getTaxTypeName = getTaxTypeName
        end if
    end function

    public function getJGubunName()
        if isNULL(F_jgubun) then Exit function

        if (F_jgubun="CC") then
            getJGubunName = "<font color=blue>수수료</font>"
        else
            getJGubunName = "매입"
        end if
    end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end class

class CElecTaxHistory

	public FMasterItemList()
	public FDetailItemList()
	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FRectRegStart
	public FRectRegEnd
	public FCurrPage

	public Fcomp
	public Fright
    public FRectonoffgubun
    public FRectOnlyComuniErr
    public FRectbiz_no
    public FRectTaxCorp

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		'redim preserve FDetailItemList(0)

		redim  FMasterItemList(0)
		redim  FDetailItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public sub datalist()

		dim sqlStr
		dim i

			if Fcomp <> "" then
				sqlStr = "Select "
			else
				sqlStr = "Select top 500"
			end if
			sqlStr = sqlStr & "idx, "				' 인덱스
			sqlStr = sqlStr & "uniq_id, "			' 계산서문서번호
			sqlStr = sqlStr & "jungsangubun, "		' 정산구분
			sqlStr = sqlStr & "makerid, "			' 브랜드ID
			sqlStr = sqlStr & "jungsanname, "		' 정산구분명
			sqlStr = sqlStr & "biz_no, "			' 사업자번호
			sqlStr = sqlStr & "corp_nm, "			' 사업자명
			sqlStr = sqlStr & "write_date, "		' 발행일
            sqlStr = sqlStr & "item_price, "		    ' 총금액
			sqlStr = sqlStr & "item_vat, "		    ' 부가세
			sqlStr = sqlStr & "item_amt, "			' 품목공급가
			sqlStr = sqlStr & "cur_dam_nm, "		' 담당자명
			sqlStr = sqlStr & "tax_no, "			' 계산서번호
			sqlStr = sqlStr & "resultmsg, "			' 결과Str
			sqlStr = sqlStr & "tax_type, "			' 과세면세영세.
			sqlStr = sqlStr & "convert(varchar(16),regdate,21) as regdate, "			' 등록일
			sqlStr = sqlStr & "resultmsg, "			' 결과Str
			sqlStr = sqlStr & "deleteyn, "			' 삭제구분
			sqlStr = sqlStr & "jgubun "			' 정산구분
			sqlStr = sqlStr & "from [db_jungsan].[dbo].[tbl_tax_history_master]"
			sqlStr = sqlStr & " where"
			if Fcomp <> "" then
			sqlStr = sqlStr & " makerid='" & Fcomp & "'"
			else
			sqlStr = sqlStr & " makerid is not null"
			end if
			if Fright <> "" then
				sqlStr = sqlStr & " and deleteyn='N'"
				sqlStr = sqlStr & " and resultmsg='OK'"
				''sqlStr = sqlStr & " and tax_no>'0'"
			end if
			if (FRectOnlyComuniErr<>"") then
			    sqlStr = sqlStr & " and ((tax_no is NULL) or (Left(tax_no,2)='TX' and (resultmsg<>'OK')))"
			    sqlStr = sqlStr & " and deleteyn='N'"
			end if

			if (FRectonoffgubun<>"") then
			    sqlStr = sqlStr & " and jungsangubun='"&FRectonoffgubun&"'"
			end if

			if (FRectbiz_no<>"") then
			    sqlStr = sqlStr & " and biz_no='"&FRectbiz_no&"'"
			end if

			if (FRectTaxCorp<>"") then
			    ''발행사 구분필요..
			    sqlStr = sqlStr & " and billsite='"&FRectTaxCorp&"'"
			end if


			sqlStr = sqlStr & " order by idx desc"

			rsget.PageSize = FPageSize
'response.write(sqlStr)
			rsget.Open sqlStr,dbget,1
			FTotalCount = rsget.RecordCount

			if (FCurrPage * FPageSize < FTotalCount) then
				FResultCount = FPageSize
			else
				FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
			end if


			FPageCount = rsget.PageCount

			FTotalPage = (FTotalCount\FPageSize)

			if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

			redim preserve FMasterItemList(FResultCount)


		if not rsget.EOF then
			rsget.absolutepage = FCurrPage

			do until (i >= FResultCount)

				set FMasterItemList(i) = new CElecTaxHistory_F
				FMasterItemList(i).F_idx			= rsget("idx")
				FMasterItemList(i).F_uniq_id		= rsget("uniq_id")
				if rsget("jungsangubun") = "ON" then
					FMasterItemList(i).F_jungsangubun	= "온라인"
				elseif rsget("jungsangubun") = "OFF" then
					FMasterItemList(i).F_jungsangubun	= "오프라인"
				elseif rsget("jungsangubun") = "FRN" then
					FMasterItemList(i).F_jungsangubun	= "가맹점"
			    elseif rsget("jungsangubun") = "OF" then
					FMasterItemList(i).F_jungsangubun	= "오프"
			    elseif rsget("jungsangubun") = "OFFSHOP" then
					FMasterItemList(i).F_jungsangubun	= "가맹점매출"
				else
					FMasterItemList(i).F_jungsangubun	= "--"
				end if
				FMasterItemList(i).F_makerid		= rsget("makerid")
				FMasterItemList(i).F_jungsanname	= db2html(rsget("jungsanname"))
				FMasterItemList(i).F_biz_no			= rsget("biz_no")
				FMasterItemList(i).F_corp_nm		= db2html(rsget("corp_nm"))
				FMasterItemList(i).F_write_date		= rsget("write_date")
				FMasterItemList(i).F_item_price     = rsget("item_price")
				FMasterItemList(i).F_item_vat       = rsget("item_vat")
				FMasterItemList(i).F_item_amt		= rsget("item_amt")
				FMasterItemList(i).F_cur_dam_nm		= rsget("cur_dam_nm")
				if rsget("tax_no") = null then
				FMasterItemList(i).F_tax_no			= "0"
				else
				FMasterItemList(i).F_tax_no			= rsget("tax_no")
				end if
				FMasterItemList(i).F_resultmsg		= rsget("resultmsg")
				FMasterItemList(i).F_deleteyn		= rsget("deleteyn")

                FMasterItemList(i).F_tax_type		= rsget("tax_type")
                FMasterItemList(i).F_regdate        = rsget("regdate")

                FMasterItemList(i).F_jgubun         = rsget("jgubun")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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

%>
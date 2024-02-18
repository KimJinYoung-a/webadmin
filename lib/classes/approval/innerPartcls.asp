<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 한용민 수정
'###########################################################

Class CInnerPartItem
        public Fidx
        public Fdivcd
        public FBIZSECTION_CD
        public FBIZSECTION_NM

        public Fscmid
        public Fuseyn
        public Fregdate

	    public function GetDivcdColor()

	        if (Fdivcd="S") then
	        	GetDivcdColor = "green"
	        elseif (Fdivcd="M") then
	        	GetDivcdColor = "#0000FF"
	        else
	            GetDivcdColor = "#000000"
	        end if

	    end function

	    public function GetDivcdName()

	        if (Fdivcd="S") then
	        	GetDivcdName = "매장"
	        elseif (Fdivcd="M") then
	        	GetDivcdName = "매입부서"
	        else
	            GetDivcdName = Fdivcd
	        end if

	    end function

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class CInnerPart
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

		public FRectIdx

		public function GetFromWhere
			dim tmpSql

			GetFromWhere = ""

			tmpSql = " from "
			tmpSql = tmpSql + "	db_partner.dbo.tbl_InternalPart p "
			tmpSql = tmpSql + "	join db_partner.dbo.tbl_TMS_BA_BIZSECTION b "
			tmpSql = tmpSql + "	on "
			tmpSql = tmpSql + "		p.BIZSECTION_CD = b.BIZSECTION_CD "
			tmpSql = tmpSql + " where "
			tmpSql = tmpSql + " 	1 = 1 "
			tmpSql = tmpSql + " 	and p.useyn = 'Y' "

			if (FRectIdx <> "") then
				tmpSql = tmpSql + " 	and p.idx = " + CStr(FRectIdx) + " "
			end if

			GetFromWhere = tmpSql

		end function

        public Sub GetInnerPartList()
            dim i,sqlStr

			'// ===============================================================
			sqlStr = " select count(p.idx) as cnt "

			sqlStr = sqlStr + GetFromWhere

			'response.write sqlStr
	        rsget.Open sqlStr, dbget, 1
	            FTotalCount = rsget("cnt")
	        rsget.Close

			'// ===============================================================
			sqlStr = " select top " + CStr(FPageSize*FCurrPage)

			sqlStr = sqlStr + " p.idx, p.divcd, p.BIZSECTION_CD, b.BIZSECTION_NM, p.scmid, p.useyn, p.regdate "

			sqlStr = sqlStr + GetFromWhere

			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	p.idx desc "

	        rsget.pagesize = FPageSize
	        rsget.Open sqlStr, dbget, 1
	        'response.write sqlStr

	        FTotalPage =  CLng(FTotalCount\FPageSize)
			if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
				FTotalPage = FtotalPage + 1
			end if
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

	        if FResultCount<1 then FResultCount=0

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
					set FItemList(i) = new CInnerPartItem

					FItemList(i).Fidx         		= rsget("idx")
					FItemList(i).Fdivcd         	= rsget("divcd")

					FItemList(i).FBIZSECTION_CD     = rsget("BIZSECTION_CD")
					FItemList(i).FBIZSECTION_NM     = rsget("BIZSECTION_NM")
					FItemList(i).Fscmid         	= rsget("scmid")
					FItemList(i).Fuseyn         	= rsget("useyn")
					FItemList(i).Fregdate         	= rsget("regdate")

					rsget.moveNext
					i=i+1
				loop
			end if

			rsget.Close
        end sub

        public Sub GetInnerPartOne()
            dim i,sqlStr

			if (FRectIdx = "") then
				'잘못된 접속
				set FOneItem = new CInnerPartItem
				Exit Sub
			end if

			Call GetInnerPartList()

			if (FResultCount < 1) then
				'잘못된 접속
				set FOneItem = new CInnerPartItem
				Exit Sub
			end if

			set FOneItem = FItemList(0)

        end sub

        public Sub GetInnerPartBlankDetail()
            dim i,sqlStr

            set FOneItem = new CInnerPartItem
        end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 20
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
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

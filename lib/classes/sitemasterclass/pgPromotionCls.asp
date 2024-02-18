<%
'#######################################################
'	History	: 2013.09.30 서동석 생성
'			  2022.07.04 한용민 수정(isms취약점수정)
'	Description : 신용카드 프로모션 관리(결제단 무이자 display)
'#######################################################

Class CCardPromotionItem
    public Fidx
	public Fcimage
	public Fpgprogbn
	public FCardCd
	Public FSDt
	Public FEDt
    Public Fconts
    Public Fcontlink
    Public FIsUsing
    Public FRegDate

    public function getMayImageName()
        if (Fpgprogbn="m") then
            getMayImageName = getCardCd2ImgURL(FCardCd)
        else
            if (Fcimage<>"") then
                getMayImageName=Fcimage
            end if
        end if
    end function

    public function getStateName()
        if (FIsUsing="N") then
            getStateName="사용중지"
        elseif (CDate(FEDt)<now()) then
            getStateName="만료"
        elseif (CDate(FSDt)>now()) then
            getStateName="예약"
        elseif (CDate(FSDt)<now()) and (CDate(FEDt)>now())then
            getStateName="Active"
        end if
    end function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CCardPromotion
	public FItemList()
    public FOneItem

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

    public FRectIdx
    public FRectStDT
    public FRectEdDT
    public FRectvalidgbn
    public FRectMatchDate
    public FRectpgprogbn
	public FRectCardCd
    public FRectIsusing

    public sub getCardPromotionOne()
        dim sqlStr
        sqlStr = "select * from db_sitemaster.dbo.tbl_pg_promotion p with (nolock)"
        sqlStr = sqlStr & " where p.idx="&FRectIdx
        rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount
        if  not rsget.EOF  then
			set FOneItem = new CCardPromotionItem

			FOneItem.Fidx     = rsget("idx")
            FOneItem.Fcimage  = rsget("cimage")
            FOneItem.Fpgprogbn= rsget("pgprogbn")
            FOneItem.FCardCd  = rsget("CardCd")
            FOneItem.FSDt     = rsget("SDt")
            FOneItem.FEDt     = rsget("EDt")
            FOneItem.Fconts   = rsget("conts")
            FOneItem.Fcontlink= db2html(rsget("contlink"))
            FOneItem.FIsUsing = db2html(rsget("IsUsing"))
            FOneItem.FRegDate = rsget("regdate")
		end if
        rsget.close()
    end Sub

    public Sub getCardPromotionList()
        dim sqlStr, sqlStrAdd

        sqlStrAdd=""
        IF (FRectIsusing<>"") then
            sqlStrAdd=sqlStrAdd&" and isusing='"&FRectIsusing&"'"
        end if

        if (FRectpgprogbn<>"") then
            sqlStrAdd=sqlStrAdd&" and pgprogbn='"&FRectpgprogbn&"'"
        end if

        if (FRectCardCd<>"") then
            sqlStrAdd=sqlStrAdd&" and cardcd='"&FRectCardCd&"'"
        end if

        ''c:현재일기준 p:특정일
        if (FRectvalidgbn<>"") then
            if (FRectvalidgbn="c") then
                FRectMatchDate=Left(CStr(now()),10)
            end if

            if (FRectMatchDate<>"") then
                sqlStrAdd=sqlStrAdd&" and isusing='Y'"
                sqlStrAdd=sqlStrAdd&" and sDt<='"&FRectMatchDate&"'"
                sqlStrAdd=sqlStrAdd&" and eDt>='"&FRectMatchDate&"'"

                sqlStrAdd=sqlStrAdd&" and ((pgprogbn='m' and idx in ("  ''날짜가 겹칠경우 최종idx기준
                sqlStrAdd=sqlStrAdd&" 	select Max(idx)"
                sqlStrAdd=sqlStrAdd&" 	from db_sitemaster.dbo.tbl_pg_promotion with (nolock)"
                sqlStrAdd=sqlStrAdd&" 	where 1=1 and isusing='Y' "
                sqlStrAdd=sqlStrAdd&" 	and sDt<='"&FRectMatchDate&"' "
                sqlStrAdd=sqlStrAdd&" 	and eDt>='"&FRectMatchDate&"' "
                sqlStrAdd=sqlStrAdd&" 	and pgprogbn in ('m')"
                sqlStrAdd=sqlStrAdd&" 	group by cardCd"
                sqlStrAdd=sqlStrAdd&" )) or (pgprogbn='a' and idx in ("
                sqlStrAdd=sqlStrAdd&" 	select Max(idx)"
                sqlStrAdd=sqlStrAdd&" 	from db_sitemaster.dbo.tbl_pg_promotion with (nolock)"
                sqlStrAdd=sqlStrAdd&" 	where 1=1 and isusing='Y' "
                sqlStrAdd=sqlStrAdd&" 	and sDt<='"&FRectMatchDate&"' "
                sqlStrAdd=sqlStrAdd&" 	and eDt>='"&FRectMatchDate&"' "
                sqlStrAdd=sqlStrAdd&" 	and pgprogbn in ('a')"
                sqlStrAdd=sqlStrAdd&" )) or (pgprogbn='b'))"


            end if
        end if


        sqlStr = " select count(*) as CNT, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") " + vbcrlf
        sqlStr = sqlStr&" from db_sitemaster.dbo.tbl_pg_promotion p with (nolock)"
        sqlStr = sqlStr&" where 1=1"
        sqlStr = sqlStr&sqlStrAdd

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

        '지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

        sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " p.* " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_pg_promotion p with (nolock)" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + sqlStrAdd + vbcrlf
		if (FRectvalidgbn<>"") then
		    sqlStr = sqlStr + " order by p.pgprogbn desc, p.CardCd asc, p.idx desc"
		else
		    sqlStr = sqlStr + " order by p.idx desc"
	    end if

        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CCardPromotionItem

				FItemList(i).Fidx     = rsget("idx")
                FItemList(i).Fcimage  = rsget("cimage")
                FItemList(i).Fpgprogbn= rsget("pgprogbn")
                FItemList(i).FCardCd= rsget("CardCd")
                FItemList(i).FSDt     = rsget("SDt")
                FItemList(i).FEDt     = rsget("EDt")
                FItemList(i).Fconts   = rsget("conts")
                FItemList(i).Fcontlink= db2html(rsget("contlink"))
                FItemList(i).FIsUsing = db2html(rsget("IsUsing"))
                FItemList(i).FRegDate = rsget("regdate")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    Private Sub Class_Initialize()
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
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Sub DrawSelectBoxCardPromoGubun(boxname,iselid,etcVal)
%>
<select name='<%=boxname%>' class='select' <%=etcVal%>>
	<option value="">선택하세요</option>
	<option value="m" <% if iselid = "m" then response.write " selected" %>>무이자</option>
	<option value="b" <% if iselid = "b" then response.write " selected" %>>배너</option>
	<option value="a" <% if iselid = "a" then response.write " selected" %>>미지정시</option>
</select>
<%
end Sub

function getCdPromotionGubunName(igbn)
    select CASE igbn
    CASE "m"
        getCdPromotionGubunName = "무이자"
    CASE "b"
        getCdPromotionGubunName = "배너"
    CASE "a"
        getCdPromotionGubunName = "미지정시"
    CASE ELSE
        getCdPromotionGubunName = ""
    end Select
end function

Sub DrawSelectBoxCardList(boxname,iselid)
%>
<select name='<%=boxname%>' class='select'>
	<option value="">선택하세요</option>
	<option value="01" <%=CHKIIF(iselid = "01","selected","") %> >신한카드</option>
	<option value="02" <%=CHKIIF(iselid = "02","selected","") %> >외환카드</option>
	<option value="03" <%=CHKIIF(iselid = "03","selected","") %> >비씨카드</option>
	<option value="04" <%=CHKIIF(iselid = "04","selected","") %> >삼성카드</option>
	<option value="05" <%=CHKIIF(iselid = "05","selected","") %> >국민카드</option>
	<option value="06" <%=CHKIIF(iselid = "06","selected","") %> >농협카드</option>
	<option value="07" <%=CHKIIF(iselid = "07","selected","") %> >하나SK카드</option>
	<option value="08" <%=CHKIIF(iselid = "08","selected","") %> >롯데카드</option>
	<option value="09" <%=CHKIIF(iselid = "09","selected","") %> >현대카드</option>
</select>
<%
End Sub

function getCardCd2Name(icdCode)
    select CASE icdCode
        CASE "01"
            getCardCd2Name = "신한카드"
        CASE "02"
            getCardCd2Name = "외환카드"
        CASE "03"
            getCardCd2Name = "비씨카드"
        CASE "04"
            getCardCd2Name = "삼성카드"
        CASE "05"
            getCardCd2Name = "국민카드"
        CASE "06"
            getCardCd2Name = "농협카드"
        CASE "07"
            getCardCd2Name = "하나SK카드"
        CASE "08"
            getCardCd2Name = "롯데카드"
        CASE "09"
            getCardCd2Name = "현대카드"
        CASE ELSE
            getCardCd2Name = ""
    end Select
end function

function getCardCd2ImgURL(icdCode)
    getCardCd2ImgURL = "http://fiximage.10x10.co.kr/web2013/cart/card_img"&icdCode&".gif"
end function

%>
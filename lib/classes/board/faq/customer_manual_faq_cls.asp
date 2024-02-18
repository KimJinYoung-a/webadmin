<%
'###########################################################
' Description : FAQ 클래스
' Hieditor : 2019.10.29 한용민 생성
'###########################################################

Class cfaq_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public ffidx
	public fgubun
	public fcontents
	public fsolution
    public fisusing
	public fregdate
    public flastupdate
	public flastadminid
	public fmanualtype
end class

class cfaq_list
	public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

    public frectfidx
    public frectisusing
    public frectgubun
    public frectcontents
    public frectsolution
	public frectmanualtype

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

    public Sub Getcustomer_manual_faq_one()
        dim sqlStr , sqlsearch

		if frectmanualtype="" or isnull(frectmanualtype) then exit sub

		if frectfidx <> "" then
			sqlsearch = sqlsearch & " and fidx = "& frectfidx &""
		end if
		if frectmanualtype <> "" then
			sqlsearch = sqlsearch & " and manualtype = '"& frectmanualtype &"'"
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " fidx,manualtype,gubun,contents,solution,isusing,regdate,lastupdate,lastadminid" & vbcrlf
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_manual_faq f with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cfaq_item
        
        if Not rsget.Eof then

    		FOneItem.ffidx = rsget("fidx")
    		FOneItem.fmanualtype = rsget("manualtype")
    		FOneItem.fgubun = rsget("gubun")
    		FOneItem.fcontents = db2html(rsget("contents"))
    		FOneItem.fsolution = db2html(rsget("solution"))
    		FOneItem.fisusing = rsget("isusing")
    		FOneItem.fregdate = rsget("regdate")
            FOneItem.flastupdate = rsget("lastupdate")
    		FOneItem.flastadminid = rsget("lastadminid")

        end if
        rsget.Close
    end Sub
    
	public sub Getcustomer_manual_faq()
		dim sqlStr,i , sqlsearch

		if frectmanualtype="" or isnull(frectmanualtype) then exit sub

        if frectisusing<>"" then
            sqlsearch = sqlsearch & " and isusing='"& frectisusing &"'" & vbcrlf
        end if
        if frectfidx<>"" then
            sqlsearch = sqlsearch & " and fidx="& frectfidx &"" & vbcrlf
        end if
        if frectgubun<>"" then
            sqlsearch = sqlsearch & " and gubun="& frectgubun &"" & vbcrlf
        end if
        if frectcontents<>"" then
            sqlsearch = sqlsearch & " and contents like '%"& frectcontents &"%'" & vbcrlf
        end if
        if frectsolution<>"" then
            sqlsearch = sqlsearch & " and solution like '%"& frectsolution &"%'" & vbcrlf
        end if
		if frectmanualtype <> "" then
			sqlsearch = sqlsearch & " and manualtype = '"& frectmanualtype &"'"
		end if

		'총 갯수 구하기
		sqlStr = "select count(fidx) as cnt" & vbcrlf
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_manual_faq f with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
					
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " fidx,manualtype,gubun,contents,solution,isusing,regdate,lastupdate,lastadminid" & vbcrlf
		sqlStr = sqlStr & " from db_cs.dbo.tbl_customer_manual_faq f with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by fidx desc" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cfaq_item

				FItemList(i).ffidx = rsget("fidx")
				FItemList(i).fmanualtype = rsget("manualtype")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fsolution = db2html(rsget("solution"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
                FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).flastadminid = rsget("lastadminid")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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

function Drawcustomerfaqgubun(selBoxName,selVal,changeFlag)
%>
    <select class='select' name="<%= selBoxName %>" <%= changeFlag %>>
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='1' <% if selVal="1" then response.write " selected" %> >업체</option>
	<option value='2' <% if selVal="2" then response.write " selected" %> >어드민</option>
	<option value='3' <% if selVal="3" then response.write " selected" %> >시스템/사이트</option>
	<option value='4' <% if selVal="4" then response.write " selected" %> >회원</option>
	<option value='17' <% if selVal="17" then response.write " selected" %> >회원가입</option>
	<option value='5' <% if selVal="5" then response.write " selected" %> >로그인</option>
    <option value='13' <% if selVal="13" then response.write " selected" %> >배송/송장</option>
    <option value='6' <% if selVal="6" then response.write " selected" %> >장바구니/주문/결제/입금</option>
	<option value='19' <% if selVal="19" then response.write " selected" %> >취소</option>
    <option value='7' <% if selVal="7" then response.write " selected" %> >반품</option>
	<option value='22' <% if selVal="22" then response.write " selected" %> >환불</option>
	<option value='23' <% if selVal="23" then response.write " selected" %> >교환/AS</option>
    <option value='8' <% if selVal="8" then response.write " selected" %> >상품/품절</option>
    <option value='9' <% if selVal="9" then response.write " selected" %> >쿠폰/마일리지/예치금</option>
	<option value='20' <% if selVal="20" then response.write " selected" %> >선물포장</option>
    <option value='10' <% if selVal="10" then response.write " selected" %> >기프트카드</option>
    <option value='11' <% if selVal="11" then response.write " selected" %> >증빙서류</option>
    <option value='12' <% if selVal="12" then response.write " selected" %> >이벤트</option>
    <option value='14' <% if selVal="14" then response.write " selected" %> >메일진</option>
    <option value='15' <% if selVal="15" then response.write " selected" %> >푸시</option>
    <option value='16' <% if selVal="16" then response.write " selected" %> >문자/카카오톡</option>
	<option value='21' <% if selVal="21" then response.write " selected" %> >오프라인매장</option>
	<option value='18' <% if selVal="18" then response.write " selected" %> >기타</option>
	</select>
<%
end Function

function getcustomerfaqgubunname(gubun)
    dim gubunname

    if gubun="1" then
        gubunname="업체"
    elseif gubun="2" then
        gubunname="어드민"
    elseif gubun="3" then
        gubunname="시스템/사이트"
    elseif gubun="4" then
        gubunname="회원"
    elseif gubun="17" then
        gubunname="회원가입"
    elseif gubun="5" then
        gubunname="로그인"
    elseif gubun="13" then
        gubunname="배송/송장"
    elseif gubun="6" then
        gubunname="장바구니/주문/결제/입금"
    elseif gubun="19" then
        gubunname="취소"
    elseif gubun="7" then
        gubunname="반품"
    elseif gubun="22" then
        gubunname="환불"
    elseif gubun="23" then
        gubunname="교환/AS"
    elseif gubun="8" then
        gubunname="상품/품절"
    elseif gubun="9" then
        gubunname="쿠폰/마일리지/예치금"
    elseif gubun="20" then
        gubunname="선물포장"
    elseif gubun="10" then
        gubunname="기프트카드"
    elseif gubun="11" then
        gubunname="증빙서류"
    elseif gubun="12" then
        gubunname="이벤트"
    elseif gubun="14" then
        gubunname="메일진"
    elseif gubun="15" then
        gubunname="푸시"
    elseif gubun="16" then
        gubunname="문자/카카오톡"
    elseif gubun="21" then
        gubunname="오프라인매장"
    elseif gubun="18" then
        gubunname="기타"
    else
        gubunname=""
    end if

    getcustomerfaqgubunname=gubunname
end Function

function Drawmdfaqgubun(selBoxName,selVal,changeFlag)
%>
    <select class='select' name="<%= selBoxName %>" <%= changeFlag %>>
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='1' <% if selVal="1" then response.write " selected" %> >업체</option>
    <option value='17' <% if selVal="17" then response.write " selected" %> >전자계약/전자서명/전자결제</option>
	<option value='2' <% if selVal="2" then response.write " selected" %> >어드민</option>
	<option value='3' <% if selVal="3" then response.write " selected" %> >시스템/사이트</option>
	<option value='4' <% if selVal="4" then response.write " selected" %> >회원</option>
	<option value='5' <% if selVal="5" then response.write " selected" %> >로그인</option>
    <option value='13' <% if selVal="13" then response.write " selected" %> >배송/송장</option>
    <option value='6' <% if selVal="6" then response.write " selected" %> >장바구니/주문/결제/입금</option>
	<option value='19' <% if selVal="19" then response.write " selected" %> >취소</option>
    <option value='7' <% if selVal="7" then response.write " selected" %> >반품</option>
	<option value='22' <% if selVal="22" then response.write " selected" %> >환불</option>
	<option value='23' <% if selVal="23" then response.write " selected" %> >교환/AS</option>
    <option value='8' <% if selVal="8" then response.write " selected" %> >상품/품절</option>
    <option value='9' <% if selVal="9" then response.write " selected" %> >쿠폰/마일리지/예치금</option>
	<option value='20' <% if selVal="20" then response.write " selected" %> >선물포장</option>
    <option value='10' <% if selVal="10" then response.write " selected" %> >기프트카드</option>
    <option value='11' <% if selVal="11" then response.write " selected" %> >증빙서류</option>
    <option value='12' <% if selVal="12" then response.write " selected" %> >이벤트</option>
    <option value='14' <% if selVal="14" then response.write " selected" %> >메일진</option>
    <option value='15' <% if selVal="15" then response.write " selected" %> >푸시</option>
    <option value='16' <% if selVal="16" then response.write " selected" %> >문자/카카오톡</option>
	<option value='21' <% if selVal="21" then response.write " selected" %> >오프라인매장</option>
	<option value='18' <% if selVal="18" then response.write " selected" %> >기타</option>
	</select>
<%
end Function

function getmdfaqgubunname(gubun)
    dim gubunname

    if gubun="1" then
        gubunname="업체"
    elseif gubun="17" then
        gubunname="전자계약/전자서명/전자결제"
    elseif gubun="2" then
        gubunname="어드민"
    elseif gubun="3" then
        gubunname="시스템/사이트"
    elseif gubun="4" then
        gubunname="회원"
    elseif gubun="5" then
        gubunname="로그인"
    elseif gubun="13" then
        gubunname="배송/송장"
    elseif gubun="6" then
        gubunname="장바구니/주문/결제/입금"
    elseif gubun="19" then
        gubunname="취소"
    elseif gubun="7" then
        gubunname="반품"
    elseif gubun="22" then
        gubunname="환불"
    elseif gubun="23" then
        gubunname="교환/AS"
    elseif gubun="8" then
        gubunname="상품/품절"
    elseif gubun="9" then
        gubunname="쿠폰/마일리지/예치금"
    elseif gubun="20" then
        gubunname="선물포장"
    elseif gubun="10" then
        gubunname="기프트카드"
    elseif gubun="11" then
        gubunname="증빙서류"
    elseif gubun="12" then
        gubunname="이벤트"
    elseif gubun="14" then
        gubunname="메일진"
    elseif gubun="15" then
        gubunname="푸시"
    elseif gubun="16" then
        gubunname="문자/카카오톡"
    elseif gubun="21" then
        gubunname="오프라인매장"
    elseif gubun="18" then
        gubunname="기타"
    else
        gubunname=""
    end if

    getmdfaqgubunname=gubunname
end Function

function Drawstafffaqgubun(selBoxName,selVal,changeFlag)
%>
    <select class='select' name="<%= selBoxName %>" <%= changeFlag %>>
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='1' <% if selVal="1" then response.write " selected" %> >업체</option>
	<option value='2' <% if selVal="2" then response.write " selected" %> >어드민</option>
	<option value='3' <% if selVal="3" then response.write " selected" %> >시스템/사이트</option>
	<option value='4' <% if selVal="4" then response.write " selected" %> >회원</option>
	<option value='17' <% if selVal="17" then response.write " selected" %> >회원가입</option>
	<option value='5' <% if selVal="5" then response.write " selected" %> >로그인</option>
    <option value='13' <% if selVal="13" then response.write " selected" %> >배송/송장</option>
    <option value='6' <% if selVal="6" then response.write " selected" %> >장바구니/주문/결제/입금</option>
	<option value='19' <% if selVal="19" then response.write " selected" %> >취소</option>
    <option value='7' <% if selVal="7" then response.write " selected" %> >반품</option>
	<option value='22' <% if selVal="22" then response.write " selected" %> >환불</option>
	<option value='23' <% if selVal="23" then response.write " selected" %> >교환/AS</option>
    <option value='8' <% if selVal="8" then response.write " selected" %> >상품/품절</option>
    <option value='9' <% if selVal="9" then response.write " selected" %> >쿠폰/마일리지/예치금</option>
	<option value='20' <% if selVal="20" then response.write " selected" %> >선물포장</option>
    <option value='10' <% if selVal="10" then response.write " selected" %> >기프트카드</option>
    <option value='11' <% if selVal="11" then response.write " selected" %> >증빙서류</option>
    <option value='12' <% if selVal="12" then response.write " selected" %> >이벤트</option>
    <option value='14' <% if selVal="14" then response.write " selected" %> >메일진</option>
    <option value='15' <% if selVal="15" then response.write " selected" %> >푸시</option>
    <option value='16' <% if selVal="16" then response.write " selected" %> >문자/카카오톡</option>
	<option value='21' <% if selVal="21" then response.write " selected" %> >오프라인매장</option>
	<option value='18' <% if selVal="18" then response.write " selected" %> >기타</option>
	</select>
<%
end Function

function getstafffaqgubunname(gubun)
    dim gubunname

    if gubun="1" then
        gubunname="업체"
    elseif gubun="2" then
        gubunname="어드민"
    elseif gubun="3" then
        gubunname="시스템/사이트"
    elseif gubun="4" then
        gubunname="회원"
    elseif gubun="17" then
        gubunname="회원가입"
    elseif gubun="5" then
        gubunname="로그인"
    elseif gubun="13" then
        gubunname="배송/송장"
    elseif gubun="6" then
        gubunname="장바구니/주문/결제/입금"
    elseif gubun="19" then
        gubunname="취소"
    elseif gubun="7" then
        gubunname="반품"
    elseif gubun="22" then
        gubunname="환불"
    elseif gubun="23" then
        gubunname="교환/AS"
    elseif gubun="8" then
        gubunname="상품/품절"
    elseif gubun="9" then
        gubunname="쿠폰/마일리지/예치금"
    elseif gubun="20" then
        gubunname="선물포장"
    elseif gubun="10" then
        gubunname="기프트카드"
    elseif gubun="11" then
        gubunname="증빙서류"
    elseif gubun="12" then
        gubunname="이벤트"
    elseif gubun="14" then
        gubunname="메일진"
    elseif gubun="15" then
        gubunname="푸시"
    elseif gubun="16" then
        gubunname="문자/카카오톡"
    elseif gubun="21" then
        gubunname="오프라인매장"
    elseif gubun="18" then
        gubunname="기타"
    else
        gubunname=""
    end if

    getstafffaqgubunname=gubunname
end Function
%>
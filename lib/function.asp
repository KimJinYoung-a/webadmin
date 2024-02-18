<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

function fnToPhoneNumber(byval rawPhoneNumber)
    fnToPhoneNumber = rawPhoneNumber
    dim midNumber

    if (rawPhoneNumber = "") then
        exit function
    end if

    if (UBound(Split(rawPhoneNumber, "-")) > 0) then
        exit function
    end if

    if Len(rawPhoneNumber) = 8 then
        '// 16441111
        fnToPhoneNumber = Left(rawPhoneNumber, 4) & "-" & Right(rawPhoneNumber, 4)
    elseif Len(rawPhoneNumber) = 10 or Len(rawPhoneNumber) = 11 then
        '// 0101112222, 01011112222
        midNumber = Left(rawPhoneNumber, Len(rawPhoneNumber) - 4)
        midNumber = Right(midNumber, Len(midNumber) - 3)
        fnToPhoneNumber = Left(rawPhoneNumber, 3) & "-" & midNumber & "-" & Right(rawPhoneNumber, 4)
    end if

end function

function fnGetItemCodeByPublicBarcode(byval ipublicBar,byRef iitemgubun,byRef iitemid,byRef iitemoption)
    dim sqlStr
    fnGetItemCodeByPublicBarcode = False

    sqlStr = "select top 1 b.* " + VbCrlf
    sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b " + VbCrlf
    sqlStr = sqlStr + " where b.barcode='" + CStr(ipublicBar) + "' " + VbCrlf

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
    	iitemgubun = rsget("itemgubun")
    	iitemid = rsget("itemid")
    	iitemoption = rsget("itemoption")
    	fnGetItemCodeByPublicBarcode = True
    end if
    rsget.Close
end function

function getDatabaseTime()
	dim sqlStr

	sqlStr = " select convert(varchar(23),getdate(),21) as currdate "
	rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
    	getDatabaseTime = rsget("currdate")
	end if
    rsget.Close
end function

'// 마진률계산
Function fnPercent(oup,inp,pnt)
	'' if oup=0 or isNull(oup) then exit function ''주석처리 2014/01/16
	if inp=0 or isNull(inp) then exit function
	fnPercent = FormatNumber((1-(CDbl(oup)/CDbl(inp)))*100,pnt) & "%"
End Function

function DDotFormat(byval str,byval n)
	DDotFormat = str
	if IsNULL(str) then Exit function

	if Len(str)> n then
		DDotFormat = Left(str,n) + "..."
	end if
end function



function replaceDelim(byval v)
	replaceDelim = replace(v,"|","")
end	function

function null2Zero(byval v)
	null2Zero = v
	if isNull(v) then null2Zero=0
end function

function MinusFont(byval v)
	MinusFont = "#000000"
	if IsNull(v) then Exit function
	if Not IsNumeric(v) then Exit function

	if (v<0) then MinusFont="#FF0000"
end function

function CsGubun2Name(byval v)
	if IsNull(v) or (v="") then
		Exit function
	end if

	if v="A000" then
		CsGubun2Name = "맞교환"
	elseif v="A001" then
		CsGubun2Name = "누락재발송"
	elseif v="A002" then
		CsGubun2Name = "서비스발송"
	elseif v="A003" then
		CsGubun2Name = "환불요청"
	elseif v="A004" then
		CsGubun2Name = "반품접수"
	elseif v="A005" then
		CsGubun2Name = "외부몰환불요청"
	elseif v="A006" then
		CsGubun2Name = "출고시유의사항"
	elseif v="A007" then
		CsGubun2Name = "신용카드취소요청"
	elseif v="A008" then
		CsGubun2Name = "배송대기중취소"
	elseif v="A009" then
		CsGubun2Name = "기타"
	else

	end if
end function

function CsState2Name(byval v)
	if IsNull(v) or (v="") then
		Exit function
	end if

	if v="0" then

	elseif v="B001" then
		CsState2Name = "접수"
	elseif v="B004" then
		CsState2Name = "운송장입력"
	elseif v="B003" then
		CsState2Name = ""
	elseif v="B005" then
		CsState2Name = "업체확인요청"
	elseif v="B006" then
		CsState2Name = "업체처리완료"
	elseif v="B007" then
		CsState2Name = "처리완료"
	else

	end if
end function

function UpCheBeasongState2Name(byval v)
	if IsNull(v) or (v="") then
		Exit function
	end if

	if v="3" then
		UpCheBeasongState2Name = "주문확인"
	elseif v="7" then
		UpCheBeasongState2Name = "배송완료"
	else

	end if

end function

function UpCheBeasongStateColor(byval v)

	if v="3" then
		UpCheBeasongStateColor = "#3333FF"
	elseif v="7" then
		UpCheBeasongStateColor = "#FF3333"
	else
		UpCheBeasongStateColor = "#FFFFFF"
	end if

end function

function Blank3NBSP(byval v)
	if IsNull(v) or (v="") then
		Blank3NBSP="&nbsp;"
	else
		Blank3NBSP=v
	end if
end function

function DeliverDivCd2Nm(byval divcd)
    if IsNULL(divcd) or divcd="" then Exit function
    DeliverDivCd2Nm = ""

    rsget.Open "Select divname from db_order.dbo.tbl_songjang_div where divcd=" & CStr(divcd) & "",dbget,1
    if Not(rsget.EOF or rsget.BOF) then
        DeliverDivCd2Nm = db2html(rsget(0))
    end if
    rsget.Close
end function

function DeliverDivTrace(byval divcd)
    if IsNULL(divcd) or divcd="" then Exit function
    DeliverDivTrace = ""

    rsget.Open "Select findurl from db_order.dbo.tbl_songjang_div where divcd=" & CStr(divcd) & "",dbget,1
    if Not(rsget.EOF or rsget.BOF) then
        DeliverDivTrace = db2html(rsget(0))
    end if
    rsget.Close
end function


public function ynColor(v)
	ynColor = "#000000"

	if v="Y" then
		ynColor = "#0000FF"
	elseif v="N" then
		ynColor = "#FF0000"
	else

	end if
end function


public function mwdivColor(v)
	mwdivColor = "#000000"

	if v="M" then
		mwdivColor = "#FF0000"
	elseif v="U" then
		mwdivColor = "#0000FF"
	elseif v="W" then
		mwdivColor = "#000000"
	else

	end if
end function


public function mwdivName(v)

	if v="M" then
		mwdivName = "매입"
	elseif v="U" then
		mwdivName = "업체"
	elseif v="W" then
		mwdivName = "위탁"
	else

	end if
end function

public function vatIncludeName(v)
	if v="Y" then
		vatIncludeName = "과세"
	elseif v="N" then
        vatIncludeName = "면세"
	end if
end function

public Function GetdeliverytypeName(v)
    Select Case v
        Case "1": GetdeliverytypeName="텐바이텐배송"
        Case "2": GetdeliverytypeName="업체 무료배송"
        Case "4": GetdeliverytypeName="텐바이텐 무료배송"
        Case "6": GetdeliverytypeName="현장수령"
        Case "7": GetdeliverytypeName="업체착불배송"
        Case "9": GetdeliverytypeName="업체조건배송"
        Case Else : GetdeliverytypeName=""
    End Select
end Function

public Function getSellYnName(v)
    Select Case v
        Case "Y": getSellYnName="판매중"
        Case "S": getSellYnName="일시품절"
        Case "N": getSellYnName="품절"
        Case Else : getSellYnName=""
    End Select
end Function

public function GetJungsanGubunName(v)
    if v="B011" then
	    GetJungsanGubunName = "위탁판매"
	elseif v="B012" then
	    GetJungsanGubunName = "업체위탁"
	elseif v="B021" then
	    GetJungsanGubunName = "오프매입"
	elseif v="B022" then
	    GetJungsanGubunName = "매장매입"
	elseif v="B023" then
	    GetJungsanGubunName = "가맹점매입"
	elseif v="B031" then
	    GetJungsanGubunName = "출고매입"
	elseif v="B032" then
	    GetJungsanGubunName = "센터매입"
	elseif v="B999" then
	    GetJungsanGubunName = "기타보정"
	elseif v="B013" then
	    GetJungsanGubunName = "출고위탁"
    else
        GetJungsanGubunName = v
    end if
end function

public function GetdeliverGubunName(v)
    if v="B011" or v="B031" or v="B013" then
		GetdeliverGubunName = "텐바이텐물류"
    elseif v="B012" or v="B022" then
		GetdeliverGubunName = "업체"
    else
        GetdeliverGubunName = v
    end if
end function

public function jumunDivColor(v)
	jumunDivColor = "#000000"

	if v="1" then
		jumunDivColor = "#000000"
	elseif v="2" then
		jumunDivColor = "#000000"
	elseif v="3" then
		jumunDivColor = "#000000"
	elseif v="4" then

	elseif v="5" then
		jumunDivColor = "#0000FF"
	elseif v="9" then
		jumunDivColor = "#FF0000"
	else

	end if
end function


public function jumunDivName(v)
	if v="1" then
		jumunDivName = "웹주문"
	elseif v="2" then
		jumunDivName = "서비스매출"
	elseif v="3" then
		jumunDivName = "예약주문"
	elseif v="4" then

	elseif v="5" then
		jumunDivName = "외부몰"
	elseif v="9" then
		jumunDivName = "마이너스"
	else

	end if
end function

public function IpkumDivColor(v)
		if v="0" then
			IpkumDivColor="#FF0000"
		elseif v="1" then
			IpkumDivColor="#FF0000"
		elseif v="2" then
			IpkumDivColor="#000000"
		elseif v="3" then
			IpkumDivColor="#000000"
		elseif v="4" then
			IpkumDivColor="#0000FF"
		elseif v="5" then
			IpkumDivColor="#444400"
		elseif v="6" then
			IpkumDivColor="#FFFF00"
		elseif v="7" then
			IpkumDivColor="#004444"
		elseif v="8" then
			IpkumDivColor="#FF00FF"
		end if
	end function

	Public function JumunMethodName(v)
		if v="7" then
			JumunMethodName="무통장"
		elseif v="14" then
			JumunMethodName="편의점결제"
		elseif v="100" then
			JumunMethodName="신용카드"
		elseif v="20" then
			JumunMethodName="실시간이체"
		elseif v="30" then
			JumunMethodName="포인트"
		elseif v="50" then
			JumunMethodName="입점몰결제"
		elseif v="80" then
			JumunMethodName="All@카드"
		elseif v="90" then
			JumunMethodName="상품권결제"
		elseif v="110" then
			JumunMethodName="OK+신용"
	    elseif v="400" then
			JumunMethodName="핸드폰결제"
	    elseif v="550" then
			JumunMethodName="기프팅"
	    elseif v="560" then
			JumunMethodName="기프티콘"
		end if
	end function

    function DrawJumunMethod(selBoxName,selVal,chplg)
%>
    <select class='select' name="<%= selBoxName %>" <%= chplg %>>
		<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
		<option value='7' <% if cstr(selVal)=cstr("7") then response.write " selected" %> >무통장</option>
        <option value='14' <% if cstr(selVal)=cstr("14") then response.write " selected" %> >편의점결제</option>
		<option value='100' <% if cstr(selVal)=cstr("100") then response.write " selected" %> >신용카드</option>
		<option value='20' <% if cstr(selVal)=cstr("20") then response.write " selected" %> >실시간이체</option>
        <option value='30' <% if cstr(selVal)=cstr("30") then response.write " selected" %> >포인트</option>
        <option value='50' <% if cstr(selVal)=cstr("50") then response.write " selected" %> >입점몰결제</option>
        <option value='80' <% if cstr(selVal)=cstr("80") then response.write " selected" %> >All@카드</option>
        <option value='90' <% if cstr(selVal)=cstr("90") then response.write " selected" %> >상품권결제</option>
        <option value='110' <% if cstr(selVal)=cstr("110") then response.write " selected" %> >OK+신용</option>
		<option value='400' <% if cstr(selVal)=cstr("400") then response.write " selected" %> >핸드폰결제</option>
        <option value='550' <% if cstr(selVal)=cstr("550") then response.write " selected" %> >기프팅</option>
        <option value='560' <% if cstr(selVal)=cstr("560") then response.write " selected" %> >기프티콘</option>
	</select>
<%
    end Function

	Public function IpkumDivName(v)
		if v="0" then
			IpkumDivName="주문대기"
		elseif v="1" then
			IpkumDivName="주문실패"
		elseif v="2" then
			IpkumDivName="주문접수"
		elseif v="3" then
			IpkumDivName="주문접수"
		elseif v="4" then
			IpkumDivName="결제완료"
		elseif v="5" then
			IpkumDivName="주문통보"
		elseif v="6" then
			IpkumDivName="상품준비"
		elseif v="7" then
			IpkumDivName="일부출고"
		elseif v="8" then
			IpkumDivName="출고완료"
		end if
	end function

    ' 주문마스터, 주문상태      ' 2020.11.11 한용민 생성
    function DrawIpkumDivName(selBoxName,selVal,chplg)
%>
    <select name="<%= selBoxName %>" class="select" <%= chplg %>>
		<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
        <option value='0' <% if cstr(selVal)=cstr("0") then response.write " selected" %> >주문대기</option>
		<option value='1' <% if cstr(selVal)=cstr("1") then response.write " selected" %> >주문실패</option>
		<option value='2' <% if cstr(selVal)=cstr("2") then response.write " selected" %> >주문접수</option>
        <option value='4' <% if cstr(selVal)=cstr("4") then response.write " selected" %> >결제완료</option>
        <option value='5' <% if cstr(selVal)=cstr("5") then response.write " selected" %> >주문통보</option>
        <option value='6' <% if cstr(selVal)=cstr("6") then response.write " selected" %> >상품준비</option>
        <option value='7' <% if cstr(selVal)=cstr("7") then response.write " selected" %> >일부출고</option>
        <option value='8' <% if cstr(selVal)=cstr("8") then response.write " selected" %> >출고완료</option>
	</select>
<%
    end Function

' PG사구분명    ' 2020.01.20 한용민 생성
function fnGetPggubunName(pggubun)
    dim tmppggubun

    Select Case pggubun
        Case "KA"
            tmppggubun = "카카오페이"
        Case "KK"
            tmppggubun = "카카오페이"
        Case "IN"
            tmppggubun = "이니시스"
        Case "DA"
            tmppggubun = "엘지데이콤"
        Case "NP"
            tmppggubun = "네이버페이"
        Case "PY"
            tmppggubun = "페이코"
        Case "TS"
            tmppggubun = "토스"
        Case "SP"
            tmppggubun = "삼성페이"
        Case "AP"
            tmppggubun = "Apple Pay"
        Case "PP"
            tmppggubun = "PIN Pay"
        Case Else
            tmppggubun = pggubun
    End Select

    fnGetPggubunName=tmppggubun
end function

' 주문마스터, PG구분    ' 허진원 생성
function DrawPggubunName(selBoxName,selVal,incAcct,chplg)
%>
    <select name="<%= selBoxName %>" class="select" <%= chplg %>>
		<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
		<option value='IN' <% if cstr(selVal)=cstr("IN") then response.write " selected" %> >이니시스</option>
        <% if incAcct="Y" then %>
		<option value='DA' <% if cstr(selVal)=cstr("DA") then response.write " selected" %> >엘지데이콤</option>
        <option value='BK' <% if cstr(selVal)=cstr("BK") then response.write " selected" %> >무통장결제</option>
        <option value='CV' <% if cstr(selVal)=cstr("CV") then response.write " selected" %> >편의점결제</option>
        <% end if%>
        <option value='KK' <% if cstr(selVal)=cstr("KK") then response.write " selected" %> >카카오페이</option>
        <option value='NP' <% if cstr(selVal)=cstr("NP") then response.write " selected" %> >네이버페이</option>
        <option value='PY' <% if cstr(selVal)=cstr("PY") then response.write " selected" %> >페이코</option>
        <option value='TS' <% if cstr(selVal)=cstr("TS") then response.write " selected" %> >토스</option>
        <option value='SP' <% if cstr(selVal)=cstr("SP") then response.write " selected" %> >삼성페이</option>
        <option value='AP' <% if cstr(selVal)=cstr("AP") then response.write " selected" %> >Apple Pay</option>
        <option value='PP' <% if cstr(selVal)=cstr("PP") then response.write " selected" %> >PIN Pay</option>
	</select>
<%
end Function

Sub gotoPageHTML(page, Pagecount, asp_name)
   Dim blockpage, i
   blockpage=Int((page-1)/10)*10+1

   '********** 이전 10 개 구문 시작 **********
   if blockPage = 1 Then
'      Response.Write "<font color= silver>[이전 10개]</font>["
      Response.Write "["
   Else
      Response.Write"<a href='"&asp_name&"?gotopage=" & blockPage-10 & "'>[이전 10개]</a> ["
   End If
   '********** 이전 10 개 구문 끝**********

   i=1
   Do Until i > 10 or blockpage > Pagecount
      If blockpage=int(page) Then
         Response.Write "<font size=2 color= gray>" & blockpage & "</font>"
      Else
         Response.Write"<a href='"&asp_name&"?gotopage=" & blockpage & "'>" & blockpage & "</a> "
      End If

      blockpage=blockpage+1
      i = i + 1
   Loop

   '********** 다음 10 개 구문 시작**********
   if blockpage > Pagecount Then
'      Response.Write "] <font color= silver>[다음 10개]</font>"
      Response.Write "]"
   Else
      Response.write"]<a href='"&asp_name&"?gotopage=" & blockpage & "'>[다음 10개]</a>"
   End If
   '********** 다음 10 개 구문 끝**********
End Sub
Sub gotoPageHTML2(page, Pagecount,table_name,site_name)
   Dim blockpage, i
   blockpage=Int((page-1)/10)*10+1

   '********** 이전 10 개 구문 시작 **********
   if blockPage = 1 Then
      Response.Write "["
   Else
      Response.Write"<a href='admin_board_list.asp?table_name="&table_name&"&site_name="&site_name&"&gotopage=" & blockPage-10 & "'>[이전 10개]</a> ["
   End If
   '********** 이전 10 개 구문 끝**********

   i=1
   Do Until i > 10 or blockpage > Pagecount
      If blockpage=int(page) Then
         Response.Write "<font size=2 color= gray>" & blockpage & "</font>"
      Else
         Response.Write"<a href='admin_board_list.asp?table_name="&table_name&"&site_name="&site_name&"&gotopage=" & blockpage & "'>" & blockpage & "</a> "
      End If

      blockpage=blockpage+1
      i = i + 1
   Loop

   '********** 다음 10 개 구문 시작**********
   if blockpage > Pagecount Then
      Response.Write "]"
   Else
      Response.write"]<a href='admin_board_list.asp?table_name="&table_name&"&site_name="&site_name&"&gotopage=" & blockpage & "'>[다음 10개]</a>"
   End If
   '********** 다음 10 개 구문 끝**********
End Sub
Sub gotoPageHTML3(page, Pagecount,table_name,site_name)
   Dim blockpage, i
   blockpage=Int((page-1)/10)*10+1

   '********** 이전 10 개 구문 시작 **********
   if blockPage = 1 Then
      Response.Write "["
   Else
      Response.Write"<a href='admin_board_list_all.asp?table_name="&table_name&"&site_name="&site_name&"&gotopage=" & blockPage-10 & "'>[이전 10개]</a> ["
   End If
   '********** 이전 10 개 구문 끝**********

   i=1
   Do Until i > 10 or blockpage > Pagecount
      If blockpage=int(page) Then
         Response.Write "<font size=2 color= gray>" & blockpage & "</font>"
      Else
         Response.Write"<a href='admin_board_list_all.asp?table_name="&table_name&"&site_name="&site_name&"&gotopage=" & blockpage & "'>" & blockpage & "</a> "
      End If

      blockpage=blockpage+1
      i = i + 1
   Loop

   '********** 다음 10 개 구문 시작**********
   if blockpage > Pagecount Then
      Response.Write "]"
   Else
      Response.write"]<a href='admin_board_list_all.asp?table_name="&table_name&"&site_name="&site_name&"&gotopage=" & blockpage & "'>[다음 10개]</a>"
   End If
   '********** 다음 10 개 구문 끝**********
End Sub

Sub drawSelectBoxPartner(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>ALL</option><%
   query1 = " select id,company_name from [db_partner].[dbo].tbl_partner where userdiv=999"
   query1 = query1 + " and isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       'rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("id")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("id")&"' "&tmp_str&">" + rsget("id") + " [" + rsget("company_name") + "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

function IsTPLMakerID(makerid)
	dim query1

	IsTPLMakerID = False

	query1 = " select top 1 tplcompanyid "
	query1 = query1 + " from "
	query1 = query1 + " [db_partner].[dbo].[tbl_partner] "
	query1 = query1 + " where id = '" & makerid & "' and tplcompanyid is not NULL "
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		IsTPLMakerID = True
	end if
	rsget.close
end function

function IsTPLItemCode(itemgubun, itemid, itemoption)
	dim query1

	IsTPLItemCode = False

	if (itemgubun = "10") then
		query1 = " select top 1 tplcompanyid "
		query1 = query1 + " from "
		query1 = query1 + " 	[db_partner].[dbo].[tbl_partner] p "
		query1 = query1 + " 	join [db_item].[dbo].[tbl_item] i "
		query1 = query1 + " 	on "
		query1 = query1 + " 		1 = 1 "
		query1 = query1 + " 		and i.makerid = p.id "
		query1 = query1 + " 		and i.itemid = " & itemid & " "
		query1 = query1 + " where "
		query1 = query1 + " 	tplcompanyid is not NULL "
	else
		query1 = " select top 1 tplcompanyid "
		query1 = query1 + " from "
		query1 = query1 + " 	[db_partner].[dbo].[tbl_partner] p "
		query1 = query1 + " 	join [db_shop].[dbo].[tbl_shop_item] o on o.makerid = p.id "
		query1 = query1 + " where "
		query1 = query1 + " 	1 = 1 "
		query1 = query1 + " 	and o.itemgubun = '" & itemgubun & "' "
		query1 = query1 + " 	and o.shopitemid = " & itemid & " "
		query1 = query1 + " 	and o.itemoption = '" & itemoption & "' "
		query1 = query1 + " 	and tplcompanyid is not NULL "
	end if

	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		IsTPLItemCode = True
	end if
	rsget.close
end function

function IsTPLIthinksoItemCode(itemgubun, itemid, itemoption)
	dim query1

	IsTPLIthinksoItemCode = False

	if (itemgubun = "10") then
		query1 = " select top 1 tplcompanyid "
		query1 = query1 + " from "
		query1 = query1 + " 	[db_partner].[dbo].[tbl_partner] p "
		query1 = query1 + " 	join [db_item].[dbo].[tbl_item] i "
		query1 = query1 + " 	on "
		query1 = query1 + " 		1 = 1 "
		query1 = query1 + " 		and i.makerid = p.id "
		query1 = query1 + " 		and i.itemid = " & itemid & " "
		query1 = query1 + " where "
		query1 = query1 + " 	p.tplcompanyid = 'tplithinkso' "
	else
		query1 = " select top 1 tplcompanyid "
		query1 = query1 + " from "
		query1 = query1 + " 	[db_partner].[dbo].[tbl_partner] p "
		query1 = query1 + " 	join [db_shop].[dbo].[tbl_shop_item] o on o.makerid = p.id "
		query1 = query1 + " where "
		query1 = query1 + " 	1 = 1 "
		query1 = query1 + " 	and o.itemgubun = '" & itemgubun & "' "
		query1 = query1 + " 	and o.shopitemid = " & itemid & " "
		query1 = query1 + " 	and o.itemoption = '" & itemoption & "' "
		query1 = query1 + " 	and p.tplcompanyid = 'tplithinkso' "
	end if

	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		IsTPLIthinksoItemCode = True
	end if
	rsget.close
end function

Sub drawSelectBoxTPLGubun(selectBoxName, selectedId)
	'// /admin/fran/offinvoice_list.asp
	'// /lib/classes/stock/offinvoicecls.asp
	dim tmp_str, sqlStr

	sqlStr = " select replace(p.id, '3pl', 'tpl') as tplcompanyid, p.company_name as tplcompanyname " & vbCrLf
	sqlStr = sqlStr & " from " & vbCrLf
	sqlStr = sqlStr & " 	[db_user].[dbo].[tbl_user_c] u with (nolock)" & vbCrLf
	sqlStr = sqlStr & " 	join [db_partner].[dbo].[tbl_partner] p with (nolock) on p.id = u.userid " & vbCrLf
	sqlStr = sqlStr & " where " & vbCrLf
	sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
	sqlStr = sqlStr & " 	and p.userdiv = '903' " & vbCrLf
	sqlStr = sqlStr & " 	and u.userdiv = '21' " & vbCrLf
	sqlStr = sqlStr & " 	and p.isusing = 'Y' " & vbCrLf
    sqlStr = sqlStr & " order by p.id "
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	%>
	<select class="select" name="<%= selectBoxName %>">
		<option value="">선택</option>
		<option value="3X" <% if (selectedId="3X") then response.write "selected" %> >3PL 제외</option>
		<option value="">------</option>
		<%
		if not rsget.EOF  then
			rsget.Movefirst

			do until rsget.EOF
				tmp_str = ""
				if Lcase(selectedId) = Lcase(rsget("tplcompanyid")) then
					tmp_str = " selected"
				end if

				response.write("<option value='"&rsget("tplcompanyid")&"' "&tmp_str&">"&db2html(rsget("tplcompanyname"))&"</option>")
				rsget.MoveNext
			loop
		end if
		rsget.close
		%>
	</select>
<%
End Sub

Sub drawSelectBoxTPLGubunNew(selectBoxName, selectedId)
	'// /admin/fran/offinvoice_list.asp
	'// /lib/classes/stock/offinvoicecls.asp
	dim tmp_str, sqlStr

	sqlStr = " select replace(p.id, '3pl', 'tpl') as tplcompanyid, p.company_name as tplcompanyname " & vbCrLf
	sqlStr = sqlStr & " from " & vbCrLf
	sqlStr = sqlStr & " 	[db_user].[dbo].[tbl_user_c] u with (nolock)" & vbCrLf
	sqlStr = sqlStr & " 	join [db_partner].[dbo].[tbl_partner] p with (nolock) on p.id = u.userid " & vbCrLf
	sqlStr = sqlStr & " where " & vbCrLf
	sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
	sqlStr = sqlStr & " 	and p.userdiv = '903' " & vbCrLf
	sqlStr = sqlStr & " 	and u.userdiv = '21' " & vbCrLf
	sqlStr = sqlStr & " 	and p.isusing = 'Y' " & vbCrLf
    sqlStr = sqlStr & " order by p.id "
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	%>
	<select class="select" name="<%= selectBoxName %>">
		<option value="">선택</option>
		<option value="3X" <% if (selectedId="3X") then response.write "selected" %> >3PL 제외</option>
		<option value="">------</option>
		<%
		if not rsget.EOF  then
			rsget.Movefirst

			do until rsget.EOF
				tmp_str = ""
				if Lcase(selectedId) = Lcase(rsget("tplcompanyid")) then
					tmp_str = " selected"
				end if

				response.write("<option value='"&rsget("tplcompanyid")&"' "&tmp_str&">"&db2html(rsget("tplcompanyname"))&"</option>")
				rsget.MoveNext
			loop
		end if
		rsget.close
		%>
	</select>
<%
End Sub

Sub drawSelectBoxOFFChargeDiv(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="2" <% if (selectedId="2") then response.write "selected" %> >텐바이텐 위탁
		<option value="4" <% if (selectedId="4") then response.write "selected" %> >텐바이텐 매입
		<option value="5" <% if (selectedId="5") then response.write "selected" %> >매입출고 정산
		<option value="6" <% if (selectedId="6") then response.write "selected" %> >업체 위탁
		<option value="8" <% if (selectedId="8") then response.write "selected" %> >업체 매입
	</select>
<%
End Sub

''정산 구분 세팅시 사용
Sub drawSelectBoxOFFJungsanCommCD(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="B011" <% if (selectedId="B011") then response.write "selected" %> >텐바이텐 위탁
		<option value="B031" <% if (selectedId="B031") then response.write "selected" %> >매입출고 정산
		<option value="B012" <% if (selectedId="B012") then response.write "selected" %> >업체 위탁
		<option value="B013" <% if (selectedId="B013") then response.write "selected" %> >출고 위탁
		<option value="B022" <% if (selectedId="B022") then response.write "selected" %> >매장 매입
		<option value="B023" <% if (selectedId="B023") then response.write "selected" %> >가맹점 매입
		<!-- option value="B021" <% if (selectedId="B021") then response.write "selected" %> >오프 매입 -->
	</select>
<%
End Sub

''정산 방식 구분
Sub drawSelectBoxJGubun(selectBoxName,selectedId)
    dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="MM" <% if (selectedId="MM") then response.write "selected" %> >매입 정산
		<option value="CC" <% if (selectedId="CC") then response.write "selected" %> >수수료 정산
		<option value="CE" <% if (selectedId="CE") then response.write "selected" %> >기타 정산
	</select>
<%
End Sub

''정산 구분 조회시 사용
Sub drawSelectBoxOFFJungsanCommCDQuery(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="B011" <% if (selectedId="B011") then response.write "selected" %> >텐바이텐 위탁
		<option value="B031" <% if (selectedId="B031") then response.write "selected" %> >매입출고 정산
		<option value="B012" <% if (selectedId="B012") then response.write "selected" %> >업체 위탁
		<option value="B022" <% if (selectedId="B022") then response.write "selected" %> >매장 매입
		<option value="B023" <% if (selectedId="B023") then response.write "selected" %> >가맹점 매입
		<option value="B021" <% if (selectedId="B021") then response.write "selected" %> >오프 매입
	</select>
<%
End Sub

Sub drawSelectBoxOFFJungsanCommCDmulti(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="B012" <% if (selectedId="B012") then response.write "selected" %> >업체 위탁
		<option value="B022" <% if (selectedId="B022") then response.write "selected" %> >매장 매입
		<option value="B088" <% if (selectedId="B088") then response.write "selected" %> >업체 위탁+매장 매입
		<option value="B031" <% if (selectedId="B031") then response.write "selected" %> >출고 매입
		<option value="B011" <% if (selectedId="B011") then response.write "selected" %> >텐바이텐위탁
		<option value="B013" <% if (selectedId="B013") then response.write "selected" %> >출고 위탁
		<option value="B099" <% if (selectedId="B099") then response.write "selected" %> >출고 매입+텐바이텐위탁
		<option value="B077" <% if (selectedId="B077") then response.write "selected" %> >출고 매입+텐바이텐위탁+출고 위탁
	</select>
<%
End Sub

'//사용안함
'//디비화 시킴 오프라인펑션(offshop_function.asp)내에 있는 drawoffshop_commoncode 사용
Sub drawSelectBoxShopDiv(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="2" <% if (selectedId="2") then response.write "selected" %> >직영[대표]
		<option value="1" <% if (selectedId="1") then response.write "selected" %> >직영
		<option value="4" <% if (selectedId="4") then response.write "selected" %> >가맹점[대표]
		<option value="3" <% if (selectedId="3") then response.write "selected" %> >가맹점
		<option value="6" <% if (selectedId="6") then response.write "selected" %> >도매[대표]
		<option value="5" <% if (selectedId="5") then response.write "selected" %> >도매
		<option value="8" <% if (selectedId="8") then response.write "selected" %> >해외[대표]
		<option value="7" <% if (selectedId="7") then response.write "selected" %> >해외
		<option value="9" <% if (selectedId="9") then response.write "selected" %> >ETC
		<option value="12" <% if (selectedId="12") then response.write "selected" %> >아이띵소[대표]
		<option value="11" <% if (selectedId="11") then response.write "selected" %> >아이띵소
	</select>
<%
End Sub

Sub drawSelectBoxWebDesigner(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="midesign" <% if (selectedId="midesign") then response.write "selected" %> >박미경
		<option value="violetiris" <% if (selectedId="violetiris") then response.write "selected" %> >유미화
		<option value="ofd0413" <% if (selectedId="ofd0413") then response.write "selected" %> >이유미
		<option value="siridoctor" <% if (selectedId="siridoctor") then response.write "selected" %> >김은주
		<option value="ea0411" <% if (selectedId="ea0411") then response.write "selected" %> >배은애
		<option value="stsunny" <% if (selectedId="stsunny") then response.write "selected" %> >이미선
		<option value="yobebedh" <% if (selectedId="yobebedh") then response.write "selected" %> >이다혜
		<option value="tigger" <% if (selectedId="tigger") then response.write "selected" %> >임은미
		<option value="tym84" <% if (selectedId="tym84") then response.write "selected" %> >탁연미
		<option value="sun6363" <% if (selectedId="sun6363") then response.write "selected" %> >송선영
		<option value="myhj23" <% if (selectedId="myhj23") then response.write "selected" %> >전혜린
		<option value="zhenghe" <% if (selectedId="zhenghe") then response.write "selected" %> >김정화
	</select>
<%
End Sub


Sub drawSelectBox(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>ALL</option><%
   query1 = " select id,company_name from [db_partner].[dbo].tbl_partner "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("id")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("id")&"' "&tmp_str&">"&rsget("company_name")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBox2(selectBoxName,user_div,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select a.userid,a.socname from [db_user].[dbo].tbl_user_c a  "
   query1 = query1 + " where a.userdiv = '" & user_div & "' "
   'query1 = query1 + " and a.isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" & rsget("userid") & "' " & tmp_str & ">" & rsget("userid") & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxChulgo(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select a.userid,a.socname from [db_user].[dbo].tbl_user_c a  "
   query1 = query1 + " where a.userdiv = '21' "
   query1 = query1 + " and a.isusing='Y'"
   query1 = query1 + " order by a.userid"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" & rsget("userid") & "' " & tmp_str & ">" & rsget("userid") & " [" & db2html(rsget("socname")) & "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxDesigner(selectBoxName,selectedId)
   dim tmp_str,query1

   NewDrawSelectBoxDesignerwithName selectBoxName,selectedId
   exit sub

   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,socname from [db_user].[dbo].tbl_user_c  "
   ''query1 = query1 & " where a.userid = b.userid "
   query1= query1 & " order by userid asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxDiscountDesigner(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select distinct c.userid,c.socname from [db_user].[dbo].tbl_user_c c, [db_item].[dbo].tbl_item i "
   query1 = query1 & " where c.userid = i.makerid "
   query1 = query1 & " and i.isusing = 'Y' "
   query1 = query1 & " and i.sailyn = 'Y' "
   query1 = query1 & " order by c.userid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

' 사용중지. NewDrawSelectBoxDesignerwithNameAndUserDIV 펑션 사용하세요.
Sub drawSelectBoxOffShop(selectBoxName,selectedId)
dim tmp_str,query1

	Call NewDrawSelectBoxDesignerwithNameAndUserDIV(selectBoxName,selectedId, "21")
	Exit Sub

	query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user"
	query1 = query1 & " where isusing='Y' "
	query1 = query1 & " and userid<>'streetshop000'"
	query1 = query1 & " and userid<>'streetshop800'"
	query1 = query1 & " and userid<>'streetshop870'"
	query1 = query1 & " and userid<>'streetshop700'"
	query1 = query1 & " and userid<>'ithinkshop000'"
	query1 = query1 & " order by convert(int,shopdiv)+10 asc"

	'response.write query1 & "<Br>"
	rsget.Open query1,dbget,1
%>
<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
   if not rsget.EOF  then
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
end sub

Sub drawSelectBoxOffShopNot000(selectBoxName,selectedId)
	dim tmp_str,query1

	Call NewDrawSelectBoxDesignerwithNameAndUserDIV(selectBoxName,selectedId, "21")
	Exit Sub

   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"
   query1 = query1 & " and userid<>'ithinkshop000'"
	query1 = query1 & " order by convert(int,shopdiv)+10 asc"

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
end sub

Sub drawSelectBoxOnIpjumShop(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
	query1 = " select c.userid,c.socname "
	query1 = query1 & " from db_user.dbo.tbl_user_c c "
	query1 = query1 & " left Join db_partner.dbo.tbl_partner_addInfo f "
	query1 = query1 & " on c.userid=f.partnerid "
	query1 = query1 & " where 1=1 "
	query1 = query1 & " and c.isusing='Y' "
	query1 = query1 & " and c.userdiv='50' "
	query1 = query1 & " order by c.userid "
   rsget.Open query1,dbget,1

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

Sub drawSelectBoxEtcMeachul(selectBoxName,selectedId)
%>
<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<option value='promotion' <%if selectedId="promotion" then response.write " selected"%>>promotion</option>
	<option value='etcAcademy' <%if selectedId="etcAcademy" then response.write " selected"%>>etcAcademy</option>
	<option value='etcithinkso' <%if selectedId="etcithinkso" then response.write " selected"%>>etcithinkso</option>
</select>
<%
end sub

Sub drawSelectBoxOffShopAll(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"
	query1 = query1 & " order by convert(int,shopdiv)+10 asc"

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

''   if Lcase(selectedId) = Lcase("cafe001") then
''   		tmp_str = " selected"
''   else
''   		tmp_str = " "
''   end if
''
''   response.write("<option value='cafe001' " + tmp_str + ">cafe001/대학로1층카페</option>")
''
''   if Lcase(selectedId) = Lcase("cafe002") then
''   		tmp_str = " selected"
''   else
''   		tmp_str = " "
''   end if
''   response.write("<option value='cafe002' " + tmp_str + ">cafe002/Zoom</option>")
''
''   if Lcase(selectedId) = Lcase("cafe003") then
''   		tmp_str = " selected"
''   else
''   		tmp_str = " "
''   end if
''   response.write("<option value='cafe003' " + tmp_str + ">cafe003/College</option>")
   response.write("</select>")
end sub

Sub drawSelectBoxOffShopNotUsingAll(selectBoxName,selectedId)
   dim tmp_str,query1

	Call NewDrawSelectBoxDesignerwithNameAndUserDIV(selectBoxName,selectedId, "21")
	Exit Sub

	query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
	query1 = query1 & " where 1=1"
	query1 = query1 & " and userid<>'streetshop000'"
	query1 = query1 & " and userid<>'streetshop800'"
	query1 = query1 & " and userid<>'streetshop870'"
	query1 = query1 & " and userid<>'streetshop700'"
	query1 = query1 & " and userid<>'cafe002'"
	query1 = query1 & " order by isusing desc, convert(int,shopdiv)+10 asc, userid asc"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
   if not rsget.EOF  then
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
end sub

Sub drawSelectBoxOffShopWith000(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   ''query1 = query1 & " where userid<>'cafe002' "
   ''query1 = query1 & " where isusing='Y' "
   query1 = query1 & " order by convert(int,shopdiv)+10 asc"

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
end sub

Sub drawSelectBoxOpenOffShop(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
''   query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u, [db_shop].[dbo].tbl_shop_designer d"
''   query1 = query1 & " where u.isusing='Y' "
''   query1 = query1 & " and u.userid=d.shopid"
''   query1 = query1 & " and d.makerid='" + session("ssBctID") + "'"
''   query1 = query1 & " and d.adminopen='Y'"
''   query1 = query1 & " and u.userid<>'streetshop000'"
''   query1 = query1 & " and u.userid<>'streetshop800'"
''   query1 = query1 & " and u.userid<>'streetshop870'"
''   query1 = query1 & " and u.userid<>'streetshop700'"

   ''매출등 조회 가능 샵.
   query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u"
   query1 = query1 & "      Join [db_shop].[dbo].tbl_shop_designer d"
   query1 = query1 & "      on u.userid=d.shopid"
   query1 = query1 & " where u.isusing='Y' "
   query1 = query1 & " and d.makerid='" + session("ssBctID") + "'"
   query1 = query1 & " and ((d.comm_cd in ('B011','B012','B022','B031','B013')) or (d.defaultbeasongdiv=2))" '''텐위, 업위, 업매 또는 매장업배 , 출고매입 추가
   query1 = query1 & " and u.userid<>'streetshop000'"
   query1 = query1 & " and u.userid<>'streetshop800'"
   query1 = query1 & " and u.userid<>'streetshop870'"
   query1 = query1 & " and u.userid<>'streetshop700'"


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
end sub

Sub drawSelectBoxOpenOffShopByMaker(selectBoxName,selectedId,makerid)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u, [db_shop].[dbo].tbl_shop_designer d"
   query1 = query1 & " where u.isusing='Y' "
   query1 = query1 & " and u.userid=d.shopid"
   query1 = query1 & " and d.makerid='" + makerid + "'"
   query1 = query1 & " and d.adminopen='Y'"
   query1 = query1 & " and u.userid<>'streetshop000'"
   query1 = query1 & " and u.userid<>'streetshop800'"
   query1 = query1 & " and u.userid<>'streetshop870'"
   query1 = query1 & " and u.userid<>'streetshop700'"

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
end sub

Sub drawSelectBoxOffShopChargeId(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
   query1 = " select chargeuser,chargename from [db_shop].[dbo].tbl_shop_chargeuser  "

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("chargeuser")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("chargeuser")&"' "&tmp_str&">"&rsget("chargeuser")&"/"&rsget("chargename")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end sub

Sub drawSelectBoxOffShopChargeId2(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
     <option value='10x10' <%if selectedId="10x10" then response.write " selected"%>>10x10</option>
   <%
   query1 = " select userid,socname from [db_user].[dbo].tbl_user_c  where userdiv<19"
   ''query1 = query1 & " and isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&db2html(rsget("socname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end sub

Sub drawSelectBoxDesignerwithName(selectBoxName,selectedId)
   dim tmp_str,query1

   NewDrawSelectBoxDesignerwithName selectBoxName,selectedId
   exit sub

   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,socname_kor from [db_user].[dbo].tbl_user_c  where userdiv<19"
   ''query1 = query1 & " where a.userid = b.userid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub NewDrawSelectBoxDesignerwithName(selectBoxName,selectedId)
    dim strRst

    strRst = "<input type=""text"" class=""text"" name=""" & selectBoxName & """ id=""[off,off,off,off][브랜드ID]"" value=""" & selectedId & """ size=""20"" >" & vbCrLf
    strRst = strRst & "<input type=""button"" class=""button"" value=""IDSearch"" onclick=""jsSearchBrandID(this.form.name,'" & selectBoxName & "');"" >"

	Response.Write strRst
End Sub

Sub NewDrawSelectBoxDesignerwithNameEvent(selectBoxName,selectedId)
    dim strRst

    strRst = "<input type=""text"" class=""text"" name=""" & selectBoxName & """ id=""[off,off,off,off][브랜드ID]"" value=""" & selectedId & """ size=""20"" >" & vbCrLf
    strRst = strRst & "<input type=""button"" class=""button"" value=""IDSearch"" onclick=""jsSearchBrandIDNew(this.form.name,'" & selectBoxName & "');"" >"

	Response.Write strRst
End Sub

Sub NewDrawSelectBoxDesignerwithNameFingers(selectBoxName,selectedId)
    dim strRst

    strRst = "<input type=""text"" class=""text"" name=""" & selectBoxName & """ id=""[off,off,off,off][브랜드ID]"" value=""" & selectedId & """ size=""20"" >" & vbCrLf
    strRst = strRst & "<input type=""button"" class=""button"" value=""IDSearch"" onclick=""jsSearchBrandID2(this.form.name,'" & selectBoxName & "');"" >"

	Response.Write strRst
End Sub

Sub NewDrawSelectBoxDesignerwithNameAndUserDIV(selectBoxName, selectedId, userdiv)
    dim strRst

    strRst = "<input type=""text"" class=""text"" id=""" & selectBoxName & """name=""" & selectBoxName & """ value=""" & selectedId & """ size=""20"" >" & vbCrLf
    strRst = strRst & "<input type=""button"" class=""button"" value=""IDSearch"" onclick=""jsSearchBrandIDwithUserDIV(this.form.name,'" & selectBoxName & "', '" + CStr(userdiv) + "');"" >"

	Response.Write strRst
End Sub

Sub NewdrawSelectBoxShopAll(selectBoxName,selectedId)
    dim strRst

    strRst = "<input type=""text"" class=""text"" name=""" & selectBoxName & """ id=""[off,off,off,off][매출처ID]"" value=""" & selectedId & """ size=""20"" >" & vbCrLf
    strRst = strRst & "<input type=""button"" class=""button"" value=""IDSearch"" onclick=""jsSearchMeachulID(this.form.name,'" & selectBoxName & "');"" >"

	Response.Write strRst
End Sub

Sub drawSelectBoxDesignerwithNameWithChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onchange="ChangeBrand(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,socname_kor from [db_user].[dbo].tbl_user_c  where userdiv<19"
   ''query1 = query1 & " where a.userid = b.userid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


Sub drawSelectBoxDesignerwithName2(selectBoxName,selectedId,selectedname)
   dim tmp_str,query1

   NewDrawSelectBoxDesignerwithName2 selectBoxName,selectedId,selectedname
   exit sub

   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,socname_kor from [db_user].[dbo].tbl_user_c  where userdiv<19"
   ''query1 = query1 & " where a.userid = b.userid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub NewDrawSelectBoxDesignerwithName2(selectBoxName,selectedId,selectedname)

Dim brid, brname
    %>
    <input type="text" class="text" name="<%= selectBoxName %>" value="<%= selectedId %>" size="60" readonly>
    <input type="hidden" class="text" name="brandkor" value="<%= selectedname %>" size="60" >
    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID2(this.form.name,'<%= selectBoxName %>');" >
    &nbsp;
    <%
End Sub

Sub NewDrawSelectBoxDesignerChangeMargin(selectBoxName,selectedId,marginDataNm,jsChgFnName)
    %>
    <input type="text" class="text" name="<%= selectBoxName %>" value="<%= selectedId %>" size="20" id="[on,off,off,off][브랜드ID]" >
    <input type="hidden" name="<%= marginDataNm %>" value="">
    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandIDchgMargin(this.form.name,'<%= selectBoxName %>','<%=marginDataNm%>','<%=jsChgFnName%>');" >
    &nbsp;
    <%
End Sub

Sub drawSelectBoxDesignerOffShopContract(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%

  	query1 = "select distinct d.makerid, c.socname_kor from [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid"
	query1 = query1 + " order by d.makerid"
	rsget.Open query1,dbget,1


   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("makerid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("makerid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub DrawAuthBox(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
     <option value='9999' <%if selectedId="9999" then response.write " selected"%>>업체</option>
	 <option value='999' <%if selectedId="999" then response.write " selected"%>>제휴사</option>
	 <option value='9000' <%if selectedId="9000" then response.write " selected"%>>강사</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>관리자</option>
     <option value='7' <%if selectedId="7" then response.write " selected"%>>마스타</option>
     <option value='5' <%if selectedId="5" then response.write " selected"%>>LV4</option>
     <option value='4' <%if selectedId="4" then response.write " selected"%>>LV3</option>
     <option value='2' <%if selectedId="2" then response.write " selected"%>>LV2</option>
     <option value='1' <%if selectedId="1" then response.write " selected"%>>LV1</option>
     <option value='500' <%if selectedId="500" then response.write " selected"%>>매장공통</option>
     <option value='501' <%if selectedId="501" then response.write " selected"%>>직영매장</option>
	 <option value='502' <%if selectedId="502" then response.write " selected"%>>수수료매장</option>
	 <option value='503' <%if selectedId="503" then response.write " selected"%>>대리점</option>
	 <option value='101' <%if selectedId="101" then response.write " selected"%>>오프샵</option>
	 <option value='111' <%if selectedId="111" then response.write " selected"%>>오프샵점장</option>
	 <option value='112' <%if selectedId="112" then response.write " selected"%>>오프샵부점장</option>
	 <option value='509' <%if selectedId="509" then response.write " selected"%>>오프매출조회</option>
	 <option value='201' <%if selectedId="201" then response.write " selected"%>>Zoom</option>
	 <option value='301' <%if selectedId="301" then response.write " selected"%>>College</option>
   </select>
   <%
end sub

Sub DrawAuthBoxTenUser(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
     <option value='1' <%if selectedId="1" then response.write " selected"%>>LV1</option>
     <option value='2' <%if selectedId="2" then response.write " selected"%>>LV2</option>
     <option value='4' <%if selectedId="4" then response.write " selected"%>>LV3</option>
     <option value='5' <%if selectedId="5" then response.write " selected"%>>LV4</option>
     <option value='7' <%if selectedId="7" then response.write " selected"%>>마스타</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>관리자</option>

	 <option value='101' <%if selectedId="101" then response.write " selected"%>>매장직원</option>
	 <option value='111' <%if selectedId="111" then response.write " selected"%>>매장점장</option>

   </select>
   <%
end sub

Sub DrawAuthBoxSimple(selectBoxName,selectedId,viewtype)
   %>
   <select class="select" name="<%=selectBoxName%>">
     <option value='' >전체</option>
     <option value='T' <%if selectedId="T" then response.write " selected"%>>직원전체</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>관리자</option>
     <option value='7' <%if selectedId="7" then response.write " selected"%>>마스타</option>
     <option value='L' <%if selectedId="L" then response.write " selected"%>>직원일반</option>

     <option value='111' <%if selectedId="111" then response.write " selected"%>>매장점장</option>
	 <option value='101' <%if selectedId="101" then response.write " selected"%>>매장직원</option>

     <% if viewtype="TT" then %>
     <option value='112' <%if selectedId="112" then response.write " selected"%>>오프샵부점장</option>
     <option value='500' <%if selectedId="500" then response.write " selected"%>>매장공통</option>
     <option value='501' <%if selectedId="501" then response.write " selected"%>>직영매장</option>
	 <option value='502' <%if selectedId="502" then response.write " selected"%>>수수료매장</option>
	 <option value='503' <%if selectedId="503" then response.write " selected"%>>대리점</option>

	 <option value='509' <%if selectedId="509" then response.write " selected"%>>오프매출조회</option>
	 <option value='201' <%if selectedId="201" then response.write " selected"%>>Zoom</option>
	 <option value='301' <%if selectedId="301" then response.write " selected"%>>College</option>

	 <option value='9000' <%if selectedId="9000" then response.write " selected"%>>강사</option>
	 <% end if %>
   </select>
   <%
end sub

Sub drawReportSelectWithEvent(selectBoxName,selectedId,targetfrm)
   %>
   <select class="select" name="<%=selectBoxName%>" onChange="SelReport(this,<%=targetfrm%>);">
     <option value='year' <%if selectedId="year" then response.write " selected"%>>year</option>
	 <option value='month' <%if selectedId="month" then response.write " selected"%>>month</option>
     <option value='day' <%if selectedId="day" then response.write " selected"%>>day</option>
   </select>
   <%
end Sub

Sub drawBeadalDiv(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>" >
   	 <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
     <option value='1' <%if selectedId="1" then response.write " selected"%>>텐바이텐배송</option>
	 <option value='2' <%if selectedId="2" OR  selectedId="5" then response.write " selected"%>>업체무료배송</option>
     <option value='4' <%if selectedId="4" then response.write " selected"%>>텐바이텐무료배송</option>
     <!--<option value='5' <%if selectedId="5" then response.write " selected"%>>업체무료배송</option>-->
     <!--<option value='6' <%if selectedId="6" then response.write " selected"%>>현장수령</option>-->
     <option value='7' <%if selectedId="7" then response.write " selected"%>>업체착불배송</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>업체조건배송</option>
   </select>
   <%
end Sub

' 배송구분  ' 2019.10.30 한용민
function getBeadalDivname(BeadalDiv)
    dim BeadalDivname

    if BeadalDiv="1" then
        BeadalDivname="텐바이텐배송"
    elseif BeadalDiv="2" or BeadalDiv="5" then
        BeadalDivname="업체무료배송"
    elseif BeadalDiv="4" then
        BeadalDivname="텐바이텐무료배송"
    elseif BeadalDiv="5" then
        BeadalDivname="업체무료배송"
    elseif BeadalDiv="6" then
        BeadalDivname="현장수령"
    elseif BeadalDiv="7" then
        BeadalDivname="업체착불배송"
    elseif BeadalDiv="9" then
        BeadalDivname="업체조건배송"
    else
        BeadalDivname=""
    end if

    getBeadalDivname=BeadalDivname
end function

' 배송방법		' 2018.06.01 한용민 생성
Sub drawdeliverfixday(selectBoxName, selectedId, changeFlag)
%>
<select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
	<option value='DEFAULT' <%if selectedId="DEFAULT" then response.write " selected"%>>택배(일반)</option>
	<option value='X' <%if selectedId="X" then response.write " selected"%>>화물</option>
	<option value='C' <%if selectedId="C" then response.write " selected"%>>플라워지정일</option>
	<option value='G' <%if selectedId="G" then response.write " selected"%>>해외직구</option>
	<option value='L' <%if selectedId="L" then response.write " selected"%>>클래스</option>
</select>
<%
end Sub

Sub drawSelectBoxWriter(byval writer)
	dim buf
	buf = "<select class='select' name='writer'>" + VbCrlf
	buf = buf + "<option selected value=''>선택</option>" + VbCrlf

	if writer="winnie" then
		buf = buf + "<option value='winnie' selected>최은희</option>" + VbCrlf
	else
		buf = buf + "<option value='winnie' >최은희</option>" + VbCrlf
	end if

	if writer="moon" then
    	buf = buf + "<option value='moon' selected>이문재</option>" + VbCrlf
    else
    	buf = buf + "<option value='moon' >이문재</option>" + VbCrlf
	end if

    buf = buf + "</select>"

    response.write buf
end sub

Sub DrawDateBox(byval yyyy1,yyyy2,mm1,mm2,dd1,dd2)
	dim buf,i

    dim today_year,today_month,monstart,MonFirstDay,lastdaytemp,result,MonLastDay

today_year = request("Year")   '이번 년
	if today_year = "" then today_year = year(date) end if
today_month = request("Month")    '이번 달
	if today_month = "" then today_month = month(date) end if
monstart=DateSerial(today_year, today_month, 1)
MonFirstDay = day(monstart)

		for lastdaytemp = 28 to 31
			result = DateSerial(today_year, today_month, lastdaytemp)
			if int(today_month) = month(result) then
               MonLastDay = lastdaytemp  '이번 달의 마지막 날..
			end if
		next


	buf = "<select class='select' name='yyyy1'>"
    for i=2001 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm1'>"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='dd1' >"

    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd1)) then
	    buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    buf = buf + "~"

    buf = buf + "<select class='select' name='yyyy2'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy2)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm2'>"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='dd2' >"
    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    response.write buf
end Sub

'//날짜함수 박스네임 동적생성	'/2012.02.07 한용민 추가
Sub DrawDateBoxdynamic(byval yyyy1,yyyy1name,yyyy2,yyyy2name,mm1,mm1name,mm2,mm2name,dd1,dd1name,dd2,dd2name)
	dim buf,i

    dim today_year,today_month,monstart,MonFirstDay,lastdaytemp,result,MonLastDay

	today_year = request("Year")   '이번 년
	if today_year = "" then today_year = year(date) end if
	today_month = request("Month")    '이번 달
	if today_month = "" then today_month = month(date) end if
	monstart=DateSerial(today_year, today_month, 1)
	MonFirstDay = day(monstart)

	for lastdaytemp = 28 to 31
		result = DateSerial(today_year, today_month, lastdaytemp)
		if int(today_month) = month(result) then
           MonLastDay = lastdaytemp  '이번 달의 마지막 날..
		end if
	next

	buf = "<select class='select' name='"&yyyy1name&"'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='"&mm1name&"'>"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='"&dd1name&"'>"

    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd1)) then
	    buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    buf = buf + "~"

    buf = buf + "<select class='select' name='"&yyyy2name&"'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy2)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='"&mm2name&"'>"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='"&dd2name&"'>"
    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    response.write buf
end Sub

'//년도 날짜함수 박스네임 동적생성	'/2012.03.06 한용민 추가
Sub DrawyearBoxdynamic(yyyy1name,yyyy1,chplg)
	dim buf,i

	buf = "<select class='select' name='"&yyyy1name&"' "&chplg&">"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    response.write buf
end Sub

'//년월일 날짜함수 박스네임 동적생성	'/2012.05.23 한용민 추가
Sub DrawOneDateBoxdynamic(yyyy1name, yyyy1, mm1name, mm1, dd1name, dd1, chplg, yyyy1id, mm1id, dd1id)
	dim buf,i

	buf = "<select class='select' name='"&yyyy1name&"' id='"&yyyy1id&"' "&chplg&">"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2002 to Year(now)+1
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='"&mm1name&"' id='"&mm1id&"' "&chplg&">"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='"&dd1name&"' id='"&dd1id&"' "&chplg&">"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub

'//년월 날짜함수 박스네임 동적생성	'/2012.11.08 한용민 추가
Sub DrawYMBoxdynamic(yyyy1name, yyyy1, mm1name, mm1, chplg)
	dim buf,i

	buf = "<select name='"&yyyy1name&"' "&chplg&">"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select name='"&mm1name&"' "&chplg&">"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub

'//시간함수 박스네임 동적생성	'/2019.05.30 한용민 추가
function DrawTimeBoxdynamic(time1name, time1, time2name, time2, chplg, time1id, time2id, timeallyn)
	dim buf,i, ti, mintime
    mintime=0

    if timeallyn="Y" then
        mintime=0
    else
        mintime=7
    end if

    buf="<select name='"& time1name &"' id='"& time1id &"' "& chplg &" >"
    buf = buf & "    <option value='' "& chkiif(Cstr(time1)= "" ,"selected","") &" >선택</option>"
    ti=0
    For ti = mintime To 23
        buf = buf & "<option value='"& Format00(2,ti) &"' "& chkiif(Cstr(Format00(2,time1)) = Cstr(Format00(2,ti)),"selected","") &">"& chkiif(ti<12,"오전","오후") &"&nbsp;"
        If ti <= 12 Then
            buf = buf & Format00(2,ti)
        ElseIf ti > 12 Then
            buf = buf & Format00(2,ti-12)
        End if
        buf = buf & "</option>"
    Next
    buf = buf & "</select>시&nbsp;"
    buf = buf & "<select name='"& time2name &"' id='"& time2id &"' "& chplg &" >"
    buf = buf & "    <option value='' "& chkiif(Cstr(Format00(2,time2))= "" ,"selected","") &" >선택</option>"
    buf = buf & "    <option value='00' "& chkiif(Cstr(Format00(2,time2))= "00" ,"selected","") &" >00</option>"
    buf = buf & "    <option value='05' "& chkiif(Cstr(Format00(2,time2))= "05" ,"selected","") &" >05</option>"
    buf = buf & "    <option value='10' "& chkiif(Cstr(Format00(2,time2))= "10" ,"selected","") &" >10</option>"
    buf = buf & "    <option value='15' "& chkiif(Cstr(Format00(2,time2))= "15" ,"selected","") &" >15</option>"
    buf = buf & "    <option value='20' "& chkiif(Cstr(Format00(2,time2))= "20" ,"selected","") &" >20</option>"
    buf = buf & "    <option value='25' "& chkiif(Cstr(Format00(2,time2))= "25" ,"selected","") &" >25</option>"
    buf = buf & "    <option value='30' "& chkiif(Cstr(Format00(2,time2))= "30" ,"selected","") &" >30</option>"
    buf = buf & "    <option value='35' "& chkiif(Cstr(Format00(2,time2))= "35" ,"selected","") &" >35</option>"
    buf = buf & "    <option value='40' "& chkiif(Cstr(Format00(2,time2))= "40" ,"selected","") &" >40</option>"
    buf = buf & "    <option value='45' "& chkiif(Cstr(Format00(2,time2))= "45" ,"selected","") &" >45</option>"
    buf = buf & "    <option value='50' "& chkiif(Cstr(Format00(2,time2))= "50" ,"selected","") &" >50</option>"
    buf = buf & "    <option value='55' "& chkiif(Cstr(Format00(2,time2))= "55" ,"selected","") &" >55</option>"
    buf = buf & "</select>분"

    response.write buf
end function

Sub DrawOneDateBox(byval yyyy1,mm1,dd1)
	dim buf,i

	buf = "<select class='select' name='yyyy1'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2002 to Year(now)+1
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm1' >"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='dd1'>"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawOneDateBox_2012(byval yyyy1,mm1,dd1)
	dim buf,i

	buf = "<select class='select' name='yyyy1'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2002 to Year(now)+1
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm1' >"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='dd1'>"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub



Sub DrawOneDateBox2(byval yyyy2,mm2,dd2)
	dim buf,i

	buf = "<select class='select' name='yyyy2'>"
    buf = buf + "<option value='" + CStr(yyyy2) +"' selected>" + CStr(yyyy2) + "</option>"
    for i=2002 to Year(now)+1
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm2' >"
    buf = buf + "<option value='" + CStr(mm2) + "' selected>" + CStr(mm2) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select class='select' name='dd2'>"
    buf = buf + "<option value='" + CStr(dd2) +"' selected>" + CStr(dd2) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawYMYMBox(byval yyyy1,mm1,yyyy2,mm2)
	dim buf,i

	buf = "<select class='select' name='yyyy1'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm1' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "~"

	buf = buf + "<select class='select' name='yyyy2'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy2)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm2' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawYMBox(byval yyyy1,mm1)
	dim buf,i

	buf = "<select class='select' name='yyyy1'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm1' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawYMBoxIdx(byval yyyy1,mm1,idx)
	dim buf,i

	buf = "<select class='select' name='yyyy"+cStr(idx)+"'>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='mm"+cStr(idx)+"' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawYMSelBox(byval fy,fm,yyyy,mm)
	dim buf,i

	buf = "<select class='select' name='" & fy & "' id='" & fy & "'>"
    buf = buf +"<option value=''>--</option>"
    for i=2002 to Year(now)+1
		if (CStr(i)=CStr(yyyy)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='select' name='" & fm & "' id='" & fm & "'>"
    buf = buf +"<option value=''>--</option>"
    for i=1 to 12
		if CStr(i)=CStr(mm) then
			buf = buf + "<option value='" & i & "' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" & i & "' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawMBox(byval mm1)
	dim buf,i

    buf = "<select class='select' name='mm1' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub

Sub drawSelectBoxCoWorker(byval selectBoxName, selectedId)
    dim tmp_str,query1
%>
    <select class='select' name="<%=selectBoxName%>">
    <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
    query1 = " select userid, username"
    query1 = query1 & " from [db_partner].[dbo].tbl_user_tenbyten with (nolock)"
    query1 = query1 & " where isusing= 1" & vbcrlf

    ' 퇴사예정자 처리	' 2018.10.16 한용민
    query1 = query1 & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
    query1 = query1 & "	and part_sn in('11','13','14','15','16') and userid <> ''" & vbcrlf
    query1 = query1 + " order by username asc"

    ' response.write query1 & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

    if  not rsget.EOF  then
        rsget.Movefirst

        do until rsget.EOF
            if Lcase(selectedId) = Lcase(rsget("userid")) then
                tmp_str = " selected"
            end if
            response.write("<option value='" + rsget("userid") + "' "&tmp_str&">" + db2html(rsget("username")) + " (" + rsget("userid") + ")</option>")
            tmp_str = ""
            rsget.MoveNext
        loop
    end if
    rsget.close
    response.write("</select>")
end Sub

Sub drawSelectBoxCoWorker_OnOff(byval selectBoxName, selectedId, onoff)
   dim tmp_str,query1, vOnOff

	If onoff = "on" Then
		vOnOff = "'15','10','11','17','21','38'" '온라인, 마케팅, 오프라인
		''권한으로 변경
		vOnOff = "'11','21','14','22','13','16'"  ''16 추가 2016/05/16

	elseif onoff ="fingers" then	'기업문화팀
		vOnOff = "'5'"
		''권한으로 변경
		vOnOff = "'16'"
	Else
		vOnOff = "'24'"
	End If
   %>
   <select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
   query1 = " select userid, username from"
   query1 = query1 + " [db_partner].[dbo].tbl_user_tenbyten "
   ''query1 = query1 + " where  isusing= 1 and statediv = 'Y' and department_id in(" + vOnOff + ") and userid <> '' "
   query1 = query1 & " where  isusing= 1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	query1 = query1 & "	and part_sn in(" + vOnOff + ") and userid <> ''" & vbcrlf
   query1 = query1 + " order by username asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" + rsget("userid") + "' "&tmp_str&">" + db2html(rsget("username")) + " (" + rsget("userid") + ")</option>")
           tmp_str = ""
           rsget.MoveNext
       Loop
       if onoff ="fingers" Then
			response.write("<option value='jjin1655'>김미진A (jjin1655)</option>")
       End If
   end if
   rsget.close
   response.write("</select>")
end Sub

Sub drawSelectBoxCoWorker_OnOffUserdiv(byval selectBoxName, selectedId, onoff, pcuserdiv)
   dim tmp_str,query1, vOnOff

	If onoff = "on" Then
		vOnOff = "'15','10','11','17','21','38'" '온라인, 마케팅, 오프라인
	elseif onoff ="fingers" then	'기업문화팀
		vOnOff = "'5'"
	elseif onoff="sell" Then    ''매출처
		if pcuserdiv="501_21" then
			vOnOff = "'15','21','38'"  '마케팅 / 오프라인팀
		elseif	pcuserdiv="503_21" then '오프라인팀 / 전략상품팀(패션/수입파트)
			vOnOff = "'17','21','38'"
		elseif  pcuserdiv="900_21"   then	'물류
			vOnOff ="'30'"
		elseif  pcuserdiv="999_50"  then '마케팅 / 전략상품팀(패션/수입파트)
			vOnOff ="'17','15'"
		end if

	Else
		vOnOff = "'21'"
	End If
   %>
   <select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
   query1 = " select userid, username from"
   query1 = query1 + " [db_partner].[dbo].tbl_user_tenbyten "
    query1 = query1 & " where isusing= 1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	query1 = query1 & "	and department_id in(" + vOnOff + ") and userid <> ''" & vbcrlf
   query1 = query1 + " order by username asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" + rsget("userid") + "' "&tmp_str&">" + db2html(rsget("username")) + " (" + rsget("userid") + ")</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end Sub

Sub drawRadioIpChulDiv(byval RadioName, selectedId)
	dim buf
	buf = "<input type=radio name=" + RadioName + " value=radiobutton value=10 "
	if selectedId="10" then buf = buf + "selected"
	buf = buf + " >매입"

    buf = buf + "<input type=radio name=" + RadioName + " value=radiobutton value=20 "
    if selectedId="20" then buf = buf + "selected"
    buf = buf + " >위탁"
    response.write buf
end Sub

Function GetValidDate(byval dt)
	'// 2014-06-31 -> 2014-06-30
	dim tmpStr, tmpDate
	tmpStr = Split(CStr(dt), "-")
	if UBound(tmpStr) <> 2 then
		GetValidDate = dt
		exit Function
	end if

	if CLng(tmpStr(2)) <= 28 then
		GetValidDate = dt
		exit Function
	end if

	do while (Month(DateSerial(tmpStr(0), tmpStr(1), tmpStr(2))) <> CLng(tmpStr(1)))
		tmpStr(2) = tmpStr(2) - 1
	loop

	tmpDate = DateSerial(tmpStr(0), tmpStr(1), tmpStr(2))
	GetValidDate = Left(tmpDate, 10)
end Function

function BaesongCd2Name(byval icd)
	if (icd="1") then
		BaesongCd2Name = "10x10"
	elseif (icd="2") then
		BaesongCd2Name = "업체"
	elseif (icd="3") then
		BaesongCd2Name = "직접수령"
	elseif (icd="4") then
		BaesongCd2Name = "10x10무료"
	elseif (icd="5") then
		BaesongCd2Name = "업체무료"
	end if

end function


Sub DrawSelectBoxDesignerTenBeadalItem(byval selectBoxName,Designer,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select itemid, itemname from [db_item].[dbo].tbl_item where makerid = '" & Designer & "' and isusing='Y' and deliverytype in ('1','3','4') order by itemid Desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("itemid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("itemid")&"' "&tmp_str&">"&rsget("itemid")&"-"&db2html(rsget("itemname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxDesignerItem(byval selectBoxName,Designer,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select itemid, itemname from [db_item].[dbo].tbl_item where makerid = '" & Designer & "' and isusing='Y' order by itemid Desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("itemid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("itemid")&"' "&tmp_str&">"&rsget("itemid")&"-"&db2html(rsget("itemname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub


Sub DrawSelectBoxStyleMid(byval selectBoxName,stylegubun,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>"  >
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select stylegubun, itemstyle, stylename from [db_item].[dbo].tbl_item_stylegubun "
   query1 = query1 + " where stylegubun='" + stylegubun + "'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("stylegubun")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("itemstyle")&"' "&tmp_str&">"&rsget("stylename")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxCategoryLargeBychannel(byval selectBoxName,selectedId,channel)
   dim tmp_str,query1
   dim categoryArr
   if (channel="02") then
   	 categoryArr = "'6'"
   elseif (channel="04") then
   	 categoryArr = "'5'"
   elseif (channel="05") then
   	 categoryArr = "'8'"
   elseif (channel="06") then
   	 categoryArr = "'7'"
   else
     categoryArr = "'1','2','3','4'"
   end if
   %><select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " and Left(code_large,1) in (" + categoryArr + ")"
   query1 = query1 + " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxDispCateLarge(byval selectBoxName,selectedId,addparam)
    dim query1
   %><select class='select' name="<%=selectBoxName%>" "&addparam&">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
    query1 = " select catecode,catename from db_item.dbo.tbl_display_cate"
    query1 = query1&" where depth=1"
    query1 = query1&" and useyn='Y'"
    query1 = query1&" order by sortno, catecode"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write("<option value='"&rsget("catecode")&"' "&CHKIIF(Cstr(selectedId) = Cstr(rsget("catecode")),"selected","")&">"& db2html(rsget("catename")) &"</option>")
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub



Sub DrawSelectBoxCategoryLarge(byval selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxCategoryMid(byval selectBoxName,largeno,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_mid, code_nm from [db_item].[dbo].tbl_Cate_mid"
   query1 = query1 & " where display_yn = 'Y'"
   query1 = query1 & " and code_large = '" & largeno & "'"
   query1 = query1 & " and code_mid<>0"
   query1 = query1 & " order by code_mid Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Not(isNull(selectedId)) then
	           if Cstr(selectedId) = Cstr(rsget("code_mid")) then
	               tmp_str = " selected"
	           end if
	       end if
           response.write("<option value='"&rsget("code_mid")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxCategoryMidNotOnchange(byval selectBoxName,largeno,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_mid, code_nm from [db_item].[dbo].tbl_Cate_mid"
   query1 = query1 & " where display_yn = 'Y'"
   query1 = query1 & " and code_large = '" & largeno & "'"
   query1 = query1 & " and code_mid<>0"
   query1 = query1 & " order by code_mid Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Not(isNull(selectedId)) then
	           if Cstr(selectedId) = Cstr(rsget("code_mid")) then
	               tmp_str = " selected"
	           end if
	       end if
           response.write("<option value='"&rsget("code_mid")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxCategorySmall(byval selectBoxName,largeno,midno,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_small, code_nm from [db_item].[dbo].tbl_cate_small"
   query1 = query1 & " where display_yn = 'Y'"
   query1 = query1 & " and code_large = '" & largeno & "'"
   query1 = query1 & " and code_mid = '" & midno & "'"
   query1 = query1 & " and code_small<>0"
   query1 = query1 & " order by code_small Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_small")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_small")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxCategoryOnlyLarge(byval selectBoxName,ByVal selectedId, ByVal strScript)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>" <%=strScript%>>
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub
Sub DrawSelectBoxItemoptionBig(byval selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%> >선택</option>
   <%
   query1 = " select optioncode01, codename from [db_item].[dbo].tbl_option_div01"
   query1 = query1 & " where optiondispyn='Y'"
   query1 = query1 & " order by disporder Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("optioncode01")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("optioncode01")&"' "&tmp_str&">"&rsget("codename")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub DrawSelectBoxItemoptionSmall(byval optionbig)
   dim tmp_str,query1

   query1 = " select optioncode02, codeview from [db_item].[dbo].tbl_option_div02"
   query1 = query1 & " where optioncode01='" & Cstr(optionbig) & "'"
   query1 = query1 & " and optiondispyn='Y'"
   query1 = query1 & " order by disporder Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write "<input type='checkbox' name='optioncode02' view='"& rsget("codeview") &"'  value='"&rsget("optioncode02")&"'>"&rsget("codeview")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close

end Sub

Sub DrawSelectStyleSiteName(byval selectBoxName,isitename)
	dim sites,i
	sites = Array("events10x10","uto","onlytenbyten","cara", _
				"emoden","dream","new10x10","category", _
				"mail1004","nanistyle","nanicollection","b2bpromotion", _
				"skdtod","miclub","preorder","petzone","tensenddream", _
				"ugi","flower","fashion","pet","beauty","boardgame")

	response.write "<select class='select' name=" + selectBoxName + ">"
	for i=0 to UBound(sites)
		if isitename=sites(i) then
			response.write("<option value=" + sites(i) + " selected>" + sites(i) + "</option>")
		else
			response.write("<option value=" + sites(i) + ">" + sites(i) + "</option>")
		end if
	next
	response.write("</select>")
end Sub

Sub DrawSelectExtSiteName(byval selectBoxName,extsitename)
	dim sqlStr
	dim styleStr

	sqlStr = " select top 150 (case when id in ('11st1010','auction1010','cjmall','ezwel','gmarket1010','gseshop','hmall1010','interpark','kakaogift','kakaostore','lfmall','lotteimall','lotteon','Mylittlewhoopee','nvstorefarm','nvstoregift','shintvshopping','skstoa','ssg','wconcept1010','WMP') then 1 else 100 end) as dispno, id from [db_partner].[dbo].tbl_partner "
	sqlStr = sqlStr + " where userdiv='999' "
	sqlStr = sqlStr + " and id<>'10x10' "
	sqlStr = sqlStr + " and isusing='Y' "
	sqlStr = sqlStr + " order by (case when id in ('11st1010','auction1010','cjmall','ezwel','gmarket1010','gseshop','hmall1010','interpark','kakaogift','kakaostore','lfmall','lotteimall','lotteon','Mylittlewhoopee','nvstorefarm','nvstoregift','shintvshopping','skstoa','ssg','wconcept1010','WMP') then 1 else 100 end), id "
	rsget.Open sqlStr,dbget,1

	response.write "<select class='select' name=" + selectBoxName + ">"
	if  not rsget.EOF  then
	    rsget.Movefirst
        do until rsget.EOF
			if (rsget("dispno") = 1) then
				styleStr = " style='color:green;' "
			else
				styleStr = ""
			end if

        	if extsitename=rsget("id") then
        		response.write("<option value=" + rsget("id") + " " + styleStr + " selected>" + rsget("id") + "</option>")
        	else
        		response.write("<option value=" + rsget("id") + " " + styleStr + " >" + rsget("id") + "</option>")
        	end if
            rsget.MoveNext
        loop
    end if
    rsget.close
    response.write("</select>")
end Sub

Sub DrawItemGubunRadio(byval selectBoxName,gubuncd)
   dim query1

   query1 = " select gubuncd,gubunname from [db_item].[dbo].tbl_item_gubun order by gubuncd"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write "<input type='radio' name='" + selectBoxName + "' value='"&rsget("gubuncd")&"'>"&rsget("gubunname") & "&nbsp;"
           rsget.MoveNext
       loop
   end if
   rsget.close
end sub

Sub DrawItemGubunRadio2(byval selectBoxName,gubuncd)
   dim query1,tmp_str
   dim itemgubun
   query1 = " select gubuncd,gubunname from [db_item].[dbo].tbl_item_gubun order by gubuncd"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
       		itemgubun = rsget("gubuncd")
	   		if IsNull(itemgubun) then itemgubun=""
	   		if IsNull(gubuncd) then gubuncd=""

       		if Cstr(gubuncd) = Cstr(itemgubun) then
               tmp_str = " checked"
           end if
           response.write "<input type='radio' name='" + selectBoxName + "' value='"&rsget("gubuncd")&"' onclick='dispSubCate(this.value);' "&tmp_str&">"&rsget("gubunname") & "&nbsp;"
           rsget.MoveNext
           tmp_str = ""
       loop
   end if
   rsget.close
end sub

Sub SelectBoxCategoryLarge(byval selectedId)
   dim tmp_str,query1
   %><select class='select' name="cdl">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   ''query1 = query1 + " and code_large<90"
   query1 = query1 + " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"&rsget("code_nm")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub SelectBoxBrandCategory(byval selectname, byval selectedId)
   dim tmp_str,query1

   if IsNULL(selectedId) then selectedId=""

   %><select class='select' name="<%= selectname %>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"&rsget("code_nm")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

Sub SelectBoxMallDiv(byval selectname, byval selectedId)
	dim tmp_str,query1
   %><select class='select' name="<%= selectname %>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select gubuncd,gubunname from [db_item].[dbo].tbl_item_gubun "
   query1 = query1 + " where gubuncd<90"
   query1 = query1 + " order by gubuncd Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("gubuncd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("gubuncd")&"' "&tmp_str&">"&rsget("gubunname")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end Sub

Sub SelectBoxUserDiv01(byval selectname, byval selectedId)
	dim tmp_str,query1
   %><select class='select' name="<%= selectname %>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select divcode,divename from [db_user].[dbo].tbl_user_div "
   query1 = query1 + " where divcode>1"
   query1 = query1 + " and divcode<19"
   query1 = query1 + " order by divcode Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("divcode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcode")&"' "&tmp_str&">"&rsget("divename")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end Sub

Sub SelectBoxCooperationName(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select id,company_name from [db_partner].[dbo].tbl_partner  where userdiv=999"
   ''query1 = query1 & " where a.userid = b.userid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("id")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("id")&"' "&tmp_str&">"&db2html(rsget("company_name"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxCaMall()
   dim tmp_str,query1
   %><select class='select' name="code_large">
     <option value=''>선택</option><%
   query1 = " select code_large,lname from [db_contents].[dbo].tbl_camall_code_large"
   query1 = query1 & " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write("<option value='"&rsget("code_large")&"'>"&db2html(rsget("lname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxLCaMall(byval selectedId)
   dim tmp_str,query1
   %><select class='select' name="code_large" onchange="GotoMid();">
     <option value=''>선택</option><%
   query1 = " select code_large,lname from [db_contents].[dbo].tbl_camall_code_large"
   query1 = query1 & " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"&db2html(rsget("lname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxMCaMall(byval code_large,selectedId)
   dim tmp_str,query1
   %><select class='select' name="code_mid" onchange="GotoMid();">
     <option value=''>선택</option><%
   query1 = " select code_mid,mname from [db_contents].[dbo].tbl_camall_code_mid"
   query1 = query1 & " where code_large='" + Cstr(code_large) + "'"
   query1 = query1 & " order by code_mid Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("code_mid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_mid")&"' "&tmp_str&">"&db2html(rsget("mname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


Sub SelectBoxCaMall2(byval selectedId)
   dim tmp_str,query1
   %><select class='select' name="code_large">
     <option value=''>선택</option><%
   query1 = " select code_large,lname from [db_contents].[dbo].tbl_camall_code_large"
   query1 = query1 & " order by code_large Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"&db2html(rsget("lname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxOffShopSuplyer(byval selectedname, selectedId,shopid, ibctdiv)
	dim restr, selectedstr, query1
	restr = "<select class='select' name=" + selectedname + ">"
	restr = restr + "<option value=''>선택</option>"

	''수정 10x10만 보이도록 설정.
	if (ibctdiv="503") or (ibctdiv="502") or (ibctdiv="501") or (ibctdiv="101") then
		if selectedId="10x10" then selectedstr = "selected"
		restr = restr + "<option value='10x10' " + selectedstr + ">10x10(텐바이텐)</option>"
	else
		if selectedId="10x10" then selectedstr = "selected"
		restr = restr + "<option value='10x10' " + selectedstr + ">10x10(텐바이텐)</option>"

		query1 = "select d.makerid, c.socname_kor from [db_shop].[dbo].tbl_shop_designer d"
		query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid"
		query1 = query1 + " where shopid='" + shopid + "'"
		query1 = query1 + " and chargediv in ('6','8')"
		query1 = query1 + " order by d.makerid"
		rsget.Open query1,dbget,1

	   if  not rsget.EOF  then
	       rsget.Movefirst

	       do until rsget.EOF
	           if Lcase(selectedId) = Lcase(rsget("makerid")) then
	               selectedstr = " selected"
	           end if
	           restr = restr + "<option value='" + rsget("makerid") + "' " + selectedstr + ">" + rsget("makerid") + "(" + db2html(rsget("socname_kor")) + ")</option>"
	           selectedstr = ""
	           rsget.MoveNext
	       loop
	   end if
	   rsget.close

	end if
	restr = restr + "</select>"

	response.write restr
End Sub

''//2010-06-07 한용민 수정
Sub drawSelectBoxShopjumunDesigner(selectBoxName,selectedId,shopid,suplyer)
   dim tmp_str,query1

   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
    query1 = " select d.makerid, c.socname_kor ,d.comm_cd from [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid"
	query1 = query1 + " where d.shopid='" + shopid + "'"
	if suplyer="10x10" then
		'query1 = query1 + " and d.chargediv in ('2','4','5')"
   	else
   		query1 = query1 + " and d.makerid='" + suplyer + "'"
   	end if

   	query1 = query1 + " order by d.makerid"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("makerid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("makerid")&"/"&db2html(rsget("socname_kor"))&"["&GetJungsanGubunName(rsget("comm_cd"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxOffjumunDesigner(selectBoxName,selectedId,shopid,suplyer)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
   <%
    query1 = " select d.makerid, c.socname_kor from [db_shop].[dbo].tbl_shop_designer d"
	query1 = query1 + " left join [db_user].[dbo].tbl_user_c c on d.makerid=c.userid"
	query1 = query1 + " where d.shopid='" + shopid + "'"
	if suplyer="10x10" then
		query1 = query1 + " and d.chargediv in ('2','4','5')"
   	else
   		query1 = query1 + " and d.makerid='" + suplyer + "'"
   	end if
   	query1 = query1 + " order by d.makerid"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("makerid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("makerid")&"' "&tmp_str&">"&rsget("makerid")&"/"&db2html(rsget("socname_kor"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxOffShopByDiv(selectBoxName,selectedId,idiv)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and shopdiv='" + idiv + "'"
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
end sub


Sub drawSelectBoxPartnerDesigner(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>ALL</option><%
   query1 = " select id,company_name from [db_partner].[dbo].tbl_partner "
   query1 = query1 + " where userdiv='9999'"
   query1 = query1 + " and isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("id")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("id")&"' "&tmp_str&">" & rsget("id") & "  [" & replace(db2html(rsget("company_name")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxDeliverCompany(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Trim(Lcase(selectedId)) = Trim(Lcase(rsget("divcd"))) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxDesignerOnlyWaitItem(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
    query1 = " select c.userid, c.socname_kor, count(w.itemid) as cnt"
	query1 = query1 + " from [db_user].[dbo].tbl_user_c c,"
	query1 = query1 + " [db_temp].[dbo].tbl_wait_item w"
	query1 = query1 + " where c.userid=w.makerid"
	query1 = query1 + " and w.currstate='1'"
	query1 = query1 + " group by c.userid, c.socname_kor"
	query1 = query1 + " order by cnt desc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
		   response.write("<option value='"&rsget("userid")&"' "&tmp_str&">" & rsget("userid") & " (" & db2html(rsget("socname_kor")) & ") - " & rsget("cnt") & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawSelectBoxDesignerOnlyWaitAndRejectItem(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
    query1 = " select c.userid, c.socname_kor, count(w.itemid) as cnt"
	query1 = query1 + " from [db_user].[dbo].tbl_user_c c,"
	query1 = query1 + " [db_temp].[dbo].tbl_wait_item w"
	query1 = query1 + " where c.userid=w.makerid"
	query1 = query1 + " and w.currstate in ('1','2')"
	query1 = query1 + " group by c.userid, c.socname_kor"
	query1 = query1 + " order by cnt desc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
		   response.write("<option value='"&rsget("userid")&"' "&tmp_str&">" & rsget("userid") & " (" & db2html(rsget("socname_kor")) & ") - " & rsget("cnt") & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub DrawSelectBoxFlowerUse(selectedId)
   dim tmp_str,query1
   %><select class='select' name="gubun" onchange="GoUse();">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select gubun,usename from [db_contents].[dbo].tbl_flower_use_category"
   query1 = query1 + " order by gubun"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("gubun")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("gubun")&"' "&tmp_str& ">" & rsget("usename") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


'' 매입처 (2)/ 출고처 (21)/ 강사 (14) : 나머지 사용금지.
'' 기존 (select divcode, divename from [db_user].[dbo].tbl_user_div)
Sub DrawBrandGubunCombo(selectedname, selectedId)
   dim tmp_str,query1
   %>
   <select class='select' name="<%= selectedname %>" >
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
  	 <option value='02' <% if (selectedId="02") or (selectedId="03") or (selectedId="04") or (selectedId="05") or (selectedId="06") or (selectedId="07") or (selectedId="08") or (selectedId="13") then response.write " selected"%>>매입처(일반)</option>
  	 <option value='14' <%if selectedId="14" then response.write " selected"%>>아카데미</option>
  	 <option value='21' <%if selectedId="21" then response.write " selected"%>>출고처</option>
  	 <option value='20' <%if selectedId="20" then response.write " selected"%>>가맹점매입처</option>
  	 <option value='50' <%if selectedId="50" then response.write " selected"%>>제휴사(온라인)</option>
  	 <option value='95' <%if selectedId="95" then response.write " selected"%>>사용안함</option>
   </select>
  <%
End Sub


Sub DrawChulgoDiv(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> >선택</option>
	<option value='1' <% if selectedId="1" then response.write " selected" %> >매입-&gt;매입</option>
	<option value='2' <% if selectedId="2" then response.write " selected" %> >위탁-&gt;매입</option>
	<option value='2' <% if selectedId="2" then response.write " selected" %> >위탁-&gt;위탁</option>
	</select>
<%
end Sub


Sub DrawMiChulgoDiv(selectedname,selectedId)
	dim varexists
	varexists = false
%>
	<select class='select' name="<%= selectedname %>" onchange="showSpecialInput(this)">
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='5일내출고' <% if selectedId="5일내출고" then response.write " selected" %> >5일내출고</option>
	<option value='재고부족' <% if selectedId="재고부족" then response.write " selected" %> >재고부족</option>
	<option value='일시품절' <% if selectedId="일시품절" then response.write " selected" %> >일시품절</option>
	<option value='단종' <% if selectedId="단종" then response.write " selected" %> >단종</option>
	<% if (selectedId<>"") and (Not varexists) then %>
	<option value="<%= selectedId %>" id=special selected ><%= selectedId %></option>
	<% else %>
	<option value='기타입력' id=special <% if selectedId="기타입력" then response.write " selected" %> >기타입력</option>
	<% end if %>
	</select>
<%
end Sub

Sub DrawBrandMWUCombo(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='M' <% if selectedId="M" then response.write " selected" %> >매입</option>
	<option value='W' <% if selectedId="W" then response.write " selected" %> >위탁</option>
	<option value='U' <% if selectedId="U" then response.write " selected" %> >업체배송</option>
	</select>
<%
end sub

'//매입+위탁 추가
Sub DrawBrandMWUPCombo(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> >-전체-</option>
	<option value='MW' <% if selectedId="MW" then response.write " selected" %> >매입+위탁</option>
	<option value='M' <% if selectedId="M" then response.write " selected" %> >매입</option>
	<option value='W' <% if selectedId="W" then response.write " selected" %> >위탁</option>
	<option value='U' <% if selectedId="U" then response.write " selected" %> >업체배송</option>
	</select>
<%
end sub

'//rdsite 추가 2017-07-03 김진영
Sub DrawRdsiteCombo(rdsitename, selectedId)
%>
	<select class='select' name="<%= rdsitename %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> >-전체-</option>
	<option value='nvshop' <% if selectedId="nvshop" then response.write " selected" %> >nvshop</option>
	<option value='coocha' <% if selectedId="coocha" then response.write " selected" %> >coocha</option>
	<option value='daumshop' <% if selectedId="daumshop" then response.write " selected" %> >daumshop</option>
	<option value='nateshop' <% if selectedId="nateshop" then response.write " selected" %> >nateshop</option>
	<option value='okcashbag' <% if selectedId="okcashbag" then response.write " selected" %> >okcashbag</option>
	<option value='gifticon' <% if selectedId="gifticon" then response.write " selected" %> >gifticon</option>
	<option value='between' <% if selectedId="between" then response.write " selected" %> >between</option>
	<option value='wmprc' <% if selectedId="wmprc" then response.write " selected" %> >wmprc</option>
    <option value='ggshop' <% if selectedId="ggshop" then response.write " selected" %> >ggshop</option>
	</select>
<%
end sub
'####### 20110831(강준구) 2.브랜드관련정보에서 계약관련사항에 브랜드기본마진에 있는 select 박스. 배송Type(DefaultDeliveryType)과 업체조건배송 을 바로 입력할 수 있도록 하기 위해 따로 사용.
Sub DrawBrandMWUCombo_2011(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" onchange="inputDeliveryType(this.value)">
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='M' <% if selectedId="M" then response.write " selected" %> >매입</option>
	<option value='W' <% if selectedId="W" then response.write " selected" %> >위탁</option>
	<option value='U' <% if selectedId="U" then response.write " selected" %> >업체배송</option>
	</select>
<%
end sub

Sub DrawBrandOffMWCombo(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='M' <% if selectedId="M" then response.write " selected" %> >매입</option>
	<option value='W' <% if selectedId="W" then response.write " selected" %> >위탁</option>
	</select>
<%
end sub

Sub DrawJungsanDateCombo(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='말일' <% if selectedId="말일" then response.write " selected" %> >말일</option>
	<option value='15일' <% if selectedId="15일" then response.write " selected" %> >15일</option>
	<option value='수시' <% if selectedId="수시" then response.write " selected" %> >수시</option>
	</select>
<%
end sub

Sub DrawBankCombo(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='경남' <% if selectedId="경남" then response.write " selected" %> >경남</option>
	<option value='광주' <% if selectedId="광주" then response.write " selected" %> >광주</option>
	<option value='국민' <% if selectedId="국민" then response.write " selected" %> >국민</option>
	<option value='기업' <% if selectedId="기업" then response.write " selected" %> >기업</option>
	<option value='농협' <% if selectedId="농협" then response.write " selected" %> >농협</option>
	<option value='단위농협' <% if selectedId="단위농협" then response.write " selected" %> >단위농협</option>
	<option value='대구' <% if selectedId="대구" then response.write " selected" %> >대구</option>
	<option value='도이치' <% if selectedId="도이치" then response.write " selected" %> >도이치</option>
	<option value='부산' <% if selectedId="부산" then response.write " selected" %> >부산</option>
	<option value='산업' <% if selectedId="산업" then response.write " selected" %> >산업</option>
	<option value='새마을금고' <% if selectedId="새마을금고" then response.write " selected" %> >새마을금고</option>
	<% if (selectedId="상호저축") then %>
	<option value='상호저축' <% if selectedId="상호저축" then response.write " selected" %> >상호저축</option>
	<% end if %>
	<% if (selectedId="수출입") then %>
	<option value='수출입' <% if selectedId="수출입" then response.write " selected" %> >수출입</option>
	<% end if %>
	<option value='수협' <% if selectedId="수협" then response.write " selected" %> >수협</option>
	<option value='신한' <% if selectedId="신한" then response.write " selected" %> >신한</option>
	<option value='KEB하나' <% if selectedId="KEB하나" then response.write " selected" %> >KEB하나</option>
	<option value='우리' <% if selectedId="우리" then response.write " selected" %> >우리</option>
	<option value='케이뱅크' <% if selectedId="케이뱅크" then response.write " selected" %> >케이뱅크</option>
	<option value='카카오뱅크' <% if selectedId="카카오뱅크" then response.write " selected" %> >카카오뱅크</option>
    <option value='토스뱅크' <% if selectedId="토스뱅크" then response.write " selected" %> >토스뱅크</option>
	<option value='우체국' <% if selectedId="우체국" then response.write " selected" %> >우체국</option>
	<option value='전북' <% if selectedId="전북" then response.write " selected" %> >전북</option>
	<option value='제일' <% if selectedId="제일" then response.write " selected" %> >제일</option>
	<% if (selectedId="조흥") then %>
	<option value='조흥' <% if selectedId="조흥" then response.write " selected" %> >조흥</option>
	<% end if %>
	<% if (selectedId="평화") then %>
	<option value='평화' <% if selectedId="평화" then response.write " selected" %> >평화</option>
	<% end if %>
	<% if (selectedId="하나") then %>
	<option value='하나' <% if selectedId="하나" then response.write " selected" %> >하나</option>
	<% end if %>
	<% if (selectedId="한국") then %>
	<option value='한국' <% if selectedId="한국" then response.write " selected" %> >한국</option>
	<% end if %>
	<% if (selectedId="한미") then %>
	<option value='한미' <% if selectedId="한미" then response.write " selected" %> >한미</option>
	<% end if %>
    <option value='한화투자증권' <% if selectedId="한화투자증권" then response.write " selected" %> >한화투자증권</option>
    <option value='미래에셋증권' <% if selectedId="미래에셋증권" then response.write " selected" %> >미래에셋증권</option>
	<option value='시티' <% if selectedId="시티" then response.write " selected" %> >시티</option>
	<option value='홍콩샹하이' <% if selectedId="홍콩샹하이" then response.write " selected" %> >홍콩샹하이</option>
	<option value='ABN암로은행' <% if selectedId="ABN암로은행" then response.write " selected" %> >ABN암로은행</option>
	<option value='UFJ은행' <% if selectedId="UFJ은행" then response.write " selected" %> >UFJ은행</option>
    <option value='Bank of America' <% if selectedId="Bank of America" then response.write " selected" %> >Bank of America</option>
	<option value='신협' <% if selectedId="신협" then response.write " selected" %> >신협</option>
    <option value='제주' <% if selectedId="제주" then response.write " selected" %> >제주</option>
    <option value='현대스위스상호저축은행' <% if selectedId="현대스위스상호저축은행" then response.write " selected" %> >현대스위스상호저축은행</option>
	<% if (selectedId="외환") then %>
    <option value='외환' <% if selectedId="외환" then response.write " selected" %> >외환</option>
	<% end if %>
    <option value='중국건설은행' <% if selectedId="중국건설은행" then response.write " selected" %> >중국건설은행</option>
	</select>
<%
end Sub

Sub SelectBoxQnaPrefaceGubun(byval masterid,selectedId)
   dim tmp_str,query1
   %><select class='select' name="gubun">
     <option value=''>선택</option><%
   query1 = " select G.code,G.cname from [db_cs].[dbo].tbl_qna_preface_gubun as G"
   query1 = query1 & "		Join [db_cs].[dbo].tbl_qna_preface as P on G.masterid=P.masterid and G.code=P.gubun "
   query1 = query1 & " where G.masterid='" + Cstr(masterid) + "' and P.isusing='Y'"
   query1 = query1 & " order by G.code Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("code")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code")&"' "&tmp_str&">"&db2html(rsget("cname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxQnaPrefaceAllGubun(byval masterid,selectedId)
   dim tmp_str,query1
   %><select class='select' name="gubun">
     <option value=''>선택</option><%
   query1 = " select G.code,G.cname from [db_cs].[dbo].tbl_qna_preface_gubun as G"
   query1 = query1 & " where G.masterid='" + Cstr(masterid) + "' "
   query1 = query1 & " order by G.code Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("code")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code")&"' "&tmp_str&">"&db2html(rsget("cname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxQnaComplimentGubun(byval selectedId)
   dim tmp_str,query1
   %><select class='select' name="gubun">
     <option value=''>선택</option><%
   query1 = " select code,cname from [db_cs].[dbo].tbl_qna_compliment_gubun"
   query1 = query1 & " order by code Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("code")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code")&"' "&tmp_str&">"&db2html(rsget("cname"))&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxQnaPreface(byval masterid)
   dim query1
   %><select class='select' name="preface" onchange="TnChangePreface(this.options[this.selectedIndex].value);">
     <option value=''>선택</option><%
   query1 = " select G.code,G.cname from [db_cs].[dbo].tbl_qna_preface_gubun as G"
   query1 = query1 & "		Join [db_cs].[dbo].tbl_qna_preface as P on G.masterid=P.masterid and G.code=P.gubun "
   query1 = query1 & " where P.isusing='Y'"

	If masterid <> "" then
		query1 = query1 & " and G.masterid='" + Cstr(masterid) + "'"
	End if

   query1 = query1 & " order by G.code Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write("<option value='"&rsget("code")&"'>"&db2html(rsget("cname"))&"</option>")
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub SelectBoxQnaCompliment(byval masterid)
   dim query1
   %><select class='select' name="gubun" onchange="TnChangeCompliment(this.options[this.selectedIndex].value);">
     <option value=''>선택</option><%
   query1 = " select code,cname from [db_cs].[dbo].tbl_qna_compliment_gubun"
	If masterid <> "" then
   query1 = query1 & " where masterid='" + Cstr(masterid) + "'"
	End if
   query1 = query1 & " order by code Asc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           response.write("<option value='"&rsget("code")&"'>"&db2html(rsget("cname"))&"</option>")
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


function fnGetUpcheDefaultSongjangDiv(byval imakerid)
    dim sqlStr

    if (imakerid="") then Exit function

    sqlStr = "select defaultsongjangdiv from [db_partner].[dbo].tbl_partner"
    sqlStr = sqlStr & " where id='" & imakerid & "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        fnGetUpcheDefaultSongjangDiv = rsget("defaultsongjangdiv")

        if IsNULL(fnGetUpcheDefaultSongjangDiv) then fnGetUpcheDefaultSongjangDiv=0
    end if
    rsget.close
end function

'//회원등급명
function getUserLevelStr(iuserlevel)
    getUserLevelStr = iuserlevel

    if (iuserlevel="0") then
        getUserLevelStr = "WHITE"
    elseif (iuserlevel="1") then
        getUserLevelStr = "RED"
    elseif (iuserlevel="2") then
        getUserLevelStr = "VIP"
    elseif (iuserlevel="3") then
        getUserLevelStr = "VIP GOLD"
    elseif (iuserlevel="4") then
        getUserLevelStr = "VVIP"
    elseif (iuserlevel="7") then
        getUserLevelStr = "STAFF"
    elseif (iuserlevel="8") then
        getUserLevelStr = "FAMILY"
    elseif (iuserlevel="9") then
        getUserLevelStr = "BIZ"
    elseif (iuserlevel="50" or iuserlevel="51") then
        getUserLevelStr = "제휴몰"
    elseif (iuserlevel="" or iuserlevel="99") then
        getUserLevelStr = "비회원/탈퇴"
    end if
end function

' 회원등급 변경 이전 데이터가 필요없을시에만 이거 쓸것. 예전 등급을 표기 해야 할때는 getUserLevelColorByDate 이거 사용해야함    2018.10.19 한용민
'//회원등급 컬러
function getUserLevelColor(iuserlevel)
    getUserLevelColor = "#000000"

    if (iuserlevel="0") then		'/ WHITE
        getUserLevelColor = "#A4A8AA"
    elseif (iuserlevel="1") then		'/ RED
        getUserLevelColor = "red"
    elseif (iuserlevel="2") then		'/ VIP
        getUserLevelColor = "#66BB66"
    elseif (iuserlevel="3") then		'/ VIP GOLD
        getUserLevelColor = "#BBBB33"
    elseif (iuserlevel="4") then		'/ VVIP
        getUserLevelColor = "#0000FF"
    elseif (iuserlevel="7") then		'/ STAFF
        getUserLevelColor = "black"
    elseif (iuserlevel="8") then		'/ FAMILY
        getUserLevelColor = "black"
    elseif (iuserlevel="9") then		'/ BIZ
        getUserLevelColor = "black"
    else
        getUserLevelColor = ""
    end if
end function

'// 기간별 회원등급명
function getUserLevelStrByDate(iuserlevel, baseDate)
    getUserLevelStrByDate = iuserlevel

    Select Case iuserlevel
        Case "0","5"
            getUserLevelStrByDate = "WHITE"
        Case "1"
            getUserLevelStrByDate = "RED"
        Case "2"
            getUserLevelStrByDate = "VIP"
        Case "3"
            getUserLevelStrByDate = "VIP GOLD"
        Case "4","6"
            getUserLevelStrByDate = "VVIP"
        Case "7"
            getUserLevelStrByDate = "STAFF"
        Case "8"
            getUserLevelStrByDate = "FAMILY"
        Case "9"
            getUserLevelStrByDate = "BIZ"
        Case "50","51"
            getUserLevelStrByDate = "제휴몰"
        Case "99"
            getUserLevelStrByDate = "비회원"
        Case Else
            getUserLevelStrByDate = "비회원"
    End Select
end function

'// 기간별 회원등급색상
function getUserLevelColorByDate(iuserlevel, baseDate)
    getUserLevelColorByDate = "#000000"

    Select Case iuserlevel
        Case "0","5"
            getUserLevelColorByDate = "#A4A8AA"	'/ WHITE
        Case "1"
            getUserLevelColorByDate = "red"	    '/ RED
        Case "2"
            getUserLevelColorByDate = "#66BB66"	'/ VIP
        Case "3"
            getUserLevelColorByDate = "#BBBB33"	'/ VIP GOLD
        Case "4","6"
            getUserLevelColorByDate = "#0000FF"	'/ VVIP
        Case "7"
            getUserLevelColorByDate = "black"	'/ STAFF
        Case "8"
            getUserLevelColorByDate = "black"	'/ FAMILY
        Case "9"
            getUserLevelColorByDate = "black"	'/ BIZ
        Case Else
            getUserLevelColorByDate = ""
    End Select
end function

'//회원등급		'/2016.06.29 한용민 생성
function DrawselectboxUserLevel(selectedname, selectedId, chplg)
%>
    <select class='select' name="<%= selectedname %>" <%= chplg %>>
	    <option value="" <% if selectedId="" then response.write " selected" %> >전체</option>
        <option value="0" <% if selectedId="0" then response.write " selected" %> >WHITE</option>
        <option value="1" <% if selectedId="1" then response.write " selected" %> >RED</option>
        <option value="2" <% if selectedId="2" then response.write " selected" %> >VIP</option>
        <option value="3" <% if selectedId="3" then response.write " selected" %> >VIP GOLD</option>
        <option value="4" <% if selectedId="4" then response.write " selected" %> >VVIP</option>
        <option value="7" <% if selectedId="7" then response.write " selected" %> >STAFF</option>
        <option value="8" <% if selectedId="8" then response.write " selected" %> >FAMILY</option>
        <option value="9" <% if selectedId="9" then response.write " selected" %> >BIZ</option>
    </select>
<%
end function

' 사용중지. <!-- #include virtual="/lib/classes/maechul/incMaechulFunction.asp"--> 에 sbGetOptPGgubun("pggubun") 함수 사용할것.
function DrawSelectBoxPGGubun(selectedname, selectedId, chplg)
%>
    <select class='select' name="<%= selectedname %>" <%= chplg %>>
	    <option value="" <% if selectedId="" then response.write " selected" %> >전체</option>
		<option value="inicis" <% if (selectedId = "inicis") then %>selected<% end if %> >inicis</option>
		<option value="uplus" <% if (selectedId = "uplus") then %>selected<% end if %> >uplus</option>
		<option value="kcp" <% if (selectedId = "kcp") then %>selected<% end if %> >kcp</option>
		<option value="kakaopay" <% if (selectedId = "kakaopay") then %>selected<% end if %> >kakaopay</option>
		<option value="newkakaopay" <% if (selectedId = "newkakaopay") then %>selected<% end if %> >newkakaopay</option>
		<option value="naverpay" <% if (selectedId = "naverpay") then %>selected<% end if %> >naverpay</option>
		<option value="payco" <% if (selectedId = "payco") then %>selected<% end if %> >payco</option>
		<option value="gifticon" <% if (selectedId = "gifticon") then %>selected<% end if %> >gifticon</option>
		<option value="giftting" <% if (selectedId = "giftting") then %>selected<% end if %> >giftting</option>
		<option value="allat" <% if (selectedId = "allat") then %>selected<% end if %> >allat</option>
		<option value="mobilians" <% if (selectedId = "mobilians") then %>selected<% end if %> >mobilians</option>
		<option value="bankipkum" <% if (selectedId = "bankipkum") then %>selected<% end if %> >bankipkum</option>
		<option value="bankrefund" <% if (selectedId = "bankrefund") then %>selected<% end if %> >bankrefund</option>
		<option value="okcashbag" <% if (selectedId = "okcashbag") then %>selected<% end if %> >okcashbag</option>
		<option value="toss" <% if (selectedId = "toss") then %>selected<% end if %> >toss</option>
        <option value="chai" <% if (selectedId = "chai") then %>selected<% end if %> >chai</option>
        <option value="balance" <% if (selectedId = "balance") then %>selected<% end if %> >balance</option>
        <option value="convinienspay" <% if (selectedId = "convinienspay") then %>selected<% end if %> >convinienspay</option>
    </select>
<%
end function

'// 기간별 배송비   ' 2018.12.27 한용민 생성
function getDefaultBeasongPayByDate(vBaseDate)
    dim vTmpBeasongPay
    vTmpBeasongPay = 2500

    if vBaseDate >= "2019-01-01" then
        vTmpBeasongPay = 2500
    else
        vTmpBeasongPay = 2000
    end if

    getDefaultBeasongPayByDate = vTmpBeasongPay
end function

'//사용중지		'/DrawselectboxUserLevel 이걸로 쓸것
Sub DrawUserLevelCombo(selectedname,selectedId)
%>
    <select class='select' name="<%= selectedname %>">
	    <option value="" <% if selectedId="" then response.write " selected" %> >전체</option>
	</select>
<%
end Sub

Sub drawSelectBoxSellYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >판매</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >일시품절</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >품절</option>
   <option value="NS" <% if selectedId="NS" then response.write "selected" %> >품절+일시품절</option>
   <option value="NN" <% if selectedId="NN" then response.write "selected" %> >품절+옵션사용안함</option>
   <option value="YS" <% if selectedId="YS" then response.write "selected" %> >판매+일시품절</option>
   </select>
   <%
End Sub

' 사용여부
Sub drawSelectBoxUsingYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">CHOICE</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >Y</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >N</option>
   </select>
   <%
End Sub

Sub drawSelectBoxDanjongYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >생산중</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >재고부족</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >단종</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >MD품절</option>
   <option value="YM" <% if selectedId="YM" then response.write "selected" %> >단종+MD품절</option>
   <option value="SN" <% if selectedId="SN" then response.write "selected" %> >단종아님</option>
   </select>
   <%
End Sub

Sub drawSelectBoxLimitYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >비한정</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >한정</option>
   <option value="Y0" <% if selectedId="Y0" then response.write "selected" %> >한정(0)</option>
   </select>
   <%
End Sub

Sub drawSelectBoxMWU(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="MW" <% if selectedId="MW" then response.write "selected" %> >매입+위탁</option>
   <option value="W" <% if selectedId="W" then response.write "selected" %> >위탁</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >매입</option>
   <option value="U" <% if selectedId="U" then response.write "selected" %> >업체</option>
   </select>
   <%
End Sub

Sub drawSelectBoxPackYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >포장가능</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >포장불가능</option>
   </select>
   <%
End Sub

Sub drawSelectBoxSailYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >할인</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >할인안함</option>
   </select>
   <%
End Sub

Sub drawSelectBoxCouponYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >쿠폰할인</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >쿠폰없음</option>
   </select>
   <%
End Sub

Sub drawSelectBoxVatYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >과세</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >면세</option>
   </select>
   <%
End Sub

Sub drawSelectBoxItemGubun(selectBoxName,selectedId)
dim tmp_str,query1
%>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">전체</option>
		<option value="10" <% if selectedId="10" then response.write "selected" %> >10</option>
		<option value="35" <% if selectedId="35" then response.write "selected" %> >35</option>
		<option value="55" <% if selectedId="55" then response.write "selected" %> >55</option>
		<option value="70" <% if selectedId="70" then response.write "selected" %> >70</option>
		<option value="75" <% if selectedId="75" then response.write "selected" %> >75</option>
		<option value="76" <% if selectedId="76" then response.write "selected" %> >76</option>
		<option value="80" <% if selectedId="80" then response.write "selected" %> >80</option>
		<option value="85" <% if selectedId="85" then response.write "selected" %> >85</option>
		<option value="90" <% if selectedId="90" then response.write "selected" %> >90</option>
		<option value="98" <% if selectedId="98" then response.write "selected" %> >98</option>
	</select>
<%
End Sub

'// 2016-06-21, skyer9
Sub drawSelectBoxItemGubunForSearch(selectBoxName,selectedId)
	dim tmp_str,query1
%>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">전체</option>
		<option value="10" <% if selectedId="10" then response.write "selected" %> >10</option>
		<option value="exc10" <% if selectedId="exc10" then response.write "selected" %> >10제외</option>
		<option value="55" <% if selectedId="55" then response.write "selected" %> >55(CS매입)</option>
		<option value="60" <% if selectedId="60" then response.write "selected" %> >60(할인권)</option>
		<option value="70" <% if selectedId="70" then response.write "selected" %> >70(소모품)</option>
		<option value="75" <% if selectedId="75" then response.write "selected" %> >75(부자재)</option>
		<option value="76" <% if selectedId="76" then response.write "selected" %> >76(핑거스부자재)</option>
		<option value="80" <% if selectedId="80" then response.write "selected" %> >80(OFF사은품)</option>
		<option value="85" <% if selectedId="85" then response.write "selected" %> >85(ON사은품)</option>
		<option value="90" <% if selectedId="90" then response.write "selected" %> >90(OFF전용)</option>
		<option value="98" <% if selectedId="98" then response.write "selected" %> >98(핑거스상품)</option>
	</select>
<%
End Sub

function GetItemGubunName(itemgubun)
	if itemgubun="10" then
		GetItemGubunName = "일반"
	elseif itemgubun="90" then
		GetItemGubunName = "오프전용"
	elseif itemgubun="60" then
		GetItemGubunName = "기타"
	elseif itemgubun="70" then
		GetItemGubunName = "소모품"
	elseif itemgubun="75" then
		GetItemGubunName = "저장품"
	elseif itemgubun="80" then
		GetItemGubunName = "사은품"
	elseif itemgubun="85" then
		GetItemGubunName = "사은품"
	elseif itemgubun="97" then
		GetItemGubunName = "강좌"
	elseif itemgubun="98" then
		GetItemGubunName = "DIY"
	elseif itemgubun="99" then
		GetItemGubunName = "일반"
	elseif itemgubun="95" then
		GetItemGubunName = "기타"
	else
		GetItemGubunName = "기타" ''itemgubun
	end if
end function

Sub drawSelectBoxIsOverSeaYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >안함</option>
   </select>
   <%
End Sub

' 상품구분  ' 2022.09.20 한용민 생성
function getItemDiv(itemDiv)
    dim resultValue

    if itemDiv="" or isnull(itemDiv) then exit function

    if itemDiv="01" then
        resultValue="일반상품"
    elseif itemDiv="06" then
        resultValue="주문제작(문구)"
    elseif itemDiv="16" then
        resultValue="주문제작"
    elseif itemDiv="08" then
        resultValue="티켓상품/클래스상품"
    elseif itemDiv="09" then
        resultValue="Present상품"
    elseif itemDiv="11" then
        resultValue="상품권상품"
    elseif itemDiv="18" then
        resultValue="여행상품"
    elseif itemDiv="07" then
        resultValue="구매제한상품"
    elseif itemDiv="82" then
        resultValue="마일리지샵 상품"
    elseif itemDiv="75" then
        resultValue="정기구독상품"
    elseif itemDiv="30" then
        resultValue="렌탈상품"
    elseif itemDiv="23" then
        resultValue="B2B상품"
    elseif itemDiv="17" then
        resultValue="마케팅전용상품"
    elseif itemDiv="21" then
        resultValue="딜상품"
    else
        resultValue=itemDiv
    end if

    getItemDiv=resultValue
end function

Sub drawSelectBoxItemDiv(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="08" <% if selectedId="08" then response.write "selected" %> >티켓/클래스 상품</option>
   <option value="09" <% if selectedId="09" then response.write "selected" %> >Present상품</option>
   <option value="11" <% if selectedId="11" then response.write "selected" %> >상품권상품</option>
   <option value="18" <% if selectedId="18" then response.write "selected" %> >여행상품</option>
   <option value="75" <% if selectedId="75" then response.write "selected" %> >정기구독상품</option>
   <option value="30" <% if selectedId="30" then response.write "selected" %> >이니렌탈상품</option>
   <option value="23" <% if selectedId="23" then response.write "selected" %> >B2B상품</option>
   <option value="">-------------</option>
   <option value="16" <% if selectedId="16" then response.write "selected" %> >주문제작</option>
   <option value="06" <% if selectedId="06" then response.write "selected" %> >주문제작(문구)</option>
   </select>
   <%
End Sub

Sub drawSelectBoxItemDivDeal(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="08" <% if selectedId="08" then response.write "selected" %> >티켓/클래스 상품</option>
   <option value="09" <% if selectedId="09" then response.write "selected" %> >Present상품</option>
   <option value="11" <% if selectedId="11" then response.write "selected" %> >상품권상품</option>
   <option value="18" <% if selectedId="18" then response.write "selected" %> >여행상품</option>
   <option value="75" <% if selectedId="75" then response.write "selected" %> >정기구독상품</option>
   <option value="">-------------</option>
   <option value="16" <% if selectedId="16" then response.write "selected" %> >주문제작</option>
   <option value="06" <% if selectedId="06" then response.write "selected" %> >주문제작(문구)</option>
    <option value="21" <% if selectedId="21" then response.write "selected" %> >딜상품</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsWeightYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >안함</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsPlusSaleItem(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="P" <% if selectedId="P" then response.write "selected" %> >플러스세일</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >일반상품</option>
   </select>
   <%
End Sub

'####### 201205808(강준구) /admin/member/popbrandinfoonly.asp 2.브랜드관련정보에 구매유형 select 박스 추가.
' 사용중지. 이함수 왜 만들었음? 다 찾아서 drawPartnerCommCodeBox 함수로 통합작업해놓음. 2022.02.09 한용민
Sub DrawBrandPurchaseType(IsAll,selectedname,selectedId,onChange)
%>
	<select class='select' name="<%= selectedname %>" <%=onChange%>>
	<% If IsAll = "Y" Then %><option value=''>-전체-</option><% End If %>
	<option value='1' <% if selectedId="1" then response.write " selected" %> >일반유통</option>
	<option value='4' <% if selectedId="4" then response.write " selected" %> >사입</option>
	<option value='5' <% if selectedId="5" then response.write " selected" %> >ODM</option>
	<option value='6' <% if selectedId="6" then response.write " selected" %> >수입</option>
	<option value='7' <% if selectedId="7" then response.write " selected" %> >브랜드수입</option>
	<option value='8' <% if selectedId="8" then response.write " selected" %> >제작</option>
    <option value='9' <% if selectedId="9" then response.write " selected" %> >해외직구</option>
    <option value='10' <% if selectedId="10" then response.write " selected" %> >B2B</option>
	<option value='101' <% if selectedId="101" then response.write " selected" %> >일반유통 제외</option>
	</select>
<%
end sub

' 사용중지. 디비에서 일괄로 쿼리해서 가져 오세요.
function getBrandPurchaseType(PurchaseType)
	dim tmpBrandPurchaseType

	if PurchaseType="1" then
		tmpBrandPurchaseType="일반유통"
	elseif PurchaseType="3" then
		tmpBrandPurchaseType="PB"
	elseif PurchaseType="4" then
		tmpBrandPurchaseType="사입"
	elseif PurchaseType="5" then
		tmpBrandPurchaseType="ODM"
	elseif PurchaseType="6" then
		tmpBrandPurchaseType="수입"
	elseif PurchaseType="7" then
		tmpBrandPurchaseType="브랜드수입"
	elseif PurchaseType="8" then
		tmpBrandPurchaseType="제작"
	elseif PurchaseType="9" then
		tmpBrandPurchaseType="해외직구"
	elseif PurchaseType="10" then
		tmpBrandPurchaseType="B2B"
	else
		tmpBrandPurchaseType=PurchaseType
	end if

	getBrandPurchaseType=tmpBrandPurchaseType
end function

'//2014-01-10 이종화 상품검색 베스트순 추가
Sub drawSelectBoxIsBestSorting(selectBoxName,selectedId)
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="new" <% IF selectedId="new" Then response.write "selected" %> >신상품순</option>
		<option value="cashH" <% IF selectedId="cashH" Then response.write "selected" %>>높은가격순</option>
		<option value="cashL" <% IF selectedId="cashL" Then response.write "selected" %>>낮은가격순</option>
		<option value="best" <% IF selectedId="best" Then response.write "selected" %>>베스트순</option>
	</select>
   <%
End Sub

''20120824 서동석 추가
Sub drawPartnerCommCodeBox(IsAllflag,comm_group,selectBoxName,selectedId,onChange)
   dim tmp_str,query1

   %><select class="select" name="<%=selectBoxName%>" <%=onChange%> >
     <% If IsAllflag Then %><option value=''>-선택-</option><% End If %>
   <%
   query1 = " select pcomm_cd,pcomm_name,pcomm_isusing "&VbCRLF
   query1 = query1 & " from [db_partner].[dbo].tbl_partner_comm_code with (nolock)"&VbCRLF
   query1 = query1 & " where pcomm_group='"&comm_group&"'"&VbCRLF
   query1 = query1 & " order by pcomm_sortno"&VbCRLF
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("pcomm_cd")) then
               tmp_str = " selected"
           end if

           if (rsget("pcomm_isusing")="Y") or (tmp_str<>"") then
               response.write("<option value='"&rsget("pcomm_cd")&"' "&tmp_str&">"&rsget("pcomm_name")&"</option>")
            end if
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

Sub drawCSCommCodeBox(IsAllflag,comm_group,selectBoxName,selectedId,onChange)
   dim tmp_str,query1

   %><select class="select" name="<%=selectBoxName%>" <%=onChange%> >
     <% If IsAllflag Then %><option value=''>-선택-</option><% End If %>
   <%

   query1 = " select cd.comm_cd, cd.comm_name, cd.comm_isDel "&VbCRLF
   query1 = query1 & " from "&VbCRLF
   query1 = query1 & " 	[db_cs].[dbo].[tbl_cs_comm_code] cm "&VbCRLF
   query1 = query1 & " 	join [db_cs].[dbo].[tbl_cs_comm_code] cd "&VbCRLF
   query1 = query1 & " 	on "&VbCRLF
   query1 = query1 & " 		1 = 1 "&VbCRLF
   query1 = query1 & " 		and cm.comm_cd = '"&comm_group&"' "&VbCRLF
   query1 = query1 & " 		and cm.comm_cd = cd.comm_group "&VbCRLF
   query1 = query1 & " where "&VbCRLF
   query1 = query1 & " 	1 = 1 "&VbCRLF
   query1 = query1 & " order by "&VbCRLF
   query1 = query1 & " 	cd.sortno "&VbCRLF
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("comm_cd")) then
               tmp_str = " selected"
           end if

           if (rsget("comm_isDel")="N") or (tmp_str<>"") then
               response.write("<option value='"&rsget("comm_cd")&"' "&tmp_str&">"&rsget("comm_name")&"</option>")
            end if
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

''20120824 서동석 추가
function getPartnerCommCodeName(comm_group,pcomm_cd)
   dim tmp_str,query1
   query1 = " select pcomm_cd,pcomm_name,pcomm_isusing "&VbCRLF
   query1 = query1 & " from [db_partner].[dbo].tbl_partner_comm_code "&VbCRLF
   query1 = query1 & " where pcomm_group='"&comm_group&"'"&VbCRLF
   query1 = query1 & " and pcomm_cd='"&pcomm_cd&"'"

   rsget.Open query1,dbget,1
   if  not rsget.EOF  then
       getPartnerCommCodeName = rsget("pcomm_name")
   end if
   rsget.Close
end function

function fndrawSaleBizSecCombo(IsAllflag,selectBoxName,selectedId,onChange)
    Dim strSql, arrrows, buf, tmp_str
    strSql = "db_partner.dbo.sp_Ten_TMS_BA_BIZSECTION_getList('','','','Y','Y')"
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	IF Not (rsget.EOF OR rsget.BOF) THEN
		arrrows = rsget.getRows()
	END IF
	rsget.close

	buf = "<select class='select' name='"&selectBoxName&"' "&onChange&" >"
    If IsAllflag Then
        buf = buf + "<option value=''>-선택-</option>"
    End If

	if isArray(arrrows) then
	    For i = 0 To UBound(arrrows,2)
	        if Lcase(selectedId) = Lcase(arrrows(0,i)) then
               tmp_str = " selected"
            end if
	        buf = buf + "<option value='"&arrrows(0,i)&"' "&tmp_str&">"&arrrows(1,i)&"</option>"

	        tmp_str=""
	    Next
	end if
	buf = buf + "</select>"
	fndrawSaleBizSecCombo = buf
end function

function fndrawSaleBizSecComboName(bizsection_cd)
	dim tmp_str,query1

    if bizsection_cd="" or isnull(bizsection_cd) then
        fndrawSaleBizSecComboName=""
        exit function
    end if

	query1 = " select top 1 BIZSECTION_NM "
	query1 = query1 & " from db_partner.dbo.tbl_TMS_BA_BIZSECTION "
	query1 = query1 & " where BIZSECTION_CD = '" + CStr(bizsection_cd) + "' "

	rsget.Open query1,dbget,1
	if  not rsget.EOF  then
		fndrawSaleBizSecComboName = rsget("BIZSECTION_NM")
	end if
	rsget.Close
end function

'// 구분에 따른 문자열 색상 지정
function fnColor(str, div)
	Select Case div
		Case "yn"
			if str<>"Y" or isNull(str) then
				fnColor = "<Font color=#F08050>" & str & "</font>"
			else
				fnColor = "<Font color=#5080F0>" & str & "</font>"
			end if
		Case "mw"
			Select Case str
				Case "M"
					fnColor = "<Font color=#F08050>매입</font>"
				Case "W"
					fnColor = "<Font color=#808080>위탁</font>"
				Case "U"
					fnColor = "<Font color=#5080F0>업체</font>"
			end Select
		Case "tx"
			if str="Y" then
				fnColor = "<Font color=#808080>과세</font>"
			elseif str="N" then
				fnColor = "<Font color=#F08050>면세</font>"
			else
			    fnColor = str
			end if
		Case "dj"
			if str="Y" then
				fnColor = "<Font color=#33CC33>단종</font>"
			elseif str="S" then
				fnColor = "<Font color=#3333CC>재고부족</font>"
			elseif str="M" then
				fnColor = "<Font color=#CC3333>MD품절</font>"
			end if
		Case "delivery"
			IF str THEN
				fnColor = "<Font color=#F08050>업체</font>"
			ELSE
				fnColor = "<Font color=#5080F0>10x10</font>"
			end IF
		Case "sellyn"
			IF str="N" THEN
				fnColor = "<Font color=#F08050>품절</font>"
			elseif str="S" then
			    fnColor = "<Font color=#3333CC>일시품절</font>"
			end IF
		Case "cancelyn"
			IF str="N" THEN
				fnColor = "<Font color=#000000>정상</font>"
			elseif str="D" then
			    fnColor = "<Font color=#FF0000>삭제</font>"
			elseif str="Y" then
			    fnColor = "<Font color=#FF0000>취소</font>"
			elseif str="A" then
			    fnColor = "<Font color=#FF0000>추가</font>"
			end IF
	end Select
end Function

Function GetCategoryName(cdl)
   rsget.Open " select code_nm from [db_item].[dbo].tbl_Cate_large where code_large='" & cdl & "'",dbget,1
   if Not(rsget.EOF or rsget.BOF) then
   	GetCategoryName = rsget(0)
   else
   	GetCategoryName = ""
   end if
   rsget.Close
End Function

'// 사은품 영수증 표시내역
Function fnComGetEventConditionStr(ByVal Fgiftkind_type, ByVal Fgift_scope,ByVal Fgift_type,ByVal Fgift_range1, ByVal Fgift_range2,ByVal FGiftName,ByVal Fgiftkind_cnt, ByVal Fgiftkind_orgcnt, ByVal Fgiftkind_limit, ByVal Fgiftkind_givecnt,ByVal FMakerid)
Dim reStr
dim remainEa

        reStr = ""
        if (FMakerid<> "") then
        	reStr = reStr + FMakerid + " "
        end if
        if (Fgift_scope="1") then
            reStr = reStr + "전체 구매 고객 "
        elseif (Fgift_scope="2") then
            reStr = reStr + "이벤트등록상품 "
        elseif (Fgift_scope="3") then
            reStr = reStr + "선택브랜드상품 "
        elseif (Fgift_scope="4") then
            reStr = reStr + "이벤트그룹상품"
        elseif (Fgift_scope="5") then
            reStr = reStr + "선택상품"
        elseif (Fgift_scope="9") then
            reStr = reStr + "다이어리샵상품포함"
        end if

        if (Fgift_type="1") then
            reStr = reStr + "모든 구매자"
        elseif (Fgift_type="2") then
            if (Fgift_range2=0) then
                reStr = reStr + CStr(Fgift_range1) + " 원 이상 구매시 "
            else
                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " 원 구매시 "
            end if
        elseif (Fgift_type="3") then
            if (Fgift_range2=0) then
                reStr = reStr + CStr(Fgift_range1) + " 개 이상 구매시 "
            else
                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " 개 구매시 "
            end if
        end if
        reStr = reStr &"'"&  FGiftName &"' "
        reStr = reStr &  Cstr(Fgiftkind_orgcnt) & " 개 "

        if (Fgiftkind_type=2) then
            reStr = reStr + "[1+1]"
             reStr = reStr & "(총 "& Cstr(Fgiftkind_cnt) & " 개)"
        elseif (Fgiftkind_type=3) then
            reStr = reStr + "[1:1]"
             reStr = reStr & "(총 "& Cstr(Fgiftkind_cnt) & " 개)"
        end if
         reStr = reStr + " 증정"


        if Fgiftkind_limit<>0 then
            reStr = reStr & " 총한정 [" & Fgiftkind_limit & "]"
            remainEa = Fgiftkind_limit-Fgiftkind_givecnt
            if (remainEa<0) then remainEa=0
             reStr = reStr & " 현재남은수량 " & remainEa
        end if
        fnComGetEventConditionStr = reStr
 End Function


	Function NullFillWith(src , data )
		if isNULL(src) or src = "" then
			if Not isNull(data) or data = "" then
				NullFillWith = data
			 else
			 	NullFillWith = 0
			end if
		else
			If Not IsNumeric(src) then
				NullFillWith = Replace(Trim(src),"'","''")
			else
				NullFillWith = src
			End if
		end if
	End Function

function fnGetSongjangURL(byval isongjangdiv)
    if IsNULL(isongjangdiv) then Exit function
    if isongjangdiv="" then Exit function

   rsget.Open " select findurl from db_order.dbo.tbl_songjang_div where divcd=" & CStr(isongjangdiv) & "",dbget,1
   if Not(rsget.EOF or rsget.BOF) then
   	fnGetSongjangURL = db2html(rsget(0))
   else
   	fnGetSongjangURL = ""
   end if
   rsget.Close
end function


function fnGetOffCurrencyUnit(byval shopid,byRef CurrencyUnit, byRef CurrencyChar, byRef ExchangeRate)
    Dim sqlStr
    sqlStr = "select U.CurrencyUnit,U.ExchangeRate,X.CurrencyChar from db_shop.dbo.tbl_shop_User U"
	sqlStr = sqlStr & " Left Join db_shop.dbo.tbl_shop_exchangeRate X"
	sqlStr = sqlStr & " on U.CurrencyUnit=X.CurrencyUnit"
    sqlStr = sqlStr & " where userid='"&shopid&"'"

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    If Not rsget.Eof then
        CurrencyUnit = rsget("CurrencyUnit")
        CurrencyChar = rsget("CurrencyChar")
        ExchangeRate = rsget("ExchangeRate")
    ELSE
        CurrencyUnit = "WON"
        CurrencyChar = "원"
        ExchangeRate = 1
    end if
    rsget.Close
end function

'//언어단위		'/2011.12.13 한용민 생성
function DrawexchangeRate_countrylangcd(selectBoxName,selectedId,loginsite,changeFlag)
	dim tmp_str,query1

	query1 = " select"
	query1 = query1 & " countrylangcd"
	query1 = query1 & " from db_item.dbo.tbl_exchangeRate"

	if loginsite<>"" and loginsite<>"SCM" then
		query1 = query1 & " where sitename='"& loginsite &"'"
	end if

	query1 = query1 & " group by countrylangcd"
	query1 = query1 & " order by countrylangcd asc"

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>CHOICE</option>
	<%

	if not rsget.EOF then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("countrylangcd")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("countrylangcd")&"' "&tmp_str&">"&rsget("countrylangcd")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	else
	%>
		<option value='KR' <%if selectedId="KR" then response.write " selected"%>>KR</option>
	<%
	end if
	rsget.close

	response.write("</select>")
end function

Function CurrURL()
	CurrURL = Request.ServerVariables("PATH_INFO")
End Function


Function CurrURLQ()
	CurrURLQ = "http://" & Request.ServerVariables("Server_name") & CurrURL()
	If Request.ServerVariables("REQUEST_METHOD") = "POST" then
		CurrURLQ = CurrURL & "?" & Request.Form
	 else
 		CurrURLQ = CurrURL & "?" & Request.QueryString
	End if
End Function

'//해당월 마지막날짜 찾기		'/2012.05.14 한용민 생성
function LastDayOfThisMonth(intYear,intMonth)
dim intNextYear ,intNextMonth

	'다음해, 다음달 확인
	If intMonth = 12 Then
	 intNextYear = intYear + 1
	 intNextMonth = 1
	Else
	 intNextMonth = intMonth + 1
	 intNextYear = intYear
	End If

	'이번달 마지막날 찾기
	If intMonth = 12 Then
		LastDayOfThisMonth = Day(Dateadd("d", -1, intNextYear & "-" & intNextMonth & "-1"))
	Else
		LastDayOfThisMonth = Day(Dateadd("d", -1, intYear & "-" & intMonth+1 & "-1"))
	End If
end function


'### 날씨 이미지(v는 날씨값,w는 width, h는 height)
Function WeatherImage(v,w,h)
	If v = "" OR v = "0" Then
		WeatherImage = ""
	Else
		WeatherImage = "<img src=""/images/weather/" & v & ".gif"" " & CHKIIF(w<>"","width="""&w&"""","") & " " & CHKIIF(h<>"","height="""&h&"""","") & ">"
	End IF
End Function

'/온라인,오프라인구분		'/2012.11.29 한용민 생성
function drawonoffgubun(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>ALL</option>
		<option value="ONLINE" <% if selVal="ONLINE" then response.write " selected" %>>온라인</option>
		<option value="OFFLINE" <% if selVal="OFFLINE" then response.write " selected" %>>오프라인</option>
	</select>
<%
end function

'//두개의 값중 큰것을 반환		'/2013.02.08 한용민 생성
function maxvalreturn(a,b)
	if (CDbl(a)> CDbl(b)) then
		maxvalreturn=a
	else
		maxvalreturn=b
	end if
end function

'/메일러구분		'/2013.03.27 한용민 생성
function drawmailergubun(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>전체</option>
        <option value="TMS" <% if selVal="TMS" then response.write " selected" %>>TMS</option>
		<option value="AMAIL" <% if selVal="AMAIL" then response.write " selected" %>>AMAIL</option>
		<option value="THUNDERMAIL" <% if selVal="THUNDERMAIL" then response.write " selected" %>>THUNDERMAIL</option>
	</select>
<%
end function

'//날짜형식 2013-01-01 오후 03:00:00 형식을 2013-01-01 15:00:00로 변환		'/2013.04.22 한용민 생성
function dateconvert(dateval)
	dim tmpval
	if dateval = "" then exit function

	tmpval = year(dateval)
	tmpval = tmpval & "-" & Format00(2,month(dateval))
	tmpval = tmpval & "-" & Format00(2,day(dateval))
	tmpval = tmpval & " " & Format00(2,hour(dateval))
	tmpval = tmpval & ":" & Format00(2,minute(dateval))
	tmpval = tmpval & ":" & Format00(2,second(dateval))

	dateconvert = left(tmpval,19)
end function

function DrawItemInfoDiv(selBoxName, selVal, isShowExists, chplg)
%>
<select name="<%=selBoxName%>" class="select" <%=chplg%> >
	<option value="" >===전체====</option>
	<% if (isShowExists) then %>
	<option value="YY" <%=chkIIF(selVal="YY","selected","")%>>품목 입력 전체</option>
	<option value="NN" <%=chkIIF(selVal="NN","selected","")%>>품목 미입력 전체</option>
	<% end if %>
	<option value="01" <%=chkIIF(selVal="01","selected","")%>>01.의류</option>
	<option value="02" <%=chkIIF(selVal="02","selected","")%>>02.구두/신발</option>
	<option value="03" <%=chkIIF(selVal="03","selected","")%>>03.가방</option>
	<option value="04" <%=chkIIF(selVal="04","selected","")%>>04.패션잡화</option>
	<option value="05" <%=chkIIF(selVal="05","selected","")%>>05.침구류/커튼</option>
	<option value="06" <%=chkIIF(selVal="06","selected","")%>>06.가구</option>
	<option value="07" <%=chkIIF(selVal="07","selected","")%>>07.영상가전</option>
	<option value="08" <%=chkIIF(selVal="08","selected","")%>>08.가정용 전기제품</option>
	<option value="09" <%=chkIIF(selVal="09","selected","")%>>09.계절가전</option>
	<option value="10" <%=chkIIF(selVal="10","selected","")%>>10.사무용기기</option>
	<option value="11" <%=chkIIF(selVal="11","selected","")%>>11.광학기기</option>
	<option value="12" <%=chkIIF(selVal="12","selected","")%>>12.소형전자</option>
	<option value="13" <%=chkIIF(selVal="13","selected","")%>>13.휴대폰</option>
	<option value="14" <%=chkIIF(selVal="14","selected","")%>>14.내비게이션</option>
	<option value="15" <%=chkIIF(selVal="15","selected","")%>>15.자동차용품</option>
	<option value="16" <%=chkIIF(selVal="16","selected","")%>>16.의료기기</option>
	<option value="17" <%=chkIIF(selVal="17","selected","")%>>17.주방용품</option>
	<option value="18" <%=chkIIF(selVal="18","selected","")%>>18.화장품</option>
	<option value="19" <%=chkIIF(selVal="19","selected","")%>>19.귀금속/보석/시계류</option>
	<option value="20" <%=chkIIF(selVal="20","selected","")%>>20.식품</option>
	<option value="21" <%=chkIIF(selVal="21","selected","")%>>21.가공식품</option>
	<option value="22" <%=chkIIF(selVal="22","selected","")%>>22.건강기능식품</option>
	<option value="23" <%=chkIIF(selVal="23","selected","")%>>23.영유아용품</option>
	<option value="24" <%=chkIIF(selVal="24","selected","")%>>24.악기</option>
	<option value="25" <%=chkIIF(selVal="25","selected","")%>>25.스포츠용품</option>
	<option value="26" <%=chkIIF(selVal="26","selected","")%>>26.서적</option>
	<option value="27" <%=chkIIF(selVal="27","selected","")%>>27.호텔/펜션 예약</option>
	<option value="28" <%=chkIIF(selVal="28","selected","")%>>28.여행패키지</option>
	<option value="29" <%=chkIIF(selVal="29","selected","")%>>29.항공권</option>
	<option value="30" <%=chkIIF(selVal="30","selected","")%>>30.자동차 대여 서비스</option>
	<option value="31" <%=chkIIF(selVal="31","selected","")%>>31.물품대여 서비스</option>
	<option value="32" <%=chkIIF(selVal="32","selected","")%>>32.물품대여 서비스</option>
	<option value="33" <%=chkIIF(selVal="33","selected","")%>>33.디지털 콘텐츠</option>
	<option value="34" <%=chkIIF(selVal="34","selected","")%>>34.상품권/쿠폰</option>
	<option value="35" <%=chkIIF(selVal="35","selected","")%>>35.기타</option>
</select>
<%
end function

'/브랜드구분 	'/2013.08.30 한용민 생성
Function drawSelectBoxbrandgubun(selectBoxName,selectedId,chplg)
	Dim tmp_str,query1

	query1 = " SELECT brandgubun, brandgubunname"
	query1 = query1 & " FROM db_brand.dbo.tbl_street_brandgubun"
	query1 = query1 & " ORDER BY orderno ASC"

	'response.write query1 & "<br>"
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
	   rsget.Movefirst
	   Do until rsget.EOF
	       If Lcase(selectedId) = Lcase(rsget("brandgubun")) then
	           tmp_str = " selected"
	       End If
	       response.write("<option value='"&rsget("brandgubun")&"' "&tmp_str&">"&rsget("brandgubunname")&" ["&rsget("brandgubun")&"]</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   Loop
	end if
	rsget.close
	response.write("</select>")
End Function

'/재고수량	 	'/2013.10.17 한용민 생성
Function drawSelectBoxexistsstock(selectBoxName,selectedId,chplg)
	Dim tmp_str
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
		<option value='1UP' <%if selectedId="1UP" then response.write " selected"%>>1개 이상</option>
		<option value='0DOWN' <%if selectedId="0DOWN" then response.write " selected"%>>0개 이하</option>
		<option value='20DOWN' <%if selectedId="20DOWN" then response.write " selected"%>>20개 이하</option>
		<option value='1UP20DOWN' <%if selectedId="1UP20DOWN" then response.write " selected"%>>1개 이상 20개 이하</option>
		<option value='20UP' <%if selectedId="20UP" then response.write " selected"%>>20개 이상</option>
		<option value='1UP3DOWN' <%if selectedId="1UP3DOWN" then response.write " selected"%>>1개 이상 3개 이하</option>
		<option value='MINUS' <%if selectedId="MINUS" then response.write " selected"%>>마이너스재고</option>
	</select>
<%
End Function

' 사용중지. <!-- #include virtual="/lib/classes/maechul/incMaechulFunction.asp"--> 에 sbGetOptPGgubun("pgid") 함수 사용할것.
Function DrawSelectBoxPGUserid(selectBoxName, selectedId, chplg)
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value=""></option>
		<option value="teenxteen3" <% if (selectedId = "teenxteen3") then %>selected<% end if %> >teenxteen3</option>
		<option value="teenxteen4" <% if (selectedId = "teenxteen4") then %>selected<% end if %> >teenxteen4</option>
		<option value="teenxteen5" <% if (selectedId = "teenxteen5") then %>selected<% end if %> >teenxteen5</option>
		<option value="teenxteen6" <% if (selectedId = "teenxteen6") then %>selected<% end if %> >teenxteen6</option>
		<option value="teenxteen8" <% if (selectedId = "teenxteen8") then %>selected<% end if %> >teenxteen8</option>
		<option value="teenxteen9" <% if (selectedId = "teenxteen9") then %>selected<% end if %> >teenxteen9</option>
		<option value="teenteen10" <% if (selectedId = "teenteen10") then %>selected<% end if %> >teenteen10</option>
		<option value="tenbyten01" <% if (selectedId = "tenbyten01") then %>selected<% end if %> >tenbyten01</option>
		<option value="tenbyten02" <% if (selectedId = "tenbyten02") then %>selected<% end if %> >tenbyten02</option>
		<option value="teenxteeha" <% if (selectedId = "teenxteeha") then %>selected<% end if %> >teenxteeha</option>
        <option value="teenxteenr" <% if (selectedId = "teenxteenr") then %>selected<% end if %> >teenxteenr</option>
        <option value="teenteensp" <% if (selectedId = "teenteensp") then %>selected<% end if %> >teenteensp</option>
        <option value="teenteenap" <% if (selectedId = "teenteenap") then %>selected<% end if %> >teenteenap</option>
		<option value="KCTEN0001m" <% if (selectedId = "KCTEN0001m") then %>selected<% end if %> >KCTEN0001m</option>
        <option value="newkakaopay" <% if (selectedId = "newkakaopay") then %>selected<% end if %> >newkakaopay</option>
        <option value="payco" <% if (selectedId = "payco") then %>selected<% end if %> >payco</option>
        <option value="naverpay" <% if (selectedId = "naverpay") then %>selected<% end if %> >naverpay</option>
		<option value="toss" <% if (selectedId = "toss") then %>selected<% end if %> >toss</option>
        <option value="chai" <% if (selectedId = "chai") then %>selected<% end if %> >chai</option>
		<option value="bankipkum" <% if (selectedId = "bankipkum") then %>selected<% end if %> >bankipkum</option>
		<option value="bankipkum_10x10" <% if (selectedId = "bankipkum_10x10") then %>selected<% end if %> >bankipkum_10x10</option>
		<option value="bankipkum_fingers" <% if (selectedId = "bankipkum_fingers") then %>selected<% end if %> >bankipkum_fingers</option>
		<option value="bankrefund" <% if (selectedId = "bankrefund") then %>selected<% end if %> >bankrefund</option>
		<option value="bankrefund_10x10" <% if (selectedId = "bankrefund_10x10") then %>selected<% end if %> >bankrefund_10x10</option>
		<option value="bankrefund_fingers" <% if (selectedId = "bankrefund_fingers") then %>selected<% end if %> >bankrefund_fingers</option>
		<option value="10x10_2" <% if (selectedId = "10x10_2") then %>selected<% end if %> >10x10_2</option>
		<option value="R5523" <% if (selectedId = "R5523") then %>selected<% end if %> >R5523</option>
		<option value="mobilians" <% if (selectedId = "mobilians") then %>selected<% end if %> >mobilians</option>
		<option value="gifticon" <% if (selectedId = "gifticon") then %>selected<% end if %> >gifticon</option>
		<option value="giftting" <% if (selectedId = "giftting") then %>selected<% end if %> >giftting</option>
		<option value="okcashbag" <% if (selectedId = "okcashbag") then %>selected<% end if %> >okcashbag</option>
        <option value="nopayment" <% if (selectedId = "nopayment") then %>selected<% end if %> >nopayment</option>
        <option value="balance" <% if (selectedId = "balance") then %>selected<% end if %> >balance</option>
        <option value="giftcard" <% if (selectedId = "giftcard") then %>selected<% end if %> >giftcard</option>
        <option value="mileage" <% if (selectedId = "mileage") then %>selected<% end if %> >mileage</option>
        <option value="XXX" <% if (selectedId = "XXX") then %>selected<% end if %> >XXX</option>
	</select>
<%
End Function

Function DrawSelectBoxPGUseridOff(selectBoxName, selectedId, chplg)
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value=""></option>
		<option value="비씨카드사" <% if (selectedId = "비씨카드사") then %>selected<% end if %> >비씨카드사</option>
		<option value="롯데카드사" <% if (selectedId = "롯데카드사") then %>selected<% end if %> >롯데카드사</option>
		<option value="삼성카드사" <% if (selectedId = "삼성카드사") then %>selected<% end if %> >삼성카드사</option>
		<option value="신한카드" <% if (selectedId = "신한카드") then %>selected<% end if %> >신한카드</option>
		<option value="하나카드" <% if (selectedId = "하나카드") then %>selected<% end if %> >하나카드</option>
		<option value="현대카드사" <% if (selectedId = "현대카드사") then %>selected<% end if %> >현대카드사</option>
		<option value="Alipay" <% if (selectedId = "Alipay") then %>selected<% end if %> >Alipay</option>
		<option value="KB국민카드" <% if (selectedId = "KB국민카드") then %>selected<% end if %> >KB국민카드</option>
		<option value="NH농협카드" <% if (selectedId = "NH농협카드") then %>selected<% end if %> >NH농협카드</option>
	</select>
<%
End Function

'//경영매출로그 매입구분 & 배송구분		'/2013.11.15 한용민 생성
function getmwdiv_beasongdivname(tmpval)
	if tmpval="" then exit function

	if trim(tmpval)="M" then
		getmwdiv_beasongdivname="매입"
	elseif trim(tmpval)="W" then
		getmwdiv_beasongdivname="위탁"
	elseif trim(tmpval)="U" then
		getmwdiv_beasongdivname="업체"
	elseif trim(tmpval)="R" then
		getmwdiv_beasongdivname="랜탈"
	elseif trim(tmpval)="TT" then
		getmwdiv_beasongdivname="텐배"
	elseif trim(tmpval)="UU" then
		getmwdiv_beasongdivname="업배"
	elseif trim(tmpval)="PP" then
		getmwdiv_beasongdivname="포장"
	elseif trim(tmpval)="LC" then
		getmwdiv_beasongdivname="강좌"
	elseif trim(tmpval)="B000" then
		getmwdiv_beasongdivname="미지정"
	elseif trim(tmpval)="B011" then
		getmwdiv_beasongdivname="위탁판매"
	elseif trim(tmpval)="B012" then
		getmwdiv_beasongdivname="업체위탁"
	elseif trim(tmpval)="B013" then
		getmwdiv_beasongdivname="출고위탁"
	elseif trim(tmpval)="B021" then
		getmwdiv_beasongdivname="오프매입"
	elseif trim(tmpval)="B022" then
		getmwdiv_beasongdivname="매장매입"
	elseif trim(tmpval)="B023" then
		getmwdiv_beasongdivname="가맹점매입"
	elseif trim(tmpval)="B031" then
		getmwdiv_beasongdivname="출고매입"
	elseif trim(tmpval)="B032" then
		getmwdiv_beasongdivname="센터매입"
	elseif trim(tmpval)="B999" then
		getmwdiv_beasongdivname="기타보정"
	end if
end function

'//경영매출로그 매입구분 & 배송구분		'/2013.11.15 한용민 생성
function drawmwdiv_beasongdiv(selectBoxName, selectedId, chplg)
   dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value="">전체</option>
		<option value="">ON ------</option>
		<option value="M" <% if selectedId="M" then response.write "selected" %> >매입</option>
		<option value="W" <% if selectedId="W" then response.write "selected" %> >위탁</option>
		<option value="U" <% if selectedId="U" then response.write "selected" %> >업체</option>
        <option value="R" <% if selectedId="R" then response.write "selected" %> >랜탈</option>
		<option value="TT" <% if selectedId="TT" then response.write "selected" %> >텐배</option>
		<option value="UU" <% if selectedId="UU" then response.write "selected" %> >업배</option>
		<option value="PP" <% if selectedId="PP" then response.write "selected" %> >포장</option>
        <option value="LC" <% if selectedId="LC" then response.write "selected" %> >강좌</option>
		<option value="">OF ------</option>
		<option value="B000" <% if selectedId="B000" then response.write "selected" %> >미지정
		<option value="B011" <% if selectedId="B011" then response.write "selected" %> >위탁판매
		<option value="B012" <% if selectedId="B012" then response.write "selected" %> >업체위탁
		<option value="B013" <% if selectedId="B013" then response.write "selected" %> >출고위탁
		<option value="B021" <% if selectedId="B021" then response.write "selected" %> >오프매입
		<option value="B022" <% if selectedId="B022" then response.write "selected" %> >매장매입
		<option value="B023" <% if selectedId="B023" then response.write "selected" %> >가맹점매입
		<option value="B031" <% if selectedId="B031" then response.write "selected" %> >출고매입
		<option value="B032" <% if selectedId="B032" then response.write "selected" %> >센터매입
		<option value="B999" <% if selectedId="B999" then response.write "selected" %> >기타보정
	</select>
   <%
End function

function draw3plMeachulComboBox(selectBoxName, selectedId)
    dim ret
	ret =""
	ret = ret&"<select name='"&selectBoxName&"' class='select'>"
'	ret = ret&"<option value='' "&CHKIIF(selectedId="","selected","")&">3pl매출제외"
'	ret = ret&"<option value='A' "&CHKIIF(selectedId="A","selected","")&">3pl매출포함"
'	ret = ret&"<option value='P' "&CHKIIF(selectedId="P","selected","")&">3pl매출만"
	ret = ret&"<option value='' "&CHKIIF(selectedId="","selected","")&">텐바이텐"
	ret = ret&"<option value='A' "&CHKIIF(selectedId="A","selected","")&">텐바이텐+아이띵소"
	ret = ret&"<option value='P' "&CHKIIF(selectedId="P","selected","")&">아이띵소"
	ret = ret&"</select>"
    response.write ret
end function

function drawSellChannelComboBox(selectBoxName, selectedId)
    dim ret
    ret =""
    ret = ret&"<select name='"&selectBoxName&"'>"
    ret = ret&"<option value='' "&CHKIIF(selectedId="","selected","")&">전체</option>"
    ret = ret&"<option value='WEB' "&CHKIIF(selectedId="WEB","selected","")&">WWW</option>"
    ret = ret&"<option value='MOB' "&CHKIIF(selectedId="MOB","selected","")&">모바일</option>"
    ret = ret&"<option value='MOBLNK' "&CHKIIF(selectedId="MOBLNK","selected","")&">모바일_제휴</option>"
    ret = ret&"<option value='APP' "&CHKIIF(selectedId="APP","selected","")&">APP</option>"
    ret = ret&"<option value='APPLNK' "&CHKIIF(selectedId="APPLNK","selected","")&">APP_제휴</option>"
    ret = ret&"<option value='OUT' "&CHKIIF(selectedId="OUT","selected","")&">제휴몰</option>"
    ret = ret&"<option value='FGN' "&CHKIIF(selectedId="FGN","selected","")&">해외몰</option>"  ''2017/01/10 분리.
    ret = ret&"<option value='3PL' "&CHKIIF(selectedId="3PL","selected","")&">3PL</option>"

    ret = ret&"<option value='TEN' "&CHKIIF(selectedId="TEN","selected","")&">TEN(해외,제휴,3PL제외)</option>"
    ret = ret&"<option value='KEY' "&CHKIIF(selectedId="KEY","selected","")&">키워드광고</option>"
    ret = ret&"</select>"
    response.write ret
end function

function getChannelvalue2ArrIDx(ivalue)
    if (ivalue="WEB") then
        getChannelvalue2ArrIDx = "1,2"
    elseif (ivalue="MOB") then
        getChannelvalue2ArrIDx = "4"
    elseif (ivalue="MOBLNK") then
        getChannelvalue2ArrIDx = "5"
    elseif (ivalue="APP") then
        getChannelvalue2ArrIDx = "7"
    elseif (ivalue="APPLNK") then
        getChannelvalue2ArrIDx = "8"
    elseif (ivalue="OUT") then
        getChannelvalue2ArrIDx = "50,51"
    elseif (ivalue="FGN") then
        getChannelvalue2ArrIDx = "80"
    elseif (ivalue="3PL") then
        getChannelvalue2ArrIDx = "90"
    elseif (ivalue="TEN") then
        getChannelvalue2ArrIDx = "1,2,4,5,7"   ''8 is between
    else
        getChannelvalue2ArrIDx = "0"
    end if
end function

public function getSellChannelDivName(ibeadaldiv)
    if (ibeadaldiv="1") or (ibeadaldiv="2") then
        getSellChannelDivName = "WEB" ''WEB			'// 왜 지웠을까... skyer9, 23017-03-02
    elseif (ibeadaldiv="4")   then
        getSellChannelDivName = "MOB"
    elseif  (ibeadaldiv="5") then
        getSellChannelDivName = "MOBLINK"
    elseif (ibeadaldiv="7") then
        getSellChannelDivName = "APP"
    elseif (ibeadaldiv="8") then
        getSellChannelDivName = "APPLINK"
    elseif (ibeadaldiv="50") or (ibeadaldiv="51") then
        getSellChannelDivName = "OUT"
    elseif (ibeadaldiv="90") then
        getSellChannelDivName = "3PL"
    elseif (ibeadaldiv="80") then
        getSellChannelDivName = "FGN"
    else
        getSellChannelDivName = "???"
    end if
end function

public function getSellChannelName(ibeadaldiv)
    if (ibeadaldiv="1") or (ibeadaldiv="2") then
        getSellChannelName = "WWW"
    elseif (ibeadaldiv="4")   then
        getSellChannelName = "모바일"
    elseif  (ibeadaldiv="5") then
        getSellChannelName = "모바일_제휴"
    elseif (ibeadaldiv="7") then
        getSellChannelName = "APP"
    elseif (ibeadaldiv="8") then
        getSellChannelName = "APP_제휴"
    elseif (ibeadaldiv="50") or (ibeadaldiv="51") then
        getSellChannelName = "제휴몰"
    elseif (ibeadaldiv="90") then
        getSellChannelName = "3PL"
    elseif (ibeadaldiv="80") then
        getSellChannelName = "해외몰"
    else
        getSellChannelName = "???"
    end if
end function

'//모바일+모바일제휴, App+App제휴 추가
function drawSellChannelComboBoxGroup(selectBoxName, selectedId)
    dim ret
    ret =""
    ret = ret&"<select name='"&selectBoxName&"'>"
    ret = ret&"<option value='' "&CHKIIF(selectedId="","selected","")&">전체"
    ret = ret&"<option value='WEB' "&CHKIIF(selectedId="WEB","selected","")&">WWW"
    ret = ret&"<option value='MOB' "&CHKIIF(selectedId="MOB","selected","")&">모바일"
    ret = ret&"<option value='MOBLNK' "&CHKIIF(selectedId="MOBLNK","selected","")&">모바일_제휴"
    ret = ret&"<option value='APP' "&CHKIIF(selectedId="APP","selected","")&">APP"
    ret = ret&"<option value='APPLNK' "&CHKIIF(selectedId="APPLNK","selected","")&">APP_제휴"
    ret = ret&"<option value='MOBALL' "&CHKIIF(selectedId="MOBALL","selected","")&">모바일+모바일_제휴"
    ret = ret&"<option value='APPALL' "&CHKIIF(selectedId="APPALL","selected","")&">APP+APP_제휴"
    ret = ret&"<option value='MOBAPPALL' "&CHKIIF(selectedId="MOBAPPALL","selected","")&">모바일(제휴포함)+APP(제휴포함)"
    ret = ret&"<option value='OUT' "&CHKIIF(selectedId="OUT","selected","")&">제휴몰"
    ret = ret&"<option value='FGN' "&CHKIIF(selectedId="FGN","selected","")&">해외몰"
    ret = ret&"<option value='3PL' "&CHKIIF(selectedId="3PL","selected","")&">3PL"

    ret = ret&"<option value='TEN' "&CHKIIF(selectedId="TEN","selected","")&">TEN(해외,제휴,3PL제외)"
    ret = ret&"</select>"
    response.write ret
end function

'//모바일+모바일제휴, App+App제휴 추가, 모바일+모바일제휴+App+App제휴 추가
function getChannelvalue2ArrIDxGroup(ivalue)
    if (ivalue="WEB") then
        getChannelvalue2ArrIDxGroup = "1,2"
    elseif (ivalue="MOB") then
        getChannelvalue2ArrIDxGroup = "4"
    elseif (ivalue="MOBLNK") then
        getChannelvalue2ArrIDxGroup = "5"
    elseif (ivalue="APP") then
        getChannelvalue2ArrIDxGroup = "7"
    elseif (ivalue="APPLNK") then
        getChannelvalue2ArrIDxGroup = "8"
    elseif (ivalue="MOBALL") then
        getChannelvalue2ArrIDxGroup = "4,5"
    elseif (ivalue="APPALL") then
        getChannelvalue2ArrIDxGroup = "7,8"
    elseif (ivalue="MOBAPPALL" ) then
        getChannelvalue2ArrIDxGroup = "4,5,7,8"
    elseif (ivalue="OUT") then
        getChannelvalue2ArrIDxGroup = "50,51"
    elseif (ivalue="FGN") then
        getChannelvalue2ArrIDxGroup = "80"
    elseif (ivalue="3PL") then
        getChannelvalue2ArrIDxGroup = "90"
    elseif (ivalue="TEN") then
        getChannelvalue2ArrIDxGroup = "1,2,4,5,7"
    else
        getChannelvalue2ArrIDxGroup = "0"
    end if
end function

'//상품고시 품목정보		'/2013.12.11 한용민 추가
Function drawSelectBoxinfodiv(selectBoxName,selectedId,chplg)
	Dim tmp_str,query1

	query1 = " SELECT top 1000 infoDiv, infoDivName"
	query1 = query1 & " FROM [db_item].dbo.tbl_item_infoDiv"
	query1 = query1 & " ORDER BY infoDiv ASC"

	'response.write query1 & "<br>"
%>
	<select class="select" name="<%=selectBoxName%>" <%= chplg %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
	rsget.Open query1,dbget,1
	If  not rsget.EOF  then
	   rsget.Movefirst
	   Do until rsget.EOF
	       If Lcase(selectedId) = Lcase(rsget("infoDiv")) then
	           tmp_str = " selected"
	       End If
	       response.write("<option value='"&rsget("infoDiv")&"' "&tmp_str&">["& rsget("infoDiv")&"] "&db2html(rsget("infoDivName"))&"</option>")
	       tmp_str = ""
	       rsget.MoveNext
	   Loop
	end if
	rsget.close
	response.write("</select>")

End Function

function fn_isDongSoongIP()
    dim tmpip : tmpip = request.ServerVariables("REMOTE_ADDR")
    dim tmp_ALLOWIPLIST
    tmp_ALLOWIPLIST = Array(  "115.94.163.42","115.94.163.43","115.94.163.44","115.94.163.45","115.94.163.46" _
                        ,"61.252.133.2","61.252.133.3","61.252.133.4","61.252.133.5","61.252.133.6" _
                        ,"61.252.133.7","61.252.133.8","61.252.133.9","61.252.133.10","61.252.133.11" _
                        ,"61.252.133.12","61.252.133.13","61.252.133.14","61.252.133.15","61.252.133.16" _
                        ,"61.252.133.17","61.252.133.18","61.252.133.19","61.252.133.20","61.252.133.21" _
                        ,"61.252.133.22","61.252.133.23","61.252.133.24","61.252.133.25","61.252.133.26" _
                        ,"61.252.133.27","61.252.133.28","61.252.133.29","61.252.133.30","61.252.133.31" _
                        ,"61.252.133.32","61.252.133.33","61.252.133.34","61.252.133.35","61.252.133.36" _
                        ,"61.252.133.37","61.252.133.38","61.252.133.39","61.252.133.40","61.252.133.41" _
                        ,"61.252.133.67","61.252.133.68","61.252.133.69","61.252.133.70" _
                        ,"61.252.133.71","61.252.133.72","61.252.133.73","61.252.133.74","61.252.133.75" _
                        ,"61.252.133.76","61.252.133.77","61.252.133.78","61.252.133.79","61.252.133.80" _
                        ,"61.252.133.81","61.252.133.82","61.252.133.83","61.252.133.84","61.252.133.85","61.252.133.86","61.252.133.91" _
                        ,"61.252.133.100","61.252.133.103","61.252.133.104","61.252.133.105","61.252.133.106","61.252.133.107" _
                        ,"61.252.133.113","61.252.133.114","61.252.133.115","61.252.133.116","61.252.133.117","61.252.133.118" _
                        ,"61.252.133.121","61.252.133.122","61.252.133.123","61.252.133.124","61.252.133.125", "61.252.133.92" _
                        ,"112.218.65.240","112.218.65.241","112.218.65.242","112.218.65.243","112.218.65.244","112.218.65.245" _
						,"112.218.65.246","112.218.65.247","112.218.65.248","112.218.65.249","112.218.65.250","112.218.65.251" _
						,"112.218.65.252","112.218.65.253","112.218.65.254" _
                      )

    dim IPCheckOK
    dim tmp_ip_i, tmp_ip_buf1
    IPCheckOK = false
    for tmp_ip_i=0 to UBound(tmp_ALLOWIPLIST)
        tmp_ip_buf1 = tmp_ALLOWIPLIST(tmp_ip_i)
        if (tmpip=tmp_ip_buf1) then
            IPCheckOK = true
            Exit For
        end if
    next

    fn_isDongSoongIP = IPCheckOK
end function

function checkDataLengthDBArr(orgStrArr,oSplitStr,MaxByteLen, byref retErrMsg)
    dim sqlStr, errExists
    dim iRow,idatalen,iVAL

    errExists = false
    orgStrArr = replace(orgStrArr,"|","/")
    orgStrArr = replace(orgStrArr,oSplitStr,"|")
    orgStrArr = replace(orgStrArr,"'","''")

    sqlStr = " select top 1 iRow,datalength(VAL) as idatalen,VAL as iVAL" & VbCRLF
    sqlStr = sqlStr & " from db_cs.[dbo].SplitStringWITHRow('"&orgStrArr&"','|')" & VbCRLF
    sqlStr = sqlStr & " where datalength(VAL)>"&MaxByteLen &VbCRLF

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
    	iRow = rsget("iRow")
    	idatalen = rsget("idatalen")
    	iVAL = rsget("iVAL")
    	errExists = true
    end if
    rsget.Close

    if (errExists) then
        retErrMsg = iRow&"번째 줄 ["&iVAL&"] 길이초과("&idatalen&" byte). 최대("&MaxByteLen&" byte) 까지 가능"
    end if
    checkDataLengthDBArr = (NOT errExists)

end function

function socialnoReplace(byval ino)
    socialnoReplace = ino
    if isNull(ino) then Exit function
    dim ret : ret = Trim(replace(ino,"-",""))
    if (LEN(ret)=13) then
        ret = LEFT(ret,6)&"-*******"
        socialnoReplace = ret
    end if

end function

function RemoveLastCariageReturn(str)
	RemoveLastCariageReturn = ""
	if IsNull(str) then
		Exit Function
	end if

	do while Len(str) >= 2
		if Right(str, 2) = vbCrLf then
			str = Left(str, Len(str) - 2)
		else
			exit do
		end if
	loop

	RemoveLastCariageReturn = str
end function

function socialnoBlank(byval ino)
    socialnoBlank = ino
    if isNull(ino) then Exit function
    dim ret : ret = Trim(replace(ino,"-",""))
    if (LEN(ret)=13) then
        socialnoBlank = ""
    end if

end function

'/소수점 올림함수	'/2016.10.24 한용민 생성
Function ceil(ByVal intParam)
	ceil = -(Int(-(intParam)))
End Function

function drawSelectBox3plcompany(selectBoxName, selectedId, chgval)
dim tmp_str
%>
	<select class="select" name="<%=selectBoxName%>" <%= chgval %>>
		<option value="">전체</option>
		<option value="tplithinkso" <% if selectedId="tplithinkso" then response.write "selected" %> >ITHINKSO</option>
        <option value="tpliconic" <% if selectedId="tpliconic" then response.write "selected" %> >코니테일</option>
        <option value="tplmmmg" <% if selectedId="tplmmmg" then response.write "selected" %> >MMMG</option>
        <option value="tplparagon" <% if selectedId="tplparagon" then response.write "selected" %> >파라곤</option>
        <option value="tplclass101" <% if selectedId="tplclass101" then response.write "selected" %> >클래스101</option>
	</select>
<%
End function

'// 안전인증
function fnSafetyDivCodeName(c)
	dim r
	select case c
		case "10" : r = "전기용품 > 안전인증"
		case "20" : r = "전기용품 > 안전확인신고"
		case "30" : r = "전기용품 > 공급자적합성확인"
		case "40" : r = "생활제품 > 안전인증"
		case "50" : r = "생활제품 > 안전확인"
		case "60" : r = "생활제품 > 공급자적합성확인"
		case "70" : r = "어린이제품 > 안전인증"
		case "80" : r = "어린이제품 > 안전확인"
		case "90" : r = "어린이제품 > 공급자적합성확인"
	end select
	fnSafetyDivCodeName = r
end function

'// 안전인증
function drawSelectBoxSafetyDivCode(selectBoxName, selectedId, safetyYn, chgval)
%>
	<select name="<%= selectBoxName %>" id="safetyDiv" <%=chkIIF(safetyYn<>"Y","disabled","")%> <%= chgval %>>
		<option value="">::안전인증구분::</option>
		<option value="10">전기용품 > 안전인증</option>
		<option value="20">전기용품 > 안전확인신고</option>
		<option value="30">전기용품 > 공급자적합성확인</option>
		<option value="40">생활제품 > 안전인증</option>
		<option value="50">생활제품 > 안전확인</option>
		<option value="60">생활제품 > 공급자적합성확인</option>
		<option value="70">어린이제품 > 안전인증</option>
		<option value="80">어린이제품 > 안전확인</option>
		<option value="90">어린이제품 > 공급자적합성확인</option>
	</select>
<%
End function

'// 전화번호, 휴대폰번호 마스킹 별처리 010-111-3333 => 010-***-3333
function AstarPhoneNumber(phoneNumber)
	Dim regEx, result
	Set regEx = New RegExp

	With regEx
		.Pattern = "([0-9]+)-([0-9]+)-([0-9]+)"
		.IgnoreCase = True
		.Global = True
	End With

	result = regEx.Replace(phoneNumber,"$1-***-$3")

	if (result = phoneNumber) then
		if (Len(phoneNumber) >= 4) then
			result = Left(phoneNumber, (Len(phoneNumber) - 4)) & "****"
		end if
	end if

	set regEx = nothing

	AstarPhoneNumber = result
end Function

'// 이름 마스킹 별처리 홍길동 => 홍*동
function AstarUserName(userName)
	Dim result

	Select Case Len(userName)
		Case 0
			''
		Case 1
			result = "*"
		Case 2
			result = Left(userName,1) & "*"
		Case Else
			''3이상
			result = Left(userName,1) & "*" & Right(userName,1)
	End Select

	AstarUserName = result
end function

'// 아이디 마스킹 별처리 홍길동 => 홍*동      ' 이상구 생성
function AstarUserid(curUserid)
	dim resultStr, leftLen, rightLen

	resultStr = ""
	If IsNull(curUserid) Then
		AstarUserid = resultStr
		Exit Function
	End If

	'// 가운데 3글자
	If Len(curUserid) <= 3 Then
		resultStr = getereg_replace(curUserid, ".", "*", True)
		AstarUserid = resultStr
		Exit Function
	End If

	If (Len(curUserid) - 3) Mod 2 = 0 Then
		leftLen = (Len(curUserid) - 3) / 2
		rightLen = Len(curUserid) - 3 - leftLen
	Else
		leftLen = Int((Len(curUserid) - 3) / 2) + 1
		rightLen = Len(curUserid) - 3 - leftLen
	End If

	resultStr = Left(curUserid, leftLen) & getereg_replace(Mid(curUserid, 3, 3), ".", "*", True) & Right(curUserid, rightLen)
	AstarUserid = resultStr
end function

' 이상구 생성
function getereg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
    ' Function replaces pattern with replacement
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    getereg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function

' 결제수단		' 2017.12.26 한용민 생성
function GetaccountdivName(accountdiv)
	dim accountdivName

    if (accountdiv="6") then
        accountdivName = "무통장입금"
    elseif (accountdiv="7") then
        accountdivName = "무통장"
    elseif (accountdiv="14") then
        accountdivName = "편의점결제"
    elseif (accountdiv="20") then
        accountdivName = "실시간이체"
    elseif (accountdiv="30") then
        accountdivName = "포인트"
    elseif (accountdiv="50") then
        accountdivName = "제휴몰결제"
    elseif (accountdiv="77") then
        accountdivName = "무통장환불"
    elseif (accountdiv="80") then
        accountdivName = "올엣카드"
    elseif (accountdiv="90") then
        accountdivName = "상품권"
    elseif (accountdiv="100") then
        accountdivName = "신용카드"
    elseif (accountdiv="110") then
        accountdivName = "OK캐시백"
    elseif (accountdiv="400") then
        accountdivName = "핸드폰"
    elseif (accountdiv="550") then
        accountdivName = "기프팅"
    elseif (accountdiv="560") then
        accountdivName = "기프티콘"
	else
		accountdivName = accountdiv
    end if

	GetaccountdivName = accountdivName
end function

' 상품고시		' 2018.03.21 한용민 생성
function DrawInfoDiv(selectBoxName, selectedId, changeFlag)
	dim tmp_str,query1

	query1 = "select" & vbcrlf
	query1 = query1 & " id.infoDiv, id.infoDivName, id.infoValidCnt, id.SafetyTargetYN, id.SafetyCertYN, id.SafetyConfirmYN, id.SafetySupplyYN" & vbcrlf
	query1 = query1 & " , id.SafetyComply, id.regdate, id.lastupdate, id.lastadminid, id.IsUsing" & vbcrlf
	query1 = query1 & " from db_item.dbo.tbl_item_infoDiv id" & vbcrlf
	query1 = query1 & " where id.IsUsing='Y'" & vbcrlf
	query1 = query1 & " order by id.infoDiv asc" & vbcrlf

	'response.write query1 &"<Br>"
	rsget.Open query1,dbget,1
	%>
	<select class="select" id="ggg" name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%

	if not rsget.EOF then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("infoDiv")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("infoDiv")&"' "&tmp_str&" SafetyTargetYN='"&rsget("SafetyTargetYN")&"' SafetyCertYN='"&rsget("SafetyCertYN")&"' SafetyConfirmYN='"&rsget("SafetyConfirmYN")&"' SafetySupplyYN='"&rsget("SafetySupplyYN")&"' SafetyComply='"&rsget("SafetyComply")&"'>"& db2html(rsget("infoDivName"))&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
	end if
	rsget.close

	response.write("</select>")
end function

'// 해외직구여부
Public Function fnUniPassNumber(orderserial)
	Dim sqlStr , uniPassNumber
	sqlStr = "EXEC [db_order].[dbo].[usp_WWW_Order_DirectPurchase_Get] " & orderserial
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.eof then
		uniPassNumber = rsget(1)
	end if
	rsget.close
	fnUniPassNumber = uniPassNumber
End Function

' 직원, 브랜드 변경로그 2018.08.14 한용민 생성
function fnChkauthlog(empno, userid, logtype, logmsg, adminid)
	dim tmp_str,query1

    if empno="" and userid="" then exit function
    if logtype="" then exit function

    query1 = "INSERT INTO db_partner.dbo.tbl_partner_authlog (empno, userid, logtype, logmsg, adminid ) values" & vbcrlf
	query1 = query1 & " ('"& empno &"', '"& userid &"', '"& logtype &"', '"& logmsg &"', '"& adminid &"')" & vbcrlf

	'response.write query1 &"<Br>"
    dbget.execute query1
end function

'// SNS 가입 구분
Function GetSNSJoinTypeName(snsgubun)
    Select Case snsgubun
        Case "fb"
            GetSNSJoinTypeName = "페이스북"
        Case "nv"
            GetSNSJoinTypeName = "네이버"
        Case "ka"
            GetSNSJoinTypeName = "카카오톡"
        Case "gl"
            GetSNSJoinTypeName = "구글"
        Case "ap"
            GetSNSJoinTypeName = "애플"
        case else
            GetSNSJoinTypeName = snsgubun
    End Select
end Function

' 로그인/개인정보 권한 IP 관리 DB화   ' 2019.09.20 한용민 생성
function fncheckAllowIPWithByDB(usescmyn, usecustomerinfoyn, uselogicsyn)
    dim icheck_UserIP : icheck_UserIP = request.ServerVariables("REMOTE_ADDR")
	dim sqlStr, found, arrIP, icheck_UserIPShort

	sqlStr = "select top 1" & vbcrlf
    sqlStr = sqlStr & " ipaddress, usescmyn, uselogicsyn, usecustomerinfoyn" & vbcrlf
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_loginIP" & vbcrlf
	sqlStr = sqlStr & " where useyn = 'Y' and ipaddress = '" & icheck_UserIP & "'" & vbcrlf

    if usescmyn<>"" then
        sqlStr = sqlStr & " and usescmyn = '"& usescmyn &"' " & vbcrlf
    end if
    if usecustomerinfoyn<>"" then
        sqlStr = sqlStr & " and usecustomerinfoyn = '"& usecustomerinfoyn &"' " & vbcrlf
    end if
    if uselogicsyn<>"" then
        sqlStr = sqlStr & " and uselogicsyn = '"& uselogicsyn &"' " & vbcrlf
    end if

	'response.write sqlStr & "<br>"
	found = False
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		found = True
	end if
	rsget.Close

	if found = False then
        '한번 더 검색
        if (application("Svr_Info")="Dev") then
            icheck_UserIPShort = "192.168.1"
        else
            arrIP = split(icheck_UserIP,".")
            icheck_UserIPShort = arrIP(0) & "." & arrIP(1) & "." & arrIP(2)
        end if
        sqlStr = "select top 1" & vbcrlf
        sqlStr = sqlStr & " ipaddress, usescmyn, uselogicsyn, usecustomerinfoyn" & vbcrlf
        sqlStr = sqlStr & " from db_partner.dbo.tbl_user_loginIP" & vbcrlf
        sqlStr = sqlStr & " where useyn = 'Y' and ipaddress = '" & icheck_UserIPShort & "'" & vbcrlf

        if usescmyn<>"" then
            sqlStr = sqlStr & " and usescmyn = '"& usescmyn &"' " & vbcrlf
        end if
        if usecustomerinfoyn<>"" then
            sqlStr = sqlStr & " and usecustomerinfoyn = '"& usecustomerinfoyn &"' " & vbcrlf
        end if
        if uselogicsyn<>"" then
            sqlStr = sqlStr & " and uselogicsyn = '"& uselogicsyn &"' " & vbcrlf
        end if
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            found = True
        end if
        rsget.Close

		'// 사설 아이피 전부 허용
'		if Left(icheck_UserIP, 8) = "192.168." then
'			found = True
'		end if
	end if

	fncheckAllowIPWithByDB = found
end function

' 쌍따옴표, 이중따옴표 생성     ' 2019.09.26 한용민 생성
Function escapedstring(vtmp)
    dim resultStr

    resultStr=vtmp
	resultStr = Replace(resultStr, Chr(34), "")		'// 이중 따옴표 제거
	resultStr = Chr(34) & resultStr & Chr(34)
	escapedstring=resultStr
End Function

'/ 입사일 대비 연차,년차 가져옴 	'/2019.11.20 한용민 추가
function getjoindayyeardiff(empno , userid)
    dim sqlStr ,sqlsearch, tmpjoinday, yyyy, today

    tmpjoinday=""

	if empno = "" and userid = "" then exit function

	if empno <> "" then
		sqlsearch = sqlsearch & " and t.empno = '"&empno&"'"
	end if
	if userid <> "" then
		sqlsearch = sqlsearch & " and p.id = '"&userid&"'"
	end if

	sqlStr = "select top 1 "
	sqlStr = sqlStr & " t.joinday"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " 	and p.isusing = 'Y'"
	sqlStr = sqlStr & " where 1=1 " & sqlsearch

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		tmpjoinday = rsget("joinday")
	end if
	rsget.close

	today = now()
	if (IsNull(tmpjoinday) or (tmpjoinday = "")) then
		getjoindayyeardiff = ""
		exit function
	end if

	yyyy = Year(today) - Year(tmpjoinday)
	 if (Month(tmpjoinday) > Month(today)) then
	 	yyyy = yyyy - 1
	 end if

	getjoindayyeardiff = yyyy
end function

'/ 어드민 직위 가져옴		'/2019.11.21 한용민 추가
function getposit_sn(empno , userid)
dim sqlStr ,sqlsearch, tmpposit_sn

	if empno = "" and userid = "" then exit function

	if empno <> "" then
		sqlsearch = sqlsearch & " and t.empno = '"&empno&"'"
	end if
	if userid <> "" then
		sqlsearch = sqlsearch & " and p.id = '"&userid&"'"
	end if

	sqlStr = "select top 1 "
	sqlStr = sqlStr & " t.posit_sn"
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " on t.userid = p.id"
	sqlStr = sqlStr & " where p.isusing = 'Y' " & sqlsearch

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		tmpposit_sn = rsget("posit_sn")
	end if
	rsget.close

    getposit_sn=tmpposit_sn
end function

' 정산구분      ' 2020.02.10 한용민 생성
function DrawJungsanGubun(selectBoxName,selectedId,comm_group, chplg)
    dim tmp_str,query1

    if comm_group="" then exit function
%>
    <select name="<%=selectBoxName%>" <%= chplg %>>
        <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
    query1 = " select comm_cd, comm_name from [db_jungsan].[dbo].tbl_jungsan_comm_code with (readuncommitted)"
    query1 = query1 + " where comm_group='"& trim(comm_group) &"'"
    query1 = query1 + " and comm_isDel='N'"
    rsget.CursorLocation = adUseClient
    rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

    if not rsget.EOF then
        do until rsget.EOF
        if Lcase(selectedId) = Lcase(rsget("comm_cd")) then
            tmp_str = " selected"
        end if
        response.write("<option value='"&rsget("comm_cd")&"' "&tmp_str&">" + rsget("comm_name") + "</option>")
        tmp_str = ""
        rsget.MoveNext
        loop
    end if
    rsget.close

    response.write("<option value='0000' "& chkIIF(selectedId="0000","selected","") &">     [미지정]</option>")
    response.write("</select>")
End function

'// AGV 피킹스테이션 선택상자 출력
Sub drawSelectPickingStation(selectBoxName, selectedId)
%>
	<select name="<%= selectBoxName %>" class="select">
        <option value="">피킹스테이션</option>
        <option value="2279" <%= CHKIIF(selectedId="2279", "selected", "") %>>2279 피킹01</option>
        <option value="2271" <%= CHKIIF(selectedId="2271", "selected", "") %>>2271 피킹02</option>
        <option value="2263" <%= CHKIIF(selectedId="2263", "selected", "") %>>2263 피킹03</option>
        <option value="2255" <%= CHKIIF(selectedId="2255", "selected", "") %>>2255 피킹04</option>
        <option value="2247" <%= CHKIIF(selectedId="2247", "selected", "") %>>2247 피킹05</option>
        <option value="2239" <%= CHKIIF(selectedId="2239", "selected", "") %>>2239 피킹06</option>
        <option value="2223" <%= CHKIIF(selectedId="2223", "selected", "") %>>2223 피킹07</option>
        <option value="2215" <%= CHKIIF(selectedId="2215", "selected", "") %>>2215 피킹08</option>
        <option value="2207" <%= CHKIIF(selectedId="2207", "selected", "") %>>2207 피킹09</option>
        <option value="2249" <%= CHKIIF(selectedId="2249", "selected", "") %>>2249</option>
        <option value="2259" <%= CHKIIF(selectedId="2259", "selected", "") %>>2259</option>
        <option value="2269" <%= CHKIIF(selectedId="2269", "selected", "") %>>2269</option>
	</select>
<%
End Sub

' 물류센터업체내역      '/2021.04.14 한용민 생성
sub drawSelectBoxSiteSeq(SiteSeqName, SiteSeqValue, chplg)
%>
    <select class="select" name="<%= SiteSeqName %>" <%= chplg %>>
        <option value="">선택</option>
        <option value="10" <% if SiteSeqValue="10" then response.write "selected" %> >텐바이텐</option>
        <option value="35" <% if SiteSeqValue="35" then response.write "selected" %> >텐바이텐(3PL)</option>
        <option value="98" <% if SiteSeqValue="98" then response.write "selected" %> >핑거스</option>
        <option value="30" <% if SiteSeqValue="30" then response.write "selected" %> >29cm</option>
        <option value="50" <% if SiteSeqValue="50" then response.write "selected" %> >탐스슈즈</option>
        <option value="99" <% if SiteSeqValue="99" then response.write "selected" %> >아이띵소</option>
    </select>
<%
End sub

' 물류센터업체내역      '/2021.04.14 한용민 생성
function getSiteSeqnamestr(SiteSeqValue)
    dim tmpsiteseqname

    if SiteSeqValue="10" then
        tmpsiteseqname="텐바이텐"
    elseif SiteSeqValue="35" then
        tmpsiteseqname="텐바이텐(3PL)"
    elseif SiteSeqValue="98" then
        tmpsiteseqname="핑거스"
    elseif SiteSeqValue="30" then
        tmpsiteseqname="29cm"
    elseif SiteSeqValue="50" then
        tmpsiteseqname="탐스슈즈"
    elseif SiteSeqValue="99" then
        tmpsiteseqname="아이띵소"
    else
        tmpsiteseqname=""
    end if

    getSiteSeqnamestr=tmpsiteseqname
End function

' 물류출고예정일    ' 2021.04.28 한용민 생성
function getchulgoscheduledate(orderserial)
    dim solar_date,query1

    if orderserial="" or isnull(orderserial) then exit function
    solar_date=""

    query1 = "declare @ipkumdate datetime"
    query1 = query1 & " select @ipkumdate=m.ipkumdate"
    query1 = query1 & " from db_order.dbo.tbl_order_master m with (nolock)"
    query1 = query1 & " where m.ipkumdiv > 3"
    query1 = query1 & " and m.cancelyn='N'"
    query1 = query1 & " and m.jumundiv not in (6,9)"
    query1 = query1 & " and m.orderserial=N'"& orderserial &"'"

    query1 = query1 & " set @ipkumdate=case when convert(varchar(10),@ipkumdate,8)<'15:00:00' then convert(varchar(10),@ipkumdate,121)"
    query1 = query1 & " 		else convert(varchar(10),dateadd(day,1,@ipkumdate),121)"
    query1 = query1 & " 		end"

    query1 = query1 & " select solar_date"
    query1 = query1 & " from db_sitemaster.dbo.LunarToSolar with(nolock)"
    query1 = query1 & " where solar_date>= @ipkumdate"
    query1 = query1 & " and solar_date<=convert(varchar(10),dateadd(day,30,@ipkumdate),121)"	' solar_date가 문자형이라서 인덱스를 안타서 최근30일로 제한하려고 넣음(로직적의미없음)
    query1 = query1 & " and not (holiday=2 or  isnull(holiday_name,'') like '%물류%휴무%' )"
    query1 = query1 & " order by convert(varchar(10),solar_date,121) ASC"

    'response.write query1 & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

    if not rsget.EOF then
        solar_date=rsget("solar_date")
    end if
    rsget.close

    getchulgoscheduledate=solar_date
End function

' 물류센터 택배사명      '/2021.05.25 한용민 생성
function getsongjangdivname(songjangdiv)
    dim sqlStr ,sqlsearch, songjangname

	if songjangdiv = "" or isnull(songjangdiv) = "" then exit function

	if songjangdiv <> "" then
		sqlsearch = sqlsearch & " and divcd = '"& songjangdiv &"'"
	end if

	sqlStr = "select"
	sqlStr = sqlStr & " divcd,divname,findurl,isUsing,isTenUsing,tel,returnURL"
	sqlStr = sqlStr & " from db_order.dbo.tbl_songjang_div with (nolock)"
	sqlStr = sqlStr & " where isusing='Y' " & sqlsearch

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		songjangname = rsget("divname")
	end if
	rsget.close

    getsongjangdivname=songjangname
End function

Sub drawOutmallSelectBox(selectBoxName, selectedId)
   dim tmp_str, sqlStr
%>
	<select class="select" name="<%=selectBoxName%>">
		<option value="" <% If selectedId = "" Then response.write " selected" %>>-선택-</option>
<%
		sqlStr = ""
		sqlStr = sqlStr & " SELECT m.sitename "
		sqlStr = sqlStr & " FROM db_order.dbo.tbl_order_master as m "
		sqlStr = sqlStr & " WHERE beadaldiv in (50, 51) "
		sqlStr = sqlStr & " and sitename <> '10x10_cs' "
		sqlStr = sqlStr & " GROUP BY sitename "
		sqlStr = sqlStr & " ORDER BY 1 "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF then
			rsget.Movefirst
			Do Until rsget.EOF
				If Lcase(selectedId) = Lcase(rsget("sitename")) Then
					tmp_str = " selected"
				End If
				response.write("<option value='"&rsget("sitename")&"' "&tmp_str&">"&rsget("sitename")&"</option>")
				tmp_str = ""
				rsget.MoveNext
			Loop
		End If
		rsget.close
	response.write("</select>")
End Sub

function drawvpnlistSelectBox(selectBoxName, selectedId, chgplg)
   dim tmp_str, sqlStr
%>
	<select name="<%=selectBoxName%>" <%= chgplg %>>
		<option value="" <% If selectedId = "" Then response.write " selected" %>>-선택-</option>
<%
        sqlStr = "select userid"
        sqlStr = sqlStr & " from [db_board].[dbo].[tbl_vpn_connect_log] with (nolock)"
        sqlStr = sqlStr & " where userid not in ('','test','test1')"
        sqlStr = sqlStr & " group by userid"
        sqlStr = sqlStr & " order by userid asc"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.EOF then
			rsget.Movefirst
			Do Until rsget.EOF
				If Lcase(selectedId) = Lcase(rsget("userid")) Then
					tmp_str = " selected"
				End If
				response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"</option>")
				tmp_str = ""
				rsget.MoveNext
			Loop
		End If
		rsget.close
	response.write("</select>")
End function

Function aspJsonStringEscape(text)
    If IsEmpty(text) Or text = "" Or IsNull(text) Then
        aspJsonStringEscape = text
        Exit Function
    End If

    Dim i, c, encoded, charcode
    ' Adapted from https://stackoverflow.com/q/2920752/1178314
    encoded = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
            'Case "'"
            '    encoded = encoded & "\'"
            Case """"
                encoded = encoded & "\"""
            Case "\"
                encoded = encoded & "\\"
            Case vbFormFeed
                encoded = encoded & "\f"
            Case vbLf
                encoded = encoded & "\n"
            Case vbCr
                encoded = encoded & "\r"
            Case vbTab
                encoded = encoded & "\t"
'            Case "<" ' This avoids breaking a <script> content, in case the string contains "<!--" or "<script" or "</script"
'               encoded = encoded & "\x3C"
            Case Else
'                charcode = AscW(c)
'                If charcode < 32 Or charcode > 127 Then
'                    encoded = encoded & GetJavascriptUnicodeEscapedChar(charcode)
'                Else
                    encoded = encoded & c
'                End If
        End Select
    Next

    aspJsonStringEscape = encoded
End Function

' Taken from https://stackoverflow.com/a/2243164/1178314
Function GetJavascriptUnicodeEscapedChar(charcode)
    charcode = Hex(charcode)
    GetJavascriptUnicodeEscapedChar = "\u" & String(4 - Len(charcode), "0") & charcode
End Function

' 과세구분      '/2022.02.23 한용민 생성
Function drawSelectBoxjungsan_gubun(jungsan_gubunName, jungsan_gubunValue, chplg)
%>
    <select class="select" name="<%= jungsan_gubunName %>" <%= chplg %>>
        <option value="">선택</option>
        <option value="일반과세" <% if jungsan_gubunValue="일반과세" then response.write "selected" %> >일반과세</option>
        <option value="간이과세" <% if jungsan_gubunValue="간이과세" then response.write "selected" %> >간이과세</option>
        <option value="원천징수" <% if jungsan_gubunValue="원천징수" then response.write "selected" %> >원천징수</option>
        <option value="면세" <% if jungsan_gubunValue="면세" then response.write "selected" %> >면세</option>
        <option value="영세(해외)" <% if jungsan_gubunValue="영세(해외)" then response.write "selected" %> >영세(해외)</option>
    </select>
<%
End Function

' 기타>>알림관리 / 알림타입      ' 2023.03.30 한용민 생성
function DrawNotificationType(selectBoxName,selectedId, chplg)
    dim tmp_str,query1

%>
    <select name="<%=selectBoxName%>" <%= chplg %>>
        <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
<%
    query1 = " select notificationType, notificationTypeName"
    query1 = query1 & " from db_partner.dbo.notificationType nt with (nolock)"
    query1 = query1 & " order by notificationSortNumber asc"

    'response.write query1 &"<Br>"
    rsget.CursorLocation = adUseClient
    rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

    if not rsget.EOF then
        do until rsget.EOF
        if Lcase(selectedId) = Lcase(rsget("notificationType")) then
            tmp_str = " selected"
        end if
        response.write("<option value='"&rsget("notificationType")&"' "&tmp_str&">" + rsget("notificationTypeName") + "</option>")
        tmp_str = ""
        rsget.MoveNext
        loop
    end if
    rsget.close

    response.write("</select>")
End function

' 물류센터 작업그룹	' 2018.03.26 한용민 생성
function DrawWorkgroup(selectBoxName, selectedId, changeFlag)
%>
   	<select name="<%= selectBoxName %>" <%= changeFlag %>>
	   	<option value="">작업그룹</option>
	   	<option value="A" <% if selectedId="A" then response.write "selected" %> >A (텐바이텐)</option>
	   	<option value="B" <% if selectedId="B" then response.write "selected" %> >B (텐바이텐)</option>
	   	<option value="C" <% if selectedId="C" then response.write "selected" %> >C (텐바이텐)</option>
	   	<option value="D" <% if selectedId="D" then response.write "selected" %> >D (텐바이텐)</option>
        <option value="E" <% if selectedId="E" then response.write "selected" %> >E (텐바이텐)</option>
	   	<option value="F" <% if selectedId="F" then response.write "selected" %> >F (텐바이텐)</option>
		<option value="K" <% if selectedId="K" then response.write "selected" %> >K (텐바이텐)</option>
		<option value="L" <% if selectedId="L" then response.write "selected" %> >L (텐바이텐)</option>
		<option value="N" <% if selectedId="N" then response.write "selected" %> >N (텐바이텐 - 단품출고)</option>
		<option value="M" <% if selectedId="M" then response.write "selected" %> >M (3PL - 단품출고)</option>
		<option value="J" <% if selectedId="J" then response.write "selected" %> >J (아이띵소 - 단품출고)</option>
		<option value="">==============</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %> >3 (3PL)</option>
		<option value="">==============</option>
	   	<option value="E" <% if selectedId="E" then response.write "selected" %> >E (텐바이텐 - EMS)</option>
        <option value="U" <% if selectedId="U" then response.write "selected" %> >U (텐바이텐 - UPS)</option>
		<option value="R" <% if selectedId="R" then response.write "selected" %> >R (텐바이텐 - KPACK)</option>
	   	<option value="G" <% if selectedId="G" then response.write "selected" %> >G (텐바이텐 - 군부대)</option>
		<option value="H" <% if selectedId="H" then response.write "selected" %> >H (텐바이텐 - 중국몰)</option>
	   	<option value="O" <% if selectedId="O" then response.write "selected" %> >O (텐바이텐 - 오프라인)</option>
	   	<option value="">==============</option>
	   	<option value="P" <% if selectedId="P" then response.write "selected" %> >P (29cm)</option>
	   	<option value="Q" <% if selectedId="Q" then response.write "selected" %> >Q (29cm)</option>
	   	<option value="">==============</option>
	   	<option value="T" <% if selectedId="T" then response.write "selected" %> >T (탐스슈즈)</option>
	   	<option value="">==============</option>
	   	<option value="I" <% if selectedId="I" then response.write "selected" %> >I (아이띵소)</option>
	   	<option value="">==============</option>
	   	<option value="Z" <% if selectedId="Z" then response.write "selected" %> >Z</option>
   	</select>
<%
end function

%>

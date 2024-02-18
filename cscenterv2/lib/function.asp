<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
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


'// 마진률계산
Function fnPercent(oup,inp,pnt)
	if oup=0 or isNull(oup) then exit function
	if inp=0 or isNull(inp) then exit function
	fnPercent = FormatNumber((1-(clng(oup)/clng(inp)))*100,pnt) & "%"
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
		if isNull(divcd) then
			DeliverDivCd2Nm = ""
			Exit function
		end if
		   if CStr(divcd) = "1" then
		    DeliverDivCd2Nm =  "한진택배"
		   elseif CStr(divcd) = "2" then
		    DeliverDivCd2Nm =  "현대택배"
		   elseif CStr(divcd) = "3" then
		    DeliverDivCd2Nm =  "대한통운"
		   elseif CStr(divcd) = "4" then
		    DeliverDivCd2Nm =  "CJ GLS"
		   elseif CStr(divcd) = "5" then
		    DeliverDivCd2Nm =  "이클라인"
		   elseif CStr(divcd) = "6" then
		    DeliverDivCd2Nm =  "HTH"
		   elseif CStr(divcd) = "7" then
		    DeliverDivCd2Nm =  "훼미리택배"
		   elseif CStr(divcd) = "8" then
		    DeliverDivCd2Nm =  "우체국"
		   elseif CStr(divcd) = "9" then
		    DeliverDivCd2Nm =  "(구)KGB"
		   elseif CStr(divcd) = "10" then
		    DeliverDivCd2Nm =  "아주택배"
		   elseif CStr(divcd) = "11" then
		    DeliverDivCd2Nm =  "오렌지택배"
		   elseif CStr(divcd) = "12" then
		    DeliverDivCd2Nm =  "한국택배"
		   elseif CStr(divcd) = "13" then
		    DeliverDivCd2Nm =  "옐로우캡"
		   elseif CStr(divcd) = "14" then
		    DeliverDivCd2Nm =  "나이스택배"
		   elseif CStr(divcd) = "15" then
		    DeliverDivCd2Nm =  "중앙택배"
		   elseif CStr(divcd) = "16" then
		    DeliverDivCd2Nm =  "주코택배"
		   elseif CStr(divcd) = "17" then
		    DeliverDivCd2Nm =  "트라넷택배"
		   elseif CStr(divcd) = "18" then
		    DeliverDivCd2Nm =  "로젠택배"
		   elseif CStr(divcd) = "19" then
		    DeliverDivCd2Nm =  "KGB특급택배"
		   elseif CStr(divcd) = "20" then
		    DeliverDivCd2Nm =  "KT로지스"
		   elseif CStr(divcd) = "21" then
		    DeliverDivCd2Nm =  "경동택배"
		   elseif CStr(divcd) = "22" then
		    DeliverDivCd2Nm =  "고려택배"
		   elseif CStr(divcd) = "23" then
		    DeliverDivCd2Nm =  "신세계SEDEX"
		   elseif CStr(divcd) = "24" then
		    DeliverDivCd2Nm =  "사가와"
		   elseif CStr(divcd) = "30" then
		    DeliverDivCd2Nm =  "이노지스"
		   elseif CStr(divcd) = "31" then
		    DeliverDivCd2Nm =  "천일택배"
		   elseif CStr(divcd) = "32" then
		    DeliverDivCd2Nm =  "사가와 임시"
		   elseif CStr(divcd) = "33" then
		    DeliverDivCd2Nm =  "호남택배"
		   elseif CStr(divcd) = "34" then
		    DeliverDivCd2Nm =  "대신화물택배"
		   elseif CStr(divcd) = "35" then
		    DeliverDivCd2Nm =  "CVSnet택배"
		   elseif CStr(divcd) = "90" then
		    DeliverDivCd2Nm =  "EMS"
		   elseif CStr(divcd) = "99" then
		    DeliverDivCd2Nm =  "기타"
		   end if

end function

function DeliverDivTrace(byval divcd)
	if isNull(divcd) then
		DeliverDivTrace = ""
		Exit function
	end if
		if CStr(divcd) = "1" then
			'한진택배
		    DeliverDivTrace =  "http://www.hanjinexpress.hanjin.net/customer/plsql/hddcw07.result?wbl_num="
		elseif CStr(divcd) = "2" then
			'현대택배
		    DeliverDivTrace =  "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo="
		elseif CStr(divcd) = "3" then
			'대한통운
		    DeliverDivTrace =  "http://www.doortodoor.co.kr/jsp/cmn/Tracking.jsp?QueryType=3&pTdNo="
		elseif CStr(divcd) = "4" then
			'CJ GLS
		    DeliverDivTrace =  "http://www.cjgls.co.kr/kor/service/service02_02.asp?slipno="
		elseif CStr(divcd) = "5" then
			'이클라인
		    DeliverDivTrace =  "http://www.sagawa-korea.co.kr/sub4/default2_2.asp?awbino="
		elseif CStr(divcd) = "6" then
			'HTH
		    DeliverDivTrace =  "http://cjhth.com/homepage/searchTraceGoods/SearchTraceDtdShtno.jhtml?dtdShtno="
		elseif CStr(divcd) = "7" then
			'훼미리택배
		    DeliverDivTrace =  "http://www.e-family.co.kr/member/delivery_search_view.jsp?item_no="
		elseif CStr(divcd) = "8" then
			'우체국
		    DeliverDivTrace =  "http://service.epost.go.kr/trace.RetrieveRegiPrclDeliv.postal?sid1="
		elseif CStr(divcd) = "9" then
			'(구)KGB
		    DeliverDivTrace =  "http://www.kgbls.co.kr/sub3/sub3_4_1.asp?f_slipno="
		elseif CStr(divcd) = "10" then
			'아주택배
		    DeliverDivTrace =  "http://www.ajuthankyou.com:8080/jsp/expr1/web_view.jsp?sheetno1="
		elseif CStr(divcd) = "11" then
			'오렌지택배
		    DeliverDivTrace =  ""
		elseif CStr(divcd) = "12" then
			'한국택배
		    DeliverDivTrace =  ""
		elseif CStr(divcd) = "13" then
			'옐로우캡
		    DeliverDivTrace =  "http://yellowcap.bizeye.co.kr/search.asp?slipno="
		elseif CStr(divcd) = "14" then
			'나이스택배
		    DeliverDivTrace =  ""
		elseif CStr(divcd) = "15" then
			'중앙택배
		    DeliverDivTrace =  ""
		elseif CStr(divcd) = "16" then
			'주코택배 - out
		    DeliverDivTrace =  ""
		elseif CStr(divcd) = "17" then
			'트라넷택배
		    DeliverDivTrace =  "http://www.etranet.co.kr/branch/chase/listbody.html?a_gb=center&a_cd=4&a_item=0&fr_slipno="
		elseif CStr(divcd) = "18" then
			'로젠택배
		    DeliverDivTrace =  "http://www.ilogen.com/customer/reserve_03_detail.asp?f_slipno="
		elseif CStr(divcd) = "19" then
			'KGB특급택배
		    DeliverDivTrace =  "http://www.kgbls.co.kr/sub3/sub3_4_1.asp?f_slipno="
		elseif CStr(divcd) = "20" then
			'KT로지스
		    DeliverDivTrace =  "http://218.153.4.42/customer/cus_trace_02.asp?searchMethod=I&invc_no="
		elseif CStr(divcd) = "21" then
			'경동택배
			DeliverDivTrace =  "http://insu.kdexp.com/insu/search.php?p_item="
		elseif CStr(divcd) = "22" then
			'고려택배
			DeliverDivTrace =  "http://www.gologis.com/delivery/s_search.asp?f_slipno="
		elseif CStr(divcd) = "23" then
			'신세계 SEDEX
			DeliverDivTrace =  "http://ptop.sedex.co.kr:8080/jsp/tr/detailSheet.jsp?iSheetNo="
		elseif CStr(divcd) = "24" then
			'사가와
		    DeliverDivTrace =  "http://www.sc-logis.co.kr/tracking/normal/default.asp?awblno="
		elseif CStr(divcd) = "99" then
		    DeliverDivTrace =  ""
		end if

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
		mwdivName = "특정"
	else

	end if
end function

public function GetJungsanGubunName(v)
    if v="B011" then
	    GetJungsanGubunName = "특정판매"
	elseif v="B012" then
	    GetJungsanGubunName = "업체특정"
	elseif v="B021" then
	    GetJungsanGubunName = "오프매입"
	elseif v="B022" then
	    GetJungsanGubunName = "매장매입"
	elseif v="B031" then
	    GetJungsanGubunName = "출고매입"
	elseif v="B032" then
	    GetJungsanGubunName = "센터매입"
	elseif v="B999" then
	    GetJungsanGubunName = "기타보정"
    else
        GetJungsanGubunName = v
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
		elseif v="100" then
			JumunMethodName="신용카드"
		elseif v="20" then
			JumunMethodName="실시간이체"
		elseif v="30" then
			JumunMethodName="포인트"
		elseif v="50" then
			JumunMethodName="외부몰"
		elseif v="90" then
			JumunMethodName="상품권"
		elseif v="110" then
			JumunMethodName="OK+신용"
	    elseif v="400" then
			JumunMethodName="핸드폰결제"
		end if
	end function

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





Sub drawSelectBoxOFFChargeDiv(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="2" <% if (selectedId="2") then response.write "selected" %> >텐바이텐 특정
		<option value="4" <% if (selectedId="4") then response.write "selected" %> >텐바이텐 매입
		<option value="5" <% if (selectedId="5") then response.write "selected" %> >매입출고 정산
		<option value="6" <% if (selectedId="6") then response.write "selected" %> >업체 특정
		<option value="8" <% if (selectedId="8") then response.write "selected" %> >업체 매입
	</select>
<%
End Sub


Sub drawSelectBoxOFFJungsanCommCD(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%=selectBoxName%>">
		<option value="">선택
		<option value="B011" <% if (selectedId="B011") then response.write "selected" %> >텐바이텐 특정
		<option value="B031" <% if (selectedId="B031") then response.write "selected" %> >매입출고 정산
		<option value="B012" <% if (selectedId="B012") then response.write "selected" %> >업체 특정
		<option value="B022" <% if (selectedId="B022") then response.write "selected" %> >업체 매입
		<!-- option value="B021" <% if (selectedId="B021") then response.write "selected" %> >오프 매입 -->
	</select>
<%
End Sub

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

Sub drawSelectBoxOffShop(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"

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

Sub drawSelectBoxOffShopNot000(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where isusing='Y' "
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"

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
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   query1 = query1 & " where 1=1"
   query1 = query1 & " and userid<>'streetshop000'"
   query1 = query1 & " and userid<>'streetshop800'"
   query1 = query1 & " and userid<>'streetshop870'"
   query1 = query1 & " and userid<>'streetshop700'"

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

Sub drawSelectBoxOffShopWith000(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
   ''query1 = query1 & " where userid<>'cafe002' "
   ''query1 = query1 & " where isusing='Y' "

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
   query1 = " select u.userid,u.shopname from [db_shop].[dbo].tbl_shop_user u, [db_shop].[dbo].tbl_shop_designer d"
   query1 = query1 & " where u.isusing='Y' "
   query1 = query1 & " and u.userid=d.shopid"
   query1 = query1 & " and d.makerid='" + session("ssBctID") + "'"
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
    %>
    <input type="text" class="text" name="<%= selectBoxName %>" value="<%= selectedId %>" size="20" >
    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'<%= selectBoxName %>');" >
    &nbsp;
    <%
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
	 <option value='101' <%if selectedId="101" then response.write " selected"%>>오프샾</option>
	 <option value='111' <%if selectedId="111" then response.write " selected"%>>오프샾점장</option>
	 <option value='112' <%if selectedId="112" then response.write " selected"%>>오프샾부점장</option>
	 <option value='509' <%if selectedId="509" then response.write " selected"%>>오프매출조회</option>
	 <option value='201' <%if selectedId="201" then response.write " selected"%>>Zoom</option>
	 <option value='301' <%if selectedId="301" then response.write " selected"%>>College</option>
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
     <option value='7' <%if selectedId="7" then response.write " selected"%>>업체착불배송</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>업체개별배송</option>
   </select>
   <%
end Sub

Sub drawSelectBoxWriter(byval writer)
	dim buf
	buf = "<select class='select' name='writer'>" + VbCrlf
	buf = buf + "<option selected value=''>선택</option>" + VbCrlf
	if writer="chwoo" then
		buf = buf + "<option value='chwoo' selected>이창우</option>" + VbCrlf
	else
		buf = buf + "<option value='chwoo' >이창우</option>" + VbCrlf
	end if

	if writer="winnie" then
		buf = buf + "<option value='winnie' selected>최은희</option>" + VbCrlf
	else
		buf = buf + "<option value='winnie' >최은희</option>" + VbCrlf
	end if

	if writer="livearc" then
    	buf = buf + "<option value='livearc' selected>백우현</option>" + VbCrlf
    else
    	buf = buf + "<option value='livearc' >백우현</option>" + VbCrlf
	end if

	if writer="moon" then
    	buf = buf + "<option value='moon' selected>이문재</option>" + VbCrlf
    else
    	buf = buf + "<option value='moon' >이문재</option>" + VbCrlf
	end if

    if writer="icommang" then
    	buf = buf + "<option value='icommang' selected>서동석</option>" + VbCrlf
    else
    	buf = buf + "<option value='icommang' >서동석</option>" + VbCrlf
	end if

    if writer="migi" then
    	buf = buf + "<option value='migi' selected>박미경</option>" + VbCrlf
    else
    	buf = buf + "<option value='migi' >박미경</option>" + VbCrlf
	end if

    if writer="mizzle" then
    	buf = buf + "<option value='mizzle' selected>최은미</option>" + VbCrlf
    else
    	buf = buf + "<option value='mizzle' >최은미</option>" + VbCrlf
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
    for i=2002 to 2017
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
    for i=2002 to 2017
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


Sub DrawOneDateBox(byval yyyy1,mm1,dd1)
	dim buf,i

	buf = "<select class='select' name='yyyy1'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2002 to 2017
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
    for i=2002 to 2017
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
    for i=2002 to 2017
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
    for i=2002 to 2017
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
    for i=2002 to 2017
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
   query1 = " select p.id, p.company_name from"
   query1 = query1 + " [db_partner].[dbo].tbl_partner p "
   query1 = query1 + " where ((p.userdiv <= 9) or (p.userdiv='111') or (p.userdiv='301'))"
   query1 = query1 + " and isusing='Y' and part_sn in('11','13','14','15','16')"
   query1 = query1 + " order by company_name asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("id")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" + rsget("id") + "' "&tmp_str&">" + db2html(rsget("company_name")) + " (" + rsget("id") + ")</option>")
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
    buf = buf + " >특정"
    response.write buf
end Sub

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
   query1 = " select code_large, code_nm from " & TABLE_CATEGORY_LARGE & " "
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

Sub DrawSelectBoxCategoryLarge(byval selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from " & TABLE_CATEGORY_LARGE & " "
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
   query1 = " select code_large, code_nm from " & TABLE_CATEGORY_LARGE & " "
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

	sqlStr = " select top 100 id from [db_partner].[dbo].tbl_partner "
	sqlStr = sqlStr + " where userdiv='999' "
	sqlStr = sqlStr + " and id<>'10x10' "
	sqlStr = sqlStr + " and isusing='Y' "
	sqlStr = sqlStr + " order by id "
	rsget.Open sqlStr,dbget,1

	response.write "<select class='select' name=" + selectBoxName + ">"
	if  not rsget.EOF  then
	        rsget.Movefirst
                do until rsget.EOF
        		if extsitename=rsget("id") then
        			response.write("<option value=" + rsget("id") + " selected>" + rsget("id") + "</option>")
        		else
        			response.write("<option value=" + rsget("id") + ">" + rsget("id") + "</option>")
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
   query1 = " select code_large, code_nm from " & TABLE_CATEGORY_LARGE & " "
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
   query1 = " select code_large, code_nm from " & TABLE_CATEGORY_LARGE & " "
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
   query1 = " select top 100 divcd,divname from " & TABLE_SONGJANG_DIV & " where isUsing='Y' "
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
  	 <option value='95' <%if selectedId="95" then response.write " selected"%>>사용안함</option>
   </select>
  <%
End Sub


Sub DrawChulgoDiv(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> >선택</option>
	<option value='1' <% if selectedId="1" then response.write " selected" %> >매입-&gt;매입</option>
	<option value='2' <% if selectedId="2" then response.write " selected" %> >특정-&gt;매입</option>
	<option value='2' <% if selectedId="2" then response.write " selected" %> >특정-&gt;특정</option>
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
	<option value='W' <% if selectedId="W" then response.write " selected" %> >특정</option>
	<option value='U' <% if selectedId="U" then response.write " selected" %> >업체배송</option>
	</select>
<%
end sub

Sub DrawBrandOffMWCombo(selectedname,selectedId)
%>
	<select class='select' name="<%= selectedname %>" >
	<option value='' <% if selectedId="" then response.write " selected" %> ></option>
	<option value='M' <% if selectedId="M" then response.write " selected" %> >매입</option>
	<option value='W' <% if selectedId="W" then response.write " selected" %> >특정</option>
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
	<!-- <option value='상호저축' <% if selectedId="상호저축" then response.write " selected" %> >상호저축</option> -->
	<!-- <option value='수출입' <% if selectedId="수출입" then response.write " selected" %> >수출입</option> -->
	<option value='수협' <% if selectedId="수협" then response.write " selected" %> >수협</option>
	<option value='신한' <% if selectedId="신한" then response.write " selected" %> >신한</option>
	<option value='외환' <% if selectedId="외환" then response.write " selected" %> >외환</option>
	<option value='우리' <% if selectedId="우리" then response.write " selected" %> >우리</option>
	<option value='우체국' <% if selectedId="우체국" then response.write " selected" %> >우체국</option>
	<option value='전북' <% if selectedId="전북" then response.write " selected" %> >전북</option>
	<option value='제일' <% if selectedId="제일" then response.write " selected" %> >제일</option>
	<option value='조흥' <% if selectedId="조흥" then response.write " selected" %> >조흥</option>
	<option value='평화' <% if selectedId="평화" then response.write " selected" %> >평화</option>
	<option value='하나' <% if selectedId="하나" then response.write " selected" %> >하나</option>
	<!-- <option value='한국' <% if selectedId="한국" then response.write " selected" %> >한국</option> -->
	<!-- <option value='한미' <% if selectedId="한미" then response.write " selected" %> >한미</option> -->
	<option value='시티' <% if selectedId="시티" then response.write " selected" %> >시티</option>
	<option value='홍콩샹하이' <% if selectedId="홍콩샹하이" then response.write " selected" %> >홍콩샹하이</option>
	<option value='ABN암로은행' <% if selectedId="ABN암로은행" then response.write " selected" %> >ABN암로은행</option>
	<option value='UFJ은행' <% if selectedId="UFJ은행" then response.write " selected" %> >UFJ은행</option>
	<option value='신협' <% if selectedId="신협" then response.write " selected" %> >신협</option>
    <option value='제주' <% if selectedId="제주" then response.write " selected" %> >제주</option>
    <option value='현대스위스상호저축은행' <% if selectedId="현대스위스상호저축은행" then response.write " selected" %> >현대스위스상호저축은행</option>
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

function getUserLevelStr(iuserlevel)
    getUserLevelStr = iuserlevel

    if (iuserlevel=0) then
        getUserLevelStr = "옐로우"
    elseif (iuserlevel=5) then
        getUserLevelStr = "오렌지"
    elseif (iuserlevel=1) then
        getUserLevelStr = "그린"
    elseif (iuserlevel=2) then
        getUserLevelStr = "블루"
    elseif (iuserlevel=3) then
        getUserLevelStr = "VIP실버"
    elseif (iuserlevel=4) then
        getUserLevelStr = "VIP골드"
    elseif (iuserlevel=7) then
        getUserLevelStr = "STAFF"
    end if
end function

function getUserLevelColor(iuserlevel)
    getUserLevelColor = "#000000"

    if (iuserlevel=0) then
        getUserLevelColor = "#BBBB33"
    elseif (iuserlevel=5) then
        getUserLevelColor = "#FF6666"
    elseif (iuserlevel=1) then
        getUserLevelColor = "#66BB66"
    elseif (iuserlevel=2) then
        getUserLevelColor = "#0000FF"
    elseif (iuserlevel=3) then
        getUserLevelColor = "#A4A8AA"
    elseif (iuserlevel=4) then
        getUserLevelColor = "#E5CC57"
    elseif (iuserlevel=7) then
        getUserLevelColor = "#000000"
    end if
end function

Sub DrawUserLevelCombo(selectedname,selectedId)
%>
    <select class='select' name="<%= selectedname %>">
	    <option value="" <% if selectedId="" then response.write " selected" %> >전체</option>
	    <option value="5" <% if selectedId="5" then response.write " selected" %> >오렌지
	    <option value="0" <% if selectedId="0" then response.write " selected" %> >옐로우
	    <option value="1" <% if selectedId="1" then response.write " selected" %> >그린
	    <option value="2" <% if selectedId="2" then response.write " selected" %> >블루
	    <option value="3" <% if selectedId="3" then response.write " selected" %> >VIP실버
	    <option value="4" <% if selectedId="4" then response.write " selected" %> >VIP골드
	    <option value="7" <% if selectedId="7" then response.write " selected" %> >STAFF
	</select>
<%
end Sub
%>


<%

Sub drawSelectBoxSellYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >판매</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >일시품절</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >품절</option>
   <option value="YS" <% if selectedId="YS" then response.write "selected" %> >판매+일시품절</option>
   </select>
   <%
End Sub

Sub drawSelectBoxUsingYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용함</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >사용안함</option>
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
   <option value="MW" <% if selectedId="MW" then response.write "selected" %> >매입+특정</option>
   <option value="W" <% if selectedId="W" then response.write "selected" %> >특정</option>
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
   <option value="10" <% if selectedId="10" then response.write "selected" %> >온라인등록상품</option>
   <option value="90" <% if selectedId="90" then response.write "selected" %> >오프전용상품</option>
   <option value="70" <% if selectedId="70" then response.write "selected" %> >각종부자재</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsOverSeaYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >안함</option>
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
					fnColor = "<Font color=#808080>특정</font>"
				Case "U"
					fnColor = "<Font color=#5080F0>업체</font>"
			end Select
		Case "tx"
			if str="Y" then
				fnColor = "<Font color=#808080>과세</font>"
			else
				fnColor = "<Font color=#F08050>면세</font>"
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
   rsget.Open " select code_nm from " & TABLE_CATEGORY_LARGE & " where code_large='" & cdl & "'",dbget,1
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

	rsget_CS.Open " select findurl from db_order.dbo.tbl_songjang_div where divcd=" & CStr(isongjangdiv) & "",dbget_CS,1
	if Not(rsget_CS.EOF or rsget_CS.BOF) then
		fnGetSongjangURL = db2html(rsget_CS(0))
	else
		fnGetSongjangURL = ""
	end if
	rsget_CS.Close
end function


function fnGetOffCurrencyUnit(byval shopid,byRef CurrencyUnit, byRef CurrencyChar, byRef ExchangeRate)
    Dim sqlStr
    sqlStr = "select U.CurrencyUnit,X.ExchangeRate,X.CurrencyChar from db_shop.dbo.tbl_shop_User U"
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
        CurrencyUnit = "KRW"
        CurrencyChar = "원"
        ExchangeRate = 1
    end if
    rsget.Close
end function

%>

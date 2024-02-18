<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.01.01 이상구 생성
'           2016.12.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/db/dbTPLHelper.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/tempOrderCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<%

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim mode, arrOutMallOrderSerial, dummyseqarr, splitedOutMallOrderSerial
dim OutMallOrderSerial, tenOrderSerial, isexist
dim sqlStr
dim i, j, k
Dim otmpOrder
dim IsForeignDLV, countryCode, ErrMsg
dim buf_totcost, buf_totvat, buf_totCpnNotAppliedcost, buf_iitemmakerid
dim ixsiteOrderSerial
dim iid, orderserial


mode = requestCheckVar(request("mode"),32)
arrOutMallOrderSerial = request("arrOutMallOrderSerial")

Function RegExResults(strTarget, strPattern)
	dim regEx
    Set regEx = New RegExp
    regEx.Pattern = strPattern
    regEx.Global = true
    RegExResults = regEx.Test(strTarget)
    Set regEx = Nothing
End Function

Function RemoveNull(str)
	Dim re
	Set re = New RegExp
	re.Global = True
	re.Pattern = "[\0]"   ' should see backslash zero inside the square braces
	RemoveNull = re.Replace(str,"")
	Set re = Nothing
End Function

if (mode = "add") then
	''response.write "TEST중<br />"
	''response.write arrOutMallOrderSerial & "<br />"
	''response.end

	if (Left(arrOutMallOrderSerial, 1) = ",") then
		arrOutMallOrderSerial = Mid(arrOutMallOrderSerial, 2, Len(arrOutMallOrderSerial))
	end if


	'// ========================================================================
	if (RegExResults(arrOutMallOrderSerial, "^[0-9,a-zA-Z-]+$") <> True) then
		response.write "<script>alert('해킹오류. 관리자 문의 요망')</script>"
		response.write "해킹오류. 관리자 문의 요망 : (" & arrOutMallOrderSerial & ")"
		dbget_TPL.close() : dbget.close()	:	response.End
	end if

	dummyseqarr = arrOutMallOrderSerial
	dummyseqarr = Replace(dummyseqarr, ", ", ",")
	dummyseqarr = Replace(dummyseqarr, ",", "','")
	dummyseqarr = "'"&dummyseqarr&"'"


	'// ========================================================================
    ''이미 전송된 주문건인지 check
    sqlStr = " select top 1 T.OutMallOrderSerial, m.orderserial from " + VbCrlf
    sqlStr = sqlStr + " db_threepl.dbo.tbl_xSite_TMPOrder T" + VbCrlf
    sqlStr = sqlStr + " 	Join db_threepl.dbo.tbl_order_master m" + VbCrlf
    sqlStr = sqlStr + " 	on T.OutMallOrderSerial=m.authcode" + VbCrlf
    sqlStr = sqlStr + " 	and m.sitename=T.sellSite" + VbCrlf
    sqlStr = sqlStr + " where T.OutMallOrderSerial in (" & dummyseqarr & ") " + VbCrlf

    rsget_TPL.Open sqlStr,dbget_TPL,1
    isexist = (not rsget_TPL.EOF)
    if (isexist = true) then
	    OutMallOrderSerial 		= rsget_TPL("OutMallOrderSerial")
	    tenOrderSerial          = rsget_TPL("orderserial")
	end if
    rsget_TPL.Close

	if (isexist = True) then
		response.write "<script>alert('주문 중복입력. 관리자 문의 요망')</script>"
		response.write "주문 중복입력. 관리자 문의 요망 : (" & OutMallOrderSerial & "," & tenOrderSerial & ")"
		dbget_TPL.close() : dbget.close()	:	response.End
	end if


	'// ========================================================================
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_threepl.dbo.tbl_XSite_TMporder "
	sqlStr = sqlStr & " WHERE  OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and datediff(m, regdate, getdate()) > 3 "
    rsget_TPL.Open sqlStr,dbget_TPL,1
	If rsget_TPL("cnt") > 0 Then
		response.write "<script>alert('3개월이상 지난 주문건은 입력 하실 수 없습니다.')</script>"
		response.write "3개월이상 지난 주문건은 입력 하실 수 없습니다"
		rsget_TPL.Close : dbget_TPL.close() : dbget.close()	:	response.End
	End If
	rsget_TPL.Close


	'// ========================================================================
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_threepl.dbo.tbl_XSite_TMporder "
	sqlStr = sqlStr & " WHERE  OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and (prdcode is NULL or orderSerial is not NULL) "
    rsget_TPL.Open sqlStr,dbget_TPL,1
	If rsget_TPL("cnt") > 0 Then
		response.write "<script>alert('상품매칭 안된 주문 또는 기입력된 주문건이 있습니다.')</script>"
		response.write "상품매칭 안된 주문 또는 기입력된 주문건이 있습니다."
		rsget_TPL.Close : dbget_TPL.close() : dbget.close()	:	response.End
	End If
	rsget_TPL.Close


	splitedOutMallOrderSerial = split(arrOutMallOrderSerial,",")
	For j = LBound(splitedOutMallOrderSerial) to UBound(splitedOutMallOrderSerial)
		ixsiteOrderSerial = Trim(splitedOutMallOrderSerial(j))
		if (ixsiteOrderSerial<>"") then
			set otmpOrder = new CTPLTempOrder
			otmpOrder.FPageSize = 200
			otmpOrder.FCurrPage = 1
			otmpOrder.FRectOutMallOrderSerial = ixsiteOrderSerial
			otmpOrder.getOnlineTmpOrderRealInputList()

			if (otmpOrder.FResultCount > 0) then
				rw otmpOrder.FItemList(0).FOutMallOrderSerial

				countryCode = otmpOrder.FItemList(0).fcountryCode
				if countryCode="" then ucase(countryCode)="KR"
				IsForeignDLV = (ucase(countryCode)<>"KR")

				ErrMsg = "[001]"

				dbget_TPL.beginTrans
        		'주문입력(마스터)
        		sqlStr = "select * from [db_threepl].[dbo].tbl_order_master where 1=0"
        		rsget_TPL.Open sqlStr,dbget_TPL,1,3
        		rsget_TPL.AddNew
        		rsget_TPL("orderserial") = Left(Left(otmpOrder.FItemList(0).FSellSite,2)&otmpOrder.FItemList(0).FOutMallOrderSerial ,11)

				rsget_TPL("reqemail") = otmpOrder.FItemList(0).forderemail
				rsget_TPL("jumundiv") = "5"
        		rsget_TPL("userid") = ""
        		rsget_TPL("ipkumdiv") = "1"
        		rsget_TPL("accountname") = ""
        		rsget_TPL("accountdiv") = "50"
        		rsget_TPL("authcode") = ixsiteOrderSerial
        		rsget_TPL("sitename") = otmpOrder.FItemList(0).FSellSite
        		rsget_TPL("DlvcountryCode") = countryCode
        		rsget_TPL("beadaldiv") = otmpOrder.FItemList(0).Fbeadaldiv
				rsget_TPL("tplcompanyid") = otmpOrder.FItemList(0).Fcompanyid
        		rsget_TPL.update
        		iid = rsget_TPL("idx")
        		rsget_TPL.close

        		orderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),1,256)
        		orderserial = orderserial & Format00(10,Right(CStr(iid),10))

				if Err then
					dbget.RollBackTrans
					response.write ErrMsg & Err.Description
					response.end
				else
					ErrMsg = "[002]"
				end if

        		sqlStr = "update M" & vbCrlf
				sqlStr = sqlStr + " set orderserial='" + CStr(orderserial) + "'," & vbCrlf
				sqlStr = sqlStr + " accountname='" + html2db(otmpOrder.FItemList(0).FOrderName) + "'," & vbCrlf
				sqlStr = sqlStr + " totalsum=0," & vbCrlf
				sqlStr = sqlStr + " ipkumdiv='4'," & vbCrlf
				sqlStr = sqlStr + " ipkumdate=getdate()," & vbCrlf
				sqlStr = sqlStr + " regdate=getdate()," & vbCrlf
				sqlStr = sqlStr + " buyname='" + html2db(otmpOrder.FItemList(0).FOrderName) + "'," & vbCrlf
				sqlStr = sqlStr + " buyphone='" + replace(otmpOrder.FItemList(0).FOrderTelNo,"'","") + "'," & vbCrlf
				sqlStr = sqlStr + " buyhp='" + replace(otmpOrder.FItemList(0).FOrderHpNo,"'","") + "'," & vbCrlf
				sqlStr = sqlStr + " buyemail=''," & vbCrlf
				sqlStr = sqlStr + " reqname='" + html2db(otmpOrder.FItemList(0).FReceiveName) + "'," & vbCrlf

				if ucase(countryCode)="KR" then
					sqlStr = sqlStr + " reqzipcode='" + Trim(otmpOrder.FItemList(0).FReceiveZipCode) + "'," & vbCrlf
				else
            		sqlStr = sqlStr + " reqzipcode='00000'," & vbCrlf
        		end if

				sqlStr = sqlStr + " reqaddress='" + TRIM(html2db(otmpOrder.FItemList(0).FReceiveAddr2)) + "'," & vbCrlf
				sqlStr = sqlStr + " reqphone='" + replace(otmpOrder.FItemList(0).FReceiveTelNo,"'","") + "'," & vbCrlf
				sqlStr = sqlStr + " reqhp='" + replace(otmpOrder.FItemList(0).FReceiveHpNo,"'","") + "'," & vbCrlf
				sqlStr = sqlStr + " comment='" + replace(TRIM(html2db(otmpOrder.FItemList(0).Fdeliverymemo)), "'", "") + "'," & vbCrlf
				sqlStr = sqlStr + " discountrate=1," & vbCrlf
				sqlStr = sqlStr + " subtotalprice=0," & vbCrlf
				sqlStr = sqlStr + " reqzipaddr='" + html2db(otmpOrder.FItemList(0).FReceiveAddr1) + "'" & vbCrlf
				sqlStr = sqlStr + " From [db_threepl].[dbo].tbl_order_master M" & vbCrlf
				sqlStr = sqlStr + " where idx=" + CStr(iid)
        		dbget_TPL.Execute sqlStr

				if Err then
					dbget_TPL.RollBackTrans
					response.write ErrMsg & Err.Description
					response.end
				else
					ErrMsg = "[003]"
				end if

				For i=0 to otmpOrder.FResultCount-1
					sqlStr = "insert into [db_threepl].[dbo].tbl_order_detail(masteridx, orderserial, prdcode, itemid," & vbCrlf
					sqlStr = sqlStr + "itemoption, itemno, itemcost, itemvat, mileage, reducedPrice, " & vbCrlf
					sqlStr = sqlStr + "orgitemcost,itemcostcouponnotApplied,bonuscouponidx,buycashcouponNotApplied, " & vbCrlf
					sqlStr = sqlStr + "itemname,itemoptionname,makerid,buycash," & vbCrlf
					sqlStr = sqlStr + "vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,requiredetail)" & vbCrlf
					sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
					sqlStr = sqlStr + " '" + orderserial + "'," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(otmpOrder.FItemList(i).Fprdcode) + "'," & vbCrlf
					sqlStr = sqlStr + " " + CStr(otmpOrder.FItemList(i).FmatchItemID) + "," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(otmpOrder.FItemList(i).FmatchItemOption) + "'," & vbCrlf
					sqlStr = sqlStr + " " + CStr(otmpOrder.FItemList(i).FItemOrderCount) + "," & vbCrlf
					sqlStr = sqlStr + " " + CStr(0) + "," & vbCrlf  '' buf_sellcash => buf_CpnNotAppliedSellcash 변경
					sqlStr = sqlStr + " 0," & vbCrlf
					sqlStr = sqlStr + " 0," & vbCrlf
					sqlStr = sqlStr + " " + CStr(0) + "," & vbCrlf               '' reducedPrice
					sqlStr = sqlStr + " 0," & vbCrlf   ''buf_onlyoptaddprice 추가 2015/05/18
					sqlStr = sqlStr + " " + CStr(0) + "," & vbCrlf
				    sqlStr = sqlStr + "NULL,"
					sqlStr = sqlStr + " 0," & vbCrlf
					sqlStr = sqlStr + " '" + RemoveNull(replace(CStr(otmpOrder.FItemList(i).FmatchItemName), "'", "")) + "'," & vbCrlf
					sqlStr = sqlStr + " '" + replace(CStr(otmpOrder.FItemList(i).FmatchItemOptionName), "'", "") + "'," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(otmpOrder.FItemList(i).Fbrandid) + "'," & vbCrlf
					sqlStr = sqlStr + " 0," & vbCrlf
					sqlStr = sqlStr + " 'Y'," & vbCrlf
					sqlStr = sqlStr + " ''," & vbCrlf
					sqlStr = sqlStr + " ''," & vbCrlf
					sqlStr = sqlStr + " ''," & vbCrlf
					sqlStr = sqlStr + " ''," & vbCrlf
					sqlStr = sqlStr + " ''," & vbCrlf
					sqlStr = sqlStr + " '" + replace(CStr(otmpOrder.FItemList(i).FrequireDetail),"'","''") + "'" & vbCrlf
					sqlStr = sqlStr + " )"
					dbget_TPL.Execute sqlStr

					if Err then
						dbget_TPL.RollBackTrans
						response.write ErrMsg & Err.Description
						response.end
					else
						ErrMsg = "[003.1]"
					end if
				next

				if Err then
					dbget_TPL.RollBackTrans
					response.write ErrMsg & Err.Description
					response.end
				else
					dbget_TPL.CommitTrans
					rw "["&orderserial&"]"
				end if

				sqlStr = "update [db_threepl].[dbo].tbl_order_master" & vbCrlf
				sqlStr = sqlStr + " set totalvat = 0" & vbCrlf
				sqlStr = sqlStr + " ,totalsum = " + CStr(0) + "" & vbCrlf  '' buf_totcost=>buf_totCpnNotAppliedcost
				sqlStr = sqlStr + " ,subtotalprice = " + CStr(0) + "" & vbCrlf
				sqlStr = sqlStr + " ,subtotalPriceCouponNotApplied = " + CStr(0) + "" & vbCrlf ''수정
				sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
				'response.write sqlStr & "<BR>"
				dbget_TPL.Execute sqlStr

				'' Flag update
				sqlStr = " update db_threepl.dbo.tbl_xSite_TMPOrder"
				sqlStr = sqlStr & " set matchState='O'"
				sqlStr = sqlStr & " ,OrderSerial='"&orderserial&"'"
				sqlStr = sqlStr & " where OutMallorderSerial='"&ixsiteOrderSerial&"'"
				sqlStr = sqlStr & " and matchState='I'"
				''rw sqlStr
				dbget_TPL.Execute sqlStr

			end if
		end if
	next


	response.write dummyseqarr & "<br />"
else
	response.write "잘못된 접근입니다."
	response.end
end if

%>
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

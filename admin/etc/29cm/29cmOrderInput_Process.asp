<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 29cm주문입력
' History : 2015.05.27 서동석 생성
'			2016.07.14 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbagirlopen.asp" -->
<!-- #include virtual="/lib/db/dbagirlHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/aGirlOrderCls.asp"-->
<%
'response.write "서팀문의요망"
'response.end

Dim cksel : cksel = request("cksel")

cksel = Trim(cksel)&","
cksel = split(cksel,",")

Dim oneAgirlOrder,i,k, cnt, sqlStr, paramInfo, brdSeq,retVal, retparmaInfo, AlreadyRegedExsiteOrder

if (application("Svr_Info")	= "Dev") then
    brdSeq = 291
else
    brdSeq = 314
end if

Dim oAgirlOrder, distinctrealcnt, maybeCnt, tmpItemList, ErrMsg, orderserial, iid, buf_totcost, buf_totvat, buf_orgprice
dim buf_sellcash, buf_sellvat, buf_mileage, buf_sellcount, buf_itemdiv, buf_iitemname, buf_sailbuycash, buf_sailsellcash
dim buf_iitemoptionname, buf_iitembuycash, buf_iitembuyvat, buf_onlyitembuycash, buf_onlyoptaddbuyprice, buf_orgsuplycash
dim buf_iitemmakerid, buf_iitemvatinclude, buf_deliverytype , buf_mwdiv, buf_sailyn
dim mayOrderDate, t_upchebeasong

For k=0 to UBound(cksel)
    oneAgirlOrder = Trim(cksel(k))
    
    if (oneAgirlOrder<>"") then
rw oneAgirlOrder
        ''전체상품 다 링크 되었는지 Check /텐바이텐 상품 존재 여부 Check
        sqlStr = "db_agirlOrder.[dbo].[usp_Back_LinkMall_OpenOrder_CheckByOrderSerial_TEN]"
        
        paramInfo = Array(Array("@RETURN_VALUE"	, adInteger	, adParamReturnValue	,		, 0) _
            ,Array("@partnerSeq" 		, adInteger	, adParamInput			,   	, 6) _
            ,Array("@brandSeq" 		, adInteger	, adParamInput			,32   	, brdSeq)	_
			,Array("@OrderSerial"			, adVarchar	, adParamInput			,13		, oneAgirlOrder)	_
		)
        
        retparmaInfo = dbagirl_fnExecSPOutput(sqlStr, paramInfo)
        
        retVal   = GetValue(retparmaInfo, "@RETURN_VALUE")
        
        if (retVal<>0) then
            response.write "<script>alert('상품 미매칭 내역 존재:"&Replace(oneAgirlOrder,"'","")&"')</script>"
            response.end
        end if
        
        ''이미 입력된 주문인지 Check
        AlreadyRegedExsiteOrder = false
        sqlStr = " select count(orderserial) as cnt"
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
        sqlStr = sqlStr + " where datediff(d,regdate,getdate())<61"
        sqlStr = sqlStr + " and sitename='29cm'"
        sqlStr = sqlStr + " and authcode='" + oneAgirlOrder + "'"
    
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
            AlreadyRegedExsiteOrder = rsget("cnt")>0
        rsget.Close
    
        if (AlreadyRegedExsiteOrder) then
            response.write "<script>alert('이미 전송된 주문 번호 입니다.- 관리자 문의 요망 Site : 29cm, 원 주문번호 : " + CStr(oneAgirlOrder) + "')</script>"
    		response.write "<script>history.back();</script>"
    		dbget.close() : response.End
    	end if

        ''주문내역 read
        Set oAgirlOrder = new aGirlOrder
        oAgirlOrder.getAgirlOneOrder(oneAgirlOrder)

        if (oAgirlOrder.FResultCount<1) then
            response.write "<script>alert('주문 내역이 올바르지 않습니다. " + CStr(oneAgirlOrder) + "')</script>"
    		response.write "<script>history.back();</script>"
    		dbget.close() : response.End
        end if
        
        '''상품Check
        tmpItemList = ""
        maybeCnt    = 0
        
''        For i=0 to oAgirlOrder.FResultCount-1
''            rw i&"["&oAgirlOrder.FItemList(i).FpartnerItemID&"]"&IsNULL(oAgirlOrder.FItemList(i).FpartnerItemID)
''            rw oAgirlOrder.FItemList(i).FItemSeq
''        NExt
''        response.end
        
        For i=0 to oAgirlOrder.FResultCount-1
            if oAgirlOrder.FItemList(i).FItemSeq<>0 then
                If Not IsNULL(oAgirlOrder.FItemList(i).FpartnerItemID) THEN
                    if tmpItemList="" then
                        tmpItemList = CStr(oAgirlOrder.FItemList(i).FpartnerItemID)
                    else
                        tmpItemList = tmpItemList +","+ CStr(oAgirlOrder.FItemList(i).FpartnerItemID)
                    end if
                    maybeCnt = maybeCnt +1
                ENd IF
            end if
        Next
        
        if (tmpItemList="") then
            response.write "<script>alert('상품 매칭 내역 없음')</script>"
            response.end
        end if
        
        sqlStr = "select count(itemid) as cnt from [db_item].[dbo].tbl_item"
    	sqlStr = sqlStr + " where itemid in (" + tmpItemList + ")"
rw sqlStr    	
    	rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    	    distinctrealcnt = rsget("cnt")
    	rsget.Close
    
    	'if (maybeCnt<>distinctrealcnt) then
    	'    response.write "<script>alert('알수 없는 상품 코드가 있습니다. " + CStr(oneAgirlOrder) + "')</script>"
    	'	response.write "<script>history.back();</script>"
    	'	dbget.close() : response.End
    	'end if
    	
''On Error Resume Next
ErrMsg = "[001]"	
        ''주문 입력.
        dbget.beginTrans
        
        sqlStr = "select * from [db_order].[dbo].tbl_order_master where 1=0"
    	rsget.Open sqlStr,dbget,1,3
    	rsget.AddNew
    	rsget("orderserial") = ""
    	rsget("jumundiv") = "5"
    	rsget("userid") = ""
    	rsget("ipkumdiv") = "1"
    	rsget("accountname") = ""
    	rsget("accountdiv") = "50"
    	rsget("authcode") = oneAgirlOrder
    	rsget("sitename") = "29cm"
    	rsget.update
    	iid = rsget("idx")
    	rsget.close
    
    	orderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
    	orderserial = orderserial & Format00(5,Right(CStr(iid),5))
    	
        if Err then
            dbget.RollBackTrans
            response.write ErrMsg & Err.Description
            response.end
        else
            ErrMsg = "[002]"
        end if
        
        sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
        sqlStr = sqlStr + " set orderserial='" + CStr(orderserial) + "'," & vbCrlf
        sqlStr = sqlStr + " accountname='" + html2db(oAgirlOrder.FItemList(0).FOrderName) + "'," & vbCrlf
        sqlStr = sqlStr + " totalsum=0," & vbCrlf
        sqlStr = sqlStr + " ipkumdiv='4'," & vbCrlf
        sqlStr = sqlStr + " ipkumdate=getdate()," & vbCrlf
        sqlStr = sqlStr + " regdate=getdate()," & vbCrlf
        sqlStr = sqlStr + " beadaldiv='1'," & vbCrlf
        sqlStr = sqlStr + " buyname='" + html2db(oAgirlOrder.FItemList(0).FOrderName) + "'," & vbCrlf
        sqlStr = sqlStr + " buyphone='" + replace(oAgirlOrder.FItemList(0).FOrderTelNo,"'","") + "'," & vbCrlf
        sqlStr = sqlStr + " buyhp='" + replace(oAgirlOrder.FItemList(0).FOrderHpNo,"'","") + "'," & vbCrlf
        sqlStr = sqlStr + " buyemail=''," & vbCrlf    ''''''''" + html2db(oAgirlOrder.FItemList(0).FOrderEmail) + "'," & vbCrlf
        sqlStr = sqlStr + " reqname='" + html2db(oAgirlOrder.FItemList(0).FReceiveName) + "'," & vbCrlf
        sqlStr = sqlStr + " reqzipcode='" + oAgirlOrder.FItemList(0).FReceiveZipCode + "'," & vbCrlf
        sqlStr = sqlStr + " reqaddress='" + html2db(oAgirlOrder.FItemList(0).FReceiveAddr2) + "'," & vbCrlf
        sqlStr = sqlStr + " reqphone='" + replace(oAgirlOrder.FItemList(0).FReceiveTelNo,"'","") + "'," & vbCrlf
        sqlStr = sqlStr + " reqhp='" + replace(oAgirlOrder.FItemList(0).FReceiveHpNo,"'","") + "'," & vbCrlf
        sqlStr = sqlStr + " comment='" + html2db(oAgirlOrder.FItemList(0).FEtcAsk) + "'," & vbCrlf
        sqlStr = sqlStr + " discountrate=1," & vbCrlf
        sqlStr = sqlStr + " subtotalprice=0," & vbCrlf
        sqlStr = sqlStr + " reqzipaddr='" + html2db(oAgirlOrder.FItemList(0).FReceiveAddr1) + "'" & vbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(iid)
    
    	dbget.Execute sqlStr
    
if Err then
    dbget.RollBackTrans
    response.write ErrMsg & Err.Description
    response.end
else
    ErrMsg = "[003]"
end if

    buf_totcost = 0
    buf_totvat = 0
    buf_iitemmakerid = ""
    
    For i=0 to oAgirlOrder.FResultCount-1
		if (oAgirlOrder.FItemList(i).FItemSeq=0) then
		    ''업체 개별배송 배송비 추가 요망.. :: brandSeq2Makerid
            buf_iitemmakerid = getbrandSeq2Makerid(oAgirlOrder.FItemList(i).FBrandSeq)
            
			buf_totcost = buf_totcost + CLng(oAgirlOrder.FItemList(i).FRealSellPrice)
			''배송옵션
			sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
			sqlStr = sqlStr + "itemoption, makerid, itemno, itemcost, orgitemcost, itemcostCouponNotApplied, buycash, itemvat, mileage, itemname, itemoptionname, reducedPrice)" & vbCrlf
			sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
			sqlStr = sqlStr + " '" & orderserial & "'," & vbCrlf
			sqlStr = sqlStr + " 0," & vbCrlf
			sqlStr = sqlStr + " '" & oAgirlOrder.FItemList(i).FOptionCode & "'," & vbCrlf
			if (Left(oAgirlOrder.FItemList(i).FpartnerOption,2)<>"09") then
			    sqlStr = sqlStr + " ''," & vbCrlf
			else
			    ''업체 조건배송
			    sqlStr = sqlStr + " '"&buf_iitemmakerid&"'," & vbCrlf
		    end if
			sqlStr = sqlStr + " 1," & vbCrlf
			sqlStr = sqlStr + "	" & oAgirlOrder.FItemList(i).FRealSellPrice & "," & vbCrlf
			sqlStr = sqlStr + "	" & oAgirlOrder.FItemList(i).FRealSellPrice & "," & vbCrlf
			sqlStr = sqlStr + "	" & oAgirlOrder.FItemList(i).FRealSellPrice & "," & vbCrlf
			
			if (Left(oAgirlOrder.FItemList(i).FpartnerOption,2)<>"09") then
			    sqlStr = sqlStr + "	0," & vbCrlf
			else
			    ''매입가
			    sqlStr = sqlStr + "	" &  oAgirlOrder.FItemList(i).FSupplyPrice & "," & vbCrlf   ''''0
			end if
			sqlStr = sqlStr + "	" & CLng(oAgirlOrder.FItemList(i).FRealSellPrice*1/11) & ","
			sqlStr = sqlStr + "	0,"
			sqlStr = sqlStr + "	'','',"
			sqlStr = sqlStr + "	" &oAgirlOrder.FItemList(i).FRealSellPrice & "" & vbCrlf
			sqlStr = sqlStr + " )"
			dbget.Execute sqlStr

			if Err then
			    dbget.RollBackTrans
			    response.write ErrMsg & Err.Description
			    response.end
			else
			    ErrMsg = "[003.1]"
			end if

		else
			sqlStr= "select top 1 convert(int, i.sellcash) as sellcash, " & vbCrlf
			sqlStr = sqlStr + " i.mileage, i.itemdiv , convert(int,i.buycash) as buycash ," & vbCrlf
			sqlStr = sqlStr + " convert(int, i.orgprice) as orgprice, convert(int,i.orgsuplycash) as orgsuplycash ," & vbCrlf
			sqlStr = sqlStr + " i.itemname, i.makerid, i.vatinclude, i.deliverytype, i.sailyn, i.mwdiv,"
			sqlStr = sqlStr + " IsNull(v.optionname,'') as codeview, IsNull(v.optaddbuyprice,0) as optaddbuyprice" & vbCrlf
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" & vbCrlf
			sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option v "
			sqlStr = sqlStr + "     on (i.itemid=v.itemid) and (v.itemoption='" + CStr(oAgirlOrder.FItemList(i).FpartnerOption) + "')" & vbCrlf
			sqlStr = sqlStr + " where i.itemid = " + CStr(oAgirlOrder.FItemList(i).FpartnerItemID) + ""

            rsget.Open sqlStr, dbget, 1
            if  not rsget.EOF  then
                buf_deliverytype = rsget("deliverytype")
				if (buf_deliverytype="2") or (buf_deliverytype="5")  or (buf_deliverytype="9") or (buf_deliverytype="7") then
					t_upchebeasong="Y"
				else
					t_upchebeasong="N"
				end if

                buf_sellcash        = rsget("sellcash")
                buf_sellvat         = CLng(buf_sellcash*11/10)-CLng(buf_sellcash)
                buf_mileage         = rsget("mileage")
                buf_itemdiv         = rsget("itemdiv")
                buf_iitemname       = replace(rsget("itemname"),"'","")
                buf_iitemoptionname = replace(rsget("codeview"),"'","")
                buf_onlyitembuycash = rsget("buycash")
                buf_onlyoptaddbuyprice = rsget("optaddbuyprice")
                buf_iitembuycash    = buf_onlyitembuycash + buf_onlyoptaddbuyprice
                buf_iitemmakerid    = rsget("makerid")
                buf_iitemvatinclude = rsget("vatinclude")
                buf_mwdiv           = rsget("mwdiv")

                buf_sailyn          = rsget("sailyn")
                buf_orgprice        = rsget("orgprice")
                buf_orgsuplycash    = rsget("orgsuplycash")
            end if
            rsget.close

            ''할인판매인경우 체크 20100201 추가 / 할인 테이블에 값이 있는경우..
            ''================================================================

            if (CLng(buf_sellcash) > CLng(oAgirlOrder.FItemList(i).FRealSellPrice))  then

'                if ((param1="dnshop") or (param1="interpark")) then
'                    mayOrderDate = Left(oAgirlOrder.FItemList(i).FSelldate,8)
'                    mayOrderDate = Left(mayOrderDate,4) & "-" & Mid(mayOrderDate,5,2) & "-" & Right(mayOrderDate,2)
'
'                    if IsDate(mayOrderDate) then
'                        sqlStr= " select top 1 itemid,saleprice,salesupplycash from db_event.dbo.tbl_saleitem"  & vbCrlf
'                        sqlStr = sqlStr + " where itemid=" & CStr(oAgirlOrder.FItemList(i).FpartnerItemID)  & vbCrlf
'                        sqlStr = sqlStr + " and convert(varchar(10),opendate,21)<='"&mayOrderDate&"'"  & vbCrlf
'                        sqlStr = sqlStr + " and convert(varchar(10),IsNULL(closedate,'2099-12-31'),21)>='"&mayOrderDate&"'"  & vbCrlf
'                        sqlStr = sqlStr + " order by saleitem_idx desc"  & vbCrlf
'
'                        rsget.Open sqlStr, dbget, 1
'                        if  not rsget.EOF  then
'                            if CLng(rsget("saleprice"))=CLng(oAgirlOrder.FItemList(i).FRealSellPrice) then
'                                buf_onlyitembuycash = rsget("salesupplycash")
'                                buf_iitembuycash    = buf_onlyitembuycash + buf_onlyoptaddbuyprice
'                                buf_sailyn   = "Y"
'                            end if
'                        end if
'                        rsget.close
'                    end if
'                end if
            elseif (CLng(buf_sellcash) < CLng(oAgirlOrder.FItemList(i).FRealSellPrice)) and (buf_sailyn="Y")  then
                if (CLng(oAgirlOrder.FItemList(i).FRealSellPrice)=buf_orgprice) then
                    buf_iitembuycash = buf_orgsuplycash + buf_onlyoptaddbuyprice
                    buf_sailyn   = "N"
                end if
            end if
            ''================================================================
            if Err then
                dbget.RollBackTrans
                response.write ErrMsg & Err.Description
                response.end
            else
                ErrMsg = "[003.1]"
            end if

            if oAgirlOrder.FItemList(i).FRealSellPrice<>0 then
            	buf_sellcash = oAgirlOrder.FItemList(i).FRealSellPrice
            end if

            buf_totcost = buf_totcost + CLng(buf_sellcash) * CLng(oAgirlOrder.FItemList(i).FOrderCount)
            buf_totvat  = buf_totvat + CLng(buf_sellvat) * CLng(oAgirlOrder.FItemList(i).FOrderCount)

			sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
			sqlStr = sqlStr + "itemoption, itemno, itemcost, itemvat, mileage, reducedPrice, " & vbCrlf
			sqlStr = sqlStr + "orgitemcost,itemcostcouponnotApplied,buycashcouponNotApplied, " & vbCrlf
			sqlStr = sqlStr + "itemname,itemoptionname,makerid,buycash," & vbCrlf
			sqlStr = sqlStr + "vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,requiredetail)" & vbCrlf
			sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
			sqlStr = sqlStr + " '" + orderserial + "'," & vbCrlf
			sqlStr = sqlStr + " " + CStr(oAgirlOrder.FItemList(i).FpartnerItemID) + "," & vbCrlf
			sqlStr = sqlStr + " '" + CStr(oAgirlOrder.FItemList(i).FpartnerOption) + "'," & vbCrlf
			sqlStr = sqlStr + " " + CStr(oAgirlOrder.FItemList(i).FOrderCount) + "," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_sellcash) + "," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_sellvat) + "," & vbCrlf
			sqlStr = sqlStr + " 0," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_sellcash) + "," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_orgprice) + "," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_sellcash) + "," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_iitembuycash) + "," & vbCrlf
			sqlStr = sqlStr + " '" + CStr(buf_iitemname) + "'," & vbCrlf
			sqlStr = sqlStr + " '" + CStr(buf_iitemoptionname) + "'," & vbCrlf
			sqlStr = sqlStr + " '" + CStr(buf_iitemmakerid) + "'," & vbCrlf
			sqlStr = sqlStr + " " + CStr(buf_iitembuycash) + "," & vbCrlf
			sqlStr = sqlStr + " '" + CStr(buf_iitemvatinclude) + "'," & vbCrlf
			sqlStr = sqlStr + " '" + t_upchebeasong + "'," & vbCrlf
			sqlStr = sqlStr + " '" + buf_sailyn + "'," & vbCrlf
			sqlStr = sqlStr + " '" + buf_itemdiv + "'," & vbCrlf
			sqlStr = sqlStr + " '" + buf_mwdiv + "'," & vbCrlf
			sqlStr = sqlStr + " '" + buf_deliverytype + "'," & vbCrlf
			sqlStr = sqlStr + " '" + replace(CStr(oAgirlOrder.FItemList(i).FAddOrderInfo),"'","''") + "'" & vbCrlf
			sqlStr = sqlStr + " )"

			dbget.Execute sqlStr
		end if
	next

	if Err then
	    dbget.RollBackTrans
	    response.write ErrMsg & Err.Description
	    response.end
	else
	    ErrMsg = "[004]"
	end if

	sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
	sqlStr = sqlStr + " set totalvat = " + CStr(buf_totvat) + "," & vbCrlf
	sqlStr = sqlStr + " totalsum = " + CStr(buf_totcost) + "," & vbCrlf
	sqlStr = sqlStr + " subtotalprice = " + CStr(buf_totcost) + "," & vbCrlf
	sqlStr = sqlStr + " subtotalPriceCouponNotApplied = " + CStr(buf_totcost) + "" & vbCrlf
	sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

	dbget.Execute sqlStr

    if Err then
        dbget.RollBackTrans
        response.write ErrMsg & Err.Description
        response.end
    else
        dbget.CommitTrans
        rw "["&orderserial&"]"
    end if

    ''aGirl Flag update 
    sqlStr = " update db_agirlOrder.dbo.tbl_OrderItem"
    sqlStr = sqlStr & " set OrderItemStatus=3"
    sqlStr = sqlStr & " ,confirmdate=getdate()"
    sqlStr = sqlStr & " where orderSerial='"&oAgirlOrder.FItemList(0).FOrderserial&"'"
    sqlStr = sqlStr & " and itemSeq<>0"
    sqlStr = sqlStr & " and isCancel<>'Y'"
    sqlStr = sqlStr & " and (BrandSeq in (select brandseq from db_agirlOrder.dbo.tbl_TenLinkBrand where UseYn='Y')"
    sqlStr = sqlStr & "     or BrandSeq in (select  brandSeq from db_agirluser.dbo.tbl_Back_Brand where partnerSEq =334)"   '''
	 sqlStr = sqlStr & "     )"
    sqlStr = sqlStr & " and IsNULL(OrderItemStatus,0)<3"
    
    dbagirl_dbget.Execute sqlStr
    
    sqlStr = " update db_agirlOrder.dbo.tbl_Order"
    sqlStr = sqlStr & " set OrderStatus=5"
    sqlStr = sqlStr & " where orderSerial='"&oAgirlOrder.FItemList(0).FOrderserial&"'"
    sqlStr = sqlStr & " and isCancel='N'"
    sqlStr = sqlStr & " and IsNULL(OrderStatus,0)<5"
    
    dbagirl_dbget.Execute sqlStr
    
    ''사은품.
    ''sqlStr = "exec [db_order].[dbo].sp_Ten_order_gift '" & orderserial & "'"
	''dbget.Execute(sqlStr)

	''현재고 업데이트
    ''sqlStr = "exec [db_summary].[dbo].sp_ten_RealtimeStock_regOrder '" & orderserial & "'"
    ''dbget.execute sqlStr

    Set oAgirlOrder = Nothing
    end if
Next
On Error Goto 0
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbagirlclose.asp" -->

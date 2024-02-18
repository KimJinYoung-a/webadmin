<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 상품등록
' History : 2008.04.04 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->

<%
Dim sMode,sCode, dSDate, dEDate ,strSql,addSql, addSqlDB ,egCode,eCode,itemidarr,sType, i
Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus ,objCmd,iResult, ErrStr : ErrStr = ""
	sMode     = requestCheckVar(Request("mode"),1)	

SELECT Case sMode
	
	Case "I"	'할인상품 추가
	itemidarr = Request("itemidarr")

	sType 		=  requestCheckVar(Request("sType"),16)
	
	sCode 		= requestCheckVar(Request("sC"),10)
	eCode 		= requestCheckVar(Request("eC"),10)
	egCode 		= requestCheckVar(Request("egC"),10)
	if egCode = "" then egCode = 0
	itemid      = request("itemid")
	itemname    = requestCheckVar(request("itemname"),32)
	makerid     = requestCheckVar(request("makerid"),32)
	sellyn      = requestCheckVar(request("sellyn"),1)
	usingyn     = requestCheckVar(request("usingyn"),1)
	danjongyn   = requestCheckVar(request("danjongyn"),1)
	limityn     = requestCheckVar(request("limityn"),1)
	mwdiv       = requestCheckVar(request("mwdiv"),2)
	sailyn      = requestCheckVar(request("sailyn"),1)
	deliverytype= requestCheckVar(request("deliverytype"),1)
	
	cdl = requestCheckVar(request("cdl"),3)
	cdm = requestCheckVar(request("cdm"),3)
	cds = requestCheckVar(request("cds"),3)
  	if itemid <> "" then
		if checkNotValidHTML(itemid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If

	addSql = ""
	addSqlDB = ""
	  IF sType = "all" THEN '검색된 모든 내용 insert  처리
	  	 '// 추가 쿼리
        if (makerid <> "") then
            addSql = addSql & " and i.makerid='" + makerid + "'"
        end if

        if (itemid <> "") then
            addSql = addSql & " and i.itemid in (" + itemid + ")"
            itemidarr = itemid
        end if

        if (itemname <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"
        end if
        
        if (sellyn <> "") then
            addSql = addSql & " and i.sellyn='" + sellyn + "'"
        end if

        if (usingyn <> "") then
            addSql = addSql & " and i.isusing='" + usingyn + "'"
        end if
        
        if danjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif danjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + danjongyn + "'"
        end if
      		
		if limityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif limityn<>"" then
            addSql = addSql + " and i.limityn='" + limityn + "'"
        end if        
        
        if mwdiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif mwdiv<>"" then
            addSql = addSql + " and i.mwdiv='" + mwdiv + "'"
        end if
		
        if cdl<>"" then
            addSql = addSql + " and i.cate_large='" + cdl + "'"
        end if
        
        if cdm<>"" then
            addSql = addSql + " and i.cate_mid='" + cdm + "'"
        end if
        
        if cds<>"" then
            addSql = addSql + " and i.cate_small='" + cds + "'"
        end if
        
        if sailyn<>"" then
            addSql = addSql + " and i.saleyn='" + sailyn + "'"
        end if  
        
         if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if
    ELSE
    	addSql = addSql & " and i.itemid in ("&trim(itemidarr)&")"	    
	END IF
		        
        if eCode <> "" then
        	addSqlDB = " , [db_academy].[dbo].[tbl_eventitem] c "
        	addSql = addSql + " and i.itemid = c.itemid and c.evt_code = "&eCode&" and c.evtgroup_code ="&egCode
        end if
	
		'- 할인중이지 않은 상품에 한해서
		'- 추가하려는 할인기간내에 할인예정이지 않은 상품에 한해서
		dim iSaleRate, iSaleMargin, iSaleMarginValue
		'- 추가하려는 할인정보의 기간 확인
		strSql = " SELECT sale_startdate, sale_enddate, sale_rate, sale_margin, sale_marginvalue, sale_status from [db_academy].[dbo].tbl_sale where sale_code= "&sCode		
		rsacademyget.Open strSql,dbacademyget 
		IF not rsacademyget.EOF THEN
			dSDate = rsacademyget("sale_startdate")
			dEDate = rsacademyget("sale_enddate")	
			iSaleRate = rsacademyget("sale_rate")	
			iSaleMargin = rsacademyget("sale_margin")	
			iSaleMarginValue = rsacademyget("sale_marginvalue")	
			saleStatus	= rsacademyget("sale_status")
		End IF
		rsacademyget.Close
		dim strStatus, arrList,intLoop
		
		IF itemidarr <> "" THEN
		strSql = "SELECT TOP 100  b.itemid, a.sale_code, a.sale_status "&_
				"   FROM  [db_academy].[dbo].tbl_sale a, [db_academy].[dbo].[tbl_saleitem] b "&_
				"   WHERE  a.sale_code = b.sale_code and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"'"&_
				"	 and a.sale_using =1 and a.sale_status <> 8 and  b.saleitem_status <> 8  and b.itemid in ("&itemidarr&")"	
		rsacademyget.Open strSql,dbacademyget 		
		IF not rsacademyget.EOF THEN
			arrList = rsacademyget.getRows()		 	
		End IF
		rsacademyget.Close
		
		'If isArray(arrList) THEN
		'	For intLoop =0 To UBound(arrList,2)
		'	IF arrList(2,intLoop) = 6 THEN 
		'		strStatus = "할인중"
		'	ELSEIF arrList(2,intLoop) = 7 THEN 
		'		strStatus = "할인예정"
		'	ELSEIF arrList(2,intLoop) = 0 THEN 
		'		strStatus = "등록대기"	
		'	END IF	
		'	
		'	ErrStr = ErrStr + "할인코드 : " + CStr(arrList(1,intLoop)) + " - 상품번호 : " + CStr(arrList(0,intLoop)) +" "+ strStatus + " \n"
		'	Next
		'END IF
	END IF	
		strSql = "INSERT INTO [db_academy].[dbo].[tbl_saleItem]([sale_code], [itemid], [saleItem_status], [saleprice],[salesupplycash]) "
		strSql = strSql&" SELECT "&sCode&", i.itemid, 7, i.orgprice-(i.orgprice*"&iSaleRate&"/100)"
	IF iSaleMargin = 1 THEN			'동일마진
		strSql = strSql&" ,(i.orgprice-(i.orgprice*"&iSaleRate&"/100))- convert(int,(i.orgprice-(i.orgprice*"&iSaleRate&"/100))*(100-convert(float,convert(int,i.orgsuplycash/i.orgprice*10000)/100))/100)"
	ELSEIF 	iSaleMargin = 2 THEN	'업체부담
		strSql = strSql&" ,(i.orgprice-(i.orgprice*"&iSaleRate&"/100)) - (i.orgprice- i.orgsuplycash)"
	ELSEIF 	iSaleMargin = 3 THEN	'반반부담
		strSql = strSql&" , i.orgsuplycash - Convert(int, (i.orgprice-(i.orgprice-(i.orgprice*"&iSaleRate&"/100)))/2)"
	ELSEIF 	iSaleMargin = 4 THEN	'10x10부담
		strSql = strSql&" , i.orgsuplycash "
	ELSEIF 	iSaleMargin = 5 THEN	'직접설정
		strSql = strSql&" , (i.orgprice-(i.orgprice*"&iSaleRate&"/100)) - convert(int, (i.orgprice-(i.orgprice*"&iSaleRate&"/100))*convert(float,"&iSaleMarginValue&")/100) "		
	END IF	
		strSql = strSql&"	FROM [db_academy].dbo.tbl_diy_item i "&addSqlDB
		strSql = strSql&"   WHERE i.itemid not in "
		strSql = strSql&" (select b.itemid from [db_academy].[dbo].tbl_sale a, [db_academy].[dbo].[tbl_saleitem] b"
		strSql = strSql&" 	where a.sale_code = b.sale_code and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"'"
		strSql = strSql&		"	 and a.sale_using =1 and a.sale_status <> 8  and  b.saleitem_status <> 8 )"&addSql							
		
		'response.write strSql &"<Br>"
		dbacademyget.execute strSql
	
		IF Err.Number <> 0 THEN			
	       Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank" 				
			dbacademyget.close()	:	response.End
		END IF	
		
		IF saleStatus = 6 THEN
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
			.ActiveConnection = dbacademyget
			.CommandType = adCmdText
			.CommandText = "{?= call [db_academy].dbo.sp_academy_item_SetPrice_RealTime ("&sCode&")}"			
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
			 iResult = objCmd(0).Value
			 Set objCmd = nothing
			IF iResult <> 1 THEN
		   			dbacademyget.RollBackTrans	
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		       		dbacademyget.close()	:	response.End	  
			End IF 
		END IF	
%>
	<script langauge="javascript">

		<%
		if ErrStr<>"" then
			ErrStr = ErrStr + "\n\n 할인은 중복설정 불가능합니다. 할인상품을 제외한 상품이 추가됩니다."
			response.write "alert('" + ErrStr + "')"
		end if
		%>		
		location.href ="about:blank";
		parent.history.go(0);
		//parent.location.reload();	

	</script>
<% 	   
		dbacademyget.close()	:	response.End			
	
	Case "U"	'할인 선택상품 수정
	Dim  dissellprice,disbuyprice,arrsaleItemStatus,saleStatus, tmpsaleItemStatus
		sCode = requestCheckVar(Request("sC"),10)
		iCurrpage 	= request("iC")
		itemid 		= split(request("itemid"),",")				
		dissellprice= split(request("iDSPrice"),",")
		disbuyprice = split(request("iDBPrice"),",")
		arrsaleItemStatus	=split(request("saleItemStatus"),",")
		saleStatus	=requestCheckVar(Request("saleStatus"),4)
		
		dbacademyget.beginTrans
		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then
								
				if Cint(trim(arrsaleItemStatus(i))) = 6 then '오픈중 상태일때 값 변경시 상태값 오픈예정으로 변경처리
					arrsaleItemStatus(i) = 7	
				end if
					
				IF trim(arrsaleItemStatus(i)) = 9 THEN	'원가로
					strSql ="UPDATE [db_academy].[dbo].[tbl_saleItem] "&_
							" SET saleitem_status =9, lastupdate = getdate()"&_
							" WHERE itemid = "&trim(itemid(i))
				ELSE	
					strSql ="UPDATE [db_academy].[dbo].[tbl_saleItem] "&_
							" SET saleprice = "&trim(dissellprice(i))&", salesupplycash="&trim(disbuyprice(i))&", saleitem_status ="&arrsaleItemStatus(i)&", lastupdate = getdate()"&_
							" WHERE itemid = "&trim(itemid(i))
				END IF
					dbacademyget.execute strSql
					
				IF Err.Number <> 0 THEN
		   			dbacademyget.RollBackTrans	
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		       		dbacademyget.close()	:	response.End	  
				End IF
				
				IF Cint(saleStatus) = 6 THEN	'오픈상태일때만 바로 적용						
					Set objCmd = Server.CreateObject("ADODB.COMMAND")
					With objCmd
					.ActiveConnection = dbacademyget
					.CommandType = adCmdText
					.CommandText = "{?= call [db_academy].dbo.sp_academy_item_SetPrice_RealTime ("&sCode&")}"			
					.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
					.Execute, , adExecuteNoRecords
					End With	
				    iResult = objCmd(0).Value
				    Set objCmd = nothing
				    
				    IF iResult <> 1 THEN
		   			dbacademyget.RollBackTrans	
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		       		dbacademyget.close()	:	response.End	  
					End IF
				END IF	
				
			end if
		next
		
		dbacademyget.CommitTrans	
		response.redirect("saleItemReg.asp?menupos="&menupos&"&sC="&sCode&"&iC="&iCurrpage)
	dbacademyget.close()	:	response.End
	
	Case "D"	'할인상품 삭제
		sCode = requestCheckVar(Request("sC"),10)		
		itemid 		= split(request("itemid"),",")		
				
		dbacademyget.beginTrans
		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then								
			strSql ="delete from [db_academy].[dbo].[tbl_saleItem] "&_
					" WHERE itemid="&trim(itemid(i)) &""&_
					" and sale_code="&Cstr(sCode)
			dbacademyget.execute strSql
			
				IF Err.Number <> 0 THEN
		   			dbacademyget.RollBackTrans	
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		       		dbacademyget.close()	:	response.End	  
				End IF
			end if
		next
		
		dbacademyget.CommitTrans	
		response.redirect("saleItemReg.asp?menupos="&menupos&"&sC="&sCode)
	dbacademyget.close()	:	response.End
	
	Case "P"	'상품테이블 적용
		sCode = requestCheckVar(Request("sC"),10)	
		iCurrpage 	= request("iC")
		IF sCode = "" THEN 
			Alert_return("파라미터값에 문제가 있습니다.")    	
		     dbacademyget.close()	:	response.End	
		END IF       		
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbacademyget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_academy].dbo.sp_academy_item_SetPrice_RealTime ("&sCode&")}"			
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With	
	    iResult = objCmd(0).Value
	    Set objCmd = nothing
	  
	'검색어 체크--------------------------------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'검색어	
	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일	
	 iCurrpage 		= requestCheckVar(Request("iC"),10)			'현재 페이지 번호
	 ssStatus		= requestCheckVar(Request("sstatus"),10)	'검색 상태
 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&salestatus="&ssStatus
 	'--------------------------------------------------------------
 
		IF iResult <> 1 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		    dbacademyget.close()	:	response.End	
		END IF
			response.redirect("saleList.asp?menupos="&menupos&"&"&strParm)
	dbacademyget.close()	:	response.End
	CASE Else
	Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")    	
       dbacademyget.close()	:	response.End
END SELECT	

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

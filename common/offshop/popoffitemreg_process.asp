<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : pos 상품관리
' Hieditor : 2011.01.13 서동석 생성
'			 2011.03.15 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
''****
''  discountsellprice 삭제, orgsellprice (소비자가) 추가, shopitemprice : 실판매가.
''****

dim mode,itemgubunarr,itemarr ,detailidxarr, currjungsanidarr ,idx
dim itemoptionarr,itempricearr,isusingarr,extbarcodearr, itemsuplyarr
dim designer,itemgubun,itemname,sellcash,suplycash ,shopbuypricearr
dim shopbuyprice ,orgsellpricearr ,sellcasharr,suplycasharr,itemnoarr,designerarr
dim chargeid,shopid,divcode,vatcode ,cksel ,i,cnt,sqlStr ,extbarcodeAlreadyExists
dim extbarcodeAlreadyExistsString ,stockitemexists
dim itemid,itemoption,orgsellprice ''itemgubun, shopbuyprice
dim shopitemprice, discountsellprice, shopsuplycash ,cd1, cd2, cd3
dim extbarcode, isusing, shopitemname, shopitemoptionname, vatinclude, makerid, centermwdiv
	mode            = requestCheckVar(request("mode"),32)
	itemgubunarr    = request("itemgubunarr")
	itemarr         = request("itemarr")
	itemoptionarr   = request("itemoptionarr")
	orgsellpricearr = request("orgsellpricearr")   ''소비자가 추가
	itempricearr    = request("itempricearr")
	itemsuplyarr    = request("itemsuplyarr")
	isusingarr      = request("isusingarr")
	extbarcodearr   = request("extbarcodearr")
	detailidxarr    = request("detailidxarr")
	currjungsanidarr = request("currjungsanidarr")
	designer        = requestCheckVar(request("designer"),32)
	itemname        = requestCheckVar(request("itemname"),124)
	sellcash        = requestCheckVar(request("sellcash"),20)
	suplycash       = requestCheckVar(request("suplycash"),20)
	shopbuyprice    = requestCheckVar(request("shopbuyprice"),20)
	''dim discountsellpricearr ''삭제
	sellcasharr     = request("sellcasharr")
	suplycasharr    = request("suplycasharr")
	itemnoarr       = request("itemnoarr")
	designerarr     = request("designerarr")
	''discountsellpricearr = request("discountsellpricearr")
	shopbuypricearr = request("shopbuypricearr")
	chargeid    = requestCheckVar(request("chargeid"),32)
	shopid      = requestCheckVar(request("shopid"),32)
	divcode     = requestCheckVar(request("divcode"),3)
	vatcode     = requestCheckVar(request("vatcode"),3)
	idx         = requestCheckVar(request("idx"),10)
	cksel = request("cksel")
    itemgubun   = requestCheckVar(request("itemgubun"),2)
    itemid      = requestCheckVar(request("itemid"),10)
    itemoption  = requestCheckVar(request("itemoption"),4)
    orgsellprice = requestCheckVar(request("orgsellprice"),20)
    shopitemprice = requestCheckVar(request("shopitemprice"),20)
    discountsellprice = requestCheckVar(request("discountsellprice"),20)
    shopsuplycash = requestCheckVar(request("shopsuplycash"),20)
    shopbuyprice = requestCheckVar(request("shopbuyprice"),20)
    extbarcode = requestCheckVar(request("extbarcode"),32)
    isusing = requestCheckVar(request("isusing"),1)
    shopitemname = requestCheckVar(html2db(request("shopitemname")),124)
    shopitemoptionname = requestCheckVar(html2db(request("shopitemoptionname")),96)
    vatinclude = requestCheckVar(request("vatinclude"),1)
    makerid = requestCheckVar(request("makerid"),32)
    centermwdiv = requestCheckVar(request("centermwdiv"),1)
    cd1 = requestCheckVar(request("cd1"),3)
    cd2 = requestCheckVar(request("cd2"),3)
    cd3 = requestCheckVar(request("cd3"),3)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim IsForeignShop
	IsForeignShop = (request("isforeignshop")="on")

if mode ="arrmodi" then
    '' 일괄수정..
	itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
	itemarr = Left(itemarr,Len(itemarr)-1)
	itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
	orgsellpricearr = Left(orgsellpricearr,Len(orgsellpricearr)-1)
	itempricearr = Left(itempricearr,Len(itempricearr)-1)
	isusingarr = Left(isusingarr,Len(isusingarr)-1)
	itemsuplyarr = Left(itemsuplyarr,Len(itemsuplyarr)-1)
	shopbuypricearr = Left(shopbuypricearr,Len(shopbuypricearr)-1)
	
	itemgubunarr = split(itemgubunarr,"|")
	itemarr = split(itemarr,"|")
	itemoptionarr = split(itemoptionarr,"|")	
	orgsellpricearr = split(orgsellpricearr,"|")
	itempricearr = split(itempricearr,"|")
	itemsuplyarr = split(itemsuplyarr,"|")
	shopbuypricearr = split(shopbuypricearr,"|")
	isusingarr = split(isusingarr,"|")
	extbarcodearr = split(extbarcodearr,"|")

	cnt = ubound(itemarr)

	for i=0 to cnt
		''CheckBarCode Already Exists
		extbarcodeAlreadyExists = false
		if extbarcodearr(i)<>"" then
			sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
			sqlStr = sqlStr + " where barcode='" + CStr(requestCheckVar(trim(extbarcodearr(i)),32)) + "'" + VbCrlf
			sqlStr = sqlStr + " and not ("
			sqlStr = sqlStr + " 	itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf
			sqlStr = sqlStr + " 	and itemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" + VbCrlf
			sqlStr = sqlStr + " 	and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" + VbCrlf
			sqlStr = sqlStr + " ) "
			
			'response.write sqlStr &"<Br>"
			rsget.Open sqlStr,dbget,1
			if Not rsget.EOF then
				extbarcodeAlreadyExists = true
				extbarcodeAlreadyExistsString = extbarcodeAlreadyExistsString + requestCheckVar(extbarcodearr(i),32) + ","
			end if
			rsget.close
		end if

		if Not extbarcodeAlreadyExists then
			sqlStr = " update [db_shop].[dbo].tbl_shop_item"
			sqlStr = sqlStr + " set shopitemprice=" + CStr(requestCheckVar(itempricearr(i),20)) + ","
			sqlStr = sqlStr + " orgsellprice=" + CStr(requestCheckVar(orgsellpricearr(i),20)) + ","     '' 소비자가 추가
			sqlStr = sqlStr + " shopsuplycash=" + CStr(requestCheckVar(itemsuplyarr(i),20)) + ","
			sqlStr = sqlStr + " extbarcode='" + CStr(requestCheckVar(extbarcodearr(i),32)) + "',"
			sqlStr = sqlStr + " isusing='" + CStr(requestCheckVar(isusingarr(i),1)) + "',"
			sqlStr = sqlStr + " shopbuyprice=" + requestCheckVar(shopbuypricearr(i),20) + ","
			sqlStr = sqlStr + " updt=getdate()"
			sqlStr = sqlStr + " where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"
			sqlStr = sqlStr + " and shopitemid=" + CStr(requestCheckVar(itemarr(i),10)) + ""
			sqlStr = sqlStr + " and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'"

			'response.write sqlStr &"<Br>"
			dbget.Execute sqlStr

            IF (IsForeignShop) and (shopid<>"") then
                sqlStr = " IF Exists(" & VbCRLF
                sqlStr = sqlStr + " 	select * from db_shop.dbo.tbl_shop_locale_item" & VbCRLF 
                sqlStr = sqlStr + " 	where shopid='"+shopid+"'" & VbCRLF
                sqlStr = sqlStr + " 	and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" & VbCRLF
                sqlStr = sqlStr + " 	and shopitemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" & VbCRLF
                sqlStr = sqlStr + " 	and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" & VbCRLF
                sqlStr = sqlStr + " )" & VbCRLF
                sqlStr = sqlStr + " BEGIN" & VbCRLF
                sqlStr = sqlStr + " 	update db_shop.dbo.tbl_shop_locale_item " & VbCRLF
                sqlStr = sqlStr + " 	set lcprice=" + CStr(requestCheckVar(itempricearr(i),20)) & VbCRLF
                sqlStr = sqlStr + " 	,lastupdate=getdate()" & VbCRLF
                sqlStr = sqlStr + " 	,exchangerate=1" & VbCRLF
                sqlStr = sqlStr + " 	,multipleRate=1" & VbCRLF
                sqlStr = sqlStr + " 	where shopid='"+shopid+"'" & VbCRLF
                sqlStr = sqlStr + " 	and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" & VbCRLF
                sqlStr = sqlStr + " 	and shopitemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" & VbCRLF
                sqlStr = sqlStr + " 	and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" & VbCRLF
                sqlStr = sqlStr + " END" & VbCRLF
                sqlStr = sqlStr + " ELSE" & VbCRLF
                sqlStr = sqlStr + " BEGIN" & VbCRLF
                sqlStr = sqlStr + " 	insert into db_shop.dbo.tbl_shop_locale_item " & VbCRLF
                sqlStr = sqlStr + " 	(shopid,itemgubun,shopitemid,itemoption,lcitemname,lcitemoptionname,lcprice,lastupdate,exchangerate,multipleRate)" & VbCRLF
                sqlStr = sqlStr + " 	select top 1 " & VbCRLF
                sqlStr = sqlStr + " 	'"+shopid+"'" & VbCRLF
                sqlStr = sqlStr + " 	,'" + requestCheckVar(itemgubunarr(i),2) + "'" & VbCRLF
                sqlStr = sqlStr + " 	," + CStr(requestCheckVar(itemarr(i),10)) + "" & VbCRLF
                sqlStr = sqlStr + " 	,'" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" & VbCRLF
                sqlStr = sqlStr + " 	,shopitemname " & VbCRLF
                sqlStr = sqlStr + " 	,shopitemoptionname" & VbCRLF
                sqlStr = sqlStr + " 	," + CStr(requestCheckVar(itempricearr(i),20)) + "" & VbCRLF
                sqlStr = sqlStr + " 	,getdate()" & VbCRLF
                sqlStr = sqlStr + " 	,1" & VbCRLF
                sqlStr = sqlStr + " 	,1" & VbCRLF
                sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "' and shopitemid=" +CStr(requestCheckVar(itemarr(i),10)) + " and itemoption='"& requestCheckVar(itemoptionarr(i),4) &"'"  & VbCRLF
                sqlStr = sqlStr + " END" & VbCRLF
                
                'response.write sqlStr &"<Br>"
                dbget.Execute sqlStr
            end if

			''바코드 테이블 확인
			if trim(CStr(itemarr(i)))<>"" then
				sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
				sqlStr = sqlStr + " where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf
				sqlStr = sqlStr + " and itemid=" + CStr(requestCheckVar(itemarr(i),10)) + "" + VbCrlf
				sqlStr = sqlStr + " and itemoption='" + CStr(requestCheckVar(itemoptionarr(i),4)) + "'" + VbCrlf
				
				'response.write sqlStr &"<Br>"
				rsget.Open sqlStr,dbget,1
				stockitemexists = (not rsget.Eof)
				rsget.close

				if (stockitemexists) then
					sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
					sqlStr = sqlStr + " set barcode='" + CStr(requestCheckVar(trim(extbarcodearr(i)),32)) + "'" + VbCrlf
					sqlStr = sqlStr + " where itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'" + VbCrlf
					sqlStr = sqlStr + " and itemid=" + CStr(requestCheckVar(itemarr(i)),10) + "" + VbCrlf
					sqlStr = sqlStr + " and itemoption='" + CStr(requestCheckVar(itemoptionarr(i)),4) + "'" + VbCrlf

					'response.write sqlStr &"<Br>"
					dbget.Execute sqlStr
				else
					sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
					sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
					sqlStr = sqlStr + " values("
					sqlStr = sqlStr + " '" + requestCheckVar(itemgubunarr(i),2) + "'," + VbCrlf
					sqlStr = sqlStr + " " + CStr(requestCheckVar(itemarr(i),10)) + "," + VbCrlf
					sqlStr = sqlStr + " '" + requestCheckVar(itemoptionarr(i),4) + "'," + VbCrlf
					sqlStr = sqlStr + " '" + requestCheckVar(trim(extbarcodearr(i)),32) + "'" + VbCrlf
					sqlStr = sqlStr + " )" + VbCrlf
					
					'response.write sqlStr &"<Br>"
					dbget.Execute sqlStr
				end if
			end if
		end if
	next

'/포스 상품 신규등록
elseif (mode="addetcoffitem") then
    if (Not IsNumeric(orgsellprice)) or (orgsellprice="") then orgsellprice =0
    if (Not IsNumeric(discountsellprice)) or (discountsellprice="") then discountsellprice =0
    if (Not IsNumeric(shopsuplycash)) or (shopsuplycash="") then shopsuplycash =0
    if (Not IsNumeric(shopbuyprice)) or (shopbuyprice="") then shopbuyprice =0
    
    if CStr(orgsellprice)="0" then orgsellprice=shopitemprice

    '''바코드 체크
    IF (extbarcode<>"") then
        sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
    	sqlStr = sqlStr + " where barcode='" + trim(extbarcode) + "'" + VbCrlf
    	
    	rsget.Open sqlStr,dbget,1
    	if Not rsget.EOF then
    		extbarcodeAlreadyExists = true
    	end if
    	rsget.close
    			
        IF (extbarcodeAlreadyExists) then
            response.write "<script type='text/javascript'>alert('이미 사용중인 바코드 입니다.-' + extbarcode + ' 등록 불가');</script>"
            response.end
        end if			
	ENd IF
	
	dbget.beginTrans ''위치변경
			
	sqlStr = " select top 1 shopitemid"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item"
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'"
	sqlStr = sqlStr + " order by shopitemid desc"
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			itemid = rsget("shopitemid")+1
		else
			itemid = 1
		end if
	rsget.close

	itemoption = "0000"

	sqlStr = " insert into [db_shop].[dbo].tbl_shop_item" + vbCrlf
	sqlStr = sqlStr + " (itemgubun,shopitemid,itemoption," + vbCrlf
	sqlStr = sqlStr + " makerid,shopitemname,shopitemoptionname,orgsellprice,shopitemprice," + vbCrlf
	sqlStr = sqlStr + " shopsuplycash,shopbuyprice, discountsellprice,"
	
	if cd1<>"" then
		sqlStr = sqlStr + "  catecdl," + vbCrlf
	end if

	if cd2<>"" then
		sqlStr = sqlStr + " catecdm," + vbCrlf
	end if

	if cd3<>"" then
		sqlStr = sqlStr + " catecdn," + vbCrlf
	end if
    
    if (centermwdiv<>"") then
        sqlStr = sqlStr + " centermwdiv," + vbCrlf
    end if
    
    if (extbarcode<>"") then
        sqlStr = sqlStr + " extbarcode," + vbCrlf
    end if
    
	sqlStr = sqlStr + " vatinclude)" + vbCrlf

	sqlStr = sqlStr + " values(" + vbCrlf
	sqlStr = sqlStr + " '" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(itemid) + "" + vbCrlf
	sqlStr = sqlStr + " ,'0000'" + vbCrlf
	sqlStr = sqlStr + " ,'" + makerid + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + shopitemname + "'" + vbCrlf
	sqlStr = sqlStr + " ,'" + shopitemoptionname + "'" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(orgsellprice) + "" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(shopitemprice) + "" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(shopsuplycash) + "" + vbCrlf
	sqlStr = sqlStr + " ," + CStr(shopbuyprice) + "" + vbCrlf
	sqlStr = sqlStr + " ,0" + vbCrlf

	if cd1<>"" then
		sqlStr = sqlStr + " ,'" + cd1 + "'" + vbCrlf
	end if

	if cd2<>"" then
		sqlStr = sqlStr + " ,'" + cd2 + "'" + vbCrlf
	end if

	if cd3<>"" then
		sqlStr = sqlStr + " ,'" + cd3 + "'" + vbCrlf
	end if
    
    if (centermwdiv<>"") then
        sqlStr = sqlStr + " ,'" + centermwdiv + "'" + vbCrlf 
    end if
    
    if (extbarcode<>"") then
        sqlStr = sqlStr + " ,'" + extbarcode + "'" + vbCrlf 
    end if
    
	sqlStr = sqlStr + " ,'" + vatinclude + "'" + vbCrlf
	sqlStr = sqlStr + " )" + vbCrlf
	
	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
       
    sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='0000'" + VbCrlf
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1
	    stockitemexists = (not rsget.Eof)
	rsget.close

	if (stockitemexists) then
		sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
		sqlStr = sqlStr + " set barcode='" + trim(extbarcode) + "'" + VbCrlf
		sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
		sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
		sqlStr = sqlStr + " and itemoption='0000'" + VbCrlf
		
		'response.write sqlStr &"<Br>"
		dbget.Execute sqlStr
	else
		sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
		sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
		sqlStr = sqlStr + " values("
		sqlStr = sqlStr + " '" + itemgubun + "'," + VbCrlf
		sqlStr = sqlStr + " " + CStr(itemid) + "," + VbCrlf
		sqlStr = sqlStr + " '0000'," + VbCrlf
		sqlStr = sqlStr + " '" + trim(extbarcode) + "'" + VbCrlf
		sqlStr = sqlStr + " )" + VbCrlf
		
		'response.write sqlStr &"<Br>"
		dbget.Execute sqlStr
	end if
    
    IF (IsForeignShop) and (shopid<>"") then
        
        sqlStr = " IF Exists(" & VbCRLF
        sqlStr = sqlStr + " 	select * from db_shop.dbo.tbl_shop_locale_item" & VbCRLF 
        sqlStr = sqlStr + " 	where shopid='"+shopid+"'" & VbCRLF
        sqlStr = sqlStr + " 	and itemgubun='" + itemgubun + "'" & VbCRLF
        sqlStr = sqlStr + " 	and shopitemid=" + CStr(itemid) + "" & VbCRLF
        sqlStr = sqlStr + " 	and itemoption='0000'" & VbCRLF
        sqlStr = sqlStr + " )" & VbCRLF
        sqlStr = sqlStr + " BEGIN" & VbCRLF
        sqlStr = sqlStr + " 	update db_shop.dbo.tbl_shop_locale_item " & VbCRLF
        sqlStr = sqlStr + " 	set lcprice=" + CStr(shopitemprice) & VbCRLF
        sqlStr = sqlStr + " 	,lcitemname='"+ shopitemname + "'" & VbCRLF
        sqlStr = sqlStr + " 	,lcitemoptionname='"+ shopitemoptionname + "'" & VbCRLF
        sqlStr = sqlStr + " 	,lastupdate=getdate()" & VbCRLF
        sqlStr = sqlStr + " 	,exchangerate=1" & VbCRLF
        sqlStr = sqlStr + " 	,multipleRate=1" & VbCRLF
        sqlStr = sqlStr + " 	where shopid='"+shopid+"'" & VbCRLF
        sqlStr = sqlStr + " 	and itemgubun='" + itemgubun + "'" & VbCRLF
        sqlStr = sqlStr + " 	and shopitemid=" + CStr(itemid) + "" & VbCRLF
        sqlStr = sqlStr + " 	and itemoption='0000'" & VbCRLF
        sqlStr = sqlStr + " END" & VbCRLF
        sqlStr = sqlStr + " ELSE" & VbCRLF
        sqlStr = sqlStr + " BEGIN" & VbCRLF
        sqlStr = sqlStr + " 	insert into db_shop.dbo.tbl_shop_locale_item " & VbCRLF
        sqlStr = sqlStr + " 	(shopid,itemgubun,shopitemid,itemoption,lcitemname,lcitemoptionname,lcprice,lastupdate,exchangerate,multipleRate)" & VbCRLF
        sqlStr = sqlStr + " 	values(" & VbCRLF
        sqlStr = sqlStr + " 	'"+shopid+"'" & VbCRLF
        sqlStr = sqlStr + " 	,'" + itemgubun + "'" & VbCRLF
        sqlStr = sqlStr + " 	," + CStr(itemid) + "" & VbCRLF
        sqlStr = sqlStr + " 	,'0000'" & VbCRLF
        sqlStr = sqlStr + " 	,'"+ shopitemname + "'" & VbCRLF
        sqlStr = sqlStr + " 	,'"+ shopitemoptionname + "'" & VbCRLF
        sqlStr = sqlStr + " 	," + CStr(shopitemprice) + "" & VbCRLF
        sqlStr = sqlStr + " 	,getdate()" & VbCRLF
        sqlStr = sqlStr + " 	,1" & VbCRLF
        sqlStr = sqlStr + " 	,1" & VbCRLF
        sqlStr = sqlStr + " 	)" & VbCRLF
        sqlStr = sqlStr + " END" & VbCRLF
        
        'response.write sqlStr &"<Br>"
        dbget.Execute sqlStr
    end if

	If Err.Number = 0 Then
	    dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장 되었습니다.');"
		response.write "	opener.location.reload();"
		
		if extbarcodeAlreadyExistsString<>"" then
			response.write "	alert('일부 상품은 저장되지 않았습니다. 이미등록된 바코드 - " & extbarcodeAlreadyExistsString &"');"
		end if
		
		response.write "	self.close();"
		response.write "</script>"
		dbget.close : response.end	    
	
	Else
	    dbget.RollBackTrans
	    
	    response.write "<script type='text/javascript'>"
	    response.write "	alert('데이터 에러 발생 \n관리자 문의하세요');"
	    response.write "	self.close();"
	    response.write "</script>"
		dbget.close : response.end
	End If

'/포스 상품 수정
elseif (mode="editetcoffitem") then
    if (Not IsNumeric(orgsellprice)) or (orgsellprice="") then orgsellprice =0
    if (Not IsNumeric(discountsellprice)) or (discountsellprice="") then discountsellprice =0
    if (Not IsNumeric(shopsuplycash)) or (shopsuplycash="") then shopsuplycash =0
    if (Not IsNumeric(shopbuyprice)) or (shopbuyprice="") then shopbuyprice =0
    
    if CStr(orgsellprice)="0" then orgsellprice=shopitemprice

	if itemid = "" or itemoption = "" or itemgubun = "" then
	    response.write "<script type='text/javascript'>"
	    response.write "	alert('[상품수정]상품정보가 잘못되었습니다');"
	    response.write "	self.close();"
	    response.write "</script>"
		dbget.close : response.end
	end if

    '''바코드 체크
    IF (extbarcode<>"") then
        sqlStr = " select top 1 * "
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
    	sqlStr = sqlStr + " where barcode='" + trim(extbarcode) + "'" + VbCrlf
    	sqlStr = sqlStr + " and barcode not in ("
    	sqlStr = sqlStr + " 	select barcode"
    	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option_stock"
    	sqlStr = sqlStr + "		where itemid = "&itemid&" and itemgubun = '"&itemgubun&"' and itemoption = '"&itemoption&"'"
    	sqlStr = sqlStr + " 	)"
    	
    	'response.write sqlStr &"<Br>"
    	rsget.Open sqlStr,dbget,1
    	if Not rsget.EOF then
    		extbarcodeAlreadyExists = true
    	end if
    	rsget.close
    			
        IF (extbarcodeAlreadyExists) then
            response.write "<script type='text/javascript'>alert('이미 사용중인 바코드 입니다.-" + extbarcode + " 등록 불가');</script>"
            dbget.close() : response.end
        end if			
	ENd IF
	
	dbget.beginTrans

	sqlStr = "update [db_shop].[dbo].tbl_shop_item set" + VbCrlf
	sqlStr = sqlStr & " makerid = '"&makerid&"'" + VbCrlf
	sqlStr = sqlStr & " ,shopitemname = '"&shopitemname&"'" + VbCrlf
	sqlStr = sqlStr & " ,shopitemoptionname = '"&shopitemoptionname&"'" + VbCrlf
	sqlStr = sqlStr & " ,orgsellprice = '"&orgsellprice&"'" + VbCrlf
	sqlStr = sqlStr & " ,shopitemprice = '"&shopitemprice&"'" + VbCrlf
	sqlStr = sqlStr & " ,shopsuplycash = '"&shopsuplycash&"'" + VbCrlf
	sqlStr = sqlStr & " ,shopbuyprice = '"&shopbuyprice&"'" + VbCrlf
	
	if (centermwdiv<>"") then
		sqlStr = sqlStr & " ,centermwdiv = '"&centermwdiv&"'" + VbCrlf
	end if
	
	if (extbarcode<>"") then
		sqlStr = sqlStr & " ,extbarcode = '"&extbarcode&"'" + VbCrlf
	end if
	
	sqlStr = sqlStr & " ,vatinclude = '"&vatinclude&"'" + VbCrlf
	sqlStr = sqlStr & " where shopitemid = "&itemid&" and itemgubun = '"&itemgubun&"' and itemoption = '"&itemoption&"'" + VbCrlf

	'response.write sqlStr &"<Br>"
	dbget.Execute sqlStr
       
    sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
	sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " and itemoption='"&itemoption&"'" + VbCrlf
	
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1
	    stockitemexists = (not rsget.Eof)
	rsget.close

	if (stockitemexists) then
		sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
		sqlStr = sqlStr + " set barcode='" + trim(extbarcode) + "'" + VbCrlf
		sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
		sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
		sqlStr = sqlStr + " and itemoption='"&itemoption&"'" + VbCrlf
		
		'response.write sqlStr &"<Br>"
		dbget.Execute sqlStr
	else
		sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
		sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
		sqlStr = sqlStr + " values("
		sqlStr = sqlStr + " '" + itemgubun + "'," + VbCrlf
		sqlStr = sqlStr + " " + CStr(itemid) + "," + VbCrlf
		sqlStr = sqlStr + " '"&itemoption&"'," + VbCrlf
		sqlStr = sqlStr + " '" + trim(extbarcode) + "'" + VbCrlf
		sqlStr = sqlStr + " )" + VbCrlf
		
		'response.write sqlStr &"<Br>"
		dbget.Execute sqlStr
	end if
    
    IF (IsForeignShop) and (shopid<>"") then
        sqlStr = " IF Exists(" & VbCRLF
        sqlStr = sqlStr + " 	select * from db_shop.dbo.tbl_shop_locale_item" & VbCRLF 
        sqlStr = sqlStr + " 	where shopid='"+shopid+"'" & VbCRLF
        sqlStr = sqlStr + " 	and itemgubun='" + itemgubun + "'" & VbCRLF
        sqlStr = sqlStr + " 	and shopitemid=" + CStr(itemid) + "" & VbCRLF
        sqlStr = sqlStr + " 	and itemoption='0000'" & VbCRLF
        sqlStr = sqlStr + " )" & VbCRLF
        sqlStr = sqlStr + " BEGIN" & VbCRLF
        sqlStr = sqlStr + " 	update db_shop.dbo.tbl_shop_locale_item " & VbCRLF
        sqlStr = sqlStr + " 	set lcprice=" + CStr(shopitemprice) & VbCRLF
        sqlStr = sqlStr + " 	,lcitemname='"+ shopitemname + "'" & VbCRLF
        sqlStr = sqlStr + " 	,lcitemoptionname='"+ shopitemoptionname + "'" & VbCRLF
        sqlStr = sqlStr + " 	,lastupdate=getdate()" & VbCRLF
        sqlStr = sqlStr + " 	,exchangerate=1" & VbCRLF
        sqlStr = sqlStr + " 	,multipleRate=1" & VbCRLF
        sqlStr = sqlStr + " 	where shopid='"+shopid+"'" & VbCRLF
        sqlStr = sqlStr + " 	and itemgubun='" + itemgubun + "'" & VbCRLF
        sqlStr = sqlStr + " 	and shopitemid=" + CStr(itemid) + "" & VbCRLF
        sqlStr = sqlStr + " 	and itemoption='"&itemoption&"'" & VbCRLF
        sqlStr = sqlStr + " END" & VbCRLF
        sqlStr = sqlStr + " ELSE" & VbCRLF
        sqlStr = sqlStr + " BEGIN" & VbCRLF
        sqlStr = sqlStr + " 	insert into db_shop.dbo.tbl_shop_locale_item " & VbCRLF
        sqlStr = sqlStr + " 	(shopid,itemgubun,shopitemid,itemoption,lcitemname,lcitemoptionname,lcprice,lastupdate,exchangerate,multipleRate)" & VbCRLF
        sqlStr = sqlStr + " 	values(" & VbCRLF
        sqlStr = sqlStr + " 	'"+shopid+"'" & VbCRLF
        sqlStr = sqlStr + " 	,'" + itemgubun + "'" & VbCRLF
        sqlStr = sqlStr + " 	," + CStr(itemid) + "" & VbCRLF
        sqlStr = sqlStr + " 	,'"&itemoption&"'" & VbCRLF
        sqlStr = sqlStr + " 	,'"+ shopitemname + "'" & VbCRLF
        sqlStr = sqlStr + " 	,'"+ shopitemoptionname + "'" & VbCRLF
        sqlStr = sqlStr + " 	," + CStr(shopitemprice) + "" & VbCRLF
        sqlStr = sqlStr + " 	,getdate()" & VbCRLF
        sqlStr = sqlStr + " 	,1" & VbCRLF
        sqlStr = sqlStr + " 	,1" & VbCRLF
        sqlStr = sqlStr + " 	)" & VbCRLF
        sqlStr = sqlStr + " END" & VbCRLF
        
        'response.write sqlStr &"<Br>"
        dbget.Execute sqlStr
    end if

	If Err.Number = 0 Then
	    dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('저장 되었습니다.');"
		
		if extbarcodeAlreadyExistsString<>"" then
			response.write "	alert('일부 상품은 저장되지 않았습니다. 이미등록된 바코드 - " & extbarcodeAlreadyExistsString &"');"
		end if

		response.write "	opener.location.reload();"
		response.write "	location.href='/common/offshop/popoffitemreg_Etc.asp?itemgubun="&itemgubun&"&itemid="&itemid&"&itemoption="&itemoption&"&makerid="&makerid&"&shopid="&shopid&"';"				
		response.write "</script>"
		dbget.close : response.end	   	    
	Else
	    dbget.RollBackTrans
	    
	    response.write "<script type='text/javascript'>"
	    response.write "	alert('데이터 에러 발생 \n관리자 문의하세요');"
	    response.write "	self.close();"
	    response.write "</script>"
		dbget.close : response.end
	End If    
end if

if (mode ="offitemreg") or (mode="arrins") then
	refer = refer + "&react=true"
end if
%>

<script type='text/javascript'>
	<% if mode="addetcoffitem" or mode="editetcoffitem" then %>
		alert('저장 되었습니다.');
		opener.location.reload();
	<% else %>
		alert('저장 되었습니다.');
	<% end if %>

	<% if extbarcodeAlreadyExistsString<>"" then %>
		alert('<%= "일부 상품은 저장되지 않았습니다. 이미등록된 바코드 - " + extbarcodeAlreadyExistsString %>');
	<% end if %>

	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
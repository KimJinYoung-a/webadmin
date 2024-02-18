<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 지역별 상품 저장
' History : 2010.08.05 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode , addSql ,sqlStr , itemoption , itemgubun ,itemidarr, itemoptionarr ,itemgubunarr , i ,shopid
dim lcitemnamearr, lcitemoptionnamearr ,lcpricearr , tmpcount , menupos , exchangeratearr , multipleratearr
dim usingyn , cdl ,cdm, cds ,designer ,itemid, itemname , shopitemname ,prdcode ,generalbarcode
dim gubun ,imageview ,nameeng ,parameter
	mode = Request("mode")
	itemid = Request("itemid")
	itemoption = Request("itemoption")
	itemgubun = Request("itemgubun")
	shopid = request("shopid")
	itemidarr = Request("ia")
	itemoptionarr = Request("ioa")
	itemgubunarr = Request("iga")		
	lcitemnamearr = request("lina")
	lcitemoptionnamearr = request("liona")
	lcpricearr = request("lpa")
	exchangeratearr = request("eratea")
	multipleratearr = request("mratea")	
	menupos = request("menupos")
	usingyn = request("usingyn")			
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	designer = request("designer")
	itemname = request("itemname")
	shopitemname = request("shopitemname")
	prdcode = request("prdcode")
	generalbarcode = request("generalbarcode")
	gubun = request("gubun")
	imageview = request("imageview")
	nameeng = request("nameeng")

parameter = "usingyn="&usingyn&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&designer="&designer&"&itemid="&itemid&"&itemname="&itemname&""
parameter = parameter & "&shopitemname="&shopitemname&"&prdcode="&prdcode&"&generalbarcode="&generalbarcode&"&gubun="&gubun&"&imageview="&imageview&""
parameter = parameter & "&nameeng="&nameeng&""
	
dim referer
referer = request.ServerVariables("HTTP_REFERER")
	
if mode = "itemadd" then
	 
	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	itemgubunarr = split(itemgubunarr,",")
	
	dbget.begintrans	 
	
	for i = 0 to ubound(itemidarr)-1
		sqlStr = ""	 
		sqlStr = " insert into db_shop.dbo.tbl_shop_locale_item" + VbCrlf
		sqlStr = sqlStr + " (shopid,itemgubun,shopitemid,itemoption,lcitemname,lcitemoptionname,lcprice,lastupdate)" + VbCrlf
		sqlStr = sqlStr + " select '" + CStr(shopid) + "', i.itemgubun,i.shopitemid"
		sqlStr = sqlStr + " , NULL as itemoption ,NULL as shopitemname" + VbCrlf
		sqlStr = sqlStr + " ,i.shopitemoptionname ,i.shopitemprice ,getdate()" + VbCrlf
		sqlStr = sqlStr + " from [db_shop].dbo.tbl_shop_item i" + VbCrlf
		sqlStr = sqlStr + " left join db_shop.dbo.tbl_shop_locale_item l" + VbCrlf
		sqlStr = sqlStr + " on i.shopitemid = l.shopitemid and i.itemoption = l.itemoption and i.itemgubun = l.itemgubun" + VbCrlf
		sqlStr = sqlStr + " and l.shopid = '"&shopid&"'" + VbCrlf
		sqlStr = sqlStr + " where l.shopitemid is null" + VbCrlf
		sqlStr = sqlStr + " and i.shopitemid = "&itemidarr(i)&"" + VbCrlf
		sqlStr = sqlStr + " and i.itemoption = '"&itemoptionarr(i)&"'" + VbCrlf
		sqlStr = sqlStr + " and i.itemgubun = '"&itemgubunarr(i)&"'" + VbCrlf		
	
		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
    next
    	
	IF Err.Number = 0 THEN
		dbget.CommitTrans    
		response.write "<script langauge='javascript'>alert('OK'); opener.location.reload(); self.close();</script>"
		dbget.close()	:	response.End
	Else
   		dbget.RollBackTrans	 
   		response.write "<script langauge='javascript'>alert('데이터 처리에 문제가 발생하였습니다.'); history.back(-1);</script>"
   		dbget.close()	:	response.End
   	end if

elseif mode = "litemadd" then

	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	itemgubunarr = split(itemgubunarr,",")
	lcitemnamearr = split(lcitemnamearr,",")
	lcitemoptionnamearr = split(lcitemoptionnamearr,",")
	lcpricearr = split(lcpricearr,",")	
	exchangeratearr = split(exchangeratearr,",")
	multipleratearr = split(multipleratearr,",")	
		
	dbget.begintrans	 
	
	for i = 0 to ubound(itemidarr)-1	 
		tmpcount = 0
		
		sqlStr = ""			
        sqlStr = "select count(shopid) as cnt"
        sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_locale_item"
        sqlStr = sqlStr & " where shopid = '"&shopid&"'" & vbcrlf        
		sqlStr = sqlStr + " and shopitemid = "&itemidarr(i)&"" + VbCrlf
		sqlStr = sqlStr + " and itemoption = '"&itemoptionarr(i)&"'" + VbCrlf
		sqlStr = sqlStr + " and itemgubun = '"&itemgubunarr(i)&"'" + VbCrlf		
	
       	'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1

        if Not rsget.Eof then
    		tmpcount = rsget("cnt")		          
        end if
        rsget.Close
        
        if tmpcount > 0 then
        	sqlStr = ""        	
        	sqlStr = "update db_shop.dbo.tbl_shop_locale_item set" + VbCrlf
        	if (lcitemnamearr(i)="") then
        	    sqlStr = sqlStr + " lcitemname = NULL"
        	else
        	    sqlStr = sqlStr + " lcitemname = '"&html2db(ReplaceRequestSpecialChar(lcitemnamearr(i)))&"'" + VbCrlf
            end if
            if (lcitemoptionnamearr(i)="") then
                sqlStr = sqlStr + " ,lcitemoptionname = NULL"
        	else
        	    sqlStr = sqlStr + " ,lcitemoptionname = '"&html2db(ReplaceRequestSpecialChar(lcitemoptionnamearr(i)))&"'" + VbCrlf
        	end if
        	sqlStr = sqlStr + " ,lcprice = "&lcpricearr(i)&"" + VbCrlf
        	sqlStr = sqlStr + " ,lastupdate = getdate()" + VbCrlf
        	sqlStr = sqlStr + " ,exchangerate = '"&exchangeratearr(i)&"'" + VbCrlf
        	sqlStr = sqlStr + " ,multiplerate = '"&multipleratearr(i)&"'" + VbCrlf
        	sqlStr = sqlStr + " where shopid = '"&shopid&"'" + VbCrlf
        	sqlStr = sqlStr + " and itemgubun = '"&itemgubunarr(i)&"'" + VbCrlf
        	sqlStr = sqlStr + " and shopitemid = "&itemidarr(i)&"" + VbCrlf
        	sqlStr = sqlStr + " and itemoption = '"&itemoptionarr(i)&"'" + VbCrlf
        			
			'response.write sqlStr &"<br>"
			dbget.execute sqlStr
        else
        	sqlStr = ""
			sqlStr = " insert into db_shop.dbo.tbl_shop_locale_item" + VbCrlf
			sqlStr = sqlStr + " (shopid, itemgubun, shopitemid, itemoption, lcitemname, lcitemoptionname, lcprice, lastupdate ,exchangerate ,multiplerate) " + VbCrlf
			sqlStr = sqlStr + " values ('"&shopid&"', '"&itemgubunarr(i)&"' ,"&itemidarr(i)&" ,'"&itemoptionarr(i)&"'" + VbCrlf
			if (lcitemnamearr(i)="") then
			    sqlStr = sqlStr + " ,NULL"+ VbCrlf
		    else
			    sqlStr = sqlStr + " ,'"&html2db(ReplaceRequestSpecialChar(lcitemnamearr(i)))&"'" + VbCrlf
		    end if
		    if (lcitemoptionnamearr(i)="") then
		        sqlStr = sqlStr + " ,NULL"+ VbCrlf
		    else
			    sqlStr = sqlStr + " ,'"&html2db(ReplaceRequestSpecialChar(lcitemoptionnamearr(i)))&"'" + VbCrlf
			end if
			sqlStr = sqlStr + " ,"&lcpricearr(i)&" , getdate()" + VbCrlf
			sqlStr = sqlStr + " ,'"&exchangeratearr(i)&"','"&multipleratearr(i)&"')" + VbCrlf
			
			'response.write sqlStr &"<br>"
			dbget.execute sqlStr
        
        end if
    next
    	
	IF Err.Number = 0 THEN
		dbget.CommitTrans    
		response.write "<script langauge='javascript'>"
		response.write "	alert('OK');"
		response.write "	location.href='/common/offshop/localeItem/localeItemList.asp?shopid="&shopid&"&menupos="&menupos&"&parameter="&parameter&"';"
		response.write "</script>"
		dbget.close()	:	response.End
	Else
   		dbget.RollBackTrans	 
   		response.write "<script langauge='javascript'>alert('데이터 처리에 문제가 발생하였습니다.'); history.back(-1);</script>"
   		dbget.close()	:	response.End
   	end if

'// 선택상품 삭제
elseif mode = "itemdel" then

	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	itemgubunarr = split(itemgubunarr,",")

	dbget.begintrans	 
	
	for i = 0 to ubound(itemidarr)-1
		
		sqlStr = "Delete From [db_shop].[dbo].tbl_eventitem_off" + VbCrlf
		sqlStr = sqlStr + " WHERE evt_code = "&evt_code&"" + VbCrlf			
		sqlStr = sqlStr + " and itemid = "&itemidarr(i)&"" + VbCrlf
		sqlStr = sqlStr + " and itemoption = '"&itemoptionarr(i)&"'" + VbCrlf
		sqlStr = sqlStr + " and itemgubun = '"&itemgubunarr(i)&"'" + VbCrlf

		'response.write sqlStr &"<br>"
		dbget.execute sqlStr
    next
    	
	IF Err.Number = 0 THEN
		dbget.CommitTrans    
		response.write "<script langauge='javascript'>alert('OK'); location.replace('" + referer + "');</script>"
		dbget.close()	:	response.End
	Else
   		dbget.RollBackTrans	 
   		response.write "<script langauge='javascript'>alert('데이터 처리에 문제가 발생하였습니다.'); history.back(-1);</script>"
   		dbget.close()	:	response.End
   	end if
	
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

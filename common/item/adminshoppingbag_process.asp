<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 온라인 & 오프라인 어드민 장바구니
' Hieditor : 2011.08.04 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->

<%
dim i , mode , bagidxarr , sql ,itemnoarr ,userid ,itemidarr ,itemgubunarr ,itemoptionarr ,shopid ,onoffgubun , menupos
    mode = requestCheckVar(request("mode"),32)
    bagidxarr = request("bagidxarr")
    itemnoarr = request("itemnoarr")    
	userid = session("ssBctId")
	onoffgubun = requestCheckVar(request("onoffgubun"),10)
	shopid = requestCheckVar(request("shopid"),32)
	itemidarr = Request("itemidarr")
	itemgubunarr = Request("itemgubunarr")
	itemoptionarr = Request("itemoptionarr")
	menupos = requestCheckVar(Request("menupos"),10)

'//장바구니 추가
if mode = "I" then
	if userid = "" or onoffgubun = "" or shopid = "" or itemgubunarr="" or itemidarr="" or itemoptionarr = "" or itemnoarr = "" then
        response.write "<script language='javascript'>"
        response.write "    alert('필요한 값이 없습니다');"
        response.write "</script>"
		dbget.close() : response.end		
	end if

	'//장바구니 추가
	'response.write userid &"/"& onoffgubun&"/"& shopid&"/"& itemgubunarr&"/"& itemidarr&"/"& itemoptionarr&"/"& itemnoarr &"<Br>"
	putadminshoppingbag_insert userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr , "parent.opener" ,menupos

'//장바구니 바로추가
elseif mode = "directbagaddarr" then
	if userid = "" or onoffgubun = "" or shopid = "" or itemgubunarr="" or itemidarr="" or itemoptionarr = "" or itemnoarr = "" then
        response.write "<script language='javascript'>"
        response.write "    alert('필요한 값이 없습니다');"
        response.write "</script>"
		dbget.close() : response.end		
	end if

	'//장바구니 추가
	'response.write userid &"/"& onoffgubun&"/"& shopid&"/"& itemgubunarr&"/"& itemidarr&"/"& itemoptionarr&"/"& itemnoarr &"<Br>"
	putadminshoppingbag_insert userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr , "" ,menupos

    response.write "<script language='javascript'>"
    response.write "    alert('장바구니에 상품이 저장되었습니다');"
    response.write "    self.close();"
    response.write "</script>"
	dbget.close() : response.end
			    		
'//장바구니 삭제모드    
elseif mode = "bagdelarr" then

	if bagidxarr = "" then
        response.write "<script language='javascript'>"
        response.write "    alert('인덱스 값이 없습니다');"
        response.write "</script>"
		dbget.close() : response.end
	end if	
	
	bagidxarr = split(bagidxarr,",")
	
	if isarray(bagidxarr) then
		for i = 0 to ubound(bagidxarr) - 1
			sql = ""		
			sql = "delete from db_temp.dbo.tbl_adminshoppingbag where" + vbcrlf
			sql = sql & " bagidx = "& requestCheckVar(bagidxarr(i),10) &""

			'response.write sql & "<Br>"
			dbget.execute sql		
		next
	end if

    response.write "<script type='text/javascript'>"
    response.write "    alert('OK');"
    response.write "	parent.location.reload();"
    response.write "</script>"
	dbget.close() : response.end

'//주문후 장바구니 삭제    
elseif mode = "baginsertdelarr" then

	if bagidxarr = "" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('인덱스 값이 없습니다');"
        response.write "</script>"
		dbget.close() : response.end
	end if	
	
	bagidxarr = split(bagidxarr,",")
	
	if isarray(bagidxarr) then
		for i = 0 to ubound(bagidxarr) - 1
			sql = ""		
			sql = "delete from db_temp.dbo.tbl_adminshoppingbag where" + vbcrlf
			sql = sql & " bagidx = "& requestCheckVar(bagidxarr(i),10) &""
			
			'response.write sql & "<Br>"
			dbget.execute sql		
		next
	end if

	dbget.close() : response.end
		
'//장바구니 수정모드	
elseif mode = "bageditarr" then

	if bagidxarr = "" or itemnoarr = "" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('인덱스 값이 없습니다');"
        response.write "</script>"
		dbget.close() : response.end
	end if
	
	bagidxarr = split(bagidxarr,",")
	itemnoarr = split(itemnoarr,",")


	if isarray(bagidxarr) then
		for i = 0 to ubound(bagidxarr) - 1		
			sql = ""
			sql = "update db_temp.dbo.tbl_adminshoppingbag" + vbcrlf
			sql = sql & " set itemno = "& requestCheckVar(itemnoarr(i),10) &" where" + vbcrlf
			sql = sql & " bagidx = "& requestCheckVar(bagidxarr(i),10) &""
			
			'response.write sql & "<Br>"
			dbget.execute sql		
		next
	end if
	
    response.write "<script type='text/javascript'>"
    response.write "    alert('OK');"
    response.write "	parent.location.reload();"
    response.write "</script>"
	dbget.close() : response.end	
end if    
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
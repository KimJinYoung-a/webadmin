<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 할인 관리
' History : 2010.12.02 한용민 생성
'####################################################
 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/sale_Cls.asp"-->

<%
Dim sMode,sCode, sale_startdate, sale_enddate ,strSql,addSql, addSqlDB ,egCode,eCode,itemidarr,sType, i
Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype
Dim  dissellprice,disbuyprice,arrsaleItemStatus,sale_status, tmpsaleItemStatus , itemoptionarr , itemgubunarr
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,strParm,ssStatus ,objCmd,iResult , shopid
dim sale_rate, sale_margin, sale_marginValue ,strStatus, intLoop ,osale , itemgubun , itemoption
dim sale_shopmargin ,sale_shopmarginvalue , idsaleshopsupplycash , point_rate ,point_ratearr, saleitem_idxarr
Dim Err_saleitemexists, Err_contractnotexists, contractoverlapyn
	sMode     = requestCheckVar(Request("mode"),2)	

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

Err_saleitemexists = ""
Err_contractnotexists = ""

'//계약조건이 겹치게 등록 가능여부.. Y 로 놓을경우 하나의 할인에 매입 특정 섞여서 등록가능하며, 매입/특정에 맞게 매입가 계산됨
contractoverlapyn = "N"
 
SELECT Case sMode
	
	'할인상품 추가
	Case "I"
		itemidarr = Request("itemidarr")
		itemgubunarr = Request("itemgubunarr")
		itemoptionarr = Request("itemoptionarr")		
		itemidarr = split(itemidarr,",")
		itemgubunarr = split(itemgubunarr,",")			
		itemoptionarr = split(itemoptionarr,",")				
		sCode 		= requestCheckVar(Request("sC"),10)
		eCode 		= requestCheckVar(Request("eC"),10)
		egCode 		= Request("egC")
		if egCode = "" then egCode = 0
		itemid      = request("itemid")
		itemname    = request("itemname")
		makerid     = request("makerid")
		sellyn      = request("sellyn")
		usingyn     = request("usingyn")
		danjongyn   = request("danjongyn")
		limityn     = request("limityn")
		mwdiv       = request("mwdiv")
		sailyn      = request("sailyn")
		deliverytype= request("deliverytype")		
		cdl = request("cdl")
		cdm = request("cdm")
		cds = request("cds")
		point_rate = request("point_rate")

		addSql = ""
		addSqlDB = ""
	    
	    '/트랜젝션
	    dbget.beginTrans

		set osale = new csale_list
		osale.frectsale_code = sCode
		
		if sCode <> "" then
			'/해당 할인 정보 가져옴
			osale.getsaledetail
			
			if osale.ftotalcount > 0 then
				sale_startdate = osale.foneitem.fsale_startdate
				sale_enddate = osale.foneitem.fsale_enddate
				sale_rate = osale.foneitem.fsale_rate
				point_rate = osale.foneitem.fpoint_rate
				sale_margin = osale.foneitem.fsale_margin
				sale_marginValue = osale.foneitem.fsale_marginValue	
				sale_status	= osale.foneitem.fsale_status
				shopid	= osale.foneitem.fshopid
				sale_shopmargin	= osale.foneitem.fsale_shopmargin
				sale_shopmarginvalue = osale.foneitem.fsale_shopmarginvalue
			end if
		end if
				
		IF ubound(itemidarr) > 0 THEN
			
			for i = 0 to ubound(itemidarr) - 1
			
			if trim(itemidarr(i))<>"" then
				
				'/할인 등록내역에서.. 할인중인 상품내역 리스트 뽑아냄, 내역을 얼럿창으로 보여줌.	
				strSql = "SELECT TOP 100"&_
						" b.itemid, a.sale_code, a.sale_status "&_
						" FROM [db_shop].[dbo].tbl_sale_off a" &_
						" join [db_shop].[dbo].[tbl_saleitem_off] b "&_
						" 	on a.sale_code = b.sale_code "&_
						" join [db_shop].[dbo].[tbl_shop_item] i "&_
						" 	on b.itemgubun = i.itemgubun "&_
						" 	and b.itemid = i.shopitemid "&_
						" 	and b.itemgubun = i.itemgubun "&_
						" 	and i.isusing='Y' "&_						
						" WHERE a.sale_startdate <= '"&sale_enddate&"'"&_
						" and a.sale_enddate >= '"&sale_startdate&"'"&_
						" and a.sale_using =1"&_
						" and a.sale_status <> 8"&_
						" and b.saleitem_status not in (8,9)"&_
						" and b.itemid = "&trim(itemidarr(i))&""&_
						" and b.itemgubun = '"&trim(itemgubunarr(i))&"'"&_
						" and b.itemoption = '"&trim(itemoptionarr(i))&"'"&_
						" and a.shopid = '"&shopid&"'"	

				'response.write strSql &"<Br>"
				rsget.Open strSql,dbget
				
				IF not rsget.EOF THEN
					IF rsget("sale_status") = 6 THEN 
						strStatus = "할인중"
					ELSEIF rsget("sale_status") = 7 THEN 
						strStatus = "할인예정"
					ELSEIF rsget("sale_status") = 0 THEN 
						strStatus = "등록대기"	
					END IF	
					
					Err_saleitemexists = Err_saleitemexists + "할인코드 : " + CStr(rsget("sale_code")) + " - 상품번호 : " + CStr(rsget("itemid")) +" "+ strStatus + " \n"									 	
				End IF
				
				rsget.Close
				
				if contractoverlapyn="N" then
					'/동일마진, 업체부담, 반반부담, 직접설정은 특정(업체특정, 출고특정, 텐바이텐특정)만 가능, 내역을 얼럿창으로 보여줌.
					if sale_margin = 1 or sale_margin = 2 or sale_margin = 3 or sale_margin = 5 then
						strSql = "SELECT TOP 100"
						strSql = strSql & " i.shopitemid"
						strSql = strSql & " from [db_shop].[dbo].[tbl_shop_item] i"
						strSql = strSql & " join db_shop.dbo.tbl_shop_designer sd"
						strSql = strSql & " 	on i.makerid=sd.makerid"
						strSql = strSql & " 	and sd.shopid = '"&shopid&"'"
						strSql = strSql & " 	and i.isusing='Y'"
						strSql = strSql & " WHERE"
						strSql = strSql & " sd.comm_cd not in ('B012','B013','B011')"
						strSql = strSql & " and i.itemgubun = '"&trim(itemgubunarr(i))&"'"
						strSql = strSql & " and i.shopitemid = "&trim(itemidarr(i))&""
						strSql = strSql & " and i.itemoption = '"&trim(itemoptionarr(i))&"'"
					
						'response.write strSql &"<Br>"
						rsget.Open strSql,dbget
						
						IF not rsget.EOF THEN												
							Err_contractnotexists = Err_contractnotexists + "상품번호 : " + CStr(rsget("shopitemid")) + " \n"									 	
						End IF
						
						rsget.Close
					end if
				end if
				
				'//현재 이벤트랑 할인과 연동해서 사용안함	'/코딩은 완성되어 있슴. 태스트후 사용은 가능..
		        if eCode <> "" then
		        	addSqlDB = " , [db_shop].[dbo].[tbl_eventitem_off] c "
		        	addSql = addSql + " and i.shopitemid = c.itemid and c.evt_code = "&eCode&""
		        end if
			
				'/이벤트 상품 등록
				strSql = "INSERT INTO [db_shop].[dbo].[tbl_saleItem_off]" + vbcrlf
				strSql = strSql & " ([sale_code], [itemid],itemgubun , itemoption, [saleItem_status], [saleprice],[salesupplycash]" + vbcrlf
				strSql = strSql & " ,saleshopsupplycash,lastadminid ,point_rate, orgcomm_cd)" + vbcrlf
				strSql = strSql & " 	SELECT "&sCode&", i.shopitemid,i.itemgubun,i.itemoption, 7, db_shop.dbo.uf_GetItemPriceCutting(  i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100) )" + vbcrlf
				
				'/매입마진
				'동일마진
				IF sale_margin = 1 THEN
					strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))- convert(int,(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*(100-convert(float,convert(int,i.shopsuplycash/i.orgsellprice*10000)/100))/100) else i.shopsuplycash end) )" + vbcrlf
				
				'업체부담
				ELSEIF sale_margin = 2 THEN				
					'strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - (i.orgsellprice- i.shopsuplycash) else i.shopsuplycash end) )" + vbcrlf
				
				'반반부담
				ELSEIF 	sale_margin = 3 THEN
					strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
				
				'10x10부담
				ELSEIF 	sale_margin = 4 THEN
					strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopsuplycash )" + vbcrlf
				
				'직접설정	
				ELSEIF sale_margin = 5 THEN	
					strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - convert(int, (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*convert(float,"&sale_marginValue&")/100) else i.shopsuplycash end) )" + vbcrlf
				
				'업체특정반반부담/나머지텐바이텐부담
				ELSEIF sale_margin = 6 THEN
					'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
				
				'업체특정,출고특정,텐바이텐특정반반부담/나머지텐바이텐부담
				ELSEIF sale_margin = 7 THEN
					'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
				END IF
				
				'/샾공급마진
				'동일마진
				IF sale_shopmargin = 1 THEN
					strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))- convert(int,(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*(100-convert(float,convert(int,i.shopbuyprice/i.orgsellprice*10000)/100))/100) else i.shopbuyprice end) )" + vbcrlf
				
				'업체부담
				ELSEIF sale_shopmargin = 2 THEN
					'strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - (i.orgsellprice- i.shopbuyprice) else i.shopbuyprice end) )" + vbcrlf
				
				'반반부담
				ELSEIF 	sale_shopmargin = 3 THEN
					strSql = strSql&"	 ,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
				
				'10x10부담
				ELSEIF 	sale_shopmargin = 4 THEN
					strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopbuyprice )" + vbcrlf
				
				'직접설정	
				ELSEIF 	sale_shopmargin = 5 THEN	
					strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - convert(int, (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*convert(float,"&sale_shopmarginvalue&")/100) else i.shopbuyprice end) )" + vbcrlf
				
				'업체특정반반부담/나머지매장부담
				ELSEIF sale_shopmargin = 6 THEN
					'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
				
				'업체특정,출고특정,텐바이텐특정반반부담/나머지텐바이텐부담
				ELSEIF sale_shopmargin = 7 THEN
					'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
				END IF
					
				strSql = strSql & "		,'"&session("ssBctId")&"', "&point_rate&", i.comm_cd" + vbcrlf
				strSql = strSql & " 	from (" + vbcrlf
				strSql = strSql & " 		select" + vbcrlf
				strSql = strSql & " 		ii.shopitemprice , ii.makerid, ii.shopitemname , ii.shopitemid ,ii.itemgubun ,ii.itemoption,sdd.comm_cd" + vbcrlf
				strSql = strSql & " 		,(CASE" + vbcrlf
				strSql = strSql & " 			when sdd.comm_cd in ('B012','B013','B011') and ii.shopsuplycash=0" + vbcrlf
				strSql = strSql & " 				THEN convert(int,ii.shopitemprice*(100-IsNULL(sdd.defaultmargin,100))/100)" + vbcrlf
				strSql = strSql & " 			ELSE ii.shopsuplycash END) as 'shopsuplycash'" + vbcrlf
				strSql = strSql & " 		,(CASE" + vbcrlf
				strSql = strSql & " 			when sdd.comm_cd in ('B012','B013','B011') and ii.shopbuyprice=0" + vbcrlf
				strSql = strSql & " 				THEN convert(int,ii.shopitemprice*(100-IsNULL(sdd.defaultsuplymargin,100))/100)" + vbcrlf
				strSql = strSql & " 			ELSE ii.shopbuyprice END) as 'shopbuyprice'" + vbcrlf
				strSql = strSql & " 		,ii.orgsellprice ,sdd.shopid" + vbcrlf
				strSql = strSql & " 		from [db_shop].dbo.tbl_shop_item ii" + vbcrlf
				strSql = strSql & " 		join db_shop.dbo.tbl_shop_designer sdd" + vbcrlf
				strSql = strSql & " 			on sdd.shopid = '"&shopid&"'" + vbcrlf
				strSql = strSql & " 			and ii.makerid=sdd.makerid" + vbcrlf
				strSql = strSql & " 			and ii.isusing='Y'" + vbcrlf
				
				if contractoverlapyn="N" then
					'/정산구분에 안맞는 상품 제낌
					'동일마진 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
					IF sale_margin = 1 THEN
						strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
						
					'업체부담 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
					ELSEIF sale_margin = 2 THEN				
						strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
						
					'반반부담 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
					ELSEIF 	sale_margin = 3 THEN
						strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
						
					'10x10부담 (따로 처리않해도 문제없을듯)
					ELSEIF 	sale_margin = 4 THEN
										
					'직접설정 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
					ELSEIF sale_margin = 5 THEN	
						strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
						
					END IF
				end if
				
				strSql = strSql & " 		where ii.orgsellprice = ii.shopitemprice" + vbcrlf		'/상품할인은 제낌(원칙상 이중할인은 안됨)
				strSql = strSql & "		) as i" + vbcrlf
				strSql = strSql & " 	left join (" + vbcrlf
				strSql = strSql & " 		select b.itemid ,b.itemgubun , b.itemoption ,a.shopid" + vbcrlf
				strSql = strSql & " 		from [db_shop].[dbo].tbl_sale_off a" + vbcrlf
				strSql = strSql & " 		join [db_shop].[dbo].[tbl_saleitem_off] b" + vbcrlf
				strSql = strSql & " 			on a.sale_code = b.sale_code" + vbcrlf
				strSql = strSql & " 		where a.sale_startdate <= '"&sale_enddate&"'"
				strSql = strSql & " 		and a.sale_enddate >= '"&sale_startdate&"'" + vbcrlf
				strSql = strSql & " 		and a.sale_using = 1"
				strSql = strSql & " 		and a.sale_status <> 8"
				strSql = strSql & " 		and b.saleitem_status not in (8,9)"
				strSql = strSql & " 		and a.shopid = '"&shopid&"'" + vbcrlf
				strSql = strSql & " 	) as t" + vbcrlf
				strSql = strSql & " 		on i.shopitemid = t.itemid" + vbcrlf
				strSql = strSql & "			and i.itemgubun = t.itemgubun" + vbcrlf
				strSql = strSql & "			and i.itemoption = t.itemoption"
				strSql = strSql & "			and i.shopid = t.shopid " & addSqlDB		
				strSql = strSql & " 	WHERE" + vbcrlf
				strSql = strSql & "		i.shopitemprice > 0" + vbcrlf
				strSql = strSql & " 	and t.itemid is null" + vbcrlf		'/현재 할인테이블에 등록 되어 있는 상품 제낌			
				strSql = strSql & " 	and i.shopitemid = "&trim(itemidarr(i))&""
				strSql = strSql & " 	and i.itemgubun = '"&trim(itemgubunarr(i))&"'"
				strSql = strSql & " 	and i.itemoption = '"&trim(itemoptionarr(i))&"'" & addSql							
				
				'response.write strSql &"<Br>"
				dbget.execute strSql
			
			end if
			
			next
		
		END IF	
		
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans			
			Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank" 				
			dbget.close()	:	response.End
		END IF
		
		dbget.CommitTrans
		
		if Err_saleitemexists<>"" then
%>
			<script langauge="javascript">
				alert('할인기간이 중복설정된 상품은 제외되고 등록됩니다.\n<%= Err_saleitemexists %>');
			</script>
<%
		end if
		
		if contractoverlapyn="N" then
			if Err_contractnotexists<>"" then
%>
				<script langauge="javascript">
					alert('매입 상품의 경우 텐바이텐부담 설정시에만 등록 가능합니다.\n정산구분과 할인매입마진이 일치 하지 않는 상품은 제외되고 등록됩니다.\n\n<%= Err_contractnotexists %>');
				</script>
<%
			end if
		end if
%>
		<script langauge="javascript">
			alert('OK\n\n상품이 사용중이 아니거나, 판매금액이 0이하인 상품은 제외되고 등록됩니다');
			location.href ="about:blank";
			//parent.close();
			//parent.history.go(0);
			//parent.location.reload();	
		</script>
<% 	   
		dbget.close()	:	response.End

	'브랜드 할인상품 추가
	Case "bi"

		shopid = Request("shopid")
		makerid = Request("makerid")
		sCode 		= requestCheckVar(Request("sC"),10)
		
		if shopid = "" then
			response.write "<script>alert('매장ID가 없습니다'); self.close();</script>"
			response.end
		end if
		if makerid = "" then
			response.write "<script>alert('브랜드ID가 없습니다'); self.close();</script>"
			response.end
		end if
		if sCode = "" then
			response.write "<script>alert('할인코드가 없습니다'); self.close();</script>"
			response.end
		end if

	    '/트랜젝션
	    dbget.beginTrans

		set osale = new csale_list
		osale.frectsale_code = sCode
		
		if sCode <> "" then
			'/해당 할인 정보 가져옴
			osale.getsaledetail
			
			if osale.ftotalcount > 0 then
				sale_startdate = osale.foneitem.fsale_startdate
				sale_enddate = osale.foneitem.fsale_enddate
				sale_rate = osale.foneitem.fsale_rate
				point_rate = osale.foneitem.fpoint_rate
				sale_margin = osale.foneitem.fsale_margin
				sale_marginValue = osale.foneitem.fsale_marginValue	
				sale_status	= osale.foneitem.fsale_status
				shopid	= osale.foneitem.fshopid
				sale_shopmargin	= osale.foneitem.fsale_shopmargin
				sale_shopmarginvalue = osale.foneitem.fsale_shopmarginvalue
			end if
		end if
				
		'/할인 등록내역에서.. 할인중인 상품내역 리스트 뽑아냄, 내역을 얼럿창으로 보여줌.	
		strSql = "SELECT TOP 100"&_
				" b.itemid, a.sale_code, a.sale_status "&_
				" FROM [db_shop].[dbo].tbl_sale_off a" &_
				" join [db_shop].[dbo].[tbl_saleitem_off] b "&_
				" 	on a.sale_code = b.sale_code "&_
				" join [db_shop].[dbo].[tbl_shop_item] i "&_
				" 	on b.itemgubun = i.itemgubun "&_
				" 	and b.itemid = i.shopitemid "&_
				" 	and b.itemgubun = i.itemgubun "&_
				" 	and i.isusing='Y' "&_
				" WHERE a.sale_startdate <= '"&sale_enddate&"'"&_
				" and a.sale_enddate >= '"&sale_startdate&"'"&_
				" and a.sale_using =1"&_
				" and a.sale_status <> 8"&_
				" and b.saleitem_status not in (8,9)"&_
				" and i.makerid = '"&makerid&"'"&_
				" and a.shopid = '"&shopid&"'"	
		
		'response.write strSql &"<Br>"
		rsget.Open strSql,dbget
		
		i=0
		IF not rsget.EOF THEN
			Do Until rsget.Eof
					
			IF rsget("sale_status") = 6 THEN 
				strStatus = "할인중"
			ELSEIF rsget("sale_status") = 7 THEN 
				strStatus = "할인예정"
			ELSEIF rsget("sale_status") = 0 THEN 
				strStatus = "등록대기"	
			END IF	
			
			Err_saleitemexists = Err_saleitemexists + "할인코드 : " + CStr(rsget("sale_code")) + " - 상품번호 : " + CStr(rsget("itemid")) +" "+ strStatus + " \n"

			rsget.movenext
			i = i + 1
			Loop									 	
		End IF
		
		rsget.Close
		
		if contractoverlapyn="N" then
			'/동일마진, 업체부담, 반반부담, 직접설정은 특정(업체특정, 출고특정, 텐바이텐특정)만 가능, 내역을 얼럿창으로 보여줌.
			if sale_margin = 1 or sale_margin = 2 or sale_margin = 3 or sale_margin = 5 then
				strSql = "SELECT TOP 100"
				strSql = strSql & " i.shopitemid"
				strSql = strSql & " from [db_shop].[dbo].[tbl_shop_item] i"
				strSql = strSql & " join db_shop.dbo.tbl_shop_designer sd"
				strSql = strSql & " 	on i.makerid=sd.makerid"
				strSql = strSql & " 	and sd.shopid = '"&shopid&"'"
				strSql = strSql & " 	and i.isusing='Y'"
				strSql = strSql & " WHERE"
				strSql = strSql & " sd.comm_cd not in ('B012','B013','B011')"
				strSql = strSql & " and i.makerid = '"&makerid&"'"
			
				'response.write strSql &"<Br>"
				rsget.Open strSql,dbget
				
				i=0
				IF not rsget.EOF THEN		
					Do Until rsget.Eof	

					Err_contractnotexists = Err_contractnotexists + "상품번호 : " + CStr(rsget("shopitemid")) + " \n"
	
					rsget.movenext
					i = i + 1
					Loop
				End IF
				
				rsget.Close
			end if
		end if
		
		'/이벤트 상품 등록
		strSql = "INSERT INTO [db_shop].[dbo].[tbl_saleItem_off]" + vbcrlf
		strSql = strSql & " ([sale_code], [itemid],itemgubun , itemoption, [saleItem_status], [saleprice],[salesupplycash]" + vbcrlf
		strSql = strSql & " ,saleshopsupplycash,lastadminid ,point_rate, orgcomm_cd)" + vbcrlf
		strSql = strSql & " 	SELECT "&sCode&", i.shopitemid,i.itemgubun,i.itemoption, 7, db_shop.dbo.uf_GetItemPriceCutting(  i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100) )" + vbcrlf
		
		'/매입마진
		'동일마진
		IF sale_margin = 1 THEN
			strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))- convert(int,(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*(100-convert(float,convert(int,i.shopsuplycash/i.orgsellprice*10000)/100))/100) else i.shopsuplycash end) )" + vbcrlf
		
		'업체부담
		ELSEIF sale_margin = 2 THEN				
			'strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - (i.orgsellprice- i.shopsuplycash) else i.shopsuplycash end) )" + vbcrlf
		
		'반반부담
		ELSEIF 	sale_margin = 3 THEN
			strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
		
		'10x10부담
		ELSEIF 	sale_margin = 4 THEN
			strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopsuplycash )" + vbcrlf
		
		'직접설정	
		ELSEIF sale_margin = 5 THEN	
			strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - convert(int, (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*convert(float,"&sale_marginValue&")/100) else i.shopsuplycash end) )" + vbcrlf
		
		'업체특정반반부담/나머지텐바이텐부담
		ELSEIF sale_margin = 6 THEN
			'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
		
		'업체특정,출고특정,텐바이텐특정반반부담/나머지텐바이텐부담
		ELSEIF sale_margin = 7 THEN
			'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopsuplycash - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopsuplycash end) )" + vbcrlf
		END IF
		
		'/샾공급마진
		'동일마진
		IF sale_shopmargin = 1 THEN
			strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))- convert(int,(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*(100-convert(float,convert(int,i.shopbuyprice/i.orgsellprice*10000)/100))/100) else i.shopbuyprice end) )" + vbcrlf
		
		'업체부담
		ELSEIF sale_shopmargin = 2 THEN
			'strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - (i.orgsellprice- i.shopbuyprice) else i.shopbuyprice end) )" + vbcrlf
		
		'반반부담
		ELSEIF 	sale_shopmargin = 3 THEN
			strSql = strSql&"	 ,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
		
		'10x10부담
		ELSEIF 	sale_shopmargin = 4 THEN
			strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  i.shopbuyprice )" + vbcrlf
		
		'직접설정	
		ELSEIF 	sale_shopmargin = 5 THEN	
			strSql = strSql&" 	,db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)) - convert(int, (i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100))*convert(float,"&sale_shopmarginvalue&")/100) else i.shopbuyprice end) )" + vbcrlf
		
		'업체특정반반부담/나머지매장부담
		ELSEIF sale_shopmargin = 6 THEN
			'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
		
		'업체특정,출고특정,텐바이텐특정반반부담/나머지텐바이텐부담
		ELSEIF sale_shopmargin = 7 THEN
			'strSql = strSql&" 	, db_shop.dbo.uf_GetItemPriceCutting(  (case when i.comm_cd = 'B012' or i.comm_cd = 'B013' or i.comm_cd = 'B011' then i.shopbuyprice - Convert(int, (i.orgsellprice-(i.shopitemprice-(i.shopitemprice*"&sale_rate&"/100)))/2) else i.shopbuyprice end) )" + vbcrlf
		END IF
			
		strSql = strSql & "		,'"&session("ssBctId")&"', "&point_rate&", i.comm_cd" + vbcrlf
		strSql = strSql & " 	from (" + vbcrlf
		strSql = strSql & " 		select" + vbcrlf
		strSql = strSql & " 		ii.shopitemprice , ii.makerid, ii.shopitemname , ii.shopitemid ,ii.itemgubun ,ii.itemoption,sdd.comm_cd" + vbcrlf
		strSql = strSql & " 		,(CASE" + vbcrlf
		strSql = strSql & " 			when sdd.comm_cd in ('B012','B013','B011') and ii.shopsuplycash=0" + vbcrlf
		strSql = strSql & " 				THEN convert(int,ii.shopitemprice*(100-IsNULL(sdd.defaultmargin,100))/100)" + vbcrlf
		strSql = strSql & " 			ELSE ii.shopsuplycash END) as 'shopsuplycash'" + vbcrlf
		strSql = strSql & " 		,(CASE" + vbcrlf
		strSql = strSql & " 			when sdd.comm_cd in ('B012','B013','B011') and ii.shopbuyprice=0" + vbcrlf
		strSql = strSql & " 				THEN convert(int,ii.shopitemprice*(100-IsNULL(sdd.defaultsuplymargin,100))/100)" + vbcrlf
		strSql = strSql & " 			ELSE ii.shopbuyprice END) as 'shopbuyprice'" + vbcrlf
		strSql = strSql & " 		,ii.orgsellprice ,sdd.shopid" + vbcrlf
		strSql = strSql & " 		from [db_shop].dbo.tbl_shop_item ii" + vbcrlf
		strSql = strSql & " 		join db_shop.dbo.tbl_shop_designer sdd" + vbcrlf
		strSql = strSql & " 			on sdd.shopid = '"&shopid&"'" + vbcrlf
		strSql = strSql & " 			and ii.makerid=sdd.makerid" + vbcrlf
		strSql = strSql & " 			and ii.isusing='Y'" + vbcrlf
		
		if contractoverlapyn="N" then
			'/정산구분에 안맞는 상품 제낌
			'동일마진 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
			IF sale_margin = 1 THEN
				strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
				
			'업체부담 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
			ELSEIF sale_margin = 2 THEN				
				strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
				
			'반반부담 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
			ELSEIF 	sale_margin = 3 THEN
				strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
				
			'10x10부담 (따로 처리않해도 문제없을듯)
			ELSEIF 	sale_margin = 4 THEN
								
			'직접설정 (업체특정, 출고특정, 텐바이텐특정만 업체부담이 가능)
			ELSEIF sale_margin = 5 THEN	
				strSql = strSql & " 			and sdd.comm_cd in ('B012','B013','B011')" + vbcrlf
				
			END IF
		end if

		strSql = strSql & " 		where ii.orgsellprice = ii.shopitemprice" + vbcrlf		'/상품할인은 제낌(원칙상 이중할인은 안됨)
		strSql = strSql & "		) as i" + vbcrlf
		strSql = strSql & " 	left join (" + vbcrlf
		strSql = strSql & " 		select b.itemid ,b.itemgubun , b.itemoption ,a.shopid" + vbcrlf
		strSql = strSql & " 		from [db_shop].[dbo].tbl_sale_off a" + vbcrlf
		strSql = strSql & " 		join [db_shop].[dbo].[tbl_saleitem_off] b" + vbcrlf
		strSql = strSql & " 			on a.sale_code = b.sale_code" + vbcrlf
		strSql = strSql & " 		where a.sale_startdate <= '"&sale_enddate&"'"
		strSql = strSql & " 		and a.sale_enddate >= '"&sale_startdate&"'" + vbcrlf
		strSql = strSql & " 		and a.sale_using = 1"
		strSql = strSql & " 		and a.sale_status <> 8"
		strSql = strSql & " 		and b.saleitem_status not in (8,9)"
		strSql = strSql & " 		and a.shopid = '"&shopid&"'" + vbcrlf
		strSql = strSql & " 	) as t" + vbcrlf
		strSql = strSql & " 		on i.shopitemid = t.itemid" + vbcrlf
		strSql = strSql & "			and i.itemgubun = t.itemgubun" + vbcrlf
		strSql = strSql & "			and i.itemoption = t.itemoption"
		strSql = strSql & "			and i.shopid = t.shopid"
		strSql = strSql & " 	WHERE" + vbcrlf
		strSql = strSql & "		i.shopitemprice > 0" + vbcrlf
		strSql = strSql & " 	and t.itemid is null" + vbcrlf		'/현재 할인테이블에 등록 되어 있는 상품 제낌			
		strSql = strSql & " 	and i.makerid = '"& makerid &"'"
		
		'response.write strSql &"<Br>"
		dbget.execute strSql
			
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans			
			Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank" 				
			dbget.close()	:	response.End
		END IF
		
		dbget.CommitTrans
		
		if Err_saleitemexists<>"" then
%>
			<script langauge="javascript">
				alert('할인기간이 중복설정된 상품은 제외되고 등록됩니다.\n<%= Err_saleitemexists %>');
			</script>
<%
		end if
		
		if contractoverlapyn="N" then
			if Err_contractnotexists<>"" then
%>
				<script langauge="javascript">
					alert('매입 상품의 경우 텐바이텐부담 설정시에만 등록 가능합니다.\n정산구분과 할인매입마진이 일치 하지 않는 상품은 제외되고 등록됩니다.\n\n<%= Err_contractnotexists %>');
				</script>
<%
			end if
		end if
%>
		<script langauge="javascript">
			alert('OK\n\n상품이 사용중이 아니거나, 판매금액이 0이하인 상품은 제외되고 등록됩니다');
			location.href ="about:blank";
			//parent.close();
			//parent.history.go(0);
			//parent.location.reload();	
		</script>
<% 	   
		dbget.close()	:	response.End
		
	'할인 선택상품 수정
	Case "U"
		sCode = requestCheckVar(Request("sC"),10)
		iCurrpage 	= request("iC")
		itemid 		= split(request("itemid"),",")		
		itemgubun = split(request("itemgubun"),",")
		itemoption = split(request("itemoption"),",")
							
		dissellprice= split(request("iDSPrice"),",")
		disbuyprice = split(request("iDBPrice"),",")
		idsaleshopsupplycash = split(request("idsaleshopsupplycash"),",")
		point_ratearr = split(request("point_ratearr"),",")
		arrsaleItemStatus	=split(request("saleItemStatus"),",")
		sale_status	=requestCheckVar(Request("sale_status"),4)
		makerid = requestCheckVar(Request("designer"),32)
		saleitem_idxarr = split(request("saleitem_idxarr"),",")

		'/트랜젝션
		dbget.beginTrans

		set osale = new csale_list
		osale.frectsale_code = sCode
		
		if sCode <> "" then
			'/해당 할인 정보 가져옴
			osale.getsaledetail
			
			if osale.ftotalcount > 0 then
				sale_startdate = osale.foneitem.fsale_startdate
				sale_enddate = osale.foneitem.fsale_enddate
				sale_rate = osale.foneitem.fsale_rate
				sale_margin = osale.foneitem.fsale_margin
				sale_marginValue = osale.foneitem.fsale_marginValue	
				sale_status	= osale.foneitem.fsale_status
				shopid	= osale.foneitem.fshopid
				sale_shopmargin	= osale.foneitem.fsale_shopmargin
				sale_shopmarginvalue = osale.foneitem.fsale_shopmarginvalue				
			end if
		end if

		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then
				
				'오픈중 상태일때 값 변경시 상태값 오픈예정으로 변경처리
				if Cint(trim(arrsaleItemStatus(i))) = 6 then
					arrsaleItemStatus(i) = 7	
				end if

				'종료 예정 일경우 종료예정으로 플래그 변경 , 아닐경우 할인가격 변경	
				IF trim(arrsaleItemStatus(i)) = 9 THEN
					strSql ="UPDATE [db_shop].[dbo].[tbl_saleItem_off] SET"&_
					" saleitem_status =9"&_
					" ,lastupdate = getdate()"&_
					" ,lastadminid='"&session("ssBctId")&"'"&_
					" WHERE itemid = "&trim(itemid(i))&""&_
					" and itemgubun = '"&trim(itemgubun(i))&"'"&_
					" and itemoption = '"&trim(itemoption(i))&"'"&_
					" and saleitem_idx = '"&trim(saleitem_idxarr(i))&"'"&_
					" and sale_code = '"&sCode&"'"
				ELSE	
					strSql ="UPDATE [db_shop].[dbo].[tbl_saleItem_off] SET"&_
					" saleprice = "&trim(dissellprice(i))&""&_
					" ,salesupplycash="&trim(disbuyprice(i))&""&_
					" ,saleitem_status ="&arrsaleItemStatus(i)&""&_
					" ,saleshopsupplycash="&trim(idsaleshopsupplycash(i))&" "&_
					" ,point_rate="&trim(point_ratearr(i))&" "&_
					" ,lastupdate = getdate()"&_					
					" ,lastadminid='"&session("ssBctId")&"'"&_
					" WHERE itemid = "&trim(itemid(i))&""&_
					" and itemgubun = '"&trim(itemgubun(i))&"'"&_
					" and itemoption = '"&trim(itemoption(i))&"'"&_
					" and saleitem_idx = '"&trim(saleitem_idxarr(i))&"'"&_
					" and sale_code = '"&sCode&"'"
				END IF
				
				'response.write strSql &"<br>"
				dbget.execute strSql
					
				IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		       		dbget.close()	:	response.End	  
				End IF	
				
			end if
		next
		
		dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('해당 상품이 수정 되었습니다.\n이미 할인중이면 리스트에서 실시간적용 버튼을 누르셔야 포스에 적용 됩니다.');"
		response.write "	location.replace('/admin/offshop/sale/saleItemReg.asp?sC="& sCode &"&designer="& makerid &"&iC="& iCurrpage &"&menupos="& menupos &"');"
		response.write "</script>"
		'response.redirect("saleItemReg.asp?menupos="&menupos&"&sC="&sCode&"&iC="&iCurrpage&"&designer="&makerid)
		dbget.close()	:	response.End

	'할인상품 삭제
	Case "D"
		sCode = requestCheckVar(Request("sC"),10)		
		itemid 		= split(request("itemid"),",")
		itemgubun = split(request("itemgubun"),",")
		itemoption = split(request("itemoption"),",")
		saleitem_idxarr = split(request("saleitem_idxarr"),",")

		'/트랜젝션
		dbget.beginTrans

		set osale = new csale_list
		osale.frectsale_code = sCode
		
		if sCode <> "" then
			'/해당 할인 정보 가져옴
			osale.getsaledetail
			
			if osale.ftotalcount > 0 then
				sale_startdate = osale.foneitem.fsale_startdate
				sale_enddate = osale.foneitem.fsale_enddate
				sale_rate = osale.foneitem.fsale_rate
				sale_margin = osale.foneitem.fsale_margin
				sale_marginValue = osale.foneitem.fsale_marginValue	
				sale_status	= osale.foneitem.fsale_status
				shopid	= osale.foneitem.fshopid
			end if
		end if

		for i=0 to UBound(itemid)
			if trim(itemid(i))<>"" then
				strSql ="UPDATE [db_shop].[dbo].[tbl_saleItem_off] "&_
						" SET saleitem_status =9"&_
						" ,lastupdate = getdate()"&_
						" ,lastadminid='"&session("ssBctId")&"'"&_
						" WHERE itemid = "&trim(itemid(i))&""&_
						" and itemgubun = '"&trim(itemgubun(i))&"'"&_
						" and itemoption = '"&trim(itemoption(i))&"'"&_
						" and saleitem_idx = '"&trim(saleitem_idxarr(i))&"'"&_
						" and sale_code = '"&sCode&"'"

				'response.write strSql &"<br>"
				'response.end
				dbget.execute strSql
			
				IF Err.Number <> 0 THEN
		   			dbget.RollBackTrans	
		   			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		       		dbget.close()	:	response.End	  
				End IF
			end if
		next
				
		dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('해당 상품이 할인 종료 예정으로 변경 되었습니다.\n이미 할인중이면 리스트에서 실시간적용 버튼을 누르셔야 포스에 적용 됩니다.');"
		response.write "	location.replace('/admin/offshop/sale/saleItemReg.asp?sC="& sCode &"&menupos="& menupos &"');"
		response.write "</script>"
		'response.redirect("saleItemReg.asp?menupos="&menupos&"&sC="&sCode)
		dbget.close()	:	response.End
	
	'할인 실시간 적용
	Case "P"
		sCode = requestCheckVar(Request("sC"),10)	
		iCurrpage 	= request("iC")
		IF sCode = "" THEN 
			Alert_return("파라미터값에 문제가 있습니다.")    	
		     dbget.close()	:	response.End	
		END IF       		
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_shop].[dbo].[sp_Ten_item_SetPrice_RealTime_off] ("&sCode&",'"&session("ssBctId")&"')}"			
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
	 	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sale_status="&ssStatus
	 	'--------------------------------------------------------------
 
		IF iResult <> 1 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
		    dbget.close()	:	response.End	
		END IF
	
		'response.redirect("saleList.asp?menupos="&menupos&"&"&strParm)
		response.write "<script>"
		response.write "	location.href='"&refer&"'"
		response.write "</script>"
	
		dbget.close()	:	response.End
		
	CASE Else
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")    	
       dbget.close()	:	response.End
END SELECT

set osale = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

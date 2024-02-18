<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 상품등록
' History : 2008.04.04 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
 Dim gCode, itemidarr, mode, sGroup,sortarr, sType, sgDelivery
 Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,deliverytype
 dim tempidarr,cnt,i,sqlStr,addSql,strSqlAdd
  dim iCurrpage
 
mode = Request("mode")

itemidarr = Request("itemidarr")
 

sType =  Request("sType")
 
gCode =request("gC")
itemid      = request("itemid")
itemname    = request("itemname")
makerid     = request("makerid")
sellyn      = request("sellyn")
usingyn     = request("usingyn")
danjongyn   = request("danjongyn")
limityn     = request("limityn")
deliverytype= request("deliverytype")
sailyn      = request("sailyn")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

iCurrpage = request("iC") 

Select Case mode
	Case "I" '// 상품추가 
	 addSql = ""
	 
	 '-- 사은품종류와 같은 배송 타입인지 체크-------------------
	sqlStr = "SELECT gift_delivery FROM [db_event].[dbo].tbl_gift  WHERE gift_code = "&gCode
	rsget.Open sqlStr, dbget
	IF not rsget.EOF THEN
		sgDelivery = rsget("gift_delivery")
	END IF	
	rsget.close	
	
	IF sgDelivery = "Y" THEN '업체배송일 겨우
		strSqlAdd = " and deliverytype not in (2,5,7,9)"
	ELSE
		strSqlAdd = " and deliverytype not in (1,4)"
	END IF	
	'------------------------------------------------------------
	 		
	  IF sType = "all" THEN '검색된 모든 내용 insert  처리
	  	 '// 추가 쿼리
        if (makerid <> "") then
            addSql = addSql & " and i.makerid='" + makerid + "'"
        end if

        if (itemid <> "") then
            addSql = addSql & " and i.itemid in (" + itemid + ")"
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
            addSql = addSql + " and i.sailyn='" + sailyn + "'"
        end if  
        
        if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if
    ELSE
    	addSql = addSql & " and itemid in ("&trim(itemidarr)&")"	    
	END IF	
	
		Dim itemCnt : itemCnt = 0
		sqlStr = " select count(i.itemid) from  [db_item].[dbo].tbl_item i where  1=1 "&addSql&strSqlAdd		
		rsget.Open sqlStr, dbget
		IF not rsget.EOF THEN
			itemCnt = rsget(0)
		END IF	
		rsget.close	
		
		IF itemCnt > 0 THEN
			Call Alert_move("사은품의 배송조건과 동일하지 않은 상품은 추가 불가능합니다. 조건 확인 후 다시 등록해주세요","about:blank")				
			dbget.close()	:	response.End	
		END IF	
			
			sqlStr = " insert into [db_event].[dbo].tbl_giftitem" + VbCrlf
			sqlStr = sqlStr + " (gift_code,itemid)" + VbCrlf
			sqlStr = sqlStr + " select " + CStr(gCode) + ", i.itemid"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " itemid not in "
			sqlStr = sqlStr + " (select itemid from [db_event].[dbo].tbl_giftitem"
			sqlStr = sqlStr + " where gift_code=" + gCode
			sqlStr = sqlStr + " )"	 + addSql							
			dbget.execute sqlStr
			 
	Case "D" '// 선택상품 삭제
			sqlStr = "Delete From  [db_event].[dbo].tbl_giftitem "&_
					"	WHERE gift_code = "&gCode&" and itemid in ("&itemidarr&") "				
			dbget.execute sqlStr 		
End Select
	

	IF Err.Number = 0 THEN
		if mode= "I" then
	%>
	<script langauge="javascript">
	<!--	
		location.href ="about:blank";
		parent.history.go(0);	
	//-->
	</script>
<% 	   
		dbget.close()	:	response.End		
		response.redirect("giftItemReg.asp?gC="&gCode&"&menupos="&menupos&"&iC="&iCurrpage)
		end if
		dbget.close()	:	response.End
	Else
   
%>
	<script language="javascript">
	<!--
	alert("데이터 처리에 문제가 발생하였습니다.");
	history.back(-1);
	//-->
	</script>
<%                
	dbget.close()	:	response.End	
End IF	
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
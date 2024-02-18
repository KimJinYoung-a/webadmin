<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/eventitem_regist.asp
' Description :  이벤트 등록 - 상품등록
' History : 2007.02.21 정윤정 생성
'           2008.10.20 상품이미지 크기 추가(허진원)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
 Dim eCode, itemidarr, mode, sGroup,sortarr, sizearr, sType
 Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype
 dim tempidarr,cnt,i,sqlStr,strSqlAdd,addSql
 dim eSort,strG
 dim iCurrpage
 dim sgDelivery : sgDelivery = ""
 Dim itemCnt 
 Dim dispCate
 
mode = Request("mode")

itemidarr = Request("itemidarr")
 
sGroup = Request("selGroup")
sType =  Request("sType")
 
eCode =request("eC")
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

dispCate = requestCheckvar(request("disp"),16)

iCurrpage = request("iC") 
strG =	 Request("selG")

		 
	dbget.beginTrans
 
Select Case mode
	Case "I" '// 상품추가 
	 addSql = ""
	
	'-- 사은품종류와 같은 배송 타입인지 체크-------------------		
	sqlStr = "SELECT gift_delivery FROM [db_event].[dbo].tbl_gift  WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup	
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

        if (itemidarr <> "") then
            addSql = addSql & " and i.itemid in (" + itemidarr + ")"
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
        end If
        
		if dispCate<>"" then
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + dispCate + "%' and isDefault='y') "
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
		'사은품이 있을 경우 이벤트등록상품  배송 확인
		IF sgDelivery <> "" THEN
				itemCnt = 0
				sqlStr = " select count(i.itemid) from  [db_item].[dbo].tbl_item i where  1=1 "&addSql	&strSqlAdd					
				rsget.Open sqlStr, dbget
				IF not rsget.EOF THEN
					itemCnt = rsget(0)
				END IF	
				rsget.close	
				
				IF itemCnt > 0 THEN
			%>
				<script language="javascript">
				<!--
				alert("사은품의 배송조건과 동일하지 않은 상품은 추가 불가능합니다. 조건 확인 후 다시 등록해주세요");
				self.location.href ="about:blank";
				//-->
				</script>
			<%               
					response.End	
				END IF	
		END IF		 
		
			Dim iChkCount
	 		sqlStr = "SELECT  count(itemid) FROM  [db_item].[dbo].tbl_item as i WHERE itemid not in (select itemid from [db_event].[dbo].tbl_eventitem where evt_code="+eCode+") "+addSql 
			rsget.Open sqlStr, dbget
			IF not rsget.EOF THEN
				iChkCount = rsget(0)
			END IF	
			rsget.close	 
			IF iChkCount>1000 THEN
					%>
				<script language="javascript">
				<!--
				alert("상품은 최대 1000건까지 가능합니다. 조건을 다시 설정해주세요");
				self.location.href ="about:blank";
				//-->
				</script>
			<%               
			response.end
			END IF
			sqlStr = " insert into [db_event].[dbo].tbl_eventitem" + VbCrlf
			sqlStr = sqlStr + " (evt_code,itemid,evtgroup_code,evtitem_sort)" + VbCrlf
			sqlStr = sqlStr + " select " + CStr(eCode) + ", i.itemid, '"&sGroup&"',50"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " itemid not in "
			sqlStr = sqlStr + " (select itemid from [db_event].[dbo].tbl_eventitem"
			sqlStr = sqlStr + " where evt_code=" + eCode
			sqlStr = sqlStr + " )"	 + addSql				
			dbget.execute sqlStr
	    
		    
		    ''다이어리 사은품 임시..
		    if (CStr(eCode)="8361" or CStr(eCode)="8362" or CStr(eCode)="8363") then
		        sqlStr = "exec db_diary_collection.dbo.ten_IMSI_diary_eventPrize"
		        dbget.execute sqlStr
		    end if
		    
		 
	Case "D" '// 선택상품 삭제
			sqlStr = "Delete From  [db_event].[dbo].tbl_eventitem "&_
					"	WHERE evt_code = "&eCode&" and itemid in ("&itemidarr&") "				
			dbget.execute sqlStr
			
		 	
		 	''다이어리 사은품 임시..
		    if (CStr(eCode)="8361" or CStr(eCode)="8362" or CStr(eCode)="8363") then
		        sqlStr = "exec db_diary_collection.dbo.ten_IMSI_diary_eventPrize"
		        dbget.execute sqlStr
		    end if
	Case "G" '//그룹이동
		
		'-- 사은품종류와 같은 배송 타입인지 체크-------------------		
			sqlStr = "SELECT gift_delivery FROM [db_event].[dbo].tbl_gift  WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup			
			rsget.Open sqlStr, dbget
			IF not rsget.EOF THEN
				sgDelivery = rsget("gift_delivery")
			END IF	
			rsget.close	
			
			IF sgDelivery <> "" THEN
				itemCnt = 0
				IF sgDelivery = "Y" THEN '업체배송일 겨우
					strSqlAdd = " and deliverytype not in (2,5,7,9)"
				ELSE
					strSqlAdd = " and deliverytype not in (1,4)"
				END IF					
				
				sqlStr = "SELECT count(itemid) FROM [db_item].[dbo].tbl_item WHERE itemid in  ( "&itemidarr&") " & strSqlAdd			
				rsget.Open sqlStr, dbget
				IF not rsget.EOF THEN
					itemCnt = rsget(0)
				END IF	
				rsget.close	
				
				IF itemCnt > 0 THEN
		%>
			<script language="javascript">
			<!--
			alert("사은품 종류의 배송타입과 다른 상품이 존재합니다. 이동 불가능합니다.");
			history.back(-1);
			//-->
			</script>
		<% 		dbget.close()	:	response.End
				END IF								
			END IF	
		
		'------------------------------------------------------------			
							
			sqlStr = "UPDATE [db_event].[dbo].tbl_eventitem SET "&_
					" evtgroup_code = "&sGroup& _
					"	WHERE evt_code = "&eCode&" and itemid in ( "&itemidarr&") "
			dbget.execute sqlStr
					
	Case "S" '//상품순서/이미지크기 저장
		Dim tmpSort, tmpSize
		sortarr = Request("sortarr")
		sizearr = Request("sizearr")

		If sortarr="" and sizearr="" THEN
			dbget.RollBackTrans
			Response.Write "<script language='javascript'>history.back(-1);</script>"
			dbget.close()	:	response.End
		end if

		'선택상품 파악
		itemidarr = split(itemidarr,",")
		cnt = ubound(itemidarr)

		'// 정렬순서 저장
		If sortarr<>"" THEN
			sortarr =  split(sortarr,",")
			
			for i=0 to cnt	
				IF sortarr(i) = "" THEN
					 tmpSort = "NULL"				
				ELSE	
					 tmpSort = sortarr(i)	
				END IF	 
				sqlStr = "UPDATE [db_event].[dbo].tbl_eventitem SET "&_
						" evtitem_sort = "&tmpSort& _
						"	WHERE evt_code = "&eCode&" and itemid =" + itemidarr(i)
				dbget.execute sqlStr
			next
		END IF

		'// 이미지 크기 저장
		If sizearr<>"" THEN			
			sizearr =  split(sizearr,",")

			for i=0 to cnt	
				IF sizearr(i) = "" THEN
					 tmpSize = "NULL"				
				ELSE	
					 tmpSize = sizearr(i)	
				END IF	 
				sqlStr = "UPDATE [db_event].[dbo].tbl_eventitem SET "&_
						" evtitem_imgsize = "&tmpSize& _
						"	WHERE evt_code = "&eCode&" and itemid =" + itemidarr(i)
				dbget.execute sqlStr
			next
		End If

End Select
	

	IF Err.Number = 0 THEN
	dbget.CommitTrans

	if mode= "I" then
%>
	<script langauge="javascript">
	<!--	
		location.href ="about:blank";
		parent.history.go(0);	
	//-->
	</script>
<%
	else		
		response.redirect("eventitem_regist.asp?eC="&eCode&"&menupos="&menupos&"&selG="&strG&"&iC="&iCurrpage)
	end if
	dbget.close()	:	response.End
	Else
   	dbget.RollBackTrans	  
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
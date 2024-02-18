<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 등록 - 상품등록
' History : 2010.09.29 한용민 생성
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
Dim eCode, itemidarr, mode, sGroup,sortarr, sizearr, sType
Dim itemid,itemname, makerid, cdl, cdm, cds, sellyn,usingyn,danjongyn,limityn,sailyn,mwdiv,deliverytype
dim tempidarr,cnt,i,sqlStr,strSqlAdd,addSql ,eSort,strG ,iCurrpage ,itemCnt
dim sgDelivery : sgDelivery = ""
	mode = RequestCheckvar(Request("mode"),2)
	itemidarr = Request("itemidarr")
	sGroup = RequestCheckvar(Request("selGroup"),10)
	sType =  RequestCheckvar(Request("sType"),10)
	eCode = RequestCheckvar(request("eC"),10)
	itemid      = RequestCheckvar(request("itemid"),10)
	itemname    = RequestCheckvar(request("itemname"),64)
	makerid     = RequestCheckvar(request("makerid"),32)
	sellyn      = RequestCheckvar(request("sellyn"),1)
	usingyn     = RequestCheckvar(request("usingyn"),1)
	danjongyn   = RequestCheckvar(request("danjongyn"),1)
	limityn     = RequestCheckvar(request("limityn"),1)
	mwdiv       = RequestCheckvar(request("mwdiv"),1)
	sailyn      = RequestCheckvar(request("sailyn"),1)
	deliverytype= RequestCheckvar(request("deliverytype"),2)
	cdl = RequestCheckvar(request("cdl"),10)
	cdm = RequestCheckvar(request("cdm"),10)
	cds = RequestCheckvar(request("cds"),10)
	iCurrpage = RequestCheckvar(request("iC"),10)
	strG =	 RequestCheckvar(Request("selG"),10)
  	if itemidarr <> "" then
		if checkNotValidHTML(itemidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
dbacademyget.beginTrans
 
Select Case mode
	
	Case "I" '// 상품추가 
	 addSql = ""
	
	'-- 사은품종류와 같은 배송 타입인지 체크-------------------		
	sqlStr = "SELECT gift_delivery FROM [db_academy].[dbo].tbl_gift WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup	
	
	'response.write sqlStr &"<br>"
	rsacademyget.Open sqlStr, dbacademyget
	
	IF not rsacademyget.EOF THEN
		sgDelivery = rsacademyget("gift_delivery")
	END IF	
	rsacademyget.close	
	
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
				sqlStr = " select count(i.itemid) from [db_academy].dbo.tbl_diy_item i where  1=1 "&addSql	&strSqlAdd					
				
				'response.write sqlStr &"<br>"
				rsacademyget.Open sqlStr, dbacademyget
				
				IF not rsacademyget.EOF THEN
					itemCnt = rsacademyget(0)
				END IF	
				rsacademyget.close	
				
				IF itemCnt > 0 THEN
			%>
				<script language="javascript">

				alert("사은품의 배송조건과 동일하지 않은 상품은 추가 불가능합니다. 조건 확인 후 다시 등록해주세요");
				self.location.href ="about:blank";

				</script>
			<%               
					response.End	
				END IF	
		END IF		 
		
			sqlStr = " insert into [db_academy].[dbo].tbl_eventitem" + VbCrlf
			sqlStr = sqlStr + " (evt_code,itemid,evtgroup_code,evtitem_sort)" + VbCrlf
			sqlStr = sqlStr + " select " + CStr(eCode) + ", i.itemid, '"&sGroup&"',50"
			sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i"
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " itemid not in ("
			sqlStr = sqlStr + " 	select itemid from [db_academy].[dbo].tbl_eventitem"
			sqlStr = sqlStr + " 	where evt_code=" + eCode
			sqlStr = sqlStr + " )"	 + addSql				
			
			'response.write sqlStr &"<br>"
			dbacademyget.execute sqlStr
	    		    
	Case "D" '// 선택상품 삭제
			sqlStr = "Delete From [db_academy].[dbo].tbl_eventitem "&_
					" WHERE evt_code = "&eCode&" and itemid in ("&itemidarr&") "				
			
			'response.write sqlStr &"<br>"
			dbacademyget.execute sqlStr
	
	Case "G" '//그룹이동		
		'-- 사은품종류와 같은 배송 타입인지 체크-------------------		
			sqlStr = "SELECT gift_delivery FROM [db_academy].[dbo].tbl_gift  WHERE gift_status < 9 and gift_using='Y' and evt_code = "&eCode&" and evtgroup_code ="&sGroup			
			
			'response.write sqlStr &"<br>"
			rsacademyget.Open sqlStr, dbacademyget
			
			IF not rsacademyget.EOF THEN
				sgDelivery = rsacademyget("gift_delivery")
			END IF	
			
			rsacademyget.close	
			
			IF sgDelivery <> "" THEN
				itemCnt = 0
				IF sgDelivery = "Y" THEN '업체배송일 겨우
					strSqlAdd = " and deliverytype not in (2,5,7,9)"
				ELSE
					strSqlAdd = " and deliverytype not in (1,4)"
				END IF					
				
				sqlStr = "SELECT count(itemid) FROM [db_academy].dbo.tbl_diy_item WHERE itemid in  ( "&itemidarr&") " & strSqlAdd			
				
				'response.write sqlStr &"<br>"
				rsacademyget.Open sqlStr, dbacademyget
				
				IF not rsacademyget.EOF THEN
					itemCnt = rsacademyget(0)
				END IF	
				
				rsacademyget.close	
				
				IF itemCnt > 0 THEN
		%>
			<script language="javascript">

			alert("사은품 종류의 배송타입과 다른 상품이 존재합니다. 이동 불가능합니다.");
			history.back(-1);

			</script>
		<% 		dbacademyget.close()	:	response.End
				END IF								
			END IF	
		
		'------------------------------------------------------------			
							
			sqlStr = "UPDATE [db_academy].[dbo].tbl_eventitem SET"&_
					" evtgroup_code = "&sGroup& _
					" WHERE evt_code = "&eCode&" and itemid in ( "&itemidarr&") "
			
			'response.write sqlStr &"<br>"
			dbacademyget.execute sqlStr
					
	Case "S" '//상품순서/이미지크기 저장
		Dim tmpSort, tmpSize
		sortarr = Request("sortarr")
		sizearr = Request("sizearr")

		If sortarr="" and sizearr="" THEN
			dbacademyget.RollBackTrans
			Response.Write "<script language='javascript'>history.back(-1);</script>"
			dbacademyget.close()	:	response.End
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
				sqlStr = "UPDATE [db_academy].[dbo].tbl_eventitem SET "&_
						" evtitem_sort = "&tmpSort& _
						" WHERE evt_code = "&eCode&" and itemid =" + itemidarr(i)
				
				'response.write sqlStr &"<br>"
				dbacademyget.execute sqlStr
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
				sqlStr = "UPDATE [db_academy].[dbo].tbl_eventitem SET "&_
						" evtitem_imgsize = "&tmpSize& _
						" WHERE evt_code = "&eCode&" and itemid =" + itemidarr(i)
				
				'response.write sqlStr &"<br>"
				dbacademyget.execute sqlStr
			next
		End If

End Select
	
IF Err.Number = 0 THEN
	dbacademyget.CommitTrans

	if mode= "I" then
%>
		<script langauge="javascript">
	
			location.href ="about:blank";
			parent.history.go(0);	

		</script>
<%
		else		
			response.redirect("eventitem_regist.asp?eC="&eCode&"&menupos="&menupos&"&selG="&strG&"&iC="&iCurrpage)
		end if
	dbacademyget.close()	:	response.End
Else
   	dbacademyget.RollBackTrans	  
%>
	<script language="javascript">
	
	alert("데이터 처리에 문제가 발생하였습니다.");
	history.back(-1);
	
	</script>
<%                
	dbacademyget.close()	:	response.End	
End IF	
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
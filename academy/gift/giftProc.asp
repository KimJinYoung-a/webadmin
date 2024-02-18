<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사은품 db 처리
' History : 2010.09.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i ,s120Img, s401Img, s402Img, s403Img, s404Img, s405Img ,strParm ,iSiteScope, sPartnerID ,eCode,gCode
Dim sMode, strSql,strSqlAdd ,iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sgStatus
Dim sTitle, dSDay, dEDay, iGiftScope, sBrand, iGroupCode, iGiftType, iGiftRange1, iGiftRange2, iGiftKindCnt, iGiftKindType, iGiftLimit
Dim sGiftKindName, itemid, sGiftKindImg, iGiftKindCode, sGiftDelivery, iGiftStatus, sGiftUsing,igStatus,sOpenDate,sCloseDate
	sMode = requestCheckVar(Request.Form("sM"),2) 
	
'===========================================================================	
'증정범위에 따른 배송 확인
Function fnChkDelivery(ByVal iGiftScope, ByVal sGiftDelivery, ByVal eCode, ByVal Brand, ByVal egCode, ByVal gCode)
	IF sGiftDelivery ="Y" THEN '업체배송인 경우
		strSqlAdd = " and deliverytype not in (2,5,7,9)"
	ELSE
		strSqlAdd = " and deliverytype not in (1,4)"
	END IF			
	
	IF 	iGiftScope = 1 THEN '모든 구매자 증정일 경우
		IF sGiftDelivery ="Y" THEN
			Alert_return("증정대상이 모든상품일 경우에는 텐바이텐배송만 가능 합니다. 조건을 다시 설정해주세요 ")    	
      	 dbget.close()	:	response.End	
		END IF	
	ELSEIF 	iGiftScope = 2 THEN '이벤트등록상품 증정일 경우
		IF eCode ="" OR eCode = "0" THEN 
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbacademyget.close()	:	response.End	
    	END IF    	
		
		strSql = " SELECT deliverytype FROM [db_academy].[dbo].[tbl_eventitem] AS A INNER JOIN [db_academy].dbo.tbl_diy_item AS B ON A.itemid = B.itemid  "&_
				"  WHERE  evt_code = "&eCode& strSqlAdd				
		rsacademyget.Open strSql, dbacademyget
		IF not (rsacademyget.EOF OR rsacademyget.BOF) THEN
			Alert_return("선택한 이벤트등록상품에 사은품배송타입과 다른 상품이 존재합니다. 조건을 다시 설정해주세요 ")    	
      	 dbacademyget.close()	:	response.End	
		END IF	
		rsacademyget.close	
	ELSEIF 	iGiftScope = 3 THEN '선택브랜드 증정일 경우					
		strSql = " SELECT deliverytype FROM  [db_item].[dbo].[tbl_Item] where makerid = '"&sBrand&"' "& strSqlAdd
		rsget.Open strSql, dbget
		IF not (rsget.EOF OR rsget.BOF) THEN
		%>
		<script language="javascript">
		<!--
		if(confirm("선택한 브랜드상품에 사은품배송타입과 다른 상품이 존재합니다. \n 그 상품에 대해서는 사은품이 발송되지 않습니다. 계속하시겠습니까?")){
			return;
		}else{
			history.back();
		}	
		//-->
		</script>				
		<%
		END IF	
		rsget.close	
	ELSEIF 	iGiftScope = 4 THEN '선택그룹상품  증정일 경우					
		strSql = " SELECT deliverytype FROM [db_academy].[dbo].[tbl_eventitem] AS A INNER JOIN [db_item].[dbo].[tbl_Item] AS B ON A.itemid = B.itemid  "&_
				"  WHERE  evt_code = "&eCode& " and evtgroup_code ="&egCode&strSqlAdd
		rsget.Open strSql, dbget
		IF not (rsget.EOF OR rsget.BOF) THEN
			Alert_return("선택한 그룹상품에  사은품배송타입과 다른 상품이 존재합니다. 조건을 다시 설정해주세요 ")    	
      	 dbget.close()	:	response.End	
		END IF	
		rsget.close		
	ELSEIF 	iGiftScope = 5 THEN '선택상품  증정일 경우					
		strSql = " SELECT deliverytype FROM [db_academy].[dbo].[tbl_giftitem] AS A INNER JOIN [db_item].[dbo].[tbl_Item] AS B ON A.itemid = B.itemid  "&_
				"  WHERE  gift_code = "&gCode&strSqlAdd
		rsget.Open strSql, dbget
		IF not (rsget.EOF OR rsget.BOF) THEN
			Alert_return("선택상품에  사은품타입과 다른 상품이 존재합니다. 조건을 다시 설정해주세요 ")    	
      	 dbget.close()	:	response.End	
		END IF	
		rsget.close		
	END IF		
End Function
'===========================================================================	

SELECT CASE sMode
	
Case "I"	'//사은품 등록
	eCode			= requestCheckVar(Request.Form("eC"),10) 
	IF eCode ="" THEN eCode = 0 
	sTitle			= html2db(requestCheckVar(Request.Form("sGN"),64))
	dSDay 			= requestCheckVar(Request.Form("sSD"),10)  
	dEDay			= requestCheckVar(Request.Form("sED"),10)  
	iGiftScope		= requestCheckVar(Request.Form("giftscope"),4)  
	sBrand			= requestCheckVar(Request.Form("ebrand"),32)  
	iGroupCode		= requestCheckVar(Request.Form("selG"),10)  	
	iGiftType		= requestCheckVar(Request.Form("gifttype"),10)  
	iGiftRange1		= requestCheckVar(Request.Form("sGR1"),10)  
	iGiftRange2		= requestCheckVar(Request.Form("sGR2"),10)  
	iGiftKindCnt	= requestCheckVar(Request.Form("iGKC"),10)  
	iGiftKindType	= requestCheckVar(Request.Form("chkKT"),10)  
	iGiftLimit		= requestCheckVar(Request.Form("iL"),10)  
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10)  	
	sGiftDelivery	= requestCheckVar(Request.Form("selD"),1)  	
	iGiftStatus		= requestCheckVar(Request.Form("giftstatus"),10)  
	sOpenDate		= requestCheckVar(Request.Form("sOD"),30)  	
	sCloseDate		= requestCheckVar(Request.Form("sCD"),30)  		
	iSiteScope		= requestCheckVar(Request.Form("eventscope"),4)	
	IF CStr(iSiteScope) = "3" THEN sPartnerID 		= requestCheckVar(Request.Form("selP"),32)	
	
	IF iGiftStatus = "7" THEN
		if sOpenDate = "" then
			 sOpenDate = "getdate()"
		else
			sOpenDate = " convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
		end if	 
	ELSEIF 	iGiftStatus = "9" THEN
		if sCloseDate = "" then
			 sCloseDate = "getdate()"
		else
			sCloseDate = " convert(nvarchar(10),'"&sCloseDate&"',21)"&"+' "&formatdatetime(sCloseDate,4)&"'"
		end if
	ELSE
		IF sOpenDate = "" THEN 
			sOpenDate = "null"	
		ELSE
			sOpenDate = " convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
		END IF	
		
		IF sCloseDate = "" THEN
			sCloseDate = "null"	
		ELSE
			sCloseDate = " convert(nvarchar(10),'"&sCloseDate&"',21)"&"+' "&formatdatetime(sCloseDate,4)&"'"
		END IF	
	END IF		
 
	IF iGiftKindType = "" THEN iGiftKindType = 1
	IF iGiftLimit ="" THEN iGiftLimit = 0
	IF iGiftType = "" THEN iGiftType =0
	IF iGiftRange1 = "" THEN iGiftRange1 = 0
	IF iGiftRange2 = "" THEN iGiftRange2 = 0
	IF iGroupCode = "" THEN iGroupCode = 0
			
	'//증정범위에 따른 배송 확인
	CALL fnChkDelivery(iGiftScope,sGiftDelivery,eCode, sBrand,iGroupCode, 0)
	
	On Error Resume Next

	'//데이터 등록
	strSql = "INSERT INTO [db_academy].[dbo].[tbl_gift] ( [gift_name], [gift_scope], [evt_code], [evtgroup_code], [makerid], [gift_type], [gift_range1], [gift_range2]"&_
			", [giftkind_code], [giftkind_type], [giftkind_cnt], [giftkind_limit], [gift_startdate], [gift_enddate],[gift_status],[gift_delivery],[adminid],opendate,lastupdate"&_
			", site_scope, partner_id)"&_
			" VALUES ('"&sTitle&"','"&iGiftScope&"','"&eCode&"','"&iGroupCode&"','"&sBrand&"','"&iGiftType&"','"&iGiftRange1&"','"&iGiftRange2&"' "&_
			",'"&iGiftKindCode&"','"&iGiftKindType&"','"&iGiftKindCnt&"','"&iGiftLimit&"','"&dSDay&"','"&dEDay&"','"&iGiftStatus&"','"&sGiftDelivery&"','"&session("ssBctId")&"',"&sOpenDate&",getdate()"&_
			", '"&iSiteScope&"','"&sPartnerID&"') "									
	dbacademyget.execute strSql
	
	IF Err.Number <> 0 THEN
		response.Write strSql
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
		dbget.close()	:	response.End	
	END IF	
	
	on error Goto 0

	IF eCode = 0 THEN eCode = ""
	response.redirect("giftList.asp?menupos="&menupos&"&eC="&eCode)
dbget.close()	:	response.End

Case "U"	'//사은품수정
	Dim strAdd : strAdd = ""
	
	'검색어 체크--------------------------------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'검색어
	 sBrand     	= requestCheckVar(Request("ebrand"),32)		'브랜드
	 sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	 sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	 sgStatus	    = requestCheckVar(Request("gstatus"),4)	'사은품 상태
	
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
 	strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&sgStatus
 	'--------------------------------------------------------------
 	
	gCode			= requestCheckVar(Request.Form("gC"),10) 
	eCode			= requestCheckVar(Request.Form("eC"),10) 
	IF eCode ="" THEN eCode = 0 
	sTitle			= html2db(requestCheckVar(Request.Form("sGN"),64))
	dSDay 			= requestCheckVar(Request.Form("sSD"),10)  
	dEDay			= requestCheckVar(Request.Form("sED"),10)  
	iGiftScope		= requestCheckVar(Request.Form("giftscope"),4)  
	sBrand			= requestCheckVar(Request.Form("ebrand"),32)  
	iGroupCode		= requestCheckVar(Request.Form("selG"),10)  	
	iGiftType		= requestCheckVar(Request.Form("gifttype"),10)  
	iGiftRange1		= requestCheckVar(Request.Form("sGR1"),10)  
	iGiftRange2		= requestCheckVar(Request.Form("sGR2"),10)  
	iGiftKindCnt	= requestCheckVar(Request.Form("iGKC"),10)  
	iGiftKindType	= requestCheckVar(Request.Form("chkKT"),10)  
	iGiftLimit		= requestCheckVar(Request.Form("iL"),10)  
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10)  	
	iGiftStatus		= requestCheckVar(Request.Form("giftstatus"),10) 
	sGiftUsing		= requestCheckVar(Request.Form("sGU"),1)  	
	sOpenDate		= requestCheckVar(Request.Form("sOD"),30)  	
	sCloseDate		= requestCheckVar(Request.Form("sCD"),30)  	
	sGiftDelivery	= requestCheckVar(Request.Form("selD"),1)  	
	iSiteScope		= requestCheckVar(Request.Form("eventscope"),4)		
	IF CStr(iSiteScope) = "3" THEN	sPartnerID 		= requestCheckVar(Request.Form("selP"),32)	
	
	IF iGiftStatus ="7" AND sOpenDate="" THEN
		strAdd = " , [opendate] = getdate()"	
	ELSEIF (iGiftStatus = "9" and sCloseDate ="" ) THEN 		
		strAdd = ", [closedate] = getdate() "	'종료처리일 설정				
	END IF	

	'종료일 이전에 종료시 종료일 현재 날짜로 변경
	IF iGiftStatus = 9 and  datediff("d",dEDay,date()) <0 THEN
			dEDay = date()
	END IF	
	
	IF iGiftKindType = "" THEN iGiftKindType = 1
	IF iGiftLimit ="" THEN iGiftLimit = 0
 	IF iGiftType = "" THEN iGiftType =0
 	IF iGiftRange1 = "" THEN iGiftRange1 = 0
	IF iGiftRange2 = "" THEN iGiftRange2 = 0
	IF iGroupCode = "" THEN iGroupCode = 0	
		
 	'//증정범위에 따른 배송 확인
 	CALL fnChkDelivery(iGiftScope,sGiftDelivery,eCode, sBrand,iGroupCode, gCode)
 	
 	'//데이터 수정
	strSql = " UPDATE [db_academy].[dbo].[tbl_gift] SET  [gift_name] = '"&sTitle&"', [gift_scope]="&iGiftScope&", [evtgroup_code] ="&iGroupCode&_
			" , [makerid]='"&sBrand&"', [gift_type]="&iGiftType&", [gift_range1]="&iGiftRange1&", [gift_range2]= "&iGiftRange2&_
			", [giftkind_code]= "&iGiftKindCode&", [giftkind_type] ="&iGiftKindType&" , [giftkind_cnt]= "&iGiftKindCnt&", [giftkind_limit]="&iGiftLimit&_
			", [gift_startdate]= '"&dSDay&"', [gift_enddate]='"&dEDay&"', [gift_status] = "&iGiftStatus&", [gift_using] = '"&sGiftUsing&"'"&_
			" , gift_delivery = '"&sGiftDelivery&"'"&_
			",[adminid]= '"&session("ssBctId")&"', [lastupdate] = getdate(), site_scope="&iSiteScope&", partner_id ='"&sPartnerID&"' "&strAdd&_
			" WHERE gift_code = "&gCode	
				
	dbacademyget.execute strSql
	
	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요1")    	
       dbget.close()	:	response.End	
	END IF	
	
	IF eCode = 0 THEN eCode = ""
	response.redirect("giftList.asp?menupos="&menupos&"&"&strParm)
dbacademyget.close()	:	response.End

Case "KI"  '//사은품 종류 등록	
	sGiftKindName 	= html2db(requestCheckVar(Request.Form("sGKN"),60))	
	itemid			= requestCheckVar(Request.Form("itemid"),10) 
	sGiftKindImg	= requestCheckVar(Request.Form("sGKImg"),100) 
	IF itemid = "" THEN itemid =0
		
	IF itemid > 0 THEN
	strSql = "SELECT itemid FROM [db_academy].dbo.tbl_diy_item where itemid = "&itemid
	rsacademyget.Open strSql, dbacademyget
	IF rsacademyget.EOF OR rsacademyget.BOF THEN
		rsacademyget.Close	
		Alert_return("존재하지 않는 상품번호입니다. 확인 후 다시 입력해주세요")    	
       dbacademyget.close()	:	response.End	
	End IF
	rsacademyget.Close	
	END IF
	strSql = "INSERT INTO [db_academy].[dbo].[tbl_giftkind] ( [giftkind_name], [giftkind_img],[itemid])"&_
			" VALUES ('"&sGiftKindName&"','"&sGiftKindImg&"',"&itemid&") "
	dbacademyget.execute strSql
	
	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbacademyget.close()	:	response.End	
	END IF	
	
	strSql = "SELECT SCOPE_IDENTITY()"
	rsacademyget.Open strSql, dbacademyget
	IF not rsacademyget.EOF THEN
		iGiftKindCode = rsacademyget(0)
	End IF
	rsacademyget.Close	

'response.redirect("popgiftkindReg.asp?sGKN="&sGiftKindName)
%>
	<script language="javascript">

		var strImg = "<%=sGiftKindImg%>";
		opener.document.all.iGK.value = "<%=iGiftKindCode%>";
		opener.document.all.sGKN.value= "<%=sGiftKindName%>";
		if(strImg !=""){
		opener.document.all.spanImg.innerHTML = "<a href=javascript:jsImgView('"+strImg+"')><img src='"+strImg+"' border=0></a>";		
		}
		window.close();	

	</script>	
	
<%
dbget.close()	:	response.End	

Case "KU"  '//사은품 종류 수정	
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10) 
	sGiftKindName 	= html2db(requestCheckVar(Request.Form("sGKN"),60))	
	itemid			= requestCheckVar(Request.Form("itemid"),10) 
	sGiftKindImg	= requestCheckVar(Request.Form("sGKImg"),100) 
	IF itemid = "" THEN itemid =0
		
	IF itemid > 0 THEN
	strSql = "SELECT itemid FROM [db_item].[dbo].[tbl_item] where itemid = "&itemid
	rsget.Open strSql, dbget
	IF rsget.EOF OR rsget.BOF THEN
		rsget.Close	
		Alert_return("존재하지 않는 상품번호입니다. 확인 후 다시 입력해주세요")    	
       dbget.close()	:	response.End	
	End IF
	rsget.Close	
	END IF
	strSql = " UPDATE [db_academy].[dbo].[tbl_giftkind] set [giftkind_name] ='"&sGiftKindName&"', [giftkind_img] ='"&sGiftKindImg&"', [itemid] ="&itemid&_
			" WHERE giftkind_code = "&iGiftKindCode		
	dbget.execute strSql
	
	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbget.close()	:	response.End	
	END IF	

response.redirect("popgiftkindReg.asp?sGKN="&sGiftKindName)
dbget.close()	:	response.End

Case "KM"  '//사은품 종류 수정	2010 추가 
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10) 
	sGiftKindName 	= html2db(requestCheckVar(Request.Form("sGKN"),60))	
	itemid			= requestCheckVar(Request.Form("itemid"),10) 
	sGiftKindImg	= requestCheckVar(Request.Form("sGKImg"),100) 
	s120Img	        = requestCheckVar(Request.Form("S120"),100) 
	s401Img         = requestCheckVar(Request.Form("S401"),100) 
	s402Img         = requestCheckVar(Request.Form("S402"),100) 
	s403Img         = requestCheckVar(Request.Form("S403"),100) 
	s404Img         = requestCheckVar(Request.Form("S404"),100) 
	s405Img         = requestCheckVar(Request.Form("S405"),100) 
	 
	IF itemid = "" THEN itemid =0
		
	IF itemid > 0 THEN
	strSql = "SELECT itemid FROM [db_item].[dbo].[tbl_item] where itemid = "&itemid
	rsget.Open strSql, dbget
	IF rsget.EOF OR rsget.BOF THEN
		rsget.Close	
		Alert_return("존재하지 않는 상품번호입니다. 확인 후 다시 입력해주세요")    	
       dbget.close()	:	response.End	
	End IF
	rsget.Close	
	END IF
	strSql = " UPDATE [db_academy].[dbo].[tbl_giftkind] " & VbCRLF
	strSql = strSql & " set [giftkind_name] ='"&sGiftKindName&"'" & VbCRLF
	strSql = strSql & " , [giftkind_img] ='"&sGiftKindImg&"'" & VbCRLF
	strSql = strSql & " , [itemid] ="&itemid & VbCRLF
	strSql = strSql & " , image120 ='"&s120Img &"'"& VbCRLF
	strSql = strSql & " WHERE giftkind_code = "&iGiftKindCode		
	dbget.execute strSql
	
	strSql = " Delete from db_academy.dbo.tbl_giftkind_AddImage " & VbCRLF
	strSql = strSql & " WHERE gift_kind_code = "&iGiftKindCode		
	dbget.execute strSql
	
	if (s401Img<>"") then
	    strSql = " Insert Into  db_academy.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode 
	    strSql = strSql & " ,1" 
	    strSql = strSql & " ,'"& s401Img& "')"
	    dbget.execute strSql
	end if
	
	if (s402Img<>"") then
	    strSql = " Insert Into  db_academy.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode 
	    strSql = strSql & " ,2" 
	    strSql = strSql & " ,'"& s402Img& "')"
	    dbget.execute strSql
	end if
	
	if (s403Img<>"") then
	    strSql = " Insert Into  db_academy.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode 
	    strSql = strSql & " ,3" 
	    strSql = strSql & " ,'"& s403Img& "')"
	    dbget.execute strSql
	end if
	
	if (s404Img<>"") then
	    strSql = " Insert Into  db_academy.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode 
	    strSql = strSql & " ,4" 
	    strSql = strSql & " ,'"& s404Img& "')"
	    dbget.execute strSql
	end if
	
	if (s405Img<>"") then
	    strSql = " Insert Into  db_academy.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode 
	    strSql = strSql & " ,5" 
	    strSql = strSql & " ,'"& s405Img& "')"
	    dbget.execute strSql
	end if
	
	
	''옵션
	Dim optCnt , gift_kind_option, gift_kind_optionName, gift_kind_Limit, gift_kind_LimitSold, gift_kind_LimitYN
	gift_kind_option = Split(request("gift_kind_option"),",")
	gift_kind_optionName = Split(request("gift_kind_optionName"),",")
	gift_kind_Limit = Split(request("gift_kind_Limit"),",")
	gift_kind_LimitSold = Split(request("gift_kind_LimitSold"),",")
	gift_kind_LimitYN = Split(request("gift_kind_LimitYN"),",")
	
	if IsArray(gift_kind_option) then
	    for i=LBound(gift_kind_option) to UBound(gift_kind_option)
	        if (Trim(gift_kind_option(i))<>"") then
	            strSql = "IF Exists(select * from db_academy.dbo.tbl_giftkind_Option where gift_kind_code="& iGiftKindCode &" and  gift_kind_option='"&Trim(gift_kind_option(i))&"' )"
	            strSql = strSql & " BEGIN"
	            strSql = strSql & " update db_academy.dbo.tbl_giftkind_Option " & VbCRLF
	            strSql = strSql & " set gift_kind_optionName='" & Trim(gift_kind_optionName(i)) & "'"  & VbCRLF
	            strSql = strSql & " ,gift_kind_Limit=" & Trim(gift_kind_Limit(i)) & ""  & VbCRLF
	            strSql = strSql & " ,gift_kind_LimitSold=" & Trim(gift_kind_LimitSold(i)) & ""  & VbCRLF
	            strSql = strSql & " ,gift_kind_optionUsing='" & Trim(request("gift_kind_optionUsing_"&Trim(gift_kind_option(i)))) & "'"  & VbCRLF
	            strSql = strSql & " ,gift_kind_LimitYN='" & Trim(gift_kind_LimitYN(i)) & "'"  & VbCRLF
	            strSql = strSql & " where gift_kind_code="& iGiftKindCode & VbCRLF
	            strSql = strSql & " and gift_kind_option='"&Trim(gift_kind_option(i))&"'" & VbCRLF
	            strSql = strSql & " END"
	            strSql = strSql & " ELSE"
	            strSql = strSql & " BEGIN"
	            strSql = strSql & " Insert Into  db_academy.dbo.tbl_giftkind_Option " & VbCRLF
	            strSql = strSql & " (gift_kind_code, gift_kind_option, gift_kind_optionName, gift_kind_Limit, gift_kind_LimitSold, gift_kind_optionUsing, gift_kind_LimitYN)"
	            strSql = strSql & " values("
	            strSql = strSql & " "& iGiftKindCode & VbCRLF
	            strSql = strSql & " ,'"&Trim(gift_kind_option(i))&"'" & VbCRLF
	            strSql = strSql & " ,'"&Trim(gift_kind_optionName(i))&"'" & VbCRLF
	            strSql = strSql & " ,"&Trim(gift_kind_Limit(i))&"" & VbCRLF
	            strSql = strSql & " ,"&Trim(gift_kind_LimitSold(i))&"" & VbCRLF
	            strSql = strSql & " ,'"&Trim(request("gift_kind_optionUsing_"&Trim(gift_kind_option(i))))&"'" & VbCRLF
	            strSql = strSql & " ,'"&Trim(gift_kind_LimitYN(i))&"'" & VbCRLF
	            strSql = strSql & " )"
	            strSql = strSql & " END"
''response.write strSql  
	            dbget.execute strSql
	        end if
	    next
    end if

	IF Err.Number <> 0 THEN
		Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")    	
       dbget.close()	:	response.End	
	END IF	

response.redirect("popgiftkindManage.asp?iGK="&iGiftKindCode)
dbget.close()	:	response.End

CASE Else
	Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요2")    	
	dbget.close()	:	response.End
END SELECT	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
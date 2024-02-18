<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 90
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/interparkXML/extsiteitemcls.asp"-->

<% if Not (IsAutoScript) then %>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<% end if %>

<%
function getReplyXMLTEST()
    dim replyXML
    replyXML = "<?xml version='1.0' encoding='euc-kr'?>"
    replyXML = replyXML + "<result>"
    replyXML = replyXML + "	<title>상품등록</title>"
    replyXML = replyXML + "	<description>상품 등록 Open API</description>"
    replyXML = replyXML + "	<totalPrdCnt>2</totalPrdCnt>"
    replyXML = replyXML + "	<totalSuccessCnt>2</totalSuccessCnt>"
    replyXML = replyXML + "	<item>"
    replyXML = replyXML + "		<sidx>1</sidx>"
    replyXML = replyXML + "		<originPrdNo>136839</originPrdNo>"
    replyXML = replyXML + "		<prdNo>17060979</prdNo>"
    replyXML = replyXML + "		<prdNm><![CDATA[[텐바이텐 단독!]마이 리틀 프랜드 스티커 2장 세트]]></prdNm>"
    replyXML = replyXML + "		<success>true</success>"
    replyXML = replyXML + "		<ecode>ECODE000</ecode>		"
    replyXML = replyXML + "		<resultMessage>성공적으로 등록 되었습니다.</resultMessage>"
    replyXML = replyXML + "	</item>"
    replyXML = replyXML + "	<item>"
    replyXML = replyXML + "		<sidx>2</sidx>"
    replyXML = replyXML + "		<originPrdNo>140236</originPrdNo>"
    replyXML = replyXML + "		<prdNo>17060980</prdNo>"
    replyXML = replyXML + "		<prdNm><![CDATA[[알뜰]쓰바 4박자 기분 스탬프 세트]]></prdNm>"
    replyXML = replyXML + "		<success>true</success>"
    replyXML = replyXML + "		<ecode>ECODE000</ecode>		"
    replyXML = replyXML + "		<resultMessage>성공적으로 등록 되었습니다.</resultMessage>"
    replyXML = replyXML + "	</item>"
    replyXML = replyXML + "</result>"
    
    getReplyXMLTEST = replyXML
end function

dim mode , pSize, vGubun
dim iParkURL
mode = request("mode")
pSize   = request("pSize")
vGubun = request("gubun")
'response.write "mode="&mode&"&gubun="&vGubun&""
'dbget.close()
'response.end

IF (application("Svr_Info")="Dev") THEN
    iParkURL = "http://sptest.interpark.com"
ELSE
    iParkURL = "http://ipss1.interpark.com"
END IF

iParkURL = iParkURL + "/openapi/product/PrdService.do"

dim iParams, dataUrl

IF (application("Svr_Info")="Dev") THEN
    if (mode="RegAll") then
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/iparkRegItems.xml" ''newRegedItem.asp"
    elseif (mode="EditAll") then
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/iparkEditItems.xml"
    elseif (mode="DelPrd") then 
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/iparkDelItems.xml"
    end if
ELSE
    if (mode="RegAll") then
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/newRegedItem.asp" 
        'dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/iparkRegItems.xml"
    elseif (mode="EditAll") then
        'dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/iparkEditItems.xml"
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/newRegedItem_n.asp?mode=EditPrd&brandid=" & request("brandid")
        if (pSize<>"") then dataUrl= dataUrl & "&pSize="&pSize
    elseif (mode="DelPrd") then
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/newRegedItem.asp?mode=DelPrd&delitemid=" & request("delitemid") 
        'dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/iparkDelItems.xml"
    elseif (mode="DelSoldOut") then
        dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/newRegedItem.asp?mode=DelSoldOut&temp=" & request("temp")
    elseif (mode="DelJaeHyu") then
    	dataUrl = "http://webadmin.10x10.co.kr/admin/etc/interparkXML/newRegedItem.asp?mode=DelJaeHyu&jaehyupagegubun=" & request("jaehyupagegubun") 
    end if
END IF
'dataUrl = server.UrlEncode(dataUrl)

if (mode="RegAll") then
    iParams = "_method=registerPrdInfo&dataUrl=" & dataUrl
elseif (mode="EditAll") or (mode="DelPrd") or (mode="DelSoldOut") or (mode="DelJaeHyu") then
    iParams = "_method=updatePrdInfo&dataUrl=" & dataUrl
end if

dim replyXML
dim i
dim ErrMsg, sqlStr, SuccCnt
dim xmlDoc, ReplyItemCcnt, ioriginPrdNoList, iInterParkPrdNoList, iecodeList, iresultMessageList

SuccCnt = 0

'response.write mode&"<br>"
'response.write iParkURL&"<br>"
'response.write "iParams::"&iParams&"<br>"
'response.end

if (mode="RegAll") then
'response.write iParkURL&"?"
'response.write iParams
    replyXML = SendReqGet(iParkURL, iParams)
    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select

    ''   replyXML = getReplyXMLTEST
'response.write "replyXML<br>"
response.write replyXML    
    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.loadXML replyXML
    
    Set ioriginPrdNoList    = xmlDoc.getElementsByTagName("originPrdNo")
    Set iInterParkPrdNoList = xmlDoc.getElementsByTagName("prdNo")
    Set iecodeList          = xmlDoc.getElementsByTagName("ecode")
    Set iresultMessageList  = xmlDoc.getElementsByTagName("resultMessage")
    
    ReplyItemCcnt = ioriginPrdNoList.length
    
    for i=0 to ReplyItemCcnt-1
    '    response.write ioriginPrdNoList(i).firstChild.nodeValue
    '    response.write iInterParkPrdNoList(i).firstChild.nodeValue
    '    response.write iecodeList(i).firstChild.nodeValue
    '    response.write iresultMessageList(i).firstChild.nodeValue
        
        if (iecodeList(i).firstChild.nodeValue="ECODE000") then
            sqlStr = "update [db_item].[dbo].tbl_interpark_reg_item"
            sqlStr = sqlStr & " set interparkregdate=getdate()"
            sqlStr = sqlStr & " ,interParkPrdNo='" & iInterParkPrdNoList(i).firstChild.nodeValue & "'"
            sqlStr = sqlStr & " ,interparklastupdate=getdate()"
            sqlStr = sqlStr & " where itemid=" & ioriginPrdNoList(i).firstChild.nodeValue
            
            dbget.execute sqlStr
            SuccCnt = SuccCnt + 1
        else
            ErrMsg = ErrMsg & "[" & ioriginPrdNoList(i).firstChild.nodeValue & "]" & iresultMessageList(i).firstChild.nodeValue & VbCrlf
        end if
    Next
    
    Set ioriginPrdNoList    = Nothing
    Set iInterParkPrdNoList = Nothing
    Set iecodeList          = Nothing
    Set iresultMessageList  = Nothing
    
    Set xmlDoc = Nothing
elseif (mode="EditAll") or (mode="DelPrd") or (mode="DelSoldOut") or (mode="DelJaeHyu") then
    replyXML = SendReqGet(iParkURL, iParams)
    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select
''   replyXML = getReplyXMLTEST
'response.write iParkURL
'response.write iParams
'dbget.close()	:	response.End

''response.write replyXML   
    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.loadXML replyXML
    
''response.write replyXML
IF InStr(replyXML,"<TITLE>▣ Interpark Partner Support System - 시스템 오류</TITLE>") >0 then
       ErrMsg =  "▣ Interpark Partner Support System - 시스템 오류"
        
end if
  
    Set ioriginPrdNoList    = xmlDoc.getElementsByTagName("originPrdNo")
    Set iInterParkPrdNoList = xmlDoc.getElementsByTagName("prdNo")
    Set iecodeList          = xmlDoc.getElementsByTagName("ecode")
    Set iresultMessageList  = xmlDoc.getElementsByTagName("resultMessage")
    
    ReplyItemCcnt = ioriginPrdNoList.length
    
    for i=0 to ReplyItemCcnt-1
    '    response.write ioriginPrdNoList(i).firstChild.nodeValue
    '    response.write iInterParkPrdNoList(i).firstChild.nodeValue
    '    response.write iecodeList(i).firstChild.nodeValue
    '    response.write iresultMessageList(i).firstChild.nodeValue
        
        if (iecodeList(i).firstChild.nodeValue="ECODE000") then
            if (mode="DelPrd") then
                sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                sqlStr = sqlStr & " where itemid=" & ioriginPrdNoList(i).firstChild.nodeValue
                
                'response.write sqlStr
                dbget.execute sqlStr
            elseif (mode="DelSoldOut") or (mode="DelJaeHyu") then
                sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                sqlStr = sqlStr & " where itemid=" & ioriginPrdNoList(i).firstChild.nodeValue
                
                'response.write sqlStr
                dbget.execute sqlStr
            else
                ''로그입력(2011-01-18)추가
                sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
                sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode)" & VbCrlf
                sqlStr = sqlStr & " select R.itemid, R.interparkprdno, i.sellcash,i.buycash,i.sellyn,'' as ErrMsg, '"&iecodeList(i).firstChild.nodeValue&"' as errCode" & VbCrlf
                sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
                sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
                sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
                sqlStr = sqlStr & " where R.itemid=" & ioriginPrdNoList(i).firstChild.nodeValue & VbCrlf
                
                dbget.execute sqlStr
                    
                '' 인터파크 가격/ 판매상태도 저장
                sqlStr = "update R" & VbCrlf
                sqlStr = sqlStr & " set interparklastupdate=getdate()" & VbCrlf
                sqlStr = sqlStr & " ,interParkPrdNo='" & iInterParkPrdNoList(i).firstChild.nodeValue & "'" & VbCrlf
                sqlStr = sqlStr & " ,mayiParkPrice=i.sellcash" & VbCrlf
                sqlStr = sqlStr & " ,mayiParkSellYn=i.sellyn" & VbCrlf
                sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item R" & VbCrlf
                sqlStr = sqlStr & "     Join  db_item.dbo.tbl_item i" & VbCrlf
                sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
                sqlStr = sqlStr & " where R.itemid=" & ioriginPrdNoList(i).firstChild.nodeValue
                
                'sqlStr = "update [db_item].[dbo].tbl_interpark_reg_item"
                'sqlStr = sqlStr & " set interparklastupdate=getdate()"
                'sqlStr = sqlStr & " ,interParkPrdNo='" & iInterParkPrdNoList(i).firstChild.nodeValue & "'"
                'sqlStr = sqlStr & " where itemid=" & ioriginPrdNoList(i).firstChild.nodeValue
                'response.write sqlStr
                dbget.execute sqlStr
            end if
            
            SuccCnt = SuccCnt + 1
        else
            
            ErrMsg = ErrMsg & "[" & ioriginPrdNoList(i).firstChild.nodeValue & "]" & iresultMessageList(i).firstChild.nodeValue & VbCrlf
            
            if (mode="DelPrd") then
                ''인터 파크에서 사용 금지 한 경우
                if (iresultMessageList(i).firstChild.nodeValue="인증전이나 판매금지의 판매상태를 임의로 변경할 수 없습니다.") then
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & ioriginPrdNoList(i).firstChild.nodeValue
                    
                    'response.write sqlStr
                    dbget.execute sqlStr
                    
                    ErrMsg = ErrMsg & "[" & ioriginPrdNoList(i).firstChild.nodeValue & "]" & "삭제"
                end if
            end if
            
            ''' 마진설정 잘못되있는 내역있음 -2010-08월 마진 원래대로 수정후 주석처리.
            if (mode="EditAll") then  
                ''''if (iresultMessageList(i).firstChild.nodeValue="해당 상품의 수수료 설정이 잘못 되었습니다. 다시 확인해 주세요.") or (iresultMessageList(i).firstChild.nodeValue="특판행사상품은 옵션정보를 수정할 수 없습니다.") or (iresultMessageList(i).firstChild.nodeValue="인증전이나 판매금지의 판매상태를 임의로 변경할 수 없습니다.") or (iresultMessageList(i).firstChild.nodeValue="상품정보가 존재하지 않습니다.") then
                    ''로그 입력 후 SKIP 하도록 수정..
                    sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
                    sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode)" & VbCrlf
                    sqlStr = sqlStr & " select R.itemid, R.interparkprdno, i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&ErrMsg&"') as ErrMsg, '"&iecodeList(i).firstChild.nodeValue&"' as errCode" & VbCrlf
                    sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
                    sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
                    sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
                    sqlStr = sqlStr & " where R.itemid=" & ioriginPrdNoList(i).firstChild.nodeValue & VbCrlf
                    
                    dbget.execute sqlStr

                    ''오류났으므로 가격/상태는 업데이트 안함
                    sqlStr = "update [db_item].[dbo].tbl_interpark_reg_item" & VbCrlf
                    sqlStr = sqlStr & " set interparklastupdate=getdate()" & VbCrlf
                    sqlStr = sqlStr & " where itemid=" & ioriginPrdNoList(i).firstChild.nodeValue & VbCrlf
                    
                    dbget.execute sqlStr
                    
                    ErrMsg = ErrMsg & "[" & ioriginPrdNoList(i).firstChild.nodeValue & "]" & "Skip"
               ''' end if
            end if
            
        end if
    Next
    
    
    Set ioriginPrdNoList    = Nothing
    Set iInterParkPrdNoList = Nothing
    Set iecodeList          = Nothing
    Set iresultMessageList  = Nothing
    
    Set xmlDoc = Nothing
end if


ErrMsg = Replace(ErrMsg,VbCr,"")
ErrMsg = Replace(ErrMsg,Vblf,"")

If vGubun = "auto" Then
	if Not (IsAutoScript) then
	    if (ErrMsg<>"") then
	        response.write Trim(ErrMsg) & "!<br>"
	    elseif (mode="RegAll") then
	        response.write "" & SuccCnt & " 건 등록 완료" & "!<br>"
	    elseif (mode="EditAll") then
	        response.write "" & SuccCnt & " 건 수정 완료" & "!<br>"
	    elseif (mode="DelPrd") then
	        response.write "" & SuccCnt & " 건 삭제 완료" & "!<br>"
	    elseif (mode="DelSoldOut") then
	        response.write "" & SuccCnt & " 건 삭제 완료" & "!<br>"
	    elseif (mode="DelJaeHyu") then
	        response.write "" & SuccCnt & " 건 삭제 완료" & "!<br>"
	    end if
	else
	    response.write "성공건수 : " & SuccCnt & "!<br>"
	    response.write ErrMsg & "!<br>"
	end if
Else
	if Not (IsAutoScript) then
	    response.write ErrMsg
	    if (ErrMsg<>"") then
	        response.write "<script language='javascript'>alert('" & Trim(ErrMsg) & "');</script>"
	    elseif (mode="RegAll") then
	        response.write "<script language='javascript'>alert('" & SuccCnt & " 건 등록 완료');</script>"
	        'response.write "<script language='javascript'>parent.location.reload();</script>"
	    elseif (mode="EditAll") then
	        response.write "<script language='javascript'>alert('" & SuccCnt & " 건 수정 완료');</script>"
	        'response.write "<script language='javascript'>parent.location.reload();</script>"
	    elseif (mode="DelPrd") then
	        response.write "<script language='javascript'>alert('" & SuccCnt & " 건 삭제 완료');</script>"
	        'response.write "<script language='javascript'>parent.location.reload();</script>"
	    elseif (mode="DelSoldOut") then
	        response.write "<script language='javascript'>alert('" & SuccCnt & " 건 삭제 완료');</script>"
	    elseif (mode="DelJaeHyu") then
	        response.write "<script language='javascript'>alert('" & SuccCnt & " 건 삭제 완료');</script>"
	    end if
	else
	    response.write "성공건수 : " & SuccCnt
	    response.write ErrMsg
	
	end if
End IF
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
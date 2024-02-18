<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 90
%>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<%

function RightCommaDel(ostr)
    dim restr
    restr = ""
    if IsNULL(ostr) then Exit function
    
    restr = Trim(ostr)
    if (Right(restr,1)=",") then restr=Left(restr,Len(restr)-1)
    RightCommaDel = restr
end function

Function GetRaiseValue(value) 
    If Fix(value) < value Then 
    GetRaiseValue = Fix(value) + 1 
    Else 
    GetRaiseValue = Fix(value) 
    End If 
End Function 

''일괄 상품등록 최대 갯수.
Dim CMaxUploadItem : CMaxUploadItem = 30 
Dim CMaxPerPage    : CMaxPerPage = 30
dim delim
delim = VbCrlf 

''dbget.close()	:	response.End

dim appPath, FileName
dim mode, IsEditMode, IsDelMode, IsDelSoldOut, IsDelJaeHyu
Dim pSize
mode    = request("mode")
pSize   = request("pSize")

if (pSize<>"") then
    if IsNumeric(pSize) then
        if (pSize<80) then 
            CMaxUploadItem = pSize 
            CMaxPerPage = pSize
        end if
    end if
end if

IsEditMode = (mode="EditPrd")
IsDelMode  = (mode="DelPrd")
IsDelJaeHyu = (mode="DelJaeHyu")
IsDelSoldOut = (mode="DelSoldOut")

appPath = server.mappath("/admin/etc/interparkXML/") + "\"

if (IsEditMode) then
    FileName = "iparkEditItems.xml"
elseif (IsDelMode) then
    FileName = "iparkDelItems.xml"
elseif (IsDelSoldOut) then
    FileName = "iparkDelSoldOutItems.xml"
elseif (IsDelJaeHyu) then
    FileName = "iparkDelSoldOutItems.xml"
else
    FileName = "iparkRegItems.xml"
end if

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

if (ref<>"222.231.7.189") and (Left(ref,10)<>"210.92.223") and (Left(ref,10)<>"61.252.133") then
    response.write "Not Valid Referer"
    dbget.close()	:	response.End
end if


'sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
'sqlStr = sqlStr + " (ref) values('" + "IPK1-" + ref + "')"
'dbget.execute sqlStr

dim i,j,k, seqIdx
dim oiParktotalpage, oiParkitem, totalpage
dim IsTheLastOption, IsOptionExists, optstr, buf, optbuf, keywordsBuf, keywordsStr, NotSoldOutOptionExists
dim ioptCodeBuf, ioptNameBuf, ioptTypeName, ioptLimitNo, ioptAddPrice
dim IsAllSoldOutOption
dim fso, tFile

seqIdx = 0

    Set oiParktotalpage = new CiParkRegItem
	oiParktotalpage.FPageSize = CMaxPerPage
	oiParktotalpage.FBrandID = Request("brandid")

	if (IsDelMode) then
	    totalpage = 1
	elseif (IsDelSoldOut) then
	    totalpage = 1
	elseif (IsDelJaeHyu) then
	    totalpage = 1
	else   
    	if (IsEditMode) then
    	    oiParktotalpage.GetIParkEditItemTotalPage
    	else
    	    oiParktotalpage.GetIParkRegItemTotalPage
    	end if
    	
    	totalpage = oiParktotalpage.FtotalPage
	end if
	
'response.write oiParktotalpage.FTotalCount
'response.write totalpage
'dbget.close()	:	response.End
	if totalpage>CLng(CMaxUploadItem/CMaxPerPage) then totalpage = CLng(CMaxUploadItem/CMaxPerPage)
	Set oiParktotalpage = Nothing

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
	tFile.WriteLine "<?xml version='1.0' encoding='euc-kr'?>"
	tFile.WriteLine "<result>"
	tFile.WriteLine "<title>상품정보 데이터</title>"
	tFile.WriteLine "<description>상품정보 등록을 위한 Open Api 데이터</description>"
     
	for j=0 to totalpage - 1
		set oiParkitem = new CiParkRegItem
		oiParkitem.FCurrPage = j+1
		oiParkitem.FPageSize = CMaxPerPage
		oiParkitem.FBrandID = Request("brandid")
		oiParkitem.FTemp = Request("temp")
		if (IsEditMode) then
	        oiParkitem.GetIParkEditItemList
	    elseif (IsDelMode) then
	        oiParkitem.GetIParkOneItemList request("delitemid")
	    elseif (IsDelSoldOut) then
	        oiParkitem.FCurrPage = 1
	        oiParkitem.FPageSize = 10
	        oiParkitem.GetIParkDelSoldOutItemList
	    elseif (IsDelJaeHyu) then
	    	oiParkitem.FJaeHyuPageGubun = request("jaehyupagegubun")
	        oiParkitem.GetIParkDelJaeHyuItemList
		else
		    oiParkitem.GetIParkRegItemList
	    end if
	
		for i=0 to oiParkitem.FResultCount-1
		    ''옵션List ---
			IsTheLastOption = false
			
			if (oiParkitem.FItemList(i).Fitemoption="0000") then
			    IsOptionExists = false
				optstr = ""
			else
			    IsOptionExists = true
			    
				if (i+1<=oiParkitem.FResultCount-1) then
					if (oiParkitem.FItemList(i).FItemID=oiParkitem.FItemList(i+1).FItemID) then
					    if (Not oiParkitem.FItemList(i).IsOptionSoldOut) then
					        ioptTypeName = oiParkitem.FItemList(i).FItemOptionTypeName
    						ioptCodeBuf = ioptCodeBuf + oiParkitem.FItemList(i).FItemOption + "," 
    						ioptNameBuf = ioptNameBuf + Replace(oiParkitem.FItemList(i).FItemOptionName,","," ") + "," 
    						ioptAddPrice = ioptAddPrice + CStr(oiParkitem.FItemList(i).Foptaddprice) + "," 
    						ioptLimitNo = ioptLimitNo + CStr(oiParkitem.FItemList(i).getOptionLimitNo) + "," 
    						NotSoldOutOptionExists = true
    					end if
					else
					    if (Not oiParkitem.FItemList(i).IsOptionSoldOut) then
					        ioptTypeName = oiParkitem.FItemList(i).FItemOptionTypeName
    					    ioptCodeBuf = ioptCodeBuf + oiParkitem.FItemList(i).FItemOption + "," 
    						ioptNameBuf = ioptNameBuf + Replace(oiParkitem.FItemList(i).FItemOptionName,","," ") + "," 
    						ioptAddPrice = ioptAddPrice + CStr(oiParkitem.FItemList(i).Foptaddprice) + "," 
    						ioptLimitNo = ioptLimitNo + CStr(oiParkitem.FItemList(i).getOptionLimitNo) + "," 
    						NotSoldOutOptionExists = true
    					end if
    					IsTheLastOption = true
    					
    					ioptNameBuf = RightCommaDel(ioptNameBuf)
    				    ioptCodeBuf = RightCommaDel(ioptCodeBuf)
    				    ioptAddPrice= RightCommaDel(ioptAddPrice)
    				    ioptLimitNo = RightCommaDel(ioptLimitNo)
    				    
    				    if (ioptTypeName="") then ioptTypeName="옵션명"
    				    
    				    optstr = ioptTypeName + "<" + ioptNameBuf + ">" 
                        if (ioptLimitNo<>"") then
                            optstr = optstr + "수량<" + ioptLimitNo + ">"
                        end if
                        optstr = optstr + "추가금액<" + ioptAddPrice + ">"
                        optstr = optstr + "옵션코드<" + ioptCodeBuf + ">"
                    
						'optstr = "옵션명<" + ioptNameBuf + ">" 
                        'optstr = optstr + "옵션코드<" + ioptCodeBuf + ">"
						ioptCodeBuf = ""
						ioptNameBuf = ""
						ioptTypeName = ""
					    ioptLimitNo = ""
					    ioptAddPrice = ""
					end if
				elseif (i=oiParkitem.FResultCount-1) then
				    if (Not oiParkitem.FItemList(i).IsOptionSoldOut) then
				        ioptTypeName = oiParkitem.FItemList(i).FItemOptionTypeName
    				    ioptCodeBuf = ioptCodeBuf + oiParkitem.FItemList(i).FItemOption + "," 
    					ioptNameBuf = ioptNameBuf + Replace(oiParkitem.FItemList(i).FItemOptionName,","," ") + "," 
    					ioptAddPrice = ioptAddPrice + CStr(oiParkitem.FItemList(i).Foptaddprice) + "," 
    					ioptLimitNo = ioptLimitNo + CStr(oiParkitem.FItemList(i).getOptionLimitNo) + "," 
                        NotSoldOutOptionExists = true
                        
                        IsTheLastOption = true
    				end if
    				
    				ioptNameBuf = RightCommaDel(ioptNameBuf)
    				ioptCodeBuf = RightCommaDel(ioptCodeBuf)
    				ioptAddPrice= RightCommaDel(ioptAddPrice)
    				ioptLimitNo = RightCommaDel(ioptLimitNo)
    				
    				if (ioptTypeName="") then ioptTypeName="옵션명"
    				
                    optstr = ioptTypeName + "<" + ioptNameBuf + ">" 
                    if (ioptLimitNo<>"") then
                        optstr = optstr + "수량<" + ioptLimitNo + ">"
                    end if
                    optstr = optstr + "추가금액<" + ioptAddPrice + ">"
                    optstr = optstr + "옵션코드<" + ioptCodeBuf + ">"
					ioptCodeBuf = ""
					ioptNameBuf = ""
					ioptTypeName = ""
					ioptAddPrice = ""
					ioptLimitNo = ""
				end if
			end if
			'' 옵션 String 끝
			
			
			buf = ""
            keywordsStr = ""
            
            if (optstr<>"") then '' and (optstr<>delim) 
                IsTheLastOption = true
            end if
            
            
            if (Not IsOptionExists) or (IsTheLastOption) then
                
                if (Right(optstr,Len("옵션코드<>"))="옵션코드<>") then
                    IsAllSoldOutOption = True
                else 
                    IsAllSoldOutOption = False
                end if
                
			    keywordsBuf = oiParkitem.FItemList(i).Fkeywords
			    keywordsBuf = Split(keywordsBuf,",")
			    
			    ''키워드 최대 4개 콤마구분
			    for k=0 to 2
			        if UBound(keywordsBuf)>k then keywordsStr = keywordsStr + Trim(keywordsBuf(k)) + ","
			    next
			    
			    keywordsStr = "텐바이텐," + keywordsStr 
			    keywordsStr = RightCommaDel(keywordsStr)
			    
			    buf = buf + "<item>" + delim
				buf = buf + "<sidx>" & seqIdx + 1 & "</sidx>" + delim  
				if (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
				    buf = buf + "<prdNo>" & + CStr(oiParkitem.FItemList(i).FInterparkPrdNo) & "</prdNo>" + delim  
				end if
				buf = buf + "<supplyEntrNo>3000010614</supplyEntrNo>" + delim   ''업체번호 고정(3000010614, 테스트, 리얼 동일)



				'####### 20111025 이전
				'IF (application("Svr_Info")="Dev") THEN
				'    buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim   
				'ELSE
				'    buf = buf + "<supplyCtrtSeq>" + CStr(oiParkitem.FItemList(i).GetSupplyCtrtSeq) + "</supplyCtrtSeq>" + delim           ''공급계약일련번호 의류(2), 잡화(3), 리빙(4)
			    'END IF
			    
				if (oiParkitem.FItemList(i).Fdeliverytype="9") then ''업체조건배송
				    buf = buf + "<supplyCtrtSeq>6</supplyCtrtSeq>" + delim
				else
				    buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim
				end if
			    
			    
			    
			    if Not (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then 
    				buf = buf + "<prdStat>01</prdStat>" + delim                     ''01 새상품
    				buf = buf + "<shopNo>0000100000</shopNo>" + delim               ''상점번호 API업체 고정
    				IF (application("Svr_Info")="Dev") THEN
    				    buf = buf + "<omDispNo>001830114002</omDispNo>" + delim
    				ELSE
    				    buf = buf + "<omDispNo>" + oiParkitem.FItemList(i).Finterparkdispcategory + "</omDispNo>" + delim                     ''인터파크 전시코드
    				END IF
			    end IF
			
				buf = buf + "<prdNm><![CDATA[[텐바이텐] " + Replace(Replace(Replace(Replace(Replace(oiParkitem.FItemList(i).FBrandNameKor + " " + CStr(oiParkitem.FItemList(i).FItemName),"'",""),Chr(34),""),"<",""),">",""),"^","") + "]]></prdNm>" + delim ''상품명
				buf = buf + "<hdelvMafcEntrNm><![CDATA[" + CStr(oiParkitem.FItemList(i).FMakerName) + "]]></hdelvMafcEntrNm>" + delim ''제조업체명
				buf = buf + "<prdOriginTp><![CDATA[" + oiParkitem.FItemList(i).GetSourcearea + "]]></prdOriginTp>" + delim       ''원산지
				buf = buf + "<taxTp>" + oiParkitem.FItemList(i).GetInterParkTaxTp + "</taxTp>" + delim      ''과세 01, 면세02, 영세 03
				buf = buf + "<ordAgeRstrYn>N</ordAgeRstrYn>" + delim            ''성인용품
				if (IsAllSoldOutOption) then 
				    buf = buf + "<saleStatTp>05</saleStatTp>" + delim  ''판매중01, 품절02, 판매중지03, 일시품절05
				elseif (IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
				    buf = buf + "<saleStatTp>03</saleStatTp>" + delim  ''판매중01, 품절02, 판매중지03, 절판04, 일시품절05
				else
				    buf = buf + "<saleStatTp>" + oiParkitem.FItemList(i).GetInterParkSaleStatTp + "</saleStatTp>" + delim  ''판매중01, 품절02, 판매중지03, 일시품절05
			    end if
			    
			    if (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
			    buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).FSellcash/10)*10) + "</saleUnitcost>" + delim  ''판매가
			    else
				buf = buf + "<saleUnitcost>" + CStr(oiParkitem.FItemList(i).FSellcash) + "</saleUnitcost>" + delim  ''판매가
				end if
				buf = buf + "<saleLmtQty>" + CStr(oiParkitem.FItemList(i).GetInterParkLmtQty)+ "</saleLmtQty>" + delim  ''한정수량
				buf = buf + "<saleStrDts>" + Replace(Left(CStr(now()),10),"-","") + "</saleStrDts>" + delim         ''판매시작일
				buf = buf + "<saleEndDts>" + oiParkitem.FItemList(i).GetSellEndDateStr + "</saleEndDts>" + delim         ''판매종료일
				
				
				
				'####### 20111025 이전
				'if (oiParkitem.FItemList(i).Fdeliverytype="4") then ''텐바이텐 무료배송만
				'    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
				'else
				'    buf = buf + "<proddelvCostUseYn>N</proddelvCostUseYn>" + delim  ''상품별배송비여부 30000/2500
				'end if
				
				if (oiParkitem.FItemList(i).Fdeliverytype="9") then ''업체조건배송
				    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
				else
					if (oiParkitem.FItemList(i).Fdeliverytype="4") then ''텐바이텐 무료배송만
					    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
					else
					    buf = buf + "<proddelvCostUseYn>N</proddelvCostUseYn>" + delim  ''상품별배송비여부 30000/2500
					end if
				end if
				


				buf = buf + "<prdBasisExplanEd><![CDATA[" + Replace(oiParkitem.FItemList(i).getItemPreInfodataHTML,"","") + Replace(oiParkitem.FItemList(i).FItemContent,"","") + Replace(oiParkitem.FItemList(i).getItemInfoImageHTML,"","") + "]]></prdBasisExplanEd>" + delim     ''상품설명
				buf = buf + "<zoomImg><![CDATA[" + oiParkitem.FItemList(i).get400Image + "]]></zoomImg>" + delim    ''대표이미지
				buf = buf + "<prdPrefix><![CDATA[" + oiParkitem.FItemList(i).GetprdPrefixStr + "]]></prdPrefix>" + delim
				''buf = buf + "<prdPostfix></prdPostfix>" + delim
				buf = buf + "<prdKeywd><![CDATA[" + Replace(keywordsStr,"'","") + "]]></prdKeywd>" + delim
				buf = buf + "<brandNm><![CDATA[" + oiParkitem.FItemList(i).Fbrandname + "]]></brandNm>" + delim
				buf = buf + "<entrPoint>" + CStr(oiParkitem.FItemList(i).GetInterParkentrPoint)+ "</entrPoint>" + delim              ''업체포인트
				buf = buf + "<minOrdQty>1</minOrdQty>" + delim              ''최소주문수량
				
				if (IsAllSoldOutOption) or (IsDelMode) or (IsDelSoldOut) or (IsDelJaeHyu) then

				else
				    buf = buf + "<prdOption><![CDATA[" + optstr + "]]></prdOption>" + delim     ''옵션.
			    end if
			
				if (oiParkitem.FItemList(i).Fdeliverytype="4") then 
				buf = buf + "<delvCost>0</delvCost>" + delim               ''상품별 배송비
				end if
				buf = buf + "<delvAmtPayTpCom>02</delvAmtPayTpCom>" + delim   ''배송비결제방식 02선불
				buf = buf + "<delvCostApplyTp>02</delvCostApplyTp>" + delim   ''배송비적용방식 개당01, 무조건02
				
				
				
				'####### 20111025 이전
				'if (oiParkitem.FItemList(i).IsFreeBeasong) then 
				'    buf = buf + "<freedelvStdCnt>1</freedelvStdCnt>" + delim             ''무료배송기준수량
				'end if
				
				if (oiParkitem.FItemList(i).Fdeliverytype <> "9") then ''업체조건배송
					if (oiParkitem.FItemList(i).IsFreeBeasong) then 
					    buf = buf + "<freedelvStdCnt>1</freedelvStdCnt>" + delim             ''무료배송기준수량
					end if
				end if
				
				
				buf = buf + "<spcaseEd><![CDATA[" + oiParkitem.FItemList(i).getOrderCommentStr + "]]></spcaseEd>" + delim
				
				buf = buf + "<pointmUseYn>" + CStr(oiParkitem.FItemList(i).GetpointmUseYn) + "</pointmUseYn>" + delim            ''포인트몰
		        buf = buf + "<ippSubmitYn>Y</ippSubmitYn>" + delim            ''가격비교등록여부
				buf = buf + "<originPrdNo>" + CStr(oiParkitem.FItemList(i).FItemID) + "</originPrdNo>" + delim     ''상품번호
				
				IF (application("Svr_Info")="Dev") THEN
				    buf = buf + "<shopDispInfo><![CDATA[전시타입<2>상점번호<0000100000>전시번호<001410038001001>]]></shopDispInfo>" + delim
				ELSE
    				IF Not IsNULL(oiParkitem.FItemList(i).Finterparkstorecategory) and (oiParkitem.FItemList(i).Finterparkstorecategory<>"") then
    				    buf = buf + "<shopDispInfo><![CDATA[전시타입<2>상점번호<0000100000>전시번호<" + oiParkitem.FItemList(i).Finterparkstorecategory + ">]]></shopDispInfo>" + delim   ''추가전시
    				END IF
    		    END IF
				
				buf = buf + "</item>" + delim
				optstr = ""
				
				NotSoldOutOptionExists = false
				
			    seqIdx = seqIdx + 1
			end if
			
			if buf<>"" then
			    tFile.WriteLine buf
		    end if
		next
		
		
		set oiParkitem = Nothing
	next

	tFile.WriteLine "</result>"

	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing


'sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
'sqlStr = sqlStr + " (ref) values('" + "IPK2-" + ref + "')"
'dbget.execute sqlStr
	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
IF (application("Svr_Info")<>"Dev") THEN
    response.redirect "/admin/etc/interparkXML/" & FileName
END IF
%>
<% @ language=vbscript %>
<%
	option explicit

	'// ��������� XML�� ��������
	Response.Clear
	Response.contentType = "text/xml; charset=euc-kr"
	Response.Write "<?xml version=""1.0"" encoding=""EUC-KR"" ?>"
	Server.ScriptTimeOut = 90
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/interparkXML/extsiteitemcls.asp"-->
<%
'// ���� �ּҸ� ����
Dim QrStr, arrQr, itemid, vMode
QrStr = Request.ServerVariables("QUERY_STRING")
arrQr = Split(QrStr,"/")
QrStr = arrQr(ubound(arrQr))
if lcase(arrQr(ubound(arrQr)-1))<>"interparkxml" then
	Response.Write "�߸��� ����!"
	dbget.close()	:Response.End
else
	'# ��ǰ��ȣ ����
	itemid = Split(Replace(Replace(lcase(arrQr(ubound(arrQr))),"tenitem",""),".xml",""),"_")(0)
	vMode = Split(Replace(Replace(lcase(arrQr(ubound(arrQr))),"tenitem",""),".xml",""),"_")(1)
	if Not(isNumeric(itemid)) then
		Response.Write "�߸��� ����!"
		dbget.close()	:Response.End
	end if
end if

'// ���IP ����
dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

if (ref<>"222.231.7.189") and (Left(ref,10)<>"210.92.223") and (Left(ref,10)<>"61.252.133") then
    response.write "Not Valid Referer"
    dbget.close()	:	response.End
end if



'--------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------
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

''�ϰ� ��ǰ��� �ִ� ����.
Dim CMaxUploadItem : CMaxUploadItem = 30 
Dim CMaxPerPage    : CMaxPerPage = 30
dim delim
delim = VbCrlf 



dim appPath, FileName
dim mode, IsEditMode, IsDelMode, IsDelSoldOut, IsDelJaeHyu
Dim pSize
mode    = request("mode")
pSize   = request("pSize")

mode = vMode

if (pSize<>"") then
    if IsNumeric(pSize) then
        if (pSize<80) then 
            CMaxUploadItem = pSize 
            CMaxPerPage = pSize
        end if
    end if
end if

IsEditMode = (mode="editprd")
IsDelMode  = (mode="delprd")
IsDelJaeHyu = (mode="deljaehyu")
IsDelSoldOut = (mode="delsoldout")



dim i,j,k, seqIdx
dim oiParktotalpage, oiParkitem, totalpage
dim IsTheLastOption, IsOptionExists, optstr, buf, optbuf, keywordsBuf, keywordsStr, NotSoldOutOptionExists
dim ioptCodeBuf, ioptNameBuf, ioptTypeName, ioptLimitNo, ioptAddPrice
dim IsAllSoldOutOption
dim fso, tFile

seqIdx = 0
'--------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------- <!-- //-->

	Response.Write "<result>"
	Response.Write "<title>��ǰ���� ������</title>"
	Response.Write "<description>��ǰ���� ����� ���� Open Api ������</description>"

		set oiParkitem = new CiParkRegItem
		oiParkitem.FCurrPage = j+1
		oiParkitem.FPageSize = CMaxPerPage
		oiParkitem.FItemID = itemid
		oiParkitem.GetIParkOneItemNew
	    
	    
		for i=0 to 0
		    ''�ɼ�List ---
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
    				    
    				    if (ioptTypeName="") then ioptTypeName="�ɼǸ�"
    				    
    				    optstr = ioptTypeName + "<" + ioptNameBuf + ">" 
                        if (ioptLimitNo<>"") then
                            optstr = optstr + "����<" + ioptLimitNo + ">"
                        end if
                        optstr = optstr + "�߰��ݾ�<" + ioptAddPrice + ">"
                        optstr = optstr + "�ɼ��ڵ�<" + ioptCodeBuf + ">"
                    
						'optstr = "�ɼǸ�<" + ioptNameBuf + ">" 
                        'optstr = optstr + "�ɼ��ڵ�<" + ioptCodeBuf + ">"
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
    				
    				if (ioptTypeName="") then ioptTypeName="�ɼǸ�"
    				
                    optstr = ioptTypeName + "<" + ioptNameBuf + ">" 
                    if (ioptLimitNo<>"") then
                        optstr = optstr + "����<" + ioptLimitNo + ">"
                    end if
                    optstr = optstr + "�߰��ݾ�<" + ioptAddPrice + ">"
                    optstr = optstr + "�ɼ��ڵ�<" + ioptCodeBuf + ">"
					ioptCodeBuf = ""
					ioptNameBuf = ""
					ioptTypeName = ""
					ioptAddPrice = ""
					ioptLimitNo = ""
				end if
			end if
			'' �ɼ� String ��
			
			
			buf = ""
            keywordsStr = ""
            
            if (optstr<>"") then '' and (optstr<>delim) 
                IsTheLastOption = true
            end if
            
            
            if (Not IsOptionExists) or (IsTheLastOption) then
                
                if (Right(optstr,Len("�ɼ��ڵ�<>"))="�ɼ��ڵ�<>") then
                    IsAllSoldOutOption = True
                else 
                    IsAllSoldOutOption = False
                end if
                
			    keywordsBuf = oiParkitem.FItemList(i).Fkeywords
			    keywordsBuf = Split(keywordsBuf,",")
			    
			    ''Ű���� �ִ� 4�� �޸�����
			    for k=0 to 2
			        if UBound(keywordsBuf)>k then keywordsStr = keywordsStr + Trim(keywordsBuf(k)) + ","
			    next
			    
			    keywordsStr = "�ٹ�����," + keywordsStr 
			    keywordsStr = RightCommaDel(keywordsStr)
			    
			    buf = buf + "<item>" + delim
				buf = buf + "<sidx>" & seqIdx + 1 & "</sidx>" + delim  
				if (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
				    buf = buf + "<prdNo>" & + CStr(oiParkitem.FItemList(i).FInterparkPrdNo) & "</prdNo>" + delim  
				end if
				buf = buf + "<supplyEntrNo>3000010614</supplyEntrNo>" + delim   ''��ü��ȣ ����(3000010614, �׽�Ʈ, ���� ����)
				IF (application("Svr_Info")="Dev") THEN
				    buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim   
				ELSE
				    buf = buf + "<supplyCtrtSeq>" + CStr(oiParkitem.FItemList(i).GetSupplyCtrtSeq) + "</supplyCtrtSeq>" + delim           ''���ް���Ϸù�ȣ �Ƿ�(2), ��ȭ(3), ����(4)
			    END IF
			    
			    if Not (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then 
    				buf = buf + "<prdStat>01</prdStat>" + delim                     ''01 ����ǰ
    				buf = buf + "<shopNo>0000100000</shopNo>" + delim               ''������ȣ API��ü ����
    				IF (application("Svr_Info")="Dev") THEN
    				    buf = buf + "<omDispNo>001830114002</omDispNo>" + delim
    				ELSE
    				    buf = buf + "<omDispNo>" + oiParkitem.FItemList(i).Finterparkdispcategory + "</omDispNo>" + delim                     ''������ũ �����ڵ�
    				END IF
			    end IF
			
				buf = buf + "<prdNm><![CDATA[[�ٹ�����] " + Replace(Replace(Replace(Replace(Replace(oiParkitem.FItemList(i).FBrandNameKor + " " + CStr(oiParkitem.FItemList(i).FItemName),"'",""),Chr(34),""),"<",""),">",""),"^","") + "]]></prdNm>" + delim ''��ǰ��
				buf = buf + "<hdelvMafcEntrNm><![CDATA[" + CStr(oiParkitem.FItemList(i).FMakerName) + "]]></hdelvMafcEntrNm>" + delim ''������ü��
				buf = buf + "<prdOriginTp><![CDATA[" + oiParkitem.FItemList(i).GetSourcearea + "]]></prdOriginTp>" + delim       ''������
				buf = buf + "<taxTp>" + oiParkitem.FItemList(i).GetInterParkTaxTp + "</taxTp>" + delim      ''���� 01, �鼼02, ���� 03
				buf = buf + "<ordAgeRstrYn>N</ordAgeRstrYn>" + delim            ''���ο�ǰ
				if (IsAllSoldOutOption) then 
				    buf = buf + "<saleStatTp>05</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, �Ͻ�ǰ��05
				elseif (IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
				    buf = buf + "<saleStatTp>03</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, ����04, �Ͻ�ǰ��05
				else
				    buf = buf + "<saleStatTp>" + oiParkitem.FItemList(i).GetInterParkSaleStatTp + "</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, �Ͻ�ǰ��05
			    end if
			    
			    if (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
			    buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).FSellcash/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
			    else
				buf = buf + "<saleUnitcost>" + CStr(oiParkitem.FItemList(i).FSellcash) + "</saleUnitcost>" + delim  ''�ǸŰ�
				end if
				buf = buf + "<saleLmtQty>" + CStr(oiParkitem.FItemList(i).GetInterParkLmtQty)+ "</saleLmtQty>" + delim  ''��������
				buf = buf + "<saleStrDts>" + Replace(Left(CStr(now()),10),"-","") + "</saleStrDts>" + delim         ''�ǸŽ�����
				buf = buf + "<saleEndDts>" + oiParkitem.FItemList(i).GetSellEndDateStr + "</saleEndDts>" + delim         ''�Ǹ�������
				if (oiParkitem.FItemList(i).Fdeliverytype="4") then ''�ٹ����� �����۸�
				    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
				else
				    buf = buf + "<proddelvCostUseYn>N</proddelvCostUseYn>" + delim  ''��ǰ����ۺ񿩺� 30000/2500
				end if
				buf = buf + "<prdBasisExplanEd><![CDATA[" + Replace(oiParkitem.FItemList(i).getItemPreInfodataHTML,"","") + Replace(oiParkitem.FItemList(i).FItemContent,"","") + Replace(oiParkitem.FItemList(i).getItemInfoImageHTML,"","") + "]]></prdBasisExplanEd>" + delim     ''��ǰ����
				buf = buf + "<zoomImg><![CDATA[" + oiParkitem.FItemList(i).get400Image + "]]></zoomImg>" + delim    ''��ǥ�̹���
				buf = buf + "<prdPrefix><![CDATA[" + oiParkitem.FItemList(i).GetprdPrefixStr + "]]></prdPrefix>" + delim
				''buf = buf + "<prdPostfix></prdPostfix>" + delim
				buf = buf + "<prdKeywd><![CDATA[" + Replace(keywordsStr,"'","") + "]]></prdKeywd>" + delim
				buf = buf + "<brandNm><![CDATA[" + oiParkitem.FItemList(i).Fbrandname + "]]></brandNm>" + delim
				buf = buf + "<entrPoint>" + CStr(oiParkitem.FItemList(i).GetInterParkentrPoint)+ "</entrPoint>" + delim              ''��ü����Ʈ
				buf = buf + "<minOrdQty>1</minOrdQty>" + delim              ''�ּ��ֹ�����
				
				buf = buf + "<optPrirTp>02</optPrirTp>" + delim
				

				If Replace(optstr," ","") <> "" Then
					buf = buf + "<prdOption><![CDATA[{" + Replace(optstr," ","") + "}]]></prdOption>" + delim     ''�ɼ�.
				End If

			
				if (oiParkitem.FItemList(i).Fdeliverytype="4") then 
				buf = buf + "<delvCost>0</delvCost>" + delim               ''��ǰ�� ��ۺ�
				end if
				buf = buf + "<delvAmtPayTpCom>02</delvAmtPayTpCom>" + delim   ''��ۺ������� 02����
				buf = buf + "<delvCostApplyTp>02</delvCostApplyTp>" + delim   ''��ۺ������� ����01, ������02
				if (oiParkitem.FItemList(i).IsFreeBeasong) then 
				    buf = buf + "<freedelvStdCnt>1</freedelvStdCnt>" + delim             ''�����۱��ؼ���
				end if
				buf = buf + "<spcaseEd><![CDATA[" + oiParkitem.FItemList(i).getOrderCommentStr + "]]></spcaseEd>" + delim
				
				buf = buf + "<pointmUseYn>" + CStr(oiParkitem.FItemList(i).GetpointmUseYn) + "</pointmUseYn>" + delim            ''����Ʈ��
		        buf = buf + "<ippSubmitYn>Y</ippSubmitYn>" + delim            ''���ݺ񱳵�Ͽ���
				buf = buf + "<originPrdNo>" + CStr(oiParkitem.FItemList(i).FItemID) + "</originPrdNo>" + delim     ''��ǰ��ȣ
				
				IF (application("Svr_Info")="Dev") THEN
				    buf = buf + "<shopDispInfo><![CDATA[����Ÿ��<2>������ȣ<0000100000>���ù�ȣ<001410038001001>]]></shopDispInfo>" + delim
				ELSE
    				IF Not IsNULL(oiParkitem.FItemList(i).Finterparkstorecategory) and (oiParkitem.FItemList(i).Finterparkstorecategory<>"") then
    				    buf = buf + "<shopDispInfo><![CDATA[����Ÿ��<2>������ȣ<0000100000>���ù�ȣ<" + oiParkitem.FItemList(i).Finterparkstorecategory + ">]]></shopDispInfo>" + delim   ''�߰�����
    				END IF
    		    END IF
				
				buf = buf + "</item>" + delim
				optstr = ""
				
				NotSoldOutOptionExists = false
				
			    seqIdx = seqIdx + 1
			end if
			
			if buf<>"" then
			    Response.Write buf
		    end if
		next
		
		Response.Write "</result>"
		
		set oiParkitem = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

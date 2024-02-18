<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 900

''API �ּ�
'' http://www.interpark.com/openapi/site/APIInsertSpecNew.jsp

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
dim delim, ref, resultRow
dim prdNo,prdNm,originPrdNo
dim mode : mode = requestCheckvar(request("mode"),32)
dim param1 : param1 = requestCheckvar(request("param1"),32)
dim param2 : param2 = requestCheckvar(request("param2"),32)
dim param3 : param3 = requestCheckvar(request("param3"),32)
dim cksel : cksel = requestCheckvar(request("cksel"),1024)
dim locNo : locNo = requestCheckvar(request("locNo"),10)
dim eventidArr : eventidArr= Trim(request("eventidArr"))
dim makeridArr : makeridArr= Trim(request("makeridArr"))

dim iParkURL, iParams, replyXML, itemidARR, itemid
dim i
dim ErrCode, ErrMsg, sqlStr, retCNT, SuccCnt, totCNT, pErrMsg
Dim xmlDoc, Nodes, SubNodes
dim dispNo , dispNm ,dispYn ,regDts ,modDts , AssignedRow, ArrRows, optArrRows
Dim optlp, optlpCode, optlpName, optlpUsing, optlpSu, optlpStr
dim regCNT, upCNT
dim errorNodes
dim oInterParkitem
dim bufStr

Dim dataUrl
Dim retVal, PrdSaleUnitcost, iPrdNm


delim = VbCrlf
ref = request.serverVariables("HTTP_REFERER")

'rw "mode:"&mode
%>
<% if Not (IsAutoScript) then %>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<% end if %>

<%
function getAPIParam(mode)
    SELECT CASE mode
        CASE "cateRcv"  : getAPIParam    = "_method=GetBasicCategoryForAPI&citeKey="&getCiteKey(mode)&"&secretKey="&getSecretKey(mode)
        CASE "regitemONE" : getAPIParam  = "_method=InsertProductAPIData&citeKey="&getCiteKey(mode)&"&secretKey="&getSecretKey(mode)&"&dataUrl="&dataUrl
        CASE "edititemONE","sellStatNONE","delitemONE" : getAPIParam = "_method=UpdateProductAPIData&citeKey="&getCiteKey(mode)&"&secretKey="&getSecretKey(mode)&"&dataUrl="&dataUrl
        CASE "CheckItemStat" : getAPIParam  = "_method=GetPrdSaleQtyForAPI&citeKey="&getCiteKey(mode)&"&secretKey="&getSecretKey(mode)&"&prdNo="
        CASE "CheckItemInfo" : getAPIParam  = "_method=GetProductInquiryForAPI&citeKey="&getCiteKey(mode)&"&secretKey="&getSecretKey(mode)&"&locNo="
        CASE "editItemStat" : getAPIParam = "_method=UpdateProductAPIStatTpQty&citeKey="&getCiteKey(mode)&"&secretKey="&getSecretKey(mode)&"&dataUrl="&dataUrl  ''������ 2015/05/13
        CASE ELSE : getAPIParam=""
    END SELECT
end function


function getCiteKey(mode)
    SELECT CASE mode
        CASE "cateRcv"  : getCiteKey = "KIQpKWSzGVladyAxxM4vAz3nCetGjAmmAXKkQotL8KQ="
        CASE "regitemONE" : getCiteKey = "Cxyso3Izaa7VNiHAauqT3ocgYfDqdiqpO6Z02j63U4w="
        CASE "edititemONE","sellStatNONE","delitemONE" : getCiteKey = "9CIgE/zSo2ZlDnPaviyqoKmRUPF6ZRea"
        CASE "CheckItemStat" : getCiteKey = "HmMTYbcJDv7aeUsOEUJ5gDCGH7eaEqrg"
        CASE "CheckItemInfo" : getCiteKey = "QhTaVJRjbpXFR0QB//XN7Yo/ek57BpYR"
        CASE "editItemStat" : getCiteKey = "h5Szn1XPevegFsnSYKfGgEOI6E1mQnQqeEjffn5p5Zo="  ''������ 2015/05/13
        CASE ELSE : getCiteKey=""
    END SELECT
end function

function getSecretKey(mode)
    SELECT CASE mode
        CASE "cateRcv"  : getSecretKey = "2FfOmboyJ6EG17kcxUnIcZF1/43iVb42"
        CASE "regitemONE" : getSecretKey = "u6r9q5YmW9nOnAuo6w6kDJF1/43iVb42"
        CASE "edititemONE","sellStatNONE","delitemONE" : getSecretKey = "MaMpPg2WSWUE1NiGGmgTm7Ax63xqcqgJ"
        CASE "CheckItemStat" : getSecretKey = "dzpAObpfn37MkqwHIXXm7aFJchN0b9Yw"
        CASE "CheckItemInfo" : getSecretKey = "LP/bHbOkXLpuU40a1Gl6fMaNAW/9kpfl"
        CASE "editItemStat" : getSecretKey = "6CxkBwV1Bk^CiWEqdQ5cV^XiFcrjBaTn"  ''������ 2015/05/13
        CASE ELSE : getSecretKey=""
    END SELECT
end function


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


''retVal is xmlURI
sub CheckFolderCreate(sFolderPath)
    dim objfile
    set objfile=Server.CreateObject("Scripting.FileSystemObject")

    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF
    set objfile=Nothing
End Sub

function getCurrDateTimeFormat()
    dim nowtimer : nowtimer= timer()
    getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
end function

function checkConfirmMatchIPark(iitemid,iPrdsaleStatTp,iPrdSaleUnitcost,iPrdNm,prdNo, regcatecode) ''������ũ �ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05
    dim sqlstr, iParkSellyn
    dim assignedRow : assignedRow=0
    dim pbuf : pbuf =""
    dim preinterparkprdno

    iPrdsaleStatTp= Trim(iPrdsaleStatTp)
    iPrdNm = Trim(replace(iPrdNm,"[�ٹ�����]",""))
    iPrdNm = replace(replace(replace(replace(iPrdNm,Chr(34),""),"<",""),">",""),"^","")

    if (iPrdsaleStatTp="01") then     ''�Ǹ���
        iParkSellyn = "Y"
    elseif (iPrdsaleStatTp="02") then ''ǰ��
        iParkSellyn = "N"
    elseif (iPrdsaleStatTp="05") then ''�Ͻ�ǰ��
        iParkSellyn = "S"
    elseif (iPrdsaleStatTp="03") then ''�Ǹ�����
        iParkSellyn = "X"
    elseif (iPrdsaleStatTp="10") or (iPrdsaleStatTp="98") then
        iParkSellyn = "X"
    end if

    sqlstr = "select (mayiparkSellyn+','+convert(varchar(10),convert(int,mayiparkPrice)) )as pbuf, isNULL(interparkprdno,'') as interparkprdno"
    sqlstr = sqlstr & " From db_item.dbo.tbl_interpark_reg_Item R" & VbCRLF
    sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
    '''sqlstr = sqlstr & " and interparkprdno='"&prdNo&"'"  ''2013/04/03 �߰�
    rsget.Open sqlstr,dbget,1
    if Not rsget.Eof then
        pbuf = rsget("pbuf")
        preinterparkprdno = rsget("interparkprdno")
    end if
    rsget.close()


    sqlstr = "Update R" & VbCRLF
    sqlstr = sqlstr & " SET mayiparkPrice="&iPrdSaleUnitcost & VbCRLF
    IF (iParkSellyn<>"") then
        sqlstr = sqlstr & " ,mayiparkSellyn='"&iParkSellyn&"'"
    ENd IF
    sqlstr = sqlstr & " ,regitemname='"&html2db(iPrdNm)&"'"
    sqlstr = sqlstr & " ,lastStatCheckDate=getdate()"
    sqlstr = sqlstr & " ,regcatecode='"&regcatecode&"'"& VbCRLF
    sqlstr = sqlstr & " ,interparkprdno=isNULL(R.interparkprdno,'"&prdNo&"')"&VbCRLF
    sqlstr = sqlstr & " From db_item.dbo.tbl_interpark_reg_Item R" & VbCRLF
    sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
    sqlstr = sqlstr & " and isNULL(interparkprdno,'') in ('','"&prdNo&"')"&VbCRLF    ''�ߺ���ϵ�CaSE ���
    sqlstr = sqlstr & " and (isNULL(mayiparkPrice,0)<>"&iPrdSaleUnitcost&"" & VbCRLF
    sqlstr = sqlstr & "     or isNULL(mayiparkSellyn,'')<>'"&iParkSellyn&"'"& VbCRLF
    sqlstr = sqlstr & "     or isNULL(regitemname,'')<>'"&html2db(iPrdNm)&"'"& VbCRLF
    if (regcatecode<>"") then
        sqlstr = sqlstr & "     or isNULL(regcatecode,'')<>'"&regcatecode&"'"& VbCRLF
    end if
    sqlstr = sqlstr & "     or isNULL(interparkprdno,'')<>'"&prdNo&"'"& VbCRLF
    sqlstr = sqlstr & " )" & VbCRLF

    ''rw sqlstr
    dbget.Execute sqlstr,assignedRow

    if (assignedRow<1) then
        if (pbuf="") and (iParkSellyn<>"X") then ''��ǰ�ڵ���°��
            rw "["&iitemid&"] STAT_ERR"&"|"&prdNo&"|"&iParkSellyn&"|"&iPrdsaleStatTp&"|"&iPrdSaleUnitcost&"|"&dispNo
            CALL Fn_AcctFailLogNone(CMALLNAME,iitemid,prdNo,iParkSellyn,iPrdSaleUnitcost,0,iPrdsaleStatTp,"STAT_ERR")
        else
            ''�ٸ��� ������ lastStatCheckDate �� ������Ʈ
            sqlstr = "Update R" & VbCRLF
            sqlstr = sqlstr & " SET lastStatCheckDate=getdate()"
            sqlstr = sqlstr & " From db_item.dbo.tbl_interpark_reg_Item R" & VbCRLF
            sqlstr = sqlstr & " where R.itemid="&iitemid & VbCRLF
            dbget.Execute sqlstr
        end if
    else
        ''�ٸ��� ������ �α�.
        CALL Fn_AcctFailLog(CMALLNAME,iitemid,iPrdsaleStatTp&","&iParkSellyn&","&iPrdSaleUnitcost&"::"&pbuf,"STAT_CHK")
    end if

    if (preinterparkprdno<>prdNo) and (preinterparkprdno<>"") then
        rw "["&iitemid&"] STAT_DUPP"&"|"&prdNo&"(����:"&preinterparkprdno&")|"&iParkSellyn&"|"&iPrdsaleStatTp&"|"&iPrdSaleUnitcost&"|"&dispNo
        CALL Fn_AcctFailLogNone(CMALLNAME,iitemid,prdNo,iParkSellyn,iPrdSaleUnitcost,0,iPrdsaleStatTp&","&preinterparkprdno,"STAT_DUPP")
    end if
end function

function makeIparkXML(mode,itemid,param2,ByRef theLastMainImage)


    Dim i,j,k
    Dim fso,tFile
    Dim opath : opath = "/admin/etc/interparkXML/newAPI/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
    Dim defaultPath : defaultPath = server.mappath(opath) + "\"
    Dim fileName : fileName = mode &"_"& getCurrDateTimeFormat&"_"&itemid&".xml"
    Dim pSize , oiParkitem
    Dim IsEditMode, IsDelMode, IsDelSoldOut, IsDelJaeHyu, IsRegMode
    Dim oneItemID

    ''�ѹ��� ó�� �� ����. �ű� API������ �Ѱ��� �ۿ� �ȵǴ���.
    pSize = 1

	set oiParkitem = new CiParkRegItem
	oiParkitem.FCurrPage = 1
	oiParkitem.FPageSize = pSize

    if (mode="edititemONE") or (mode="sellStatNONE") then
        IsEditMode = true
        oiParkitem.GetIParkOneItemList itemid, (mode="sellStatNONE")  '''���� �����ΰ��  �� �Լ� ���.
    elseif (mode="regitemONE") then
        IsRegMode = true
        oiParkitem.FRectItemIdARR = itemid
	    oiParkitem.GetIParkRegItemList
    elseif (mode="delitemONE")  then
        IsDelMode = true
        ''oiParkitem.FRectItemIdARR = itemid
        oiParkitem.GetIParkOneItemList itemid, (mode="delitemONE")
'    elseif (IsDelSoldOut) then
'        oiParkitem.FCurrPage = 1
'        oiParkitem.FPageSize = 10
'        oiParkitem.GetIParkDelSoldOutItemList
'    elseif (IsDelJaeHyu) then
'    	oiParkitem.FJaeHyuPageGubun = request("jaehyupagegubun")
'        oiParkitem.GetIParkDelJaeHyuItemList
	else
	    oiParkitem.GetIParkRegItemList
    end if

    dim IsTheLastOption, IsOptionExists, optstr, buf, optbuf, keywordsBuf, keywordsStr, NotSoldOutOptionExists
    dim ioptCodeBuf, ioptNameBuf, ioptTypeName, ioptLimitNo, ioptAddPrice
    dim IsAllSoldOutOption

    IF (oiParkitem.FResultCount<1) then
        makeIparkXML = ""
        exit function
    end if

    CALL CheckFolderCreate(defaultPath)

    Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(defaultPath & FileName )
	tFile.WriteLine "<?xml version='1.0' encoding='euc-kr'?>"
	tFile.WriteLine "<result>"
	tFile.WriteLine "<title>Interpark Product API</title>"
	IF (IsRegMode) then
	    tFile.WriteLine "<description>��ǰ ���</description>"
	ELSE
	    tFile.WriteLine "<description>��ǰ ����</description>"
    END IF

	'' New API�� �ִ°͵�.
	'' inOpt :: �Է��� ����ǰ �ɼ�
    '' detailImg ::  ���̹��� - ���̹��� URL, ����/���� ����, JPG�� GIF�� ���� �ִ� 4���� �̹�������, �޸�(,)�� �����Ͽ� ���.
	for i=0 to oiParkitem.FResultCount-1
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
				        ioptTypeName = Replace(Replace(Trim(oiParkitem.FItemList(i).FItemOptionTypeName)," ",""),"����","����")
						ioptCodeBuf = ioptCodeBuf + oiParkitem.FItemList(i).FItemOption + ","
						ioptNameBuf = ioptNameBuf + Replace(Replace(Replace(Replace(Trim(oiParkitem.FItemList(i).FItemOptionName),",",".")," ",""),"<","("),">",")") + ","  ''�ɼǳ��뿡 ���� ������ �ȵ�.//������ �ɼ� �����Ϳ� ������
						ioptAddPrice = ioptAddPrice + CStr(oiParkitem.FItemList(i).Foptaddprice) + ","
						ioptLimitNo = ioptLimitNo + CStr(oiParkitem.FItemList(i).getOptionLimitNo) + ","
						NotSoldOutOptionExists = true
					end if
				else
				    if (Not oiParkitem.FItemList(i).IsOptionSoldOut) then
				        ioptTypeName = Replace(Replace(Trim(oiParkitem.FItemList(i).FItemOptionTypeName)," ",""),"����","����")
					    ioptCodeBuf = ioptCodeBuf + oiParkitem.FItemList(i).FItemOption + ","
						ioptNameBuf = ioptNameBuf + Replace(Replace(Replace(Replace(Trim(oiParkitem.FItemList(i).FItemOptionName),",",".")," ",""),"<","("),">",")") + ","
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
			        ioptTypeName = Replace(Replace(Trim(oiParkitem.FItemList(i).FItemOptionTypeName)," ",""),"����","����")
				    ioptCodeBuf = ioptCodeBuf + oiParkitem.FItemList(i).FItemOption + ","
					ioptNameBuf = ioptNameBuf + Replace(Replace(Replace(Replace(Trim(oiParkitem.FItemList(i).FItemOptionName),",",".")," ",""),"<","("),">",")") + ","
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

        optstr = Replace(optstr,VbTab,"")
  'rw  optstr
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

		    ''Ű���� �ִ� 3�� �޸����� :: Ű���� ������ ������ �ʰ�. �ִ������ 100 byte 'prdKeywd'
		    IF (mode="sellStatNONE") THEN
		        if UBound(keywordsBuf)>0 then keywordsStr = keywordsStr + Trim(keywordsBuf(0)) + ","
		    ELSE
    		    for k=0 to 2
    		        if UBound(keywordsBuf)>k then keywordsStr = keywordsStr + Trim(keywordsBuf(k)) + ","
    		    next
    		ENd IF
    ''Ű���� ���� (ī�װ��� �־��ٰ�)
		    keywordsStr = "�ٹ�����," + keywordsStr
		    keywordsStr = RightCommaDel(keywordsStr)

		    ''keywordsStr=chrbyte(keywordsStr,90,"")
		    if (oiParkitem.FItemList(i).FItemID=486220) or (oiParkitem.FItemList(i).FItemID=486222)  then
		        keywordsStr=""
		    end if

		    if (mode="sellStatNONE") then
		        keywordsStr=""
		    end if

		    buf = buf + "<item>" + delim

			''''buf = buf + "<sidx>" & seqIdx + 1 & "</sidx>" + delim      '''new API ������ ��� ����
			if (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
			    buf = buf + "<prdNo>" & + CStr(oiParkitem.FItemList(i).FInterparkPrdNo) & "</prdNo>" + delim
			end if
			'''buf = buf + "<supplyEntrNo>3000010614</supplyEntrNo>" + delim   ''��ü��ȣ ����(3000010614, �׽�Ʈ, ���� ����)'''new API ������ ��� ����

			'####### 20111025 ����
			'IF (application("Svr_Info")="Dev") THEN
			'    buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim
			'ELSE
			'    buf = buf + "<supplyCtrtSeq>" + CStr(oiParkitem.FItemList(i).GetSupplyCtrtSeq) + "</supplyCtrtSeq>" + delim           ''���ް���Ϸù�ȣ �Ƿ�(2), ��ȭ(3), ����(4)
		    'END IF

		    '''������ 2���� ��.
			if (FALSE) and (oiParkitem.FItemList(i).Fdeliverytype="9") then ''��ü���ǹ��
			    IF (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then  '''���
			        ''if (oiParkitem.FItemList(i).Finterparkregdate>CDate("2011-11-01 20:40:00")) then
			        if (oiParkitem.FItemList(i).FSupplyCtrtSeq=6) then
			            buf = buf + "<supplyCtrtSeq>6</supplyCtrtSeq>" + delim
			        else
			            buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim
			        end if
			    ELSE
			        if (oiParkitem.FItemList(i).FSellcash>=oiParkitem.FItemList(i).FdefaultfreeBeasongLimit) then  ''' ��ü ���� �������̸� 2
			            buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim
			        else
			            buf = buf + "<supplyCtrtSeq>6</supplyCtrtSeq>" + delim
			        end if
			    END IF
			else
			    ''buf = buf + "<supplyCtrtSeq>2</supplyCtrtSeq>" + delim  '''new API ������ ��� ����
			end if



		    if Not (IsDelMode or IsDelSoldOut or IsDelJaeHyu) then              '''IsEditMode or
				buf = buf + "<prdStat>01</prdStat>" + delim                     ''01 ����ǰ
				buf = buf + "<shopNo>0000100000</shopNo>" + delim               ''������ȣ API��ü ����  ''' �ʼ�
				IF (application("Svr_Info")="Dev") THEN
				    buf = buf + "<omDispNo>001830114002</omDispNo>" + delim
				ELSE
				    buf = buf + "<omDispNo>" + Trim(oiParkitem.FItemList(i).Finterparkdispcategory) + "</omDispNo>" + delim  ''������ũ �����ڵ�

				    ''rw oiParkitem.FItemList(i).Finterparkdispcategory
				END IF
		    end IF
			buf = buf + "<prdNm><![CDATA[" + oiParkitem.FItemList(i).getItemNameFormat + "]]></prdNm>" + delim ''��ǰ��
			buf = buf + "<hdelvMafcEntrNm><![CDATA[" + CStr(oiParkitem.FItemList(i).FMakerName) + "]]></hdelvMafcEntrNm>" + delim ''������ü��
			buf = buf + "<prdOriginTp><![CDATA[" + oiParkitem.FItemList(i).GetSourcearea + "]]></prdOriginTp>" + delim       ''������
			buf = buf + "<taxTp>" + oiParkitem.FItemList(i).GetInterParkTaxTp + "</taxTp>" + delim      ''���� 01, �鼼02, ���� 03
			buf = buf + "<ordAgeRstrYn>N</ordAgeRstrYn>" + delim            ''���ο�ǰ
			if (IsAllSoldOutOption) then
			    buf = buf + "<saleStatTp>05</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, �Ͻ�ǰ��05
			    rw "saleStatTp:05"
			elseif (IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
			    buf = buf + "<saleStatTp>03</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, ����04, �Ͻ�ǰ��05
			    rw "saleStatTp:03"
			elseif (mode="sellStatNONE") then
			    buf = buf + "<saleStatTp>02</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, ����04, �Ͻ�ǰ��05
			    rw "saleStatTp:02"
			else
			    buf = buf + "<saleStatTp>" + oiParkitem.FItemList(i).GetInterParkSaleStatTp + "</saleStatTp>" + delim  ''�Ǹ���01, ǰ��02, �Ǹ�����03, �Ͻ�ǰ��05
			    rw "saleStatTp::"&oiParkitem.FItemList(i).GetInterParkSaleStatTp
		    end if

		    if (IsEditMode or IsDelMode or IsDelSoldOut or IsDelJaeHyu) then
		        if (mode="sellStatNONE")  then
		            if (oiParkitem.FItemList(i).FlastErrStr="[000|00]�ش��ǰ�� ���������� �̹� ����Ǿ����ϴ�.����ȸ �� �ٽ� �������ּ���.") _
		                or (oiParkitem.FItemList(i).FlastErrStr="�ش��ǰ�� ���������� �̹� ����Ǿ����ϴ�.����ȸ �� �ٽ� �������ּ���.") then  ''2013/09/30 �߰�
						If CLng(10000-oiParkitem.FItemList(i).Fbuycash/oiParkitem.FItemList(i).Fsellcash*100*100)/100 < 15 Then
                        	buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).Forgprice/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
						Else
							buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).Fmayiparkprice/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
						End If
                        rw "��������:"&oiParkitem.FItemList(i).Fmayiparkprice
                    else
						If CLng(10000-oiParkitem.FItemList(i).Fbuycash/oiParkitem.FItemList(i).Fsellcash*100*100)/100 < 15 Then
                        	buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).Forgprice/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
						Else
							buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).FSellcash/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
						End If
                    end if
		        else
					If CLng(10000-oiParkitem.FItemList(i).Fbuycash/oiParkitem.FItemList(i).Fsellcash*100*100)/100 < 15 Then
	                	buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).Forgprice/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
					Else
						 buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).FSellcash/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
					End If
		        end if
		    else
				If CLng(10000-oiParkitem.FItemList(i).Fbuycash/oiParkitem.FItemList(i).Fsellcash*100*100)/100 < 15 Then
                	buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).Forgprice/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
				Else
					buf = buf + "<saleUnitcost>" + CStr(GetRaiseValue(oiParkitem.FItemList(i).FSellcash/10)*10) + "</saleUnitcost>" + delim  ''�ǸŰ�
				End If
			end if

			buf = buf + "<saleLmtQty>" + CStr(oiParkitem.FItemList(i).GetInterParkLmtQty)+ "</saleLmtQty>" + delim  ''��������
			buf = buf + "<saleStrDts>" + Replace(Left(CStr(now()),10),"-","") + "</saleStrDts>" + delim         ''�ǸŽ�����
			buf = buf + "<saleEndDts>" + oiParkitem.FItemList(i).GetSellEndDateStr + "</saleEndDts>" + delim         ''�Ǹ�������



			'####### 20111025 ����
			'if (oiParkitem.FItemList(i).Fdeliverytype="4") then ''�ٹ����� �����۸�
			'    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
			'else
			'    buf = buf + "<proddelvCostUseYn>N</proddelvCostUseYn>" + delim  ''��ǰ����ۺ񿩺� 30000/2500
			'end if

'				if (oiParkitem.FItemList(i).Fdeliverytype="9") then ''��ü���ǹ�� ��ǰ�� ��ۺ� �ΰ� 2500
'				    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
'				else
				if (oiParkitem.FItemList(i).Fdeliverytype="4") then ''�ٹ����� �����۸�
				    buf = buf + "<proddelvCostUseYn>Y</proddelvCostUseYn>" + delim
				else
				    buf = buf + "<proddelvCostUseYn>N</proddelvCostUseYn>" + delim  ''��ǰ����ۺ񿩺� 30000/2500
				end if
'				end if

			''ȭ����� ���� ��ǰ ��ۺ�
			if (oiParkitem.FItemList(i).IsTruckReturnDlvExists) then
                buf = buf + "<prdrtnCostUseYn>Y</prdrtnCostUseYn>" + delim ''��ǰ ��ǰ�ù�� ��뿩�� - ��ǰ��ǰ�ù����:Y, ��ü��ǰ�ù����:N //2013/02/27
                buf = buf + "<rtndelvCost>"&oiParkitem.FItemList(i).getTruckReturnDlvPrice&"</rtndelvCost>" + delim ''��ǰ ��ǰ�ù��. prdrtnCostUseYn �� 'Y' �� ��� �ʼ���
            end if

			oiParkitem.FItemList(i).FItemContent = Replace(oiParkitem.FItemList(i).FItemContent,"","")'2013-11-05 ������ ���� 000|03 ����..��ǰ ����Ʈ�� �̻��� ���� ����

			buf = buf + "<prdBasisExplanEd><![CDATA[" + Replace(oiParkitem.FItemList(i).getItemPreInfodataHTML,"","") + Replace(Replace(oiParkitem.FItemList(i).FItemContent,"",""),"","")+ Replace(oiParkitem.FItemList(i).getItemInfoImageHTML,"","") + "]]></prdBasisExplanEd>" + delim     ''��ǰ����
			buf = buf + "<zoomImg><![CDATA[" + oiParkitem.FItemList(i).get400Image + "]]></zoomImg>" + delim    ''��ǥ�̹���
			theLastMainImage = oiParkitem.FItemList(i).getBasicImage
			If IsEditMode Then
			    if (oiParkitem.FItemList(i).isImageChanged) then  ''�̹����� ����Ǿ����� Ȯ�� �ʿ�(����)
			    	buf = buf + "<detailImg>"&oiParkitem.FItemList(i).getAddimageInfo&"</detailImg>" + delim		'2014-12-01 ������ �߰�(��ǰ��Ͻÿ��� �Է��ϴ���..)
				    buf = buf + "<imgUpdateYn>Y</imgUpdateYn>" + delim    ''��ǥ�̹����� ���̹��� �����Ϸ��� Y�� �ؾߵ�..����Ʈ N��	'2013-07-23 ������ �߰�
				    rw "imgUpdateYn:Y"
				    rw "zoomImg:" + oiParkitem.FItemList(i).get400Image 
				else
				    rw "imgUpdateYn:N"
			    end if
			End If
			buf = buf + "<prdPrefix><![CDATA[" + oiParkitem.FItemList(i).GetprdPrefixStr + "]]></prdPrefix>" + delim
			''buf = buf + "<prdPostfix></prdPostfix>" + delim
			buf = buf + "<prdKeywd><![CDATA[" + Replace(keywordsStr,"'","") + "]]></prdKeywd>" + delim
			buf = buf + "<brandNm><![CDATA[" + oiParkitem.FItemList(i).Fbrandname + "]]></brandNm>" + delim
			buf = buf + "<entrPoint>" + CStr(oiParkitem.FItemList(i).GetInterParkentrPoint)+ "</entrPoint>" + delim              ''��ü����Ʈ
			buf = buf + "<minOrdQty>1</minOrdQty>" + delim              ''�ּ��ֹ�����

			if (IsAllSoldOutOption) or (IsDelMode) or (IsDelSoldOut) or (IsDelJaeHyu) then

			else

			''2013-10-11 15:38�� ������..�ٹ����� �ɼǼ��� outmallreged�ɼ� ���̺��� ���ؼ� �ɼ� ����
			''2013-11-15 11:43�� ������..�ֹ����ۻ�ǰ�� �� prdOptoin�� ���� ������ ����
			''2013-12-23 10:19�� ������..�ֹ����ۻ�ǰ + ���տɼ� + �߰��ݾױ��� �ִ� Ǯ�ɼǻ��� itemid:935358�����ؼ� 552���� �ּ�ó��
			''2014-06-30 15:23�� ������..�ֹ����ۻ�ǰ�� �� prdOptoin�� ���� ������ ���� oiParkitem.FItemList(i).Fitemoption ���� �ѹ� �� �ɸ�
'				If oiParkitem.FItemList(i).Fitemdiv<>"06" Then 
		    	If oiParkitem.FItemList(i).FoptionCnt = 0 AND oiParkitem.FItemList(i).FregOptCnt > 0 Then
					sqlStr = ""
				    sqlStr = sqlStr & " SELECT itemoption, outmallOptName "&VbCRLF
				    sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption"&VbCRLF
				    sqlStr = sqlStr & " WHERE itemid="&oiParkitem.FItemList(i).FItemID&VbCRLF
				    sqlStr = sqlStr & " and mallid = 'interpark' "
				    rsget.Open sqlStr,dbget,1
				    if not rsget.Eof then
				        optArrRows = rsget.getRows()
				    end if
				    rsget.close
				    For optlp =0 To UBound(optArrRows,2)
				    	optlpName	= optlpName & optArrRows(1,optlp) & ","
				    	optlpCode	= optlpCode & optArrRows(0,optlp) & ","
				    	optlpSu		= optlpSu & "0,"
				    	optlpUsing	= optlpUsing & "N,"
					Next
					optlpName	= RightCommaDel(optlpName)
					optlpCode	= RightCommaDel(optlpCode)
					optlpSu		= RightCommaDel(optlpSu)
					optlpUsing	= RightCommaDel(optlpUsing)

					optlpName	= "�ɼ�<" & optlpName & ">"
					optlpCode	= "�ɼ��ڵ�<" & optlpCode & ">"
					optlpSu		= "����<" & optlpSu & ">"
					optlpUsing	= "��뿩��<" & optlpUsing & ">"
					optlpStr = optlpName & optlpSu & optlpCode & optlpUsing

					If oiParkitem.FItemList(i).Fitemoption = "0000" Then
						buf = buf + "<prdOption><![CDATA[" + optlpStr + "]]></prdOption>" + delim
					Else
						buf = buf + "<prdOption><![CDATA[����<��ǰ>��뿩��<N>]]></prdOption>" + delim     ''�ɼ�.	
					End If
			    else
			    	IF (oiParkitem.FItemList(i).Fitemdiv="06") THEN
			    	    if (optstr<>"") then
			    		buf = buf + "<prdOption><![CDATA[{" + optstr + "}]]></prdOption>" + delim     ''2013-10-18 ������ ����..�ɼǿ� ���ȣ�� ����� �����Ǵ�..
			    	    else
			    	    buf = buf + "<prdOption><![CDATA[" + optstr + "]]></prdOption>" + delim   ''2014/05/30
			            end if
			    		buf = buf + "<optPrirTp><![CDATA[01]]></optPrirTp>" + delim ''�ɼǳ��� ���� ����/01-��ϼ�, 02-�����ټ�. ������ �ɼǸ� �����. 2013-10-18 ������ �߰�
			    	Else
			        	buf = buf + "<prdOption><![CDATA[" + optstr + "]]></prdOption>" + delim     ''�ɼ�.
			    	End If
			    end if
			End If
'		    end if

		    if (FALSE) and (oiParkitem.FItemList(i).Fdeliverytype="9") then
		        if (oiParkitem.FItemList(i).FSellcash>=oiParkitem.FItemList(i).FdefaultfreeBeasongLimit) then
		            buf = buf + "<delvCost>0</delvCost>" + delim
		        else
		            buf = buf + "<delvCost>2500</delvCost>" + delim             '' ���ǹ���ΰ�� 2500
		        end if
			elseif (oiParkitem.FItemList(i).Fdeliverytype="4") then
			    buf = buf + "<delvCost>0</delvCost>" + delim               ''��ǰ�� ��ۺ�
			end if
			buf = buf + "<delvAmtPayTpCom>"&oiParkitem.FItemList(i).delvAmtPayTpCom&"</delvAmtPayTpCom>" + delim   ''��ۺ������� 02����
			buf = buf + "<delvCostApplyTp>02</delvCostApplyTp>" + delim   ''��ۺ������� ����01, ������02


			'####### 20111025 ����
			'if (oiParkitem.FItemList(i).IsFreeBeasong) then
			'    buf = buf + "<freedelvStdCnt>1</freedelvStdCnt>" + delim             ''�����۱��ؼ���
			'end if

			''if (oiParkitem.FItemList(i).Fdeliverytype <> "9") then ''��ü���ǹ��
				if (oiParkitem.FItemList(i).IsFreeBeasong) then
				    buf = buf + "<freedelvStdCnt>1</freedelvStdCnt>" + delim             ''�����۱��ؼ���
				end if
'			    else
'			        '' ��ü ���� ����ΰ�� / ��ǰ�� ��ۺ��� ��� ������ ���� ���� �־�� ��.
'			        if (oiParkitem.FItemList(i).FSellcash>=oiParkitem.FItemList(i).FdefaultfreeBeasongLimit) then
'			            buf = buf + "<freedelvStdCnt>1</freedelvStdCnt>" + delim
'			        else
'			            buf = buf + "<freedelvStdCnt>"&Round((oiParkitem.FItemList(i).FdefaultfreeBeasongLimit/oiParkitem.FItemList(i).FSellcash) + 0.49)&"</freedelvStdCnt>" + delim
'			        end if
'				end if


			buf = buf + "<spcaseEd><![CDATA[" + oiParkitem.FItemList(i).getOrderCommentStr + "]]></spcaseEd>" + delim

			buf = buf + "<pointmUseYn>" + CStr(oiParkitem.FItemList(i).GetpointmUseYn) + "</pointmUseYn>" + delim            ''����Ʈ��
	        buf = buf + "<ippSubmitYn>Y</ippSubmitYn>" + delim            ''���ݺ񱳵�Ͽ���
			buf = buf + "<originPrdNo>" + CStr(oiParkitem.FItemList(i).FItemID) + "</originPrdNo>" + delim     ''��ǰ��ȣ

'			IF (application("Svr_Info")="Dev") THEN
'			    buf = buf + "<shopDispInfo><![CDATA[����Ÿ��<2>������ȣ<0000100000>���ù�ȣ<001410038001001>]]></shopDispInfo>" + delim
'			ELSE
''    				IF Not IsNULL(oiParkitem.FItemList(i).Finterparkstorecategory) and (oiParkitem.FItemList(i).Finterparkstorecategory<>"") then
''    				    IF (Left(oiParkitem.FItemList(i).Finterparkstorecategory,5)<>"00143")  THEN
''    				        buf = buf + "<shopDispInfo><![CDATA[����Ÿ��<2>������ȣ<0000100000>���ù�ȣ<001430026003012>]]></shopDispInfo>" + delim   ''�߰�����
''    				    ELSE
''        				    buf = buf + "<shopDispInfo><![CDATA[����Ÿ��<2>������ȣ<0000100000>���ù�ȣ<" + CHKIIF(oiParkitem.FItemList(i).FSupplyCtrtSeq<>2,"001430026003012",Trim(oiParkitem.FItemList(i).Finterparkstorecategory)) + ">]]></shopDispInfo>" + delim   ''�߰�����
''        				END IF
''    				END IF
'		    END IF

			'''201204�߰� �Ķ����//2013-10-18 ������ ����  or IsEditMode�߰���
			IF (oiParkitem.FItemList(i).Fitemdiv="06") THEN
			    IF (IsRegMode or IsEditMode) then
			        ''if (CStr(oiParkitem.FItemList(i).FItemID)<>"1033148") then
			        buf = buf + "<inOpt>"&oiParkitem.FItemList(i).getInOptTitle&"</inOpt>" + delim  ''������ ���ɵǴ��� Ȯ��// �����ȵ�.
			        ''end if
			    END IF
		    END IF

		    IF (IsRegMode) then ''��Ͻÿ��� �ϴ�.
			    buf = buf + "<detailImg>"&oiParkitem.FItemList(i).getAddimageInfo&"</detailImg>" + delim
			    buf = buf + "<imgUpdateYn>"&"Y"&"</imgUpdateYn>" + delim  '''Y /N
			END IF

			'���� ��ǰǰ����� �ڵ� ���� 2012-11-12 ����
			If (IsRegMode or IsEditMode) Then
			    buf = buf + oiParkitem.FItemList(i).getInterparkItemInfoCdToReg()
			    buf = buf + oiParkitem.FItemList(i).getInterparkItemsafetyReg()
		    End If
			'���� ��ǰǰ����� �ڵ� ���� ��

			buf = buf + "</item>" + delim
			optstr = ""
			NotSoldOutOptionExists = false
		    '''seqIdx = seqIdx + 1
		end if
		if buf<>"" then
		    tFile.WriteLine buf
	    end if
	next


	set oiParkitem = Nothing

	tFile.WriteLine "</result>"
	tFile.Close
	Set tFile = nothing
    Set fso = nothing

    makeIparkXML = "http://webadmin.10x10.co.kr"&opath&FileName
end function

function getiParkItemIdByTenItemID(iitemid)
    dim sqlStr, retVal
    sqlStr = " select isNULL(interparkPrdNo,'') as interparkPrdNo "&VbCRLF
    sqlStr = sqlStr & " from db_item.dbo.tbl_interpark_reg_Item"&VbCRLF
    sqlStr = sqlStr & " where itemid="&iitemid&VbCRLF

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    retVal = rsget("interparkPrdNo")
	end if
	rsget.Close

	if IsNULL(retVal) then retVal=""
	getiParkItemIdByTenItemID = retVal
end function

function iParkGetItemInfoArrNewAPI(mode,locNo,byREF ErrMsg,byREF ErrCode)
    Dim iSuccCnt, iParkURL, iParams, replyXML
    Dim sqlStr
    Dim prdNo,prdNm,saleUnitcost,saleStatTp,externalPrdNo,saleLmtQty,dispNo,givePointYn
    Dim xmlDoc,errorNodes,Nodes,SubNodes

    iParkURL = "http://ipss1.interpark.com/openapi/product/ProductAPIService.do"
    iParams  = getAPIParam(mode) & locNo
    iParams = iParams & "&saleStatTp=01"  ''�Ǹ����λ�ǰ�� (2013/07/09)

    replyXML = SendReqGet(iParkURL, iParams)

    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select

    if (ErrMsg<>"") then
        iParkGetItemInfoArrNewAPI = ""
        exit function
    end if

    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.loadXML replyXML

    Set errorNodes = xmlDoc.getElementsByTagName("error")
    Set Nodes = xmlDoc.getElementsByTagName("item")

    If Not (errorNodes(0) is Nothing ) THEN
        ErrMsg = errorNodes(0).getElementsByTagName("explanation")(0).Text
        ErrCode = errorNodes(0).getElementsByTagName("code")(0).Text
        iParkGetItemInfoArrNewAPI = ""
        exit function
    END IF

    if (ErrMsg<>"") then
        iParkGetItemInfoArrNewAPI = ""
        exit function
    end if

    iSuccCnt = 0

    For each SubNodes in Nodes

        prdNo           = SubNodes.getElementsByTagName("prdNo")(0).Text                  '' ������ũ ��ǰ�ڵ�
        prdNm           = SubNodes.getElementsByTagName("prdNm")(0).Text
        saleUnitcost    = SubNodes.getElementsByTagName("saleUnitcost")(0).Text         ''�ǸŰ�
        saleStatTp      = SubNodes.getElementsByTagName("saleStatTp")(0).Text         ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05
        externalPrdNo   = SubNodes.getElementsByTagName("externalPrdNo")(0).Text         ''10x10 ��ǰ�ڵ�
        saleLmtQty      = SubNodes.getElementsByTagName("saleLmtQty")(0).Text
        dispNo          = SubNodes.getElementsByTagName("dispNo")(0).Text        ''���ù�ȣ
        givePointYn     = SubNodes.getElementsByTagName("givePointYn")(0).Text



        CALL checkConfirmMatchIPark(externalPrdNo,saleStatTp,saleUnitcost,prdNm,prdNo,dispNo)

    Next
end function

function iParkOneItemGetNewAPI(itemid,mode,byREF dataUrl,byREF ErrMsg,byREF ErrCode,byref prdNo,byref PrdSaleUnitcost,byref iPrdNm)
    Dim iSuccCnt, iParkURL, iParams, replyXML
    Dim sqlStr
    Dim prdNm,saleUnitcost,saleStatTp, optStkMgtYn, externalPrdNo, saleLmtQty, salePossRestQty
    Dim isOption
    Dim dispNo
    Dim prdOrOptNo

    prdNo = getiParkItemIdByTenItemID(itemid)

    ''if (itemid="823414") then prdNo="1602093913" ''�ɼ� ���� CASE
    ''if (itemid="823335") then prdNo="1602106923" ''�ɼ� �ִ� CASE
  ''rw "prdNo="&prdNo

    iParkURL = "http://ipss1.interpark.com/openapi/product/ProductAPIService.do"
    iParams  = getAPIParam(mode) & prdNo
  ''rw ""
'rw iParkURL&"?"&iParams
    replyXML = SendReqGet(iParkURL, iParams)

    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select

    if (ErrMsg<>"") then
        iParkOneItemGetNewAPI = ""
        exit function
    end if

    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.loadXML replyXML
'response.write replyXML
'response.end
    Set errorNodes = xmlDoc.getElementsByTagName("error")
    Set Nodes = xmlDoc.getElementsByTagName("item")

    If Not (errorNodes(0) is Nothing ) THEN
        ErrMsg = errorNodes(0).getElementsByTagName("explanation")(0).Text
        ErrCode = errorNodes(0).getElementsByTagName("code")(0).Text
        iParkOneItemGetNewAPI = ""
        exit function
    END IF

    if (ErrMsg<>"") then
        iParkOneItemGetNewAPI = ""
        exit function
    end if

    Dim PrdsaleStatTp

	''2013-10-11 ������ �߰�..outmall_regedoption ���̺� ���������� Ȯ��
	Dim strSQL, regedoptionCnt, MasterPrice

	regedoptionCnt = True
	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as cnt FROM db_item.dbo.tbl_outmall_regedoption WHERE itemid = '"&itemid&"' AND mallid = 'interpark' "
	rsget.Open strSQL, dbget, 1
	If rsget("cnt") > 0 Then
		regedoptionCnt = True
	Else
		regedoptionCnt = False
	End If
	rsget.Close
	''2013-10-11 15:17 ������ �߰�..outmall_regedoption ���̺� ���������� Ȯ�� ��

    '' �ɼǰ��� Ȯ��
    iSuccCnt = 0
    For each SubNodes in Nodes
        isOption = false
        optStkMgtYn =""

        prdOrOptNo = SubNodes.getElementsByTagName("prdNo")(0).Text                  ''�ɼ��ΰ�� ������ũ ��ǰ�ڵ�
        prdNm = SubNodes.getElementsByTagName("prdNm")(0).Text                  ''��ǰ�� �Ǵ� �ɼ�
        saleUnitcost = SubNodes.getElementsByTagName("saleUnitcost")(0).Text    ''�ǸŰ�, �ɼ��ΰ�� �ɼ��߰��ݾ��� ���ѱݾ�
        externalPrdNo = SubNodes.getElementsByTagName("externalPrdNo")(0).Text  ''TEN ��ǰ��ȣ �Ǵ� �ɼǹ�ȣ
        ''dispNo = SubNodes.getElementsByTagName("dispNo")(0).Text  ''���ù�ȣ

        On Error Resume Next
        optStkMgtYn = SubNodes.getElementsByTagName("optStkMgtYn")(0).Text      ''�ɼ������� ��뿩�� - Y:�����, N:������
                                                                                 '' 'Y' �� ���� �ɼ� ��ǰ�� ������ ���, 'N' �� ���� �θ��ǰ�� ������ ���
                                                                                 '' ���ʵ尡 ������ ��ǰ ������ �ɼ�.
        If (ERR) Then
        	isOption = true
        Else
        	isOption = false
        	MasterPrice = saleUnitcost
    	End If
        On Error Goto 0

        if (Not isOption) then
            saleStatTp  = SubNodes.getElementsByTagName("saleStatTp")(0).Text       ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05
            PrdsaleStatTp = saleStatTp
            PrdSaleUnitcost = saleUnitcost
            iPrdNm           = prdNm
        end if

        if ((Not isOption) and (optStkMgtYn="N")) or (isOption) then
            saleLmtQty       = SubNodes.getElementsByTagName("saleLmtQty")(0).Text       ''�Ǹ�(����)����, Ư���� ���� Ư�� ��������
            salePossRestQty  = SubNodes.getElementsByTagName("salePossRestQty")(0).Text  ''��������, Ư���� ���� Ư�� ��������
           	''2013-10-11 15:17 ������ �߰�..outmall_regedoption ���̺� ���������� Ȯ�� �� ����
           	''2013-11-25 09:27 ������ �߰�..Trim(SplitValue(prdNm,"/",1))	Ʈ���κп� html2db �Լ� �߰�..����ǥ ����
			If (regedoptionCnt = False) AND (isOption) Then
				strSQL = ""
				strSQL = strSQL & " SELECT count(*) as Opcnt FROM db_item.dbo.tbl_outmall_regedoption " & VbCrlf
				strSQL = strSQL & " WHERE mallid = 'interpark' and itemid = '"&itemid&"' and itemoption = '"&externalPrdNo&"' " & VbCrlf
				rsget.Open strSQL, dbget, 1
				If rsget("Opcnt") = 0 Then
					strSQL = ""
					strSQL = strSQL & " INSERT INTO db_item.dbo.tbl_outmall_regedoption " & VbCrlf
					strSQL = strSQL & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice, lastupdate)" & VbCrlf
					strSQL = strSQL & " SELECT '"&itemid&"', '"&externalPrdNo&"', 'interpark', '"&prdOrOptNo&"', '"&html2db(Trim(SplitValue(prdNm,"/",1)))&"', '"&Chkiif(saleStatTp="01","Y","N")&"', limityn, '"&salePossRestQty&"', "&saleUnitcost - MasterPrice&", getdate() " & VbCrlf
					strSQL = strSQL & " FROM db_item.dbo.tbl_item " & VbCrlf
					strSQL = strSQL & " WHERE itemid = '"&itemid&"' " & VbCrlf
					dbget.execute strSQL
				End If
				rsget.Close
			End If
			''2013-10-11 ������ �߰�..outmall_regedoption ���̺� ���������� Ȯ�� �� ���� ��
        end if

'        rw "-----------------------------------"
'		 rw "MasterPrice="&MasterPrice
'        rw "prdNo="&prdNo
'        rw "prdOrOptNo="&prdOrOptNo
'        rw "prdNm="&prdNm
'        rw "externalPrdNo="&externalPrdNo
'        rw "saleUnitcost="&saleUnitcost
'        rw "saleStatTp="&saleStatTp
'        rw "optStkMgtYn="&optStkMgtYn
'        rw "saleLmtQty="&saleLmtQty
'        rw "salePossRestQty="&salePossRestQty
        iSuccCnt = iSuccCnt + 1
    Next
    iParkOneItemGetNewAPI = PrdsaleStatTp

end function

''xml ���� ����
function DelAPITMPFile(iFileURI)
    dim iFullPath
    iFullPath = server.mappath(replace(iFileURI,"http://webadmin.10x10.co.kr",""))

    dim FSO, iFile
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set iFile = FSO.GetFile(iFullPath)
    if (iFile<>"") then iFile.Delete
    Set iFile = Nothing
    Set FSO = Nothing

end function

function iParkOneItemProcNewAPI(itemid,mode,byREF dataUrl,byREF ErrMsg,byREF ErrCode,byref originPrdNo,byref prdNo,byREF prdNm)
    Dim iSuccCnt, iParkURL, iParams, replyXML
    Dim sqlStr
    Dim theLastMainImage
	Dim marginCHK

    dataUrl = makeIparkXML(mode,itemid,"", theLastMainImage)
    if (dataUrl="") then
        ErrMsg = "["&itemid&"]"&"���/������ ��ǰ�� �����ϴ�. ī�װ� ����/���޻� ��ϰ��� ���� Ȯ�ο��."
        iParkOneItemProcNewAPI = 0
        exit function
    end if
''response.end

    iParkURL = "http://ipss1.interpark.com/openapi/product/ProductAPIService.do"
    iParams  = getAPIParam(mode)
    replyXML = SendReqGet(iParkURL, iParams)
    ''''replyXML = getReplyXMLTEST_regItem
    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select

CALL DelAPITMPFile(dataUrl)
    if (ErrMsg<>"") then
        iParkOneItemProcNewAPI = 0

        exit function
    end if

    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.loadXML replyXML

    Set errorNodes = xmlDoc.getElementsByTagName("error")
    Set Nodes = xmlDoc.getElementsByTagName("success")

    If Not (errorNodes(0) is Nothing ) THEN
        ErrMsg = errorNodes(0).getElementsByTagName("explanation")(0).Text
        ErrCode = errorNodes(0).getElementsByTagName("code")(0).Text
        iParkOneItemProcNewAPI = 0
        exit function
    END IF

    if (ErrMsg<>"") then
        iParkOneItemProcNewAPI = 0
        exit function
    end if

    '' �Ѱ����� �Ǵ��� ������ �������� ���� Ȯ��. ==>��� ���� ���� �Ѱ����� ������..
    iSuccCnt = 0
    For each SubNodes in Nodes
        prdNo = SubNodes.getElementsByTagName("prdNo")(0).Text
        prdNm = SubNodes.getElementsByTagName("prdNm")(0).Text
        originPrdNo = SubNodes.getElementsByTagName("originPrdNo")(0).Text

        '' ������ũ ����/ �ǸŻ��µ� ����
        IF (mode="regitemONE") then

            sqlStr = "update R" & VbCrlf
            sqlStr = sqlStr & " set interparkregdate=getdate()" & VbCrlf
            sqlStr = sqlStr & " ,interParkPrdNo='" & prdNo & "'" & VbCrlf
            sqlStr = sqlStr & " ,interparklastupdate=getdate()"
            sqlStr = sqlStr & " ,mayiParkPrice=i.sellcash" & VbCrlf
            sqlStr = sqlStr & " ,mayiParkSellYn=i.sellyn" & VbCrlf
            sqlStr = sqlStr & " ,accFailCNT=0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
            if (theLastMainImage<>"") then ''2013/09/01 �߰�
                sqlStr = sqlStr & " ,regimageName='"&theLastMainImage&"'"& VbCrlf
            end if
            sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item R" & VbCrlf
            sqlStr = sqlStr & "     Join  db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid=" & originPrdNo

            dbget.execute sqlStr

            sqlStr = " update R"
            sqlStr = sqlStr & " set interparkSupplyCtrtSeq=2"                   '''������ 2���� ���..
            sqlStr = sqlStr & " , interparkStoreCategory=D.interparkStoreCategory"
            sqlStr = sqlStr & " , Pinterparkdispcategory=D.interparkdispcategory"
            sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item R"
            sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item i"
            sqlStr = sqlStr & " 	on R.itemid=i.itemid"
            sqlStr = sqlStr & " 	Join [db_user].[dbo].tbl_user_c c"
            sqlStr = sqlStr & " 	on i.makerid=c.userid"
            sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_interpark_dspcategory_mapping D"
            sqlStr = sqlStr & " 	on D.tencdl=i.cate_large"
            sqlStr = sqlStr & " 	and D.tencdm=i.cate_mid"
            sqlStr = sqlStr & " 	and D.tencdn=i.cate_small"
            sqlStr = sqlStr & " where D.SupplyCtrtSeq is Not NULL"
            sqlStr = sqlStr & " and i.itemid="& originPrdNo & VbCrlf
            sqlStr = sqlStr & " and R.interParkPrdNo is Not NULL"

            dbget.execute sqlStr

            iSuccCnt = iSuccCnt + 1
        ELSEIF (mode="edititemONE") then
             ''�α��Է�(2011-01-18)�߰�
            sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
            sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode)" & VbCrlf
            sqlStr = sqlStr & " select R.itemid, R.interparkprdno, i.sellcash,i.buycash,i.sellyn,'' as ErrMsg, '' as errCode" & VbCrlf
            sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
            sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid=" & originPrdNo & VbCrlf
            'rw sqlStr
            dbget.execute sqlStr


			sqlStr = ""
			sqlStr = sqlStr & " SELECT top 1 buycash, sellcash FROM db_item.dbo.tbl_item WHERE itemid = " & originPrdNo
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) then
				If CLng(10000-rsget("buycash")/rsget("sellcash")*100*100)/100 < 15 Then
                	marginCHK = False
				Else
					marginCHK = True
				End If
			End If
			rsget.Close

            '' ������ũ ����/ �ǸŻ��µ� ����
            sqlStr = "update R" & VbCrlf
            sqlStr = sqlStr & " set interparklastupdate=getdate()" & VbCrlf
            sqlStr = sqlStr & " ,interParkPrdNo='" & prdNo & "'" & VbCrlf
			If marginCHK = False Then
            	sqlStr = sqlStr & " ,mayiParkPrice=i.orgprice" & VbCrlf
			Else
				sqlStr = sqlStr & " ,mayiParkPrice=i.sellcash" & VbCrlf
			End If
			sqlStr = sqlStr & " ,mayiParkSellYn=i.sellyn" & VbCrlf
            sqlStr = sqlStr & " ,accFailCNT=0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
            if (theLastMainImage<>"") then ''2013/09/01 �߰�
                sqlStr = sqlStr & " ,regimageName='"&theLastMainImage&"'"& VbCrlf
            end if
            sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item R" & VbCrlf
            sqlStr = sqlStr & "     Join  db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid=" & originPrdNo
            'rw sqlStr
            dbget.execute sqlStr

            '''ī�װ� �������� ���� ����. :: ī�װ��� �ٲ� ������� �ʰ�.. // ������ �ٲ�Ⱦȵ�.
            sqlStr = " update R"
            sqlStr = sqlStr & " set interparkStoreCategory=D.interparkStoreCategory"
            sqlStr = sqlStr & " , Pinterparkdispcategory=D.interparkdispcategory"
            sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item R"
            sqlStr = sqlStr & " 	Join [db_item].[dbo].tbl_item i"
            sqlStr = sqlStr & " 	on R.itemid=i.itemid"
            sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_interpark_dspcategory_mapping D"
            sqlStr = sqlStr & " 	on D.tencdl=i.cate_large"
            sqlStr = sqlStr & " 	and D.tencdm=i.cate_mid"
            sqlStr = sqlStr & " 	and D.tencdn=i.cate_small"
            sqlStr = sqlStr & " where D.SupplyCtrtSeq is Not NULL"
            sqlStr = sqlStr & " and i.itemid="& originPrdNo & VbCrlf
            sqlStr = sqlStr & " and R.interParkPrdNo is Not NULL"
            'rw sqlStr
            dbget.execute sqlStr

            iSuccCnt = iSuccCnt + 1
        ELSEIF (mode="delitemONE") or (mode="sellStatNONE") then

             ''�α��Է�(2011-01-18)�߰�
            sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
            sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode)" & VbCrlf
            sqlStr = sqlStr & " select R.itemid, R.interparkprdno, i.sellcash,i.buycash,i.sellyn,'' as ErrMsg, '' as errCode" & VbCrlf
            sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
            sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid=" & originPrdNo & VbCrlf
            'rw sqlStr
            dbget.execute sqlStr


			sqlStr = ""
			sqlStr = sqlStr & " SELECT top 1 buycash, sellcash FROM db_item.dbo.tbl_item WHERE itemid = " & originPrdNo
			rsget.Open sqlStr,dbget,1
			If Not(rsget.EOF or rsget.BOF) then
				If CLng(10000-rsget("buycash")/rsget("sellcash")*100*100)/100 < 15 Then
                	marginCHK = False
				Else
					marginCHK = True
				End If
			End If
			rsget.Close


            '' ������ũ ����/ �ǸŻ��� ����(N)
            sqlStr = "update R" & VbCrlf
            sqlStr = sqlStr & " set interparklastupdate=getdate()" & VbCrlf
            sqlStr = sqlStr & " ,interParkPrdNo='" & prdNo & "'" & VbCrlf
			If marginCHK = False Then
            	sqlStr = sqlStr & " ,mayiParkPrice=i.orgprice" & VbCrlf
			Else
				sqlStr = sqlStr & " ,mayiParkPrice=i.sellcash" & VbCrlf
			End If

            IF (mode="delitemONE") THEN
                sqlStr = sqlStr & " ,mayiParkSellYn='X'" & VbCrlf
            ELSE
                sqlStr = sqlStr & " ,mayiParkSellYn='N'" & VbCrlf
            END IF
            sqlStr = sqlStr & " ,accFailCNT=0" & VbCrlf                 ''����ȸ�� �ʱ�ȭ
            sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item R" & VbCrlf
            sqlStr = sqlStr & "     Join  db_item.dbo.tbl_item i" & VbCrlf
            sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
            sqlStr = sqlStr & " where R.itemid=" & originPrdNo
            'rw sqlStr
            dbget.execute sqlStr


            iSuccCnt = iSuccCnt + 1
        ELSE
            iSuccCnt = iSuccCnt + 1
        END IF
    Next

    Set Nodes = Nothing
    Set xmlDoc = Nothing

    iParkOneItemProcNewAPI = iSuccCnt
end function

function getReplyXMLTEST()
    dim replyXML
    replyXML = "<?xml version='1.0' encoding='euckr' ?>"
    replyXML = replyXML&"<result>"
    replyXML = replyXML&"<title>Interpark Product API</title> "
    replyXML = replyXML&"<description>�⺻�������� ��ȸ</description>"
    replyXML = replyXML&"<item>"
    replyXML = replyXML&"<idx>1</idx>"
    replyXML = replyXML&"<shopNo>0000100000</shopNo>"
    replyXML = replyXML&"<dispNo>001110101001001</dispNo>"
    replyXML = replyXML&"<dispNm>"
    replyXML = replyXML&"<![CDATA[ ��ǻ��/��Ʈ��/������>��Ʈ��/UMPC>LG XNOTE>30Cm(12��ġ)����]]> "
    replyXML = replyXML&"</dispNm>"
    replyXML = replyXML&"<dispYn>Y</dispYn> "
    replyXML = replyXML&"<regDts>20080128165752</regDts> "
    replyXML = replyXML&"<modDts>20100120133200</modDts> "
    replyXML = replyXML&"</item>"
    replyXML = replyXML&"<item>"
    replyXML = replyXML&"<idx>2</idx> "
    replyXML = replyXML&"<shopNo>0000100000</shopNo> "
    replyXML = replyXML&"<dispNo>001110101001002</dispNo> "
    replyXML = replyXML&"<dispNm>"
    replyXML = replyXML&"<![CDATA[ ��ǻ��/��Ʈ��/������>��Ʈ��/UMPC>LG XNOTE>33Cm(13��ġ)]]> "
    replyXML = replyXML&"</dispNm>"
    replyXML = replyXML&"<dispYn>Y</dispYn> "
    replyXML = replyXML&"<regDts>20080128165752</regDts> "
    replyXML = replyXML&"<modDts>20100120133200</modDts> "
    replyXML = replyXML&"</item>"
    replyXML = replyXML&"</result>"

    getReplyXMLTEST = replyXML
end function

function getReplyXMLTEST_regItem()
    dim replyXML
    replyXML = "<?xml version='1.0' encoding='euc-kr' ?>"
    replyXML = replyXML&"<result>"
    replyXML = replyXML&"<title>Interpark Product API</title>"
    replyXML = replyXML&"<description>API ȣ�� �Ϸ�</description>"
    replyXML = replyXML&"<success>"
    replyXML = replyXML&"<prdNo>71790033</prdNo>"
    replyXML = replyXML&"<prdNm>������Ī������������ P361</prdNm>"
    replyXML = replyXML&"<originPrdNo>2550440</originPrdNo>"
    replyXML = replyXML&"</success>"
    replyXML = replyXML&"</result>"

    getReplyXMLTEST_regItem = replyXML
end function

IF mode="cateRcv" then
    iParkURL = "http://ipss1.interpark.com/openapi/product/ProductAPIService.do"
    iParams  = getAPIParam(mode)
    if (param1<>"") then
        iParams = iParams & "&strDt=" & param1 ''[�Ⱓ����] YYYYMMDD
    end if

    if (param2<>"") then
        iParams = iParams & "&endDt=" & param2 ''[�Ⱓ����] YYYYMMDD
    end if

    if (param3<>"") then
        iParams = iParams & "&dispYn=" & param3 ''[���ÿ���] YN
    end if

    replyXML = SendReqGet(iParkURL, iParams)
    ''replyXML = getReplyXMLTEST
'rw replyXML
'IF InStr(replyXML,"<TITLE>�� Interpark Partner Support System - �ý��� ����</TITLE>")>0 then
'       ErrMsg =  "�� Interpark Partner Support System - �ý��� ����"
'end if

    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select


    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
    xmlDoc.loadXML replyXML

    Set errorNodes = xmlDoc.getElementsByTagName("error")
    Set Nodes = xmlDoc.getElementsByTagName("item")

    If Not (errorNodes(0) is Nothing ) THEN
        ErrMsg = errorNodes(0).getElementsByTagName("explanation")(0).Text
    END IF

    For each SubNodes in Nodes
        dispNo = SubNodes.getElementsByTagName("dispNo")(0).Text
        dispNm = SubNodes.getElementsByTagName("dispNm")(0).Text
        dispYn = SubNodes.getElementsByTagName("dispYn")(0).Text
        regDts = SubNodes.getElementsByTagName("regDts")(0).Text
        modDts = SubNodes.getElementsByTagName("modDts")(0).Text

        sqlStr = "update db_temp.dbo.tbl_interpark_Tmp_DispCategory"
        sqlStr = sqlStr & " set DispCateName=convert(varchar(255),'"&dispNm&"')"
        sqlStr = sqlStr & " ,dispYn='"&dispYn&"'"
        sqlStr = sqlStr & " ,iParkregDts='"&regDts&"'"
        sqlStr = sqlStr & " ,iParkmodDts='"&modDts&"'"
        sqlStr = sqlStr & " where DispcateCode='"&dispNo&"'"

        dbget.Execute sqlStr, AssignedRow

        upCNT = upCNT + AssignedRow
        if (AssignedRow<1) and (dispYn<>"N") then  ''������ΰŸ� �Է�
            sqlStr = "Insert Into db_temp.dbo.tbl_interpark_Tmp_DispCategory"
            sqlStr = sqlStr & " (DispcateCode,DispCateName,dispYn,lastRegdate,iParkregDts,iParkmodDts)"
            sqlStr = sqlStr & " values('"&dispNo&"'"
            sqlStr = sqlStr & " ,convert(varchar(255),'"&dispNm&"')"
            sqlStr = sqlStr & " ,'"&dispYn&"'"
            sqlStr = sqlStr & " ,getdate()"
            sqlStr = sqlStr & " ,'"&regDts&"'"
            sqlStr = sqlStr & " ,'"&modDts&"'"
            sqlStr = sqlStr & " )"
            dbget.Execute sqlStr, AssignedRow

            regCNT = regCNT + AssignedRow
        end if
    Next

    Set Nodes = Nothing
    Set xmlDoc = Nothing

    rw "ErrMsg : "&ErrMsg
    rw "������Ʈ : "& upCNT
    rw "�űԵ�� : "& regCNT

ELSEIF mode="regitemONE" then
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""
    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��Ͻ���</font> - ("&ErrCode&")"&ErrMsg&"<br>"


            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] ��ϼ��� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
    rw "<br><a href='javascript:history.back();'><font color=blue>BACK</font></a>"

ELSEIF mode="edititemONE" then
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""
    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                 call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") _
                    or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.")   then
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSE
                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

			'2014-07-24 15:35 ������ �߰�. ��ǰ ������ mayiparksellyn�� sellyn�� �������� ġȯ�ϱ⿡ ��ǰ ���� �� ���� �ǸŻ���Ȯ�� �ϱ�.
			retVal = iParkOneItemGetNewAPI(itemid,"CheckItemStat",dataUrl,ErrMsg,ErrCode,prdNo,PrdSaleUnitcost,iPrdNm) ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05
			if (retVal<>"") then
			    CALL checkConfirmMatchIPark(itemid,retVal,PrdSaleUnitcost,iPrdNm,prdNo,"")
			end if
        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
    rw "<br><a href='javascript:history.back();'><font color=blue>BACK</font></a>"
elseif (mode="CheckItemStatBatch") then  ''�ǸŻ���Ȯ�ι�ġ
    sqlStr = " select top 10 r.itemid"
    sqlStr = sqlStr & " from db_item.dbo.tbl_interpark_reg_Item r"
    sqlStr = sqlStr & " where interparkPrdno is Not NULL"
    sqlStr = sqlStr & " order by r.lastStatCheckDate, (CASE WHEN r.mayiParkSellyn='X' THEN '0' ELSE r.mayiParkSellyn END), r.interparkLastUpdate , r.itemid desc"

'    sqlStr = " select top 10 r.itemid"
'    sqlStr = sqlStr & " from db_item.dbo.tbl_interpark_reg_Item r"
'    sqlStr = sqlStr & " where interparkPrdno is Not NULL"
'    sqlStr = sqlStr & " order by r.lastStatCheckDate, r.mayiParkSellyn, r.interparkLastUpdate , r.itemid desc"

    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        ArrRows = rsget.getRows()
    end if
    rsget.close

    mode="CheckItemStat"

    if isArray(ArrRows) then
        For i =0 To UBound(ArrRows,2)
            itemid = CStr(ArrRows(0,i))
            IF (itemid<>"") then
                totCNT = totCNT + 1
                ErrMsg = ""
                ErrCode = ""

                retVal = iParkOneItemGetNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,prdNo,PrdSaleUnitcost,iPrdNm) ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05

                if (retVal<>"") then
                    ''if (iLotteItemStat="10") then
                        rw "["&itemid&"] :"&retVal&":"& PrdSaleUnitcost& ":"&iPrdNm&":"& ErrMsg
                    ''end if

                    CALL checkConfirmMatchIPark(itemid,retVal,PrdSaleUnitcost,iPrdNm,prdNo,"")
                else
                    rw "["&itemid&"] : ["&prdNo&"]:"&retVal&":"& PrdSaleUnitcost& ":"&iPrdNm&":"& ErrMsg

                    if (ErrMsg="�������� �ʴ� ��ǰ��ȣ") then
                        sqlStr = " update db_item.dbo.tbl_interpark_reg_Item"
                        sqlStr = sqlStr & " set lastStatCheckDate=getdate()"
                        sqlStr = sqlStr & " ,mayiParkSellyn='X'"
                        sqlStr = sqlStr & " where itemid="&itemid
                        dbget.Execute sqlStr

                        rw "X flag ó��"
                    end if

                end if

                ''succCNT = succCNT + retCNT
                'rw "prdNo="&prdNo
                'rw "retVal="&retVal
                'rw "PrdSaleUnitcost="&PrdSaleUnitcost

                if (ErrMsg<>"") or (ErrCode<>"") then
                    rw "ErrMsg="&ErrMsg
                    rw "ErrCode="&ErrCode
                end if
            END IF
        Next
    end if
elseif (mode="CheckItemStat") then  ''�ǸŻ���Ȯ��.(��ǰ ����)
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    for i=LBound(cksel) to UBound(cksel)

        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retVal = iParkOneItemGetNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,prdNo,PrdSaleUnitcost,iPrdNm) ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05
            if (retVal<>"") then
                ''if (iLotteItemStat="10") then
                    rw "["&itemid&"] :"&retVal&":"& PrdSaleUnitcost& ":"&iPrdNm&":"& ErrMsg
                ''end if

                CALL checkConfirmMatchIPark(itemid,retVal,PrdSaleUnitcost,iPrdNm,prdNo,"")
            else
                rw "["&itemid&"] :"&retVal&":"& PrdSaleUnitcost& ":"&iPrdNm&":"& ErrMsg
            end if

            ''succCNT = succCNT + retCNT
            rw "prdNo="&prdNo
            rw "retVal="&retVal
            rw "PrdSaleUnitcost="&PrdSaleUnitcost

            rw "ErrMsg="&ErrMsg
            rw "ErrCode="&ErrCode
        END IF
    next
ELSEIF (mode="CheckItemInfo") then ''��ǰ ���� ��ȸ API
    rw "<a href='/admin/etc/iParkAPI_Process.asp?mode=CheckItemInfo&locNo="&locNo+1&"'>next"&locNo+1&"</a>"
    Call iParkGetItemInfoArrNewAPI(mode,locNo,ErrMsg,ErrCode) ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05
    rw "["&ErrCode&"]"&ErrMsg
ELSEIF (mode="chkNdelitem") then ''üũ �� ����
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")
    succCNT = 0
    mode="CheckItemStat"
    for i=LBound(cksel) to UBound(cksel)

        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retVal = iParkOneItemGetNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,prdNo,PrdSaleUnitcost,iPrdNm) ''�ǸŻ��� - �Ǹ���:01, ǰ��:02, �Ǹ�����:03, �Ͻ�ǰ��:05

            if (retVal<>"") then
                if (retVal="03") then
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow

					sqlStr = ""
	                sqlStr = sqlStr & " delete from db_item.dbo.tbl_Outmall_regedoption "
	                sqlStr = sqlStr & " where itemid=" & itemid
	                sqlStr = sqlStr & " and mallid = 'interpark' "
	                dbget.Execute sqlStr
                    succCNT = succCNT + 1
                else
                    CALL checkConfirmMatchIPark(itemid,retVal,PrdSaleUnitcost,iPrdNm,prdNo,"")
                end if

            else
                rw "["&itemid&"] :"&retVal&":"& PrdSaleUnitcost& ":"&iPrdNm&":"& ErrMsg
            end if

            ''succCNT = succCNT + retCNT
'            rw "prdNo="&prdNo
'            rw "retVal="&retVal
'            rw "PrdSaleUnitcost="&PrdSaleUnitcost
'
'            rw "ErrMsg="&ErrMsg
'            rw "ErrCode="&ErrCode
        END IF
    next

    rw "<script>alert('"&succCNT&"�� ������')</script>"
ELSEIF (mode="sellStatNONE") or (mode="delitemONE") then
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""
    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                 call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") _
                    or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSEIF  (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")  then
                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set mayiParkSellYn='X'" & VbCrlf
                    sqlStr = sqlStr & " ,interparklastupdate=getdate()" & VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-X flag<br>"
                ELSE

                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set accfailCnt=accfailCnt+1" & VbCrlf
                    sqlStr = sqlStr & " ,lastErrStr=convert(varchar(100),'"&Trim(html2db(ErrMsg))&"')"& VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr

                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
    rw "<br><a href='javascript:history.back();'><font color=blue>BACK</font></a>"
ELSEIF mode="edititemPrice" then
    cksel =""
    set oInterParkitem = new CExtSiteItem

    oInterParkitem.FPageSize       = 20 '20
    oInterParkitem.FCurrPage       = 1
    oInterParkitem.FRectExtNotReg  = ""                 ''�������
    oInterParkitem.FRectExpensive10x10 = "on"           ''���ݺ�Ѱ�
    oInterParkitem.FRectInteryes10x10no = ""          ''ǰ�����
    oInterParkitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
    ''oInterParkitem.FRectNotInc_NotEditItemid = "on"
    oInterParkitem.GetInterParkRegedItemList

    for i=0 to oInterParkitem.FResultCount - 1
        If (InStr(cksel,CStr(oInterParkitem.FItemList(i).FItemID)&",")<1) then
            cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
        end if
    next

    rw cksel
    SET oInterParkitem=Nothing

    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""

    mode="edititemONE"

    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSE
                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
ELSEIF mode="edititemPrice2" then					''2013-10-01 ������ 15�۹̸� ��ǰ �����ϱ�
    cksel =""
    set oInterParkitem = new CExtSiteItem

    oInterParkitem.FPageSize       = 20 '20
    oInterParkitem.FCurrPage       = 1
    oInterParkitem.FRectExtNotReg  = "F"                 ''��ϿϷ�
    oInterParkitem.FRectMinusMigin15 = "N"          	''15�۹̸����� ��ǰ����
    oInterParkitem.FRectInteryes10x10no = ""          ''ǰ�����
    oInterParkitem.FRectSellYn = "Y"				''�Ǹ����� ��ǰ��
    oInterParkitem.FRectOrdType = "MG"        		  ''������ũ ��Ʈ������Ʈ ����
    oInterParkitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
    oInterParkitem.GetInterParkRegedItemList

    for i=0 to oInterParkitem.FResultCount - 1
        If (InStr(cksel,CStr(oInterParkitem.FItemList(i).FItemID)&",")<1) then
            cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
        end if
    next

    rw cksel
    SET oInterParkitem=Nothing

    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""

    mode="edititemONE"

    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSE
                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
ELSEIF mode="edititemAuto" then
    cksel =""

    ''oInterParkitem.FRectMatchCate  = MatchCate
    ''oInterParkitem.FRectMinusMigin = showminusmagin
    ''oInterParkitem.FRectMinusMigin15 = showminusmagin15
    ''oInterParkitem.FRectIsSoldOut  = onlysoldout
    ''oInterParkitem.FRectOnreginotmapping = onreginotmapping


    set oInterParkitem = new CExtSiteItem
    
    if (param1="1") then
        oInterParkitem.FPageSize       = 10 ''10
        oInterParkitem.FCurrPage       = 1
        oInterParkitem.FRectExtNotReg  = ""                 ''�������
        oInterParkitem.FRectExpensive10x10 = ""             ''���ݺ�Ѱ�
        oInterParkitem.FRectInteryes10x10no = "on"          ''ǰ�����
        oInterParkitem.FRectFailCntOverExcept="5"       '' 3ȸ �̻� ���г��� ����.
        ''oInterParkitem.FRectNotInc_NotEditItemid = "on"
        oInterParkitem.GetInterParkRegedItemList
    
        for i=0 to oInterParkitem.FResultCount - 1
            cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
        next
    end if
    
    if (param1="2") then
        oInterParkitem.FPageSize       = 10 '20
        oInterParkitem.FCurrPage       = 1
        oInterParkitem.FRectExtNotReg  = ""                 ''�������
        oInterParkitem.FRectExpensive10x10 = "on"           ''���ݺ�Ѱ�
        oInterParkitem.FRectInteryes10x10no = ""          ''ǰ�����
        oInterParkitem.FRectFailCntOverExcept="5"       '' 3ȸ �̻� ���г��� ����.
        ''oInterParkitem.FRectNotInc_NotEditItemid = "on"
        oInterParkitem.GetInterParkRegedItemList
    
        for i=0 to oInterParkitem.FResultCount - 1
            If (InStr(cksel,CStr(oInterParkitem.FItemList(i).FItemID)&",")<1) then
                cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
            end if
        next
    end if
    
    if (param1="3") then
        oInterParkitem.FPageSize       = 10 '10
        oInterParkitem.FCurrPage       = 1
        oInterParkitem.FRectExtNotReg  = "R"                 ''�������
        oInterParkitem.FRectExpensive10x10 = ""             ''���ݺ�Ѱ�
        oInterParkitem.FRectInteryes10x10no = ""            ''ǰ�����
        oInterParkitem.FRectFailCntOverExcept="5"       '' 3ȸ �̻� ���г��� ����.
        oInterParkitem.FRectNotInc_NotEditItemid = "on"
        oInterParkitem.FRectLimitYn="Y"                 ''����
        oInterParkitem.GetInterParkRegedItemList
    
        for i=0 to oInterParkitem.FResultCount - 1
            If (InStr(cksel,CStr(oInterParkitem.FItemList(i).FItemID)&",")<1) then
                cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
            end if
        next
    end if


    if (UBound(split(cksel,","))<10) then
        oInterParkitem.FPageSize       = 10 '10
        oInterParkitem.FCurrPage       = 1
        oInterParkitem.FRectExtNotReg  = "R"                 ''�������
        oInterParkitem.FRectExpensive10x10 = ""             ''���ݺ�Ѱ�
        oInterParkitem.FRectInteryes10x10no = ""            ''ǰ�����
        oInterParkitem.FRectFailCntOverExcept="5"       '' 3ȸ �̻� ���г��� ����.
        oInterParkitem.FRectNotInc_NotEditItemid = "on"
        oInterParkitem.FRectLimitYn=""
        oInterParkitem.GetInterParkRegedItemList

        for i=0 to oInterParkitem.FResultCount - 1
            If (InStr(cksel,CStr(oInterParkitem.FItemList(i).FItemID)&",")<1) then
                cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
            end if
        next
    end if

rw cksel
    SET oInterParkitem=Nothing

    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""

    mode="edititemONE"

    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSE
                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg

ELSEIF mode="regitemAuto" then
    cksel =""

    set oInterParkitem = new CExtSiteItem
    oInterParkitem.FPageSize       = 10 ''20  '' too slow (������ ���Ⱦ��� : 2013/09/01 ����)
    oInterParkitem.FCurrPage       = 1
    oInterParkitem.FRectExtNotReg  = "M"                ''��Ͽ���
    oInterParkitem.FRectMatchCate  = "Y"
    oInterParkitem.FRectAvailReg   = "on"
    oInterParkitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
    oInterParkitem.GetInterParkRegedItemList

    for i=0 to oInterParkitem.FResultCount - 1
        cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
    next

    SET oInterParkitem=Nothing

    IF RIGHT(cksel,1)="" THEN itemidARR= itemidARR&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""

    mode="regitemONE"

    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��Ͻ���</font> - ("&ErrCode&")"&ErrMsg&"<br>"

                ''���н� �α��Է�
                sqlStr = " insert into db_log.dbo.tbl_interparkEdit_log" & VbCrlf
                sqlStr = sqlStr & " (itemid, interparkprdno,sellcash,buycash,sellyn, ErrMsg, errCode)" & VbCrlf
                sqlStr = sqlStr & " select R.itemid, IsNULL(R.interparkprdno,''), i.sellcash,i.buycash,i.sellyn,convert(varchar(100),'"&html2db(ErrMsg)&"'),'"&ErrCode&"' " & VbCrlf
                sqlStr = sqlStr & "  from db_item.dbo.tbl_interpark_reg_item R" & VbCrlf
                sqlStr = sqlStr & " 	Join db_item.dbo.tbl_item i" & VbCrlf
                sqlStr = sqlStr & " 	on R.itemid=i.itemid" & VbCrlf
                sqlStr = sqlStr & " where R.itemid=" & itemid & VbCrlf
                'rw sqlStr
                dbget.execute sqlStr

                IF (ErrMsg="��ǰ��,ī�װ�,�ǸŰ��� ������ ��ǰ ��� �Ұ�") or (ErrMsg="������ �����Ͱ� ���� 'prdOriginTp'") then
                    ''sqlStr = "delete from db_item.dbo.tbl_interpark_reg_item where itemid="&itemid&" and interparkprdno is NULL"
                    ''dbget.execute sqlStr
                    ''pErrMsg = pErrMsg & " : ����"
                    '' �����ϸ� �ȵ�.. ���� ��ǰ�� �������� ����
                end if
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] ��ϼ��� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF

        ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg


elseif (mode="delitemAuto") or (mode="expireitemAuto")  then  '' ���� ǰ��ó�� // ���޻����� �Ǹ�����
    cksel =""

    set oInterParkitem = new CExtSiteItem
    oInterParkitem.FPageSize       = 20
    oInterParkitem.FCurrPage       = 1
    oInterParkitem.FRectExtNotReg  = ""                ''��Ͽ���
    oInterParkitem.FRectMatchCate  = ""
    oInterParkitem.FRectExtSellYn   = "Y"            '' �������ܷ� ���� /2013/05/23
    if (mode="expireitemAuto") then
        oInterParkitem.FRectExtSellYn   = "YN"
        oInterParkitem.FRectOnlyNotUsingCheck ="on"
    end if
    ''oInterParkitem.FRectFailCntOverExcept="3"       '' 3ȸ �̻� ���г��� ����.
    oInterParkitem.GetInterParkExpireItemList


    for i=0 to oInterParkitem.FResultCount - 1
        cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
    next

    SET oInterParkitem=Nothing
rw cksel
    IF RIGHT(cksel,1)="" THEN cksel= cksel&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""

    if (mode="expireitemAuto") then
        mode="delitemONE"     '' X
        bufStr = "�Ǹű���"
    else
        mode="sellStatNONE"   ''�Ǹ�����
        bufStr = "�Ǹ�����"
    end if



    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                 call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") _
                    or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSEIF  (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")  then
                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set mayiParkSellYn='X'" & VbCrlf
                    sqlStr = sqlStr & " ,interparklastupdate=getdate()" & VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-X flag<br>"
                ELSE

                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set accfailCnt=accfailCnt+1" & VbCrlf
                    sqlStr = sqlStr & " ,lastErrStr=convert(varchar(100),'"&Trim(html2db(ErrMsg))&"')"& VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr

                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] "&bufStr&"���� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF
          ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�����Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
elseif (mode="infoDivNone")  then
    cksel =""

    set oInterParkitem = new CExtSiteItem
    oInterParkitem.FPageSize       = 10
    oInterParkitem.FCurrPage       = 1
    oInterParkitem.FRectExtNotReg  = "F"                ''��ϿϷ�
    oInterParkitem.FRectMatchCate  = ""
    oInterParkitem.FRectExtSellYn   = "Y"
    oInterParkitem.FRectInfoDivYn = "N"
    ''oInterParkitem.FRectFailCntExists = "on"            ''���� �ּ� ó��
    oInterParkitem.GetInterParkRegedItemList


    for i=0 to oInterParkitem.FResultCount - 1
        cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
    next

    SET oInterParkitem=Nothing
rw cksel

    IF RIGHT(cksel,1)="" THEN cksel= cksel&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""


    mode="sellStatNONE"                     '''ǰ��
    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                 call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") _
                    or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSEIF  (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")  then
                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set mayiParkSellYn='X'" & VbCrlf
                    sqlStr = sqlStr & " ,interparklastupdate=getdate()" & VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-X flag<br>"
                ELSE

                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set accfailCnt=accfailCnt+1" & VbCrlf
                    sqlStr = sqlStr & " ,lastErrStr=convert(varchar(100),'"&Trim(html2db(ErrMsg))&"')"& VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr

                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �Ǹ��������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF
          ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�Ǹ������Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
elseif (mode="iparkmarginNotSaleItem")  then
	cksel =""
	set oInterParkitem = new CExtSiteItem
	oInterParkitem.FPageSize       = 10
	oInterParkitem.FCurrPage       = 1
	oInterParkitem.FRectExtNotReg  = "F"                ''��ϿϷ�
	oInterParkitem.FRectMatchCate  = ""
	oInterParkitem.FRectSailYn	   = "N"
	oInterParkitem.FRectMinusMigin15 = "N"
	oInterParkitem.FRectExtSellYn   = "Y"
	oInterParkitem.GetInterParkRegedItemList


    for i=0 to oInterParkitem.FResultCount - 1
        cksel = cksel & oInterParkitem.FItemList(i).FItemID & ","
    next

    SET oInterParkitem=Nothing
rw cksel

    IF RIGHT(cksel,1)="" THEN cksel= cksel&","
    cksel = split(cksel,",")

    ''��� ������ �Ѱ����� �����ϴ���.
    succCNT = 0
    totCNT  = 0
    pErrMsg = ""


    mode="sellStatNONE"                     '''ǰ��
    for i=LBound(cksel) to UBound(cksel)
        itemid = TRIM(cksel(i))
        IF (itemid<>"") then
            totCNT = totCNT + 1
            ErrMsg = ""
            ErrCode = ""

            retCNT = iParkOneItemProcNewAPI(itemid,mode,dataUrl,ErrMsg,ErrCode,originPrdNo,prdNo,prdNm)
            succCNT = succCNT + retCNT

            IF (retCNT=0) then
                 call Fn_AcctFailTouch(CMALLNAME,itemid,"["&ErrCode&"]"&ErrMsg)
                pErrMsg = pErrMsg & "["&itemid&"] <font color=red>��������</font> - ("&itemid&":"&ErrCode&")"&ErrMsg&"<br>"
                IF (Trim(ErrMsg)="�ش� ��ü�� ��ǰ�� �ƴ� 'prdNo'") _
                    or (Trim(ErrMsg)="�Ǹű���, TNS�Ǹű��� ������ ��ǰ�� ��ǰ���������� �Ұ����մϴ�.") then ''or (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")
                    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-����<br>"
                ELSEIF  (Trim(ErrMsg)="�������� �ʴ� ��ǰ��ȣ")  then
                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set mayiParkSellYn='X'" & VbCrlf
                    sqlStr = sqlStr & " ,interparklastupdate=getdate()" & VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr,assignedRow
                    IF (assignedRow>0) then pErrMsg = pErrMsg & "-X flag<br>"
                ELSE

                    sqlStr = "update  [db_item].[dbo].tbl_interpark_reg_item"
                    sqlStr = sqlStr & " set accfailCnt=accfailCnt+1" & VbCrlf
                    sqlStr = sqlStr & " ,lastErrStr=convert(varchar(100),'"&Trim(html2db(ErrMsg))&"')"& VbCrlf
                    sqlStr = sqlStr & " where itemid=" & itemid
                    dbget.Execute sqlStr

                    rw ErrMsg
                End IF
            ELSe
                pErrMsg = pErrMsg & "["&itemid&"] �Ǹ��������� - ("&prdNo&")"&prdNm&" <a target=_blank href='http://www.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo="&prdNo&"'><font color=blue>[����]</font></a><br>"
            end IF
          ENd IF
    Next

    pErrMsg = "�� ��û�Ǽ�:"&totCNT&"<br>"&"�Ǹ������Ǽ�:"&succCNT&"<br>"&pErrMsg
    rw pErrMsg
ELSEIF (mode="regitemIMSIArr") then ''�ӽõ��
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="," THEN itemidARR= LEFT(itemidARR,Len(itemidARR)-1)

    sqlStr = "insert into [db_item].[dbo].tbl_interpark_reg_item " + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid) " + VbCrlf
    sqlStr = sqlStr + " select top 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + "     left join  [db_item].[dbo].tbl_interpark_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='110' and  i.cate_mid in ('010','020','030','040','050','070','090'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','055','060','070','075','080','090','100'))" + VbCrlf
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " and Not (i.cate_large='110' and i.cate_mid='030'  and i.cate_small='040')" + VbCrlf  ''����

    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.itemid in (" + itemidArr + ")" + VbCrlf
    sqlStr = sqlStr + " and sellcash<>0" + VbCrlf
    sqlStr = sqlStr + " and ((sellcash-buycash)/sellcash)*100>=" & CMAXMARGIN & VbCrlf

    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf
    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')" + VbCrlf

    ''��Ͻ� ����.
    sqlStr = sqlStr + " and i.makerid<>'haba'" + VbCrlf
    sqlStr = sqlStr + " and ((i.deliverytype<6) or " + VbCrlf
    sqlStr = sqlStr + "     ((i.deliverytype=9) " + VbCrlf
    sqlStr = sqlStr + "     and " + VbCrlf
    sqlStr = sqlStr + "     i.sellcash>=10000 " + VbCrlf ''' ���ǹ���� 1���� �̻�¥����..
    sqlStr = sqlStr + " ))" + VbCrlf
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"

    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
    response.write "<script >document.location.href='"&ref&"';</script>"
ELSEIF (mode="delitemIMSIArr") then ''�ӽõ�ϻ���
    itemidARR = Trim(cksel)
    IF RIGHT(cksel,1)="," THEN itemidARR= LEFT(itemidARR,Len(itemidARR)-1)

    sqlStr = "DELETE R " + VbCrlf
    sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item R " + VbCrlf
    sqlStr = sqlStr + " where R.itemid in (" + itemidArr + ")" + VbCrlf
    sqlStr = sqlStr + " and interparkregdate is NULL" + VbCrlf
    sqlStr = sqlStr + " and interparkPrdNo is NULL" + VbCrlf

    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ���� �����Ǿ����ϴ�.')</script>"
    response.write "<script >document.location.href='"&ref&"';</script>"

ELSEIF (mode="regByEventIDarr") then ''�̺�Ʈ �ڵ�� �ӽõ��
    sqlStr = "insert into [db_item].[dbo].tbl_interpark_reg_item" + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid)" + VbCrlf
    sqlStr = sqlStr + " select top 1000 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_event].[dbo].tbl_eventitem e," + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + " left join  [db_item].[dbo].tbl_interpark_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where e.evt_code in (" + eventidArr + ")" + VbCrlf
    sqlStr = sqlStr + " and e.itemid=i.itemid" + VbCrlf
    sqlStr = sqlStr + " and (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='110' and  i.cate_mid in ('010','020','030','040','050','070','090'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','055','060','070','075','080','090','100'))" + VbCrlf
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " and Not (i.cate_large='110' and i.cate_mid='030'  and i.cate_small='040')" + VbCrlf  ''����

    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.sellcash<>0" + VbCrlf
    sqlStr = sqlStr + " and (( i.sellcash- i.buycash)/ i.sellcash)*100>=" & CMAXMARGIN & "" + VbCrlf

    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf

    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')" + VbCrlf

    sqlStr = sqlStr + " and ((i.deliverytype<6) or " + VbCrlf
    sqlStr = sqlStr + "     ((i.deliverytype=9) " + VbCrlf
    sqlStr = sqlStr + "     and " + VbCrlf
    sqlStr = sqlStr + "     i.sellcash>=10000 " + VbCrlf ''' ���ǹ���� 1���� �̻�¥����..
    sqlStr = sqlStr + " ))" + VbCrlf

    ''��Ͻ� ����.
    sqlStr = sqlStr + " and i.makerid<>'haba'" + VbCrlf
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"

    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
    response.write "<script >document.location.href='"&ref&"';</script>"
elseif (mode="regByMakerid") then
    sqlStr = "insert into [db_item].[dbo].tbl_interpark_reg_item" + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid)" + VbCrlf
    sqlStr = sqlStr + " select top 200 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item_contents c,[db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + " left join  [db_item].[dbo].tbl_interpark_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where i.itemid=c.itemid" + VbCrlf
    sqlStr = sqlStr + " and i.makerid ='" & makeridArr & "'" + VbCrlf
    sqlStr = sqlStr + " and (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='110' and  i.cate_mid in ('010','020','030','040','050','070','090'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','055','060','070','075','080','090','100'))" + VbCrlf
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " and Not (i.cate_large='110' and i.cate_mid='030'  and i.cate_small='040')" + VbCrlf  ''����


    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.sellyn='Y'" + VbCrlf
    sqlStr = sqlStr + " and ((i.limityn='N') or (i.limityn='Y' and i.limitno-i.limitsold>=10))" + VbCrlf
    sqlStr = sqlStr + " and i.regdate>'2007-01-01'" + VbCrlf
    sqlStr = sqlStr + " and i.sellcash>500" + VbCrlf
    sqlStr = sqlStr + " and ((i.sellcash-i.buycash)/i.sellcash)*100>=" & CMAXMARGIN & "" & VbCrlf
    sqlStr = sqlStr + " and i.basicimage is not null"
    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf

    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')" + VbCrlf

    sqlStr = sqlStr + " and ((i.deliverytype<6) or " + VbCrlf
    sqlStr = sqlStr + "     ((i.deliverytype=9) " + VbCrlf
    sqlStr = sqlStr + "     and " + VbCrlf
    sqlStr = sqlStr + "     i.sellcash>=10000 " + VbCrlf ''' ���ǹ���� 1���� �̻�¥����..
    sqlStr = sqlStr + " ))" + VbCrlf

    ''��Ͻ� ����.
    sqlStr = sqlStr + " and i.makerid<>'haba'" + VbCrlf
    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"

    sqlStr = sqlStr + " order by i.itemid desc"

    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
    response.write "<script >document.location.href='"&ref&"';</script>"
elseif (mode="recentBestSeller") then
    sqlStr = "insert into [db_item].[dbo].tbl_interpark_reg_item" + VbCrlf
    sqlStr = sqlStr + " (itemid,reguserid)" + VbCrlf
    sqlStr = sqlStr + " select top 100 i.itemid,'" + session("ssBctID") + "'" + VbCrlf
    sqlStr = sqlStr + "  from" + VbCrlf
    sqlStr = sqlStr + " [db_item].[dbo].tbl_item_contents c,[db_item].[dbo].tbl_item i" + VbCrlf
    sqlStr = sqlStr + " left join  [db_item].[dbo].tbl_interpark_reg_item t on i.itemid=t.itemid" + VbCrlf
    sqlStr = sqlStr + " where i.itemid=c.itemid" + VbCrlf
    sqlStr = sqlStr + " and (" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large='110' and  i.cate_mid in ('010','020','030','040','050','070','090'))" + VbCrlf
    sqlStr = sqlStr + " 	or" + VbCrlf
    sqlStr = sqlStr + " 	(i.cate_large in ('010','020','025','030','035','040','045','050','055','060','070','075','080','090','100'))" + VbCrlf
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " and Not (i.cate_large='110' and i.cate_mid='030'  and i.cate_small='040')" + VbCrlf  ''����


    sqlStr = sqlStr + " and t.itemid is null" + VbCrlf
    sqlStr = sqlStr + " and i.sellyn='Y'" + VbCrlf
    sqlStr = sqlStr + " and ((i.limityn='N') or (i.limityn='Y' and i.limitno-i.limitsold>=10))" + VbCrlf
    sqlStr = sqlStr + " and (c.recentsellcount>=1 or c.recentfavcount>=5)" + VbCrlf
    ''sqlStr = sqlStr + " and c.recentfavcount>=1" + VbCrlf
    sqlStr = sqlStr + " and c.sellcount>=1" + VbCrlf
    sqlStr = sqlStr + " and i.regdate>'2008-01-01'" + VbCrlf
    sqlStr = sqlStr + " and i.sellcash>500" + VbCrlf
    sqlStr = sqlStr + " and ((i.sellcash-i.buycash)/i.sellcash)*100>=20" + VbCrlf
    sqlStr = sqlStr + " and i.basicimage is not null"

    ''Ư����ǰ����
    sqlStr = sqlStr + " and i.itemid<>114039" + VbCrlf

    sqlStr = sqlStr + " and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark')" + VbCrlf

    sqlStr = sqlStr + " and ((i.deliverytype<6) or " + VbCrlf
    sqlStr = sqlStr + "     ((i.deliverytype=9) " + VbCrlf
    sqlStr = sqlStr + "     and " + VbCrlf
    sqlStr = sqlStr + "     i.sellcash>=10000 " + VbCrlf ''' ���ǹ���� 1���� �̻�¥����..
    sqlStr = sqlStr + " ))" + VbCrlf

    ''��Ͻ� ����.
    sqlStr = sqlStr + " and i.makerid<>'haba'" + VbCrlf

    ''���� �������ΰ� �ɷ���. isExtusing = 'N'
    sqlStr = sqlStr + " and i.isExtusing = 'Y'"

    sqlStr = sqlStr + " order by i.itemid desc"

    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� ��ϵǾ����ϴ�.')</script>"
    response.write "<script >document.location.href='"&ref&"';</script>"
ELSE
    response.write "������ mode=="&mode
end if

%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
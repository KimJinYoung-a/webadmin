<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 900
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim cksel : cksel = request("cksel")

Dim ord_no          : ord_no = requestCheckVar(request("ord_no"),32)
Dim ord_dtl_sn      : ord_dtl_sn = requestCheckVar(request("ord_dtl_sn"),32)
Dim inv_no          : inv_no = requestCheckVar(request("inv_no"),32)
Dim sendQnt         : sendQnt = requestCheckVar(request("sendQnt"),10)
Dim sendDate        : sendDate = requestCheckVar(request("sendDate"),10)
Dim outmallGoodsID  : outmallGoodsID = requestCheckVar(request("outmallGoodsID"),32)
Dim hdc_cd          : hdc_cd = requestCheckVar(request("hdc_cd"),10)
Dim subcmd          : subcmd = requestCheckVar(request("subcmd"),10)

Dim xmlDOM, CateInfo, ConfirmResult, GoodseDtInfo
Dim L_CODE, L_NAME, M_CODE, M_NAME, S_CODE, S_NAME, D_CODE, D_NAME
Dim sqlStr, AssignedRow, AssignedTTL
Dim i, iitemid, ret, ierrStr
Dim RESULT_MSG, GODDS_B2BCODE, ENTP_GOODS_CODE, GOODS_BUY_PRICE, GOODS_SALE_PRICE, REG_DATE
Dim GOODSDT_CODE, ENTP_DT_CODE, GOODS_INFO, GOODS_MAX_STC, SALE_GB
Dim SubNodes, SubSubNodes
Dim SuccCNT, FailCNT, alertMsg
Dim ArrRows, bufStr
Dim isValidDel

IF (cmdparam="getdispcate") then
    set xmlDOM= getLotteiMallXMLReq(cmdparam,False,ierrStr,"")
    ''set xmlDOM= getLotteiMallXMLReqTestFile(cmdparam,False)

    if (Not (xmlDOM is Nothing)) then
        sqlStr = "delete from db_temp.dbo.tbl_LTiMall_Category_BUF where cateGbn in ('D','B')"
        dbget.Execute sqlStr
        Set CateInfo = xmlDOM.getElementsByTagName("CategoryInfo")

        for each SubNodes in CateInfo
			D_CODE	= Trim(SubNodes.getElementsByTagName("D_CODE").item(0).text)		'카테고리 코드
			D_NAME	= Trim(SubNodes.getElementsByTagName("D_NAME").item(0).text)		'카테고리명(세세분류)
			L_CODE	= Trim(SubNodes.getElementsByTagName("L_CODE").item(0).text)		'대분류명
			L_NAME	= Trim(SubNodes.getElementsByTagName("L_NAME").item(0).text)		'중분류명
			M_CODE	= Trim(SubNodes.getElementsByTagName("M_CODE").item(0).text)		'소분류명
			M_NAME	= Trim(SubNodes.getElementsByTagName("M_NAME").item(0).text)		'세분류명
			S_CODE	= Trim(SubNodes.getElementsByTagName("S_CODE").item(0).text)		'소분류명
			S_NAME	= Trim(SubNodes.getElementsByTagName("S_NAME").item(0).text)		'세분류명

			sqlStr = "Insert Into db_temp.dbo.tbl_LTiMall_Category_BUF"
			sqlStr = sqlStr & " (CateKey,cateGbn,L_CODE,M_CODE,S_CODE,D_CODE,L_NAME,M_NAME,S_NAME,D_NAME)"
			sqlStr = sqlStr & " values('"&D_CODE&"'"
			sqlStr = sqlStr & " ,'D'"
			sqlStr = sqlStr & " ,'"&L_CODE&"'"
			sqlStr = sqlStr & " ,'"&M_CODE&"'"
			sqlStr = sqlStr & " ,'"&S_CODE&"'"
			sqlStr = sqlStr & " ,'"&D_CODE&"'"
			sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(L_NAME)&"')"
			sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(M_NAME)&"')"
			sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(S_NAME)&"')"
			sqlStr = sqlStr & " ,convert(varchar(100),'"&html2db(D_NAME)&"')"
			sqlStr = sqlStr & " )"
			dbget.Execute sqlStr, AssignedRow

			AssignedTTL = AssignedTTL + AssignedRow
        Next
		Set CateInfo = Nothing

		''
		if (AssignedTTL>5000) then  ''Maybe Over 5000 Rows


		    sqlStr = "update C"
		    sqlStr = sqlStr & " set isusing=(CASE WHEN B.CateKey is NULL THEN 'N' ELSE 'Y' END)"
		    sqlStr = sqlStr & " ,lastupdate=getdate()"
		    sqlStr = sqlStr & " from db_temp.dbo.tbl_LTiMall_Category C"
		    sqlStr = sqlStr & "     left join db_temp.dbo.tbl_LTiMall_Category_BUF B"
		    sqlStr = sqlStr & "     on C.CateKey=B.CateKey"
		    ''sqlStr = sqlStr & "     and C.cateGbn=B.cateGbn"
		    sqlStr = sqlStr & " where C.cateGbn in ('D','B')"
		    dbget.Execute sqlStr

		    sqlStr = "insert into db_temp.dbo.tbl_LTiMall_Category"
		    sqlStr = sqlStr & " (CateKey,cateGbn,L_CODE,M_CODE,S_CODE,D_CODE,L_NAME,M_NAME,S_NAME,D_NAME)"
		    sqlStr = sqlStr & " select B.CateKey,'D',B.L_CODE,B.M_CODE,B.S_CODE,B.D_CODE,B.L_NAME,B.M_NAME,B.S_NAME,B.D_NAME"
		    sqlStr = sqlStr & " from db_temp.dbo.tbl_LTiMall_Category_BUF B"
		    sqlStr = sqlStr & "     left join db_temp.dbo.tbl_LTiMall_Category C"
		    sqlStr = sqlStr & "     on C.CateKey=B.CateKey"
		    ''sqlStr = sqlStr & "     and C.cateGbn=B.cateGbn"
		    sqlStr = sqlStr & " where C.CateKey is NULL"

		    dbget.Execute sqlStr, AssignedRow

		    sqlStr = "update  db_temp.dbo.tbl_LTiMall_Category"
            sqlStr = sqlStr & " set isUsing='N'"
            sqlStr = sqlStr & " where (M_NAME like '%DCX 디자인소품%'"
            sqlStr = sqlStr & " or  M_NAME like '%바보사랑%'"
            sqlStr = sqlStr & " or M_NAME like '%하이모리 디자인샵%'"
            sqlStr = sqlStr & " or S_NAME like '%이즈워즈%')"
            sqlStr = sqlStr & " and isusing='Y'"
            dbget.Execute sqlStr

            ''전문카테고리 구분변경
            sqlStr = "update db_temp.dbo.tbl_LTiMall_Category_BUF"
            sqlStr = sqlStr&" set CateGbn='B'"
            sqlStr = sqlStr&" where L_CODE='10500000'"
            sqlStr = sqlStr&" and M_Code='201200078827'"
            dbget.Execute sqlStr

            ''2013/04/29 추가
            sqlStr = "update db_temp.dbo.tbl_LTiMall_Category"
            sqlStr = sqlStr&" set CateGbn='B'"
            sqlStr = sqlStr&" where L_CODE='50000000'"
            sqlStr = sqlStr&" and M_Code='201300115948'"
            dbget.Execute sqlStr

            ''20120820 카테고리 변경 / 임시등록건.
            sqlStr = "update  db_temp.dbo.tbl_LTiMall_Category"
            sqlStr = sqlStr & " set isUsing='Y'"
            sqlStr = sqlStr & " where L_Code='201200082559'"
            sqlStr = sqlStr & " and M_CODe='201200095507'"
            dbget.Execute sqlStr

            sqlStr = " update db_temp.dbo.tbl_LTiMall_Category"
            sqlStr = sqlStr & " set isusing='N'"
            sqlStr = sqlStr & " where NOT ( (L_Code='201200082559' and M_CODe='201200095507')"
            sqlStr = sqlStr & " or (L_CODE='50000000' and M_Code='201300115948')"
            sqlStr = sqlStr & " )"
            sqlStr = sqlStr & " and isusing='Y'"
            dbget.Execute sqlStr

		    response.write AssignedRow&"건 반영됨"
		end if
    else
        rw ierrStr
        rw "ERR:xmlDOM is Nothing"
    end if
    set xmlDOM= Nothing
    response.write "<script>alert('완료');</script>"
ELSEIF (cmdparam="CheckItemStatAuto") then ''판매상태 확인
    SuccCNT = 0

    sqlStr = "select top 20 r.itemid "
    sqlStr = sqlStr & "	from db_item.dbo.tbl_LtiMall_regitem r"
    sqlStr = sqlStr & "	where 1=1"
    sqlStr = sqlStr & "	and r.itemid not in (621817)" '' 확인
    sqlStr = sqlStr & "	and LtimallStatCd>3" '' 1 전송시도,3 승인대기
    sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.LTImallsellyn='X' THEN '0' ELSE r.LTImallsellyn END),  r.LtiMallLastUpdate , r.itemid desc"
    ''sqlStr = sqlStr & "	order by r.lastStatCheckDate, (CASE WHEN r.LTImallsellyn='X' THEN '0' ELSE r.LTImallsellyn END), isNULL(rctSellcnt,0) desc,  isNULL(r.regedoptCnt,2) desc , r.LtiMallLastUpdate , r.itemid desc" ''isNULL(r.accfailcnt,0) desc,

    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        ArrRows = rsget.getRows()
    end if
    rsget.close

    bufStr=""
    if isArray(ArrRows) then

        For i =0 To UBound(ArrRows,2)
            ierrStr = ""
            iitemid = CStr(ArrRows(0,i))
            if (Not chkLotteiMallOneItem(cmdparam, iitemid,  ierrStr,  SuccCNT, isValidDel)) then
                bufStr = bufStr + iitemid + ","
            end if

            if (ierrStr<>"") then
                rw ierrStr
            end if
        next

        if (bufStr<>"") then rw "ERR:"&bufStr
    end if
'    if (SuccCNT>0) then
'        alertMsg = ""&SuccCNT&"건 승인 "
'    else
'        alertMsg = "승인건 없음!"
'    end if
'    if (FailCNT>0) then
'        alertMsg = alertMsg & ""&FailCNT&"건 실패 "
'    end if
'
'    if (alertMsg<>"") then
'        'response.write "<script>if (confirm('"&alertMsg&"\n\nreload?')){parent.location.reload();};</script>"
'        'dbget.close() : response.end
'
'    end if
ELSEIF (cmdparam="getconfirmList") then ''등록대기상품 확인
    ''if (request("subcmd"))<>"arrconfirm") then cksel=""
    SuccCNT = 0
    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        ierrStr =""

        call chkLotteiMallOneItem(cmdparam, iitemid,  ierrStr,  SuccCNT, isValidDel)

        if (ierrStr<>"") then
            rw ierrStr
        end if
    next

    if (SuccCNT>0) then
        alertMsg = ""&SuccCNT&"건 승인 "
    else
        alertMsg = "승인건 없음!"
    end if
    if (FailCNT>0) then
        alertMsg = alertMsg & ""&FailCNT&"건 실패 "
    end if

     if (alertMsg<>"") then
        'response.write "<script>if (confirm('"&alertMsg&"\n\nreload?')){parent.location.reload();};</script>"
        'dbget.close() : response.end

    end if
ELSEIF (cmdparam="RegSelectWait") then   ''선택상품 예정등록.
    cksel = Trim(cksel)
    if Right(cksel,1)="," then cksel=Left(cksel,Len(cksel)-1)

    sqlStr = "Insert into db_item.dbo.tbl_LTiMall_regItem"
    sqlStr = sqlStr & " (itemid,regdate,reguserid,LtiMallStatCD)"
    sqlStr = sqlStr & " select i.itemid,getdate(),'"&session("SSBctID")&"',0"
    sqlStr = sqlStr & " from db_item.dbo.tbl_item i"
    sqlStr = sqlStr & "     left join db_item.dbo.tbl_LTiMall_regItem R"
    sqlStr = sqlStr & "     on i.itemid=R.itemid"
    sqlStr = sqlStr & " where i.itemid in ("&cksel&")"
    ''sqlStr = sqlStr & " and i.sellyn='Y'"
    sqlStr = sqlStr & " and R.itemid is NULL"
''rw  sqlStr
    dbget.Execute sqlStr,AssignedRow

    response.write "<script>alert('"&AssignedRow&"건 예정등록됨.');parent.location.reload();</script>"
ELSEIF (cmdparam="DelSelectWait") then   ''선택상품 예정등록삭제.
    cksel = Trim(cksel)
    if Right(cksel,1)="," then cksel=Left(cksel,Len(cksel)-1)

    sqlStr = "delete from db_item.dbo.tbl_LTiMall_regItem"
    sqlStr = sqlStr & " where LtimallStatCD in (0,-1)"
    sqlStr = sqlStr & " and itemid in ("&cksel&")"
''rw  sqlStr
    dbget.Execute sqlStr,AssignedRow

    response.write "<script>alert('"&AssignedRow&"건 예정 삭제됨.');parent.location.reload();</script>"

ELSEIF (cmdparam="CheckNDel") then   ''체크 후 삭제

    ierrStr = ""
    iitemid = Trim(cksel)
    if (Not chkLotteiMallOneItem("CheckItemStatAuto", iitemid,  ierrStr,  SuccCNT, isValidDel)) then
        sqlStr = "delete from db_item.dbo.tbl_LTiMall_regItem"
        sqlStr = sqlStr & " where LtimallStatCD in (0,-1,1)"  ''등록대기, 실패, 전송시도
        sqlStr = sqlStr & " and itemid in ("&cksel&")"

        dbget.Execute sqlStr,AssignedRow

        response.write "<script>alert('"&AssignedRow&"건 삭제됨.');</script>"
    end if

    if (ierrStr<>"") then
        rw ierrStr
    end if
ELSEIF (cmdparam="CheckNDelReged") then   ''체크 후 삭제 등록된 상품

    ierrStr = ""
    iitemid = Trim(cksel)
    if (chkLotteiMallOneItem("CheckItemStatAuto", iitemid,  ierrStr,  SuccCNT, isValidDel)) then
        if (isValidDel) then
            sqlStr = "delete from db_item.dbo.tbl_LTiMall_regItem"
            sqlStr = sqlStr & " where itemid in ("&cksel&")"
            sqlStr = sqlStr & " and Ltimallsellyn='X'"

            dbget.Execute sqlStr,AssignedRow

            if (AssignedRow<1) then
                response.write "<script>alert('삭제 불가. 먼저 imall어드민에서 판매중단 등록 후 사용바람');</script>"
            else
                response.write "<script>alert('"&AssignedRow&"건 삭제됨.');</script>"
            end if
        else
            response.write "<script>alert('삭제 불가. 먼저 imall어드민에서 판매중단 등록 후 사용바람');</script>"
        end if
    end if

    if (ierrStr<>"") then
        rw ierrStr
    end if
ELSEIF (cmdparam="RegSelect") then   ''선택상품 실등록.
    SuccCNT = 0
    FailCNT = 0
    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        ret = regLotteiMallOneItem(iitemid, ierrStr)

        if (Not ret) then
            FailCNT = FailCNT +1
            rw ierrStr
        else
            SuccCNT = SuccCNT +1
        end if
    next

    alertMsg = ""&SuccCNT&"건 등록 "
    if (FailCNT>0) then
        alertMsg = alertMsg & ""&FailCNT&"건 실패 "
    end if


ELSEIF (cmdparam="EditSelect") then ''선택상품 실 수정.
    ''response.write "수정중"
    ''response.end
    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        ret = editLotteiMallOneItem(iitemid, ierrStr)
        if (Not ret) then
            rw ierrStr
        end if
        Call chkLotteiMallOneItem("CheckItemStatAuto", iitemid, ierrStr, SuccCNT, isValidDel)  ''2013/03/28 추가 아이몰 판매상태 check
    next
ELSEIF (cmdparam="EdSaleDTSel") then ''선택상품 단품 수정.
    '' response.write "수정중"
    ''response.end
    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        ret = editDTLotteiMallOneItem(iitemid, ierrStr)

        if (Not ret) then
            rw ierrStr
        end if
        Call chkLotteiMallOneItem("CheckItemStatAuto", iitemid, ierrStr, SuccCNT, isValidDel)  ''2013/03/28 추가 아이몰 판매상태 check
    next
ELSEIF (cmdparam="EditSellYn") then ''선택상품 판매상태 수정
    rw subcmd

    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        SuccCNT = 0
        Call chkLotteiMallOneItem(cmdparam, iitemid, ierrStr, SuccCNT, isValidDel)  ''2013/03/27 추가 아이몰 판매상태 check
        ret = editSOLDOUTLotteiMallOneItem(iitemid, ierrStr)

        if (Not ret) then
            rw ierrStr
        end if
    next
ELSEIF (cmdparam="songjangip")  then ''송장입력
    ''rw ord_no&"songjangip"
    ''rw hdc_cd
    ''rw inv_no

    if (hdc_cd="99") and Len(replace(inv_no,"-",""))>15 then inv_no=Left(replace(inv_no,"-",""),15)
    if (inv_no="11시배송완료") then inv_no="11시배송"
    if (inv_no="핸드폰으로전송예정:)") then inv_no="기타"
    	

    If instr(inv_no,"-") > 0 Then			'2015-07-10 김진영 추가
    	inv_no = replace(inv_no, "-", "")
    End If


    CAll regLotteiMallSongjang(ord_no,ord_dtl_sn,hdc_cd,inv_no,sendQnt,sendDate,outmallGoodsID, ierrStr)
    ''rw "FIN"
'2013/02/28 진영추가'
ELSEIF (cmdparam="updateSendState") then
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');opener.close();window.close()</script>"
else
    rw "미지정 ["&cmdparam&"]"
end if


if (alertMsg<>"") then
    IF (IsAutoScript) then
        rw alertMsg
    else
        response.write "<script>alert('"&alertMsg&"');</script>"
    end if
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
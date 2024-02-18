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
			D_CODE	= Trim(SubNodes.getElementsByTagName("D_CODE").item(0).text)		'ī�װ� �ڵ�
			D_NAME	= Trim(SubNodes.getElementsByTagName("D_NAME").item(0).text)		'ī�װ���(�����з�)
			L_CODE	= Trim(SubNodes.getElementsByTagName("L_CODE").item(0).text)		'��з���
			L_NAME	= Trim(SubNodes.getElementsByTagName("L_NAME").item(0).text)		'�ߺз���
			M_CODE	= Trim(SubNodes.getElementsByTagName("M_CODE").item(0).text)		'�Һз���
			M_NAME	= Trim(SubNodes.getElementsByTagName("M_NAME").item(0).text)		'���з���
			S_CODE	= Trim(SubNodes.getElementsByTagName("S_CODE").item(0).text)		'�Һз���
			S_NAME	= Trim(SubNodes.getElementsByTagName("S_NAME").item(0).text)		'���з���

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
            sqlStr = sqlStr & " where (M_NAME like '%DCX �����μ�ǰ%'"
            sqlStr = sqlStr & " or  M_NAME like '%�ٺ����%'"
            sqlStr = sqlStr & " or M_NAME like '%���̸� �����μ�%'"
            sqlStr = sqlStr & " or S_NAME like '%�������%')"
            sqlStr = sqlStr & " and isusing='Y'"
            dbget.Execute sqlStr

            ''����ī�װ� ���к���
            sqlStr = "update db_temp.dbo.tbl_LTiMall_Category_BUF"
            sqlStr = sqlStr&" set CateGbn='B'"
            sqlStr = sqlStr&" where L_CODE='10500000'"
            sqlStr = sqlStr&" and M_Code='201200078827'"
            dbget.Execute sqlStr

            ''2013/04/29 �߰�
            sqlStr = "update db_temp.dbo.tbl_LTiMall_Category"
            sqlStr = sqlStr&" set CateGbn='B'"
            sqlStr = sqlStr&" where L_CODE='50000000'"
            sqlStr = sqlStr&" and M_Code='201300115948'"
            dbget.Execute sqlStr

            ''20120820 ī�װ� ���� / �ӽõ�ϰ�.
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

		    response.write AssignedRow&"�� �ݿ���"
		end if
    else
        rw ierrStr
        rw "ERR:xmlDOM is Nothing"
    end if
    set xmlDOM= Nothing
    response.write "<script>alert('�Ϸ�');</script>"
ELSEIF (cmdparam="CheckItemStatAuto") then ''�ǸŻ��� Ȯ��
    SuccCNT = 0

    sqlStr = "select top 20 r.itemid "
    sqlStr = sqlStr & "	from db_item.dbo.tbl_LtiMall_regitem r"
    sqlStr = sqlStr & "	where 1=1"
    sqlStr = sqlStr & "	and r.itemid not in (621817)" '' Ȯ��
    sqlStr = sqlStr & "	and LtimallStatCd>3" '' 1 ���۽õ�,3 ���δ��
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
'        alertMsg = ""&SuccCNT&"�� ���� "
'    else
'        alertMsg = "���ΰ� ����!"
'    end if
'    if (FailCNT>0) then
'        alertMsg = alertMsg & ""&FailCNT&"�� ���� "
'    end if
'
'    if (alertMsg<>"") then
'        'response.write "<script>if (confirm('"&alertMsg&"\n\nreload?')){parent.location.reload();};</script>"
'        'dbget.close() : response.end
'
'    end if
ELSEIF (cmdparam="getconfirmList") then ''��ϴ���ǰ Ȯ��
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
        alertMsg = ""&SuccCNT&"�� ���� "
    else
        alertMsg = "���ΰ� ����!"
    end if
    if (FailCNT>0) then
        alertMsg = alertMsg & ""&FailCNT&"�� ���� "
    end if

     if (alertMsg<>"") then
        'response.write "<script>if (confirm('"&alertMsg&"\n\nreload?')){parent.location.reload();};</script>"
        'dbget.close() : response.end

    end if
ELSEIF (cmdparam="RegSelectWait") then   ''���û�ǰ �������.
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

    response.write "<script>alert('"&AssignedRow&"�� ������ϵ�.');parent.location.reload();</script>"
ELSEIF (cmdparam="DelSelectWait") then   ''���û�ǰ ������ϻ���.
    cksel = Trim(cksel)
    if Right(cksel,1)="," then cksel=Left(cksel,Len(cksel)-1)

    sqlStr = "delete from db_item.dbo.tbl_LTiMall_regItem"
    sqlStr = sqlStr & " where LtimallStatCD in (0,-1)"
    sqlStr = sqlStr & " and itemid in ("&cksel&")"
''rw  sqlStr
    dbget.Execute sqlStr,AssignedRow

    response.write "<script>alert('"&AssignedRow&"�� ���� ������.');parent.location.reload();</script>"

ELSEIF (cmdparam="CheckNDel") then   ''üũ �� ����

    ierrStr = ""
    iitemid = Trim(cksel)
    if (Not chkLotteiMallOneItem("CheckItemStatAuto", iitemid,  ierrStr,  SuccCNT, isValidDel)) then
        sqlStr = "delete from db_item.dbo.tbl_LTiMall_regItem"
        sqlStr = sqlStr & " where LtimallStatCD in (0,-1,1)"  ''��ϴ��, ����, ���۽õ�
        sqlStr = sqlStr & " and itemid in ("&cksel&")"

        dbget.Execute sqlStr,AssignedRow

        response.write "<script>alert('"&AssignedRow&"�� ������.');</script>"
    end if

    if (ierrStr<>"") then
        rw ierrStr
    end if
ELSEIF (cmdparam="CheckNDelReged") then   ''üũ �� ���� ��ϵ� ��ǰ

    ierrStr = ""
    iitemid = Trim(cksel)
    if (chkLotteiMallOneItem("CheckItemStatAuto", iitemid,  ierrStr,  SuccCNT, isValidDel)) then
        if (isValidDel) then
            sqlStr = "delete from db_item.dbo.tbl_LTiMall_regItem"
            sqlStr = sqlStr & " where itemid in ("&cksel&")"
            sqlStr = sqlStr & " and Ltimallsellyn='X'"

            dbget.Execute sqlStr,AssignedRow

            if (AssignedRow<1) then
                response.write "<script>alert('���� �Ұ�. ���� imall���ο��� �Ǹ��ߴ� ��� �� ���ٶ�');</script>"
            else
                response.write "<script>alert('"&AssignedRow&"�� ������.');</script>"
            end if
        else
            response.write "<script>alert('���� �Ұ�. ���� imall���ο��� �Ǹ��ߴ� ��� �� ���ٶ�');</script>"
        end if
    end if

    if (ierrStr<>"") then
        rw ierrStr
    end if
ELSEIF (cmdparam="RegSelect") then   ''���û�ǰ �ǵ��.
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

    alertMsg = ""&SuccCNT&"�� ��� "
    if (FailCNT>0) then
        alertMsg = alertMsg & ""&FailCNT&"�� ���� "
    end if


ELSEIF (cmdparam="EditSelect") then ''���û�ǰ �� ����.
    ''response.write "������"
    ''response.end
    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        ret = editLotteiMallOneItem(iitemid, ierrStr)
        if (Not ret) then
            rw ierrStr
        end if
        Call chkLotteiMallOneItem("CheckItemStatAuto", iitemid, ierrStr, SuccCNT, isValidDel)  ''2013/03/28 �߰� ���̸� �ǸŻ��� check
    next
ELSEIF (cmdparam="EdSaleDTSel") then ''���û�ǰ ��ǰ ����.
    '' response.write "������"
    ''response.end
    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        ret = editDTLotteiMallOneItem(iitemid, ierrStr)

        if (Not ret) then
            rw ierrStr
        end if
        Call chkLotteiMallOneItem("CheckItemStatAuto", iitemid, ierrStr, SuccCNT, isValidDel)  ''2013/03/28 �߰� ���̸� �ǸŻ��� check
    next
ELSEIF (cmdparam="EditSellYn") then ''���û�ǰ �ǸŻ��� ����
    rw subcmd

    cksel = split(cksel,",")
    For i=0 To UBound(cksel)
        iitemid=Trim(cksel(i))
        SuccCNT = 0
        Call chkLotteiMallOneItem(cmdparam, iitemid, ierrStr, SuccCNT, isValidDel)  ''2013/03/27 �߰� ���̸� �ǸŻ��� check
        ret = editSOLDOUTLotteiMallOneItem(iitemid, ierrStr)

        if (Not ret) then
            rw ierrStr
        end if
    next
ELSEIF (cmdparam="songjangip")  then ''�����Է�
    ''rw ord_no&"songjangip"
    ''rw hdc_cd
    ''rw inv_no

    if (hdc_cd="99") and Len(replace(inv_no,"-",""))>15 then inv_no=Left(replace(inv_no,"-",""),15)
    if (inv_no="11�ù�ۿϷ�") then inv_no="11�ù��"
    if (inv_no="�ڵ����������ۿ���:)") then inv_no="��Ÿ"
    	

    If instr(inv_no,"-") > 0 Then			'2015-07-10 ������ �߰�
    	inv_no = replace(inv_no, "-", "")
    End If


    CAll regLotteiMallSongjang(ord_no,ord_dtl_sn,hdc_cd,inv_no,sendQnt,sendDate,outmallGoodsID, ierrStr)
    ''rw "FIN"
'2013/02/28 �����߰�'
ELSEIF (cmdparam="updateSendState") then
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');opener.close();window.close()</script>"
else
    rw "������ ["&cmdparam&"]"
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
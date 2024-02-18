<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 상품등록
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
''======2010 추가=====================

function optKindSeq2Code(iseq)
    dim ascCode
    optKindSeq2Code = CStr(iseq)

    if (iseq>9) then
        iseq = iseq + 55
        if (iseq>64) and (iseq<91) then
            optKindSeq2Code = CHR(iseq)
        end if
    end if
end function
''======2010 추가=====================

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim itemid, itemoption
dim mode
dim arritemoption, arritemoptionname
dim optionTypename, optionName
dim optaddprice, optaddBuyprice

dim i, j, k, index, sqlStr, foundcount, found, ErrStr
dim TypeSeq, KindSeq

dim TypeCnt, OptCnt

itemid              = requestCheckVar(request("itemid"),10)
itemoption          = requestCheckVar(request("itemoption"),4)
mode                = requestCheckVar(request("mode"),32)
optionTypename      = requestCheckVar(request("optionTypename"),32)
arritemoption       = request("arritemoption")
arritemoptionname   = request("arritemoptionname")

TypeSeq             = requestCheckVar(request("TypeSeq"),10)
KindSeq             = requestCheckVar(request("KindSeq"),10)

arritemoption = Split(arritemoption, "|")
arritemoptionname = Split(arritemoptionname, "|")

dim IsUpchebeasong, itemLimitYn
IsUpchebeasong =false
itemLimitYn = "N"

if itemid="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('코드가 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.End
end if

''업체배송인경우 입출/판매 관계없이 삭제
sqlStr = " select limityn, deliverytype "
sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item"
sqlStr = sqlStr & " where itemid=" & CStr(itemid)

rsget.Open sqlStr,dbget,1
if not rsget.EOF then
    itemLimitYn = rsget("limityn")
    IsUpchebeasong = (rsget("deliverytype") = "2") or (rsget("deliverytype") = "5")
end if
rsget.Close


function ReMatchMultiOption(itemid)
    dim sqlStr
    dim MultiLevel

    MultiLevel = 0

    sqlStr = " select TypeSeq, Count(KindSeq) as KindCnt "
    sqlStr = sqlStr + " from [db_temp].[dbo].tbl_wait_item_option_Multiple "
    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    sqlStr = sqlStr + " group by TypeSeq"
    sqlStr = sqlStr + " order by TypeSeq"

    rsget.Open sqlStr, dbget, 1
	    MultiLevel = rsget.RecordCount
	rsget.close

    ''기존 2차 옵션인 경우 삭제.
    if (MultiLevel=3) then
        sqlStr = " delete from [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
        sqlStr = sqlStr + " where itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'" &VbCRLF
        sqlStr = sqlStr + " and Right(itemoption,1)='0'" &VbCRLF

        dbget.Execute sqlStr
    end if

    if (MultiLevel=2) then
        sqlStr = " delete from [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
        sqlStr = sqlStr + " where itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'" &VbCRLF
        sqlStr = sqlStr + " and Right(itemoption,1)='00'" &VbCRLF

        dbget.Execute sqlStr
    end if


    ''옵션 재작성.
'   --Only 1중옵션.
    if (MultiLevel=1) then
        ''-- 전 옵션 삭제;
        sqlStr = " delete from [db_temp].[dbo].tbl_wait_item_option_Multiple" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        dbget.Execute sqlStr

        sqlStr = " delete from [db_temp].[dbo].tbl_wait_itemoption" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        sqlStr = sqlStr & " and Left(itemoption,1)='Z'"
        dbget.Execute sqlStr

''        sqlStr = " insert into [db_temp].[dbo].tbl_wait_itemoption"
''        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
''        sqlStr = sqlStr + " select T.itemid, ('Z' + convert(varchar(1),T.KindSeq) + '0' + '0') as itemoption,"
''        sqlStr = sqlStr + " T.optionTypeName, T.optionKindName, 'Y','Y','" + itemLimitYn + "', 0, 0,"
''        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice"
''        sqlStr = sqlStr + " from [db_temp].[dbo].tbl_wait_item_option_Multiple T"
''        sqlStr = sqlStr + "     left join [db_temp].[dbo].tbl_wait_itemoption o "
''        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid)
''        sqlStr = sqlStr + "     and T.itemid=o.itemid "
''        sqlStr = sqlStr + "     and ('Z' + convert(varchar(1),T.KindSeq) + '0' + '0')=o.itemoption "
''        sqlStr = sqlStr + " where  o.itemid is NULL"
''
''        dbget.Execute sqlStr
''
''        '' 옵션명/ 가격 등이 변경된 경우
''        sqlStr = " update [db_temp].[dbo].tbl_wait_itemoption"
''        sqlStr = sqlStr + " set optionname=T.optionname"
''        sqlStr = sqlStr + " , optaddprice=T.optaddprice"
''        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice"
''        sqlStr = sqlStr + " from [db_temp].[dbo].tbl_wait_item_option_Multiple T "
''        sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_itemoption.itemid=T.itemid"
''        sqlStr = sqlStr + " and [db_temp].[dbo].tbl_wait_itemoption.itemoption=T.itemoption"
''        sqlStr = sqlStr + " and ("
''        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_itemoption.optionname<>T.optionname"
''        sqlStr = sqlStr + "     or [db_temp].[dbo].tbl_wait_itemoption.optaddprice<>T.optaddprice"
''        sqlStr = sqlStr + "     or [db_temp].[dbo].tbl_wait_itemoption.optaddbuyprice<>T.optaddbuyprice"
''        sqlStr = sqlStr + " )"
''
''        dbget.Execute sqlStr
    end if

'   --Only 2중옵션.
    if (MultiLevel=2) then
        sqlStr = " insert into [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) " &VbCRLF
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '복합옵션' as optionTypeName," &VbCRLF
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0," &VbCRLF
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_temp].[dbo].tbl_wait_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_item_option_Multiple b" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T" &VbCRLF
        sqlStr = sqlStr + "     left join [db_temp].[dbo].tbl_wait_itemoption o " &VbCRLF
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and T.itemid=o.itemid " &VbCRLF
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption " &VbCRLF
        sqlStr = sqlStr + " where  o.itemid is NULL"

        dbget.Execute sqlStr

        '' 옵션명/ 가격 등이 변경된 경우
        sqlStr = " update [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)" &VbCRLF
        sqlStr = sqlStr + " , optaddprice=T.optaddprice" &VbCRLF
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName ) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_temp].[dbo].tbl_wait_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_item_option_Multiple b" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T " &VbCRLF
        sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_itemoption.itemid=T.itemid" &VbCRLF
        sqlStr = sqlStr + " and [db_temp].[dbo].tbl_wait_itemoption.itemoption=T.itemoption" &VbCRLF
        sqlStr = sqlStr + " and (" &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_itemoption.optionname<>T.optionname" &VbCRLF
        sqlStr = sqlStr + "     or [db_temp].[dbo].tbl_wait_itemoption.optaddprice<>T.optaddprice" &VbCRLF
        sqlStr = sqlStr + "     or [db_temp].[dbo].tbl_wait_itemoption.optaddbuyprice<>T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr
    end if

'    --Only 3중옵션
    if (MultiLevel=3) then
        sqlStr = " insert into [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) " &VbCRLF
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '복합옵션' as optionTypeName," &VbCRLF
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0," &VbCRLF
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_temp].[dbo].tbl_wait_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_item_option_Multiple b," &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_item_option_Multiple c" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and b.itemid=c.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T " &VbCRLF
        sqlStr = sqlStr + "     left join [db_temp].[dbo].tbl_wait_itemoption o " &VbCRLF
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and T.itemid=o.itemid " &VbCRLF
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption " &VbCRLF
        sqlStr = sqlStr + " where  o.itemid is NULL" &VbCRLF

        dbget.Execute sqlStr


        '' 옵션명/ 가격 등이 변경된 경우
        sqlStr = " update [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)" &VbCRLF
        sqlStr = sqlStr + " , optaddprice=T.optaddprice" &VbCRLF
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_temp].[dbo].tbl_wait_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_item_option_Multiple b," &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_item_option_Multiple c" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and b.itemid=c.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T " &VbCRLF
        sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_itemoption.itemid=T.itemid" &VbCRLF
        sqlStr = sqlStr + " and [db_temp].[dbo].tbl_wait_itemoption.itemoption=T.itemoption" &VbCRLF
        sqlStr = sqlStr + " and (" &VbCRLF
        sqlStr = sqlStr + "     [db_temp].[dbo].tbl_wait_itemoption.optionname<>T.optionname" &VbCRLF
        sqlStr = sqlStr + "     or [db_temp].[dbo].tbl_wait_itemoption.optaddprice<>T.optaddprice" &VbCRLF
        sqlStr = sqlStr + "     or [db_temp].[dbo].tbl_wait_itemoption.optaddbuyprice<>T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr
    end if

end function


'' 옵션수정 - 이중옵션
if (mode="editOptionMultiple") then
    ''TypeCnt, OptCnt
    TypeCnt = request("optionTypename").count

    for i=1 to TypeCnt
        optionTypename  = requestCheckVar(Trim(request("optionTypename")(i)),32)
        TypeSeq         = requestCheckVar(Trim(request("TypeSeqTmp")(i)),10)

        sqlStr = "update [db_temp].[dbo].tbl_wait_item_option_Multiple" &VbCRLF
        sqlStr = sqlStr + " set optionTypeName='" + html2Db(optionTypename) + "'" &VbCRLF
        sqlStr = sqlStr + " where itemid=" & CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and TypeSeq=" & CStr(TypeSeq) &VbCRLF
        sqlStr = sqlStr + " and optionTypeName<>'" + html2Db(optionTypename) + "'" &VbCRLF

        dbget.Execute sqlStr

    next

    OptCnt  = request("KindSeq").count
    for i=1 to OptCnt
        TypeSeq     = requestCheckVar(Trim(request("TypeSeq")(i)),10)
        KindSeq     = requestCheckVar(Trim(request("KindSeq")(i)),10)
        optionName  = requestCheckVar(Trim(request("optionName")(i)),96)
        optaddprice = requestCheckVar(Trim(request("optaddprice")(i)),20)
        optaddBuyprice = requestCheckVar(Trim(request("optaddBuyprice")(i)),20)

        if optaddprice < 0  then '추가가격 있는 경우 마이너스 금액은 입력할수 없다 2018-12-11
            response.write "<script language='javascript'>alert('추가금액 판매가에 마이너스 금액을 입력할수 없습니다1. (추가금액이 없으면 0) '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
        if optaddBuyprice < 0  then
            response.write "<script language='javascript'>alert('추가금액 공급가에 마이너스 금액을 입력할수 없습니다1. (추가금액이 없으면 0) '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if

		if 	optaddprice ="" then optaddprice =0
		if optaddBuyprice ="" then optaddBuyprice =0
		if optaddprice >0 and optaddBuyprice =0 then
			 response.write "<script language='javascript'>alert('옵션추가급액의 공급가가 0원입니다.확인해주세요'); </script>"
  			  response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    		dbget.close()	:	response.End
	   end if

        sqlStr = "update [db_temp].[dbo].tbl_wait_item_option_Multiple" &VbCRLF
        sqlStr = sqlStr + " set optionKindName='" + html2Db(optionName) + "'" &VbCRLF
        sqlStr = sqlStr + " ,optaddprice=" & CStr(optaddprice) &VbCRLF
        sqlStr = sqlStr + " ,optaddBuyprice=" & CStr(optaddBuyprice) &VbCRLF
        sqlStr = sqlStr + " where itemid=" & CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and TypeSeq=" & CStr(TypeSeq) &VbCRLF
        sqlStr = sqlStr + " and KindSeq='" & CStr(KindSeq) &"'" &VbCRLF
        sqlStr = sqlStr + " and (" &VbCRLF
        sqlStr = sqlStr + "     (optionKindName<>'" + html2Db(optionName) + "')" &VbCRLF
        sqlStr = sqlStr + "     or (optaddprice<>" & CStr(optaddprice) & ")" &VbCRLF
        sqlStr = sqlStr + "     or (optaddBuyprice<>" & CStr(optaddBuyprice) & ")" &VbCRLF
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr

    next

    Call ReMatchMultiOption(itemid)
    response.write "<script language='javascript'>alert('수정 되었습니다.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if

'' 옵션수정 - 단일옵션
if (mode="editOption") then

	if (request("itemoption").count <> request("optionName").count) or (request("itemoption").count <> request("optaddprice").count) or (request("itemoption").count <> request("optaddBuyprice").count) then
        response.write "<script language='javascript'>alert('비정상적인 접근입니다. (파라미터 갯수 불일치!) '); </script>"
        response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
        dbget.close()	:	response.End
	end if

    sqlStr = "update [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
    sqlStr = sqlStr + " set optionTypeName='" + html2Db(optionTypeName) + "'" &VbCRLF
    sqlStr = sqlStr + " where itemid=" & CStr(itemid) &VbCRLF
    sqlStr = sqlStr + " and optionTypeName<>'" + html2Db(optionTypeName) + "'" &VbCRLF
    dbget.Execute sqlStr

    OptCnt = request("itemoption").count

    for i=1 to OptCnt
        itemoption = requestCheckVar(Trim(request("itemoption")(i)),4)
        optionName = requestCheckVar(Trim(request("optionName")(i)),96)
        optaddprice = requestCheckVar(Trim(request("optaddprice")(i)),20)
        optaddBuyprice = requestCheckVar(Trim(request("optaddBuyprice")(i)),20)

     if optaddBuyprice < 0  then '추가가격 있는 경우 마이너스 금액은 입력할수 없다 2018-12-11
        response.write "<script language='javascript'>alert('추가금액에 마이너스 금액을 입력할수 없습니다1. (추가금액이 없으면 0) '); </script>"
        response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
        dbget.close()	:	response.End
        end if


    if optaddprice ="" then optaddprice =0
   	if optaddBuyprice ="" then optaddBuyprice =0
		if optaddprice >0 and optaddBuyprice =0 then
			 response.write "<script language='javascript'>alert('옵션추가급액의 공급가가 0원입니다.확인해주세요'); </script>"
  			  response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    		dbget.close()	:	response.End
	   end if

        if (Len(itemoption)=4) and (Len(optionName)>0) then
            sqlStr = "update [db_temp].[dbo].tbl_wait_itemoption" &VbCRLF
            sqlStr = sqlStr + " set optionName='" + html2Db(optionName) + "'" &VbCRLF
            sqlStr = sqlStr + " ,optaddprice=" & CStr(optaddprice) &VbCRLF
            sqlStr = sqlStr + " ,optaddBuyprice=" & CStr(optaddBuyprice) &VbCRLF
            sqlStr = sqlStr + " where itemid=" & CStr(itemid) &VbCRLF
            sqlStr = sqlStr + " and itemoption='" & itemoption & "'" &VbCRLF
            sqlStr = sqlStr + " and (" &VbCRLF
            sqlStr = sqlStr + "     (optionName<>'" + html2Db(optionName) + "')" &VbCRLF
            sqlStr = sqlStr + "     or (optaddprice<>" & CStr(optaddprice) & ")" &VbCRLF
            sqlStr = sqlStr + "     or (optaddBuyprice<>" & CStr(optaddBuyprice) & ")" &VbCRLF
            sqlStr = sqlStr + " )"

            dbget.Execute sqlStr

        end if
    next

    response.write "<script language='javascript'>alert('수정 되었습니다.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if


'''옵션 삭제. - 단일옵션
if (mode = "deleteoption") then
	'' 등록 대기 이므로 삭제.

	'// 옵션삭제
	sqlStr = "delete from [db_temp].[dbo].tbl_wait_itemoption" + VbCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(Trim(itemoption)) + "'" + VbCrlf
	rsget.Open sqlStr, dbget, 1

	'// 색상옵션 삭제
	sqlStr = "delete from db_temp.dbo.tbl_wait_item_colorOption" + VbCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(Trim(itemoption)) + "'" + VbCrlf
	rsget.Open sqlStr, dbget, 1

	''상품옵션수량
	sqlStr = "update [db_temp].[dbo].tbl_wait_item" + VBCrlf
	sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
	sqlStr = sqlStr + " 	from [db_temp].[dbo].tbl_wait_itemoption" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_item.itemid=" + CStr(itemid) + VBCrlf
	rsget.Open sqlStr, dbget, 1

	''상품한정수량
	sqlStr = "update [db_temp].[dbo].tbl_wait_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from [db_temp].[dbo].tbl_wait_itemoption" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_item.itemid=" + CStr(itemid) + VBCrlf
	rsget.Open sqlStr, dbget, 1

	response.write "<script language='javascript'>alert('삭제되었습니다.'); </script>"
	response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
	dbget.close()	:	response.End
end if

'' 옵션삭제 - 이중옵션
if (mode = "deleteMultipleOption") then
    'TypeSeq
    'KindSeq

	    '// 이중옵션 정보 삭제
	    sqlStr = "delete from [db_temp].[dbo].tbl_wait_item_option_Multiple" + VbCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	    sqlStr = sqlStr + " and TypeSeq=" + CStr(TypeSeq)
	    sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq) + "'"

	    dbget.Execute sqlStr

		'// 상품 옵션 삭제
		sqlStr = "delete from [db_temp].[dbo].tbl_wait_itemoption" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
		if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,2)='Z" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=2) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and LEFT(RIGHT(itemoption,2),1)='" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=3) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and RIGHT(itemoption,1)='" + CStr(KindSeq) + "'"
    	else
    	    sqlStr = sqlStr + " and 1=0"
    	end if

    	dbget.Execute sqlStr

		'// 색상옵션 삭제(색상은 TypeSeq:1 만 적용)
		sqlStr = "delete from db_temp.dbo.tbl_wait_item_colorOption" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
		if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,2)='Z" + CStr(KindSeq) + "'"
    	else
    	    sqlStr = sqlStr + " and 1=0"
    	end if

    	dbget.Execute sqlStr

    	'' 3차옵션->2차로 변경 or 2차옵션 ->1차로 변경 등..
    	Call ReMatchMultiOption(itemid)

		''상품옵션수량
		sqlStr = "update [db_temp].[dbo].tbl_wait_item" + VBCrlf
		sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
		sqlStr = sqlStr + " from (" + VBCrlf
		sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
		sqlStr = sqlStr + " 	from [db_temp].[dbo].tbl_wait_itemoption" + VBCrlf
		sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
		sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
		sqlStr = sqlStr + " ) T" + VBCrlf
		sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_item.itemid=" + CStr(itemid) + VBCrlf
		rsget.Open sqlStr, dbget, 1

		''상품한정수량
		sqlStr = "update [db_temp].[dbo].tbl_wait_item" + VBCrlf
		sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
		sqlStr = sqlStr + " from (" + VBCrlf
		sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
		sqlStr = sqlStr + " 	from [db_temp].[dbo].tbl_wait_itemoption" + VBCrlf
		sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
		sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
		sqlStr = sqlStr + " ) T" + VBCrlf
		sqlStr = sqlStr + " where [db_temp].[dbo].tbl_wait_item.itemid=" + CStr(itemid) + VBCrlf
		rsget.Open sqlStr, dbget, 1

		response.write "<script language='javascript'>alert('삭제되었습니다.'); </script>"
		response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End
end if


'' 단일 옵션 추가
if (mode = "addoptionCustom") then
    foundcount = 0

    for i = 0 to ubound(arritemoption)
        if (Trim(arritemoption(i)) <> "") then
            sqlStr = " select itemid from [db_temp].[dbo].tbl_wait_itemoption "
            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and ((itemoption = '" + CStr(requestCheckVar(Trim(arritemoption(i)),4)) + "') or (optionname='" + requestCheckVar(html2db(arritemoptionname(i)),96) + "'))"

            rsget.Open sqlStr,dbget,1

            if not rsget.EOF then
                found = true
                foundcount = foundcount + 1
            else
                found = false
            end if
            rsget.close

            ''한정 구분은 상품 한정 구분과 동일
            if (found = false) then
                sqlStr = " insert into [db_temp].[dbo].tbl_wait_itemoption(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(requestCheckVar(arritemoption(i),4)) + "', '" + html2db(optionTypename) + "', '" + CStr(requestCheckVar(html2db(arritemoptionname(i)),96)) + "', 'Y', 'Y', '" + itemLimitYn + "', 0, 0) "

                dbget.Execute sqlStr

            end if
        end if
    next

    ''옵션 구분명은 동일

    sqlStr = " update [db_temp].[dbo].tbl_wait_itemoption " &VbCRLF
    sqlStr = sqlStr + " set optionTypeName='" + html2db(optionTypename) + "'" &VbCRLF
    sqlStr = sqlStr + " where itemid=" + cStr(itemid) &VbCRLF
    sqlStr = sqlStr + " and optionTypeName<>'" + html2db(optionTypename) + "'" &VbCRLF

    dbget.Execute sqlStr

    sqlStr = " update [db_temp].[dbo].tbl_wait_item " &VbCRLF
    sqlStr = sqlStr + " set optioncnt = IsNULL(T.cnt,0) " &VbCRLF
    sqlStr = sqlStr + " from " &VbCRLF
    sqlStr = sqlStr + " ( " &VbCRLF
    sqlStr = sqlStr + "     select count(itemid) as cnt " &VbCRLF
    sqlStr = sqlStr + "     from [db_temp].[dbo].tbl_wait_itemoption where itemid = " + CStr(itemid) + " and isusing = 'Y' " &VbCRLF
    sqlStr = sqlStr + " ) T " &VbCRLF
    sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "

    dbget.Execute sqlStr

    if (foundcount > 0) then
        response.write "<script>alert('일부 옵션은 기존에 있는 옵션과 중복되어 무시되었습니다.');</script>"
    end if

    response.write "<script>alert('옵션이 추가되었습니다.'); opener.location.reload(); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if




''이중옵션 추가
if (mode="addDoubleOption") then

    dim optionTypename1, optionTypename2, optionTypename3
    dim itemoption1, itemoption2, itemoption3
    dim optionName1, optionName2, optionName3
    dim retErrMsg, chkbyteStr, chkSplitChar : chkSplitChar="!@@!@@!"

    optionTypename1 = requestCheckVar(Trim(request("optionTypename1")),32)
    optionTypename2 = requestCheckVar(Trim(request("optionTypename2")),32)
    optionTypename3 = requestCheckVar(Trim(request("optionTypename3")),32)
    itemoption1     = requestCheckVar(request("itemoption1"),4)
    itemoption2     = requestCheckVar(request("itemoption2"),4)
    itemoption3     = requestCheckVar(request("itemoption3"),4)
    optionName1     = requestCheckVar(request("optionName1"),96)
    optionName2     = requestCheckVar(request("optionName2"),96)
    optionName3     = requestCheckVar(request("optionName3"),96)

    dim Lv1cnt, Lv2cnt, Lv3cnt
    dim Val1cnt, Val2cnt, Val3cnt
    dim option1, option2, option3
    dim optName1, optName2, optName3
    dim Valid1, Valid2, Valid3
    dim buf, ErrMsg, AssignedOption

    Lv1cnt = request("itemoption1").count
    Lv2cnt = request("itemoption2").count
    Lv3cnt = request("itemoption3").count

    Val1cnt = 0
    Val2cnt = 0
    Val3cnt = 0

    '' 체크 조건추가 //2016/04/19
    if ((Lv1cnt>35) or (Lv2cnt>35) or (Lv3cnt>35)) then
        ErrMsg = ErrMsg & "이중옵션은 옵션구분당 최대 35개까지 가능합니다.\n"
    end if

    chkbyteStr = ""
    for i=1 to Lv1cnt
        buf = requestCheckVar(Trim(request("optionName1")(i)),96)
        if Len(buf)>0 then Val1cnt = Val1cnt + 1

        if Len(buf)>0 then
            chkbyteStr = chkbyteStr&buf&chkSplitChar
        end if
    next

    if Right(chkbyteStr,LEN(chkSplitChar))=chkSplitChar then chkbyteStr=Left(chkbyteStr,Len(chkbyteStr)-LEN(chkSplitChar))
    if (NOT checkDataLengthDBArr(chkbyteStr,chkSplitChar,32,retErrMsg)) then
        ErrMsg = ErrMsg & retErrMsg
    end if

    chkbyteStr = ""
    for i=1 to Lv2cnt
        buf = requestCheckVar(Trim(request("optionName2")(i)),96)
        if Len(buf)>0 then Val2cnt = Val2cnt + 1

        if Len(buf)>0 then
            chkbyteStr = chkbyteStr&buf&chkSplitChar
        end if
    next

    if Right(chkbyteStr,LEN(chkSplitChar))=chkSplitChar then chkbyteStr=Left(chkbyteStr,Len(chkbyteStr)-LEN(chkSplitChar))
    if (NOT checkDataLengthDBArr(chkbyteStr,chkSplitChar,32,retErrMsg)) then
        ErrMsg = ErrMsg & retErrMsg
    end if

    chkbyteStr = ""
    for i=1 to Lv3cnt
        buf = requestCheckVar(Trim(request("optionName3")(i)),96)
        if Len(buf)>0 then Val3cnt = Val3cnt + 1

        if Len(buf)>0 then
            chkbyteStr = chkbyteStr&buf&chkSplitChar
        end if
    next

    if Right(chkbyteStr,LEN(chkSplitChar))=chkSplitChar then chkbyteStr=Left(chkbyteStr,Len(chkbyteStr)-LEN(chkSplitChar))
    if (NOT checkDataLengthDBArr(chkbyteStr,chkSplitChar,32,retErrMsg)) then
        ErrMsg = ErrMsg & retErrMsg
    end if

    if (optionTypename1=optionTypename2) or (optionTypename1=optionTypename3) or (optionTypename2=optionTypename3) then
        ErrMsg = ErrMsg & "옵션구분명이 동일할 수 없습니다.\n"
    end if

    if (Len(optionTypename1)<1) and (Len(optionTypename2)<1) and (Len(optionTypename3)<1) then
        ErrMsg = ErrMsg & "옵션구분명이 입력되지 않았습니다.\n"
    end if

    if (Val1cnt>0) and (Len(optionTypename1)<1) then
        ErrMsg = ErrMsg & "옵션구분명1이 입력되지 않았습니다.\n"
    end if

    if (Val2cnt>0) and (Len(optionTypename2)<1) then
        ErrMsg = ErrMsg & "옵션구분명2이 입력되지 않았습니다.\n"
    end if

    if (Val3cnt>0) and (Len(optionTypename3)<1) then
        ErrMsg = ErrMsg & "옵션구분명3이 입력되지 않았습니다.\n"
    end if

    if (Val1cnt<1) and (Len(optionTypename1)>0) then
        ErrMsg = ErrMsg & "옵션구분명1에 대한 옵션이 입력되지 않았습니다.\n"
    end if

    if (Val2cnt<1) and (Len(optionTypename2)>0) then
        ErrMsg = ErrMsg & "옵션구분명2에 대한 옵션이 입력되지 않았습니다.\n"
    end if

    if (Val3cnt<1) and (Len(optionTypename3)>0) then
        ErrMsg = ErrMsg & "옵션구분명3에 대한 옵션이 입력되지 않았습니다.\n"
    end if

    if ((Val1cnt<1) and (Val2cnt<1)) or ((Val2cnt<1) and (Val3cnt<1)) or ((Val1cnt<1) and (Val3cnt<1)) then
        ErrMsg = ErrMsg & "이중옵션으로 등록 하시려면 옵션구분은 2개 이상 등록하셔야 합니다.\n"
    end if

    ''순서대로 입력해야 가능
    if ((Val1cnt<1) or (Val2cnt<1)) then
        ErrMsg = ErrMsg & "이중옵션으로 등록 하시려면 옵션구분 1 부터 구분을 2개 이상 등록하셔야 합니다.\n"
    end if

    if (Len(ErrMsg)>0) then
        response.write "<script>alert('" + replace(ErrMsg,"'","") + "'); history.back();</script>"
        dbget.close()	:	response.End
    end if

    foundcount=0


    ''0번은 입력 없음. N까지
    for i=0 to Lv1cnt-1
        for j=0 to Lv2cnt-1
            for k=0 to Lv3cnt-1
                optName1 = requestCheckVar(Trim(request("optionName1")(i+1)),96)
                optName2 = requestCheckVar(Trim(request("optionName2")(j+1)),96)
                optName3 = requestCheckVar(Trim(request("optionName3")(k+1)),96)

                option1  = requestCheckVar(Trim(request("itemoption1")(i+1)),4)
                option2  = requestCheckVar(Trim(request("itemoption2")(j+1)),4)
                option3  = requestCheckVar(Trim(request("itemoption3")(k+1)),4)

                Valid1  = (Len(optionTypename1)>0) and (Len(optName1)>0)
                Valid2  = (Len(optionTypename2)>0) and (Len(optName2)>0)
                Valid3  = (Len(optionTypename3)>0) and (Len(optName3)>0)

                AssignedOption = "Z"
                if (Not Valid1) then
                    AssignedOption = AssignedOption + "0"
                else
                    AssignedOption = AssignedOption + optKindSeq2Code(CStr(i+1))
                end if

                if (Not Valid2) then
                    AssignedOption = AssignedOption + "0"
                else
                    AssignedOption = AssignedOption + optKindSeq2Code(CStr(j+1))
                end if

                if (Not Valid3) then
                    AssignedOption = AssignedOption + "0"
                else
                    AssignedOption = AssignedOption + optKindSeq2Code(CStr(k+1))
                end if

                optionName = optName1 + "," + optName2 + "," + optName3
                ''콤마제거
                optionName = Replace(optionName,",,",",")
                if Right(optionName,1)="," then optionName=Left(optionName,Len(optionName)-1)

                if (Valid1 and Valid2) or (Valid1 and Valid3) or (Valid2 and Valid3) then
                    if ((i=0) or (Valid1)) and  ((j=0) or (Valid2)) and ((k=0) or (Valid3))  then
                        ''같은 옵션이 존재하는지 Check.

                        found = false

                        if (Len(option1)<1) and (Len(optName1)>0) and (Len(optionTypename1)>0) then
                            sqlStr = " select itemid from [db_temp].[dbo].tbl_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                            sqlStr = sqlStr + " and ((optionTypeName=convert(varchar(32),'" + html2db(optionTypename1) + "') and optionKindName=convert(varchar(32),'" + html2db(optName1) + "')))"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsget.Close

                            if (Not found) then
                                sqlStr = " insert into [db_temp].[dbo].tbl_wait_item_option_Multiple" &VbCRLF
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)" &VbCRLF
                                sqlStr = sqlStr + " values(" &VbCRLF
                                sqlStr = sqlStr + " " & itemid &VbCRLF
                                sqlStr = sqlStr + " ,1" &VbCRLF
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(i+1)) & "'" &VbCRLF
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename1) & "')" &VbCRLF
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName1) & "')" &VbCRLF
                                sqlStr = sqlStr + " )" &VbCRLF

                                dbget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName1 + "," + optionTypename1 + "," + CStr(1) + "," + CStr(i+1) + "<br>"
                            end if
                        end if

                        found = false
                        if (Len(option2)<1) and (Len(optName2)>0) and (Len(optionTypename2)>0) then
                            sqlStr = " select itemid from [db_temp].[dbo].tbl_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                            sqlStr = sqlStr + " and ((optionTypeName=convert(varchar(32),'" + html2db(optionTypename2) + "') and optionKindName=convert(varchar(32),'" + html2db(optName2) + "')))"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsget.Close

                            if (Not found) then
                                sqlStr = " insert into [db_temp].[dbo].tbl_wait_item_option_Multiple" &VbCRLF
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)" &VbCRLF
                                sqlStr = sqlStr + " values(" &VbCRLF
                                sqlStr = sqlStr + " " & itemid &VbCRLF
                                sqlStr = sqlStr + " ,2" &VbCRLF
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(j+1))&"'" &VbCRLF
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename2) & "')" &VbCRLF
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName2) & "')" &VbCRLF        '' convert(varchar(32) 추가..
                                sqlStr = sqlStr + " )"

                                dbget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName2 + "," + optionTypename2 + "," + CStr(2) + "," + CStr(j+1) + "<br>"
                            end if
                        end if

                        found = false
                        if (Len(option3)<1) and (Len(optName3)>0) and (Len(optionTypename3)>0) then
                            sqlStr = " select itemid from [db_temp].[dbo].tbl_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                            sqlStr = sqlStr + " and ((optionTypeName=convert(varchar(32),'" + html2db(optionTypename3) + "') and optionKindName=convert(varchar(32),'" + html2db(optName3) + "')))"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsget.Close

                            if (Not found) then
                                sqlStr = " insert into [db_temp].[dbo].tbl_wait_item_option_Multiple" &VbCRLF
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)" &VbCRLF
                                sqlStr = sqlStr + " values(" &VbCRLF
                                sqlStr = sqlStr + " " & itemid &VbCRLF
                                sqlStr = sqlStr + " ,3" &VbCRLF
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(k+1))&"'" &VbCRLF
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename3) & "')" &VbCRLF
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName3) & "')" &VbCRLF
                                sqlStr = sqlStr + " )"

                                dbget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName3 + "," + optionTypename3 + "," + CStr(3) + "," + CStr(k+1) + "<br>"
                            end if
                        end if

                        found = false
                        sqlStr = " select itemid from [db_temp].[dbo].tbl_wait_itemoption "
                        sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                        sqlStr = sqlStr + " and ((itemoption = '" + CStr(AssignedOption) + "') or (optionTypeName='복합옵션' and optionname=convert(varchar(96),'" + html2db(optionName) + "')))"

                        rsget.Open sqlStr,dbget,1
                        if Not rsget.Eof then
                            found = true
                        end if
                        rsget.Close

                        if (Not found) then
                            sqlStr = " insert into [db_temp].[dbo].tbl_wait_itemoption(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                            sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(AssignedOption) + "', '복합옵션', convert(varchar(96),'" + CStr(html2db(optionName)) + "'), 'Y', 'Y', '" + itemLimitYn + "', 0, 0) "

                            dbget.Execute sqlStr
                            ''response.write AssignedOption + ":" +  optName1 + "," + optName2 + "," + optName3 + "<BR>"
                        end if
                    end if

                end if
            next
        next
    next

    '' 2차옵션->3차로 변경  등..
    Call ReMatchMultiOption(itemid)


    ''옵션 총수 저장
    sqlStr = " update [db_temp].[dbo].tbl_wait_item " &VbCRLF
    sqlStr = sqlStr + " set optioncnt = IsNULL(T.cnt,0) " &VbCRLF
    sqlStr = sqlStr + " from " &VbCRLF
    sqlStr = sqlStr + " ( " &VbCRLF
    sqlStr = sqlStr + "     select count(itemid) as cnt " &VbCRLF
    sqlStr = sqlStr + "     from [db_temp].[dbo].tbl_wait_itemoption where itemid = " + CStr(itemid) + " and isusing = 'Y' " &VbCRLF
    sqlStr = sqlStr + " ) T " &VbCRLF
    sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "

    dbget.Execute sqlStr

    response.write "<script>alert('옵션이 추가되었습니다.'); opener.location.reload(); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

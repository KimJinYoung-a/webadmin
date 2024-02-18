<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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

itemid              = RequestCheckvar(request("itemid"),10)
itemoption          = RequestCheckvar(request("itemoption"),10)
mode                = RequestCheckvar(request("mode"),32)
optionTypename      = RequestCheckvar(request("optionTypename"),32)
arritemoption       = request("arritemoption")
arritemoptionname   = request("arritemoptionname")

TypeSeq             = RequestCheckvar(request("TypeSeq"),10)
KindSeq             = RequestCheckvar(request("KindSeq"),10)

arritemoption = Split(arritemoption, "|")
arritemoptionname = Split(arritemoptionname, "|")

dim IsUpchebeasong, itemLimitYn
IsUpchebeasong =false
itemLimitYn = "N"

''업체배송인경우 입출/판매 관계없이 삭제
sqlStr = " select limityn, deliverytype "
sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item"
sqlStr = sqlStr & " where itemid=" & CStr(itemid)

rsACADEMYget.Open sqlStr,dbACADEMYget,1
if not rsACADEMYget.EOF then
    itemLimitYn = rsACADEMYget("limityn")
    IsUpchebeasong = (rsACADEMYget("deliverytype") = "2") or (rsACADEMYget("deliverytype") = "5")
end if
rsACADEMYget.Close


function ReMatchMultiOption(itemid)
    dim sqlStr
    dim MultiLevel
    
    MultiLevel = 0
    
    sqlStr = " select TypeSeq, Count(KindSeq) as KindCnt "
    sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    sqlStr = sqlStr + " group by TypeSeq"
    sqlStr = sqlStr + " order by TypeSeq"
    
    rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	    MultiLevel = rsACADEMYget.RecordCount
	rsACADEMYget.close
    	
    ''기존 2차 옵션인 경우 삭제.
    if (MultiLevel=3) then 
        sqlStr = " delete from db_academy.dbo.tbl_diy_wait_item_option"
        sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'"
        sqlStr = sqlStr + " and Right(itemoption,1)='0'"
        
        dbACADEMYget.Execute sqlStr
    end if
    
    if (MultiLevel=2) then 
        sqlStr = " delete from db_academy.dbo.tbl_diy_wait_item_option"
        sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'"
        sqlStr = sqlStr + " and Right(itemoption,1)='00'"
        
        dbACADEMYget.Execute sqlStr
    end if 
    
    
    ''옵션 재작성.
'   --Only 1중옵션.
    if (MultiLevel=1) then 
        ''-- 전 옵션 삭제;
        sqlStr = " delete from db_academy.dbo.tbl_diy_wait_item_option_Multiple" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        dbACADEMYget.Execute sqlStr
        
        sqlStr = " delete from db_academy.dbo.tbl_diy_wait_item_option" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        sqlStr = sqlStr & " and Left(itemoption,1)='Z'"
        dbACADEMYget.Execute sqlStr
        
    end if
    
'   --Only 2중옵션.
    if (MultiLevel=2) then 
        sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option"
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '복합옵션' as optionTypeName,"
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0,"
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_wait_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option_Multiple b"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + " ) T"
        sqlStr = sqlStr + "     left join db_academy.dbo.tbl_diy_wait_item_option o "
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and T.itemid=o.itemid "
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption "
        sqlStr = sqlStr + " where  o.itemid is NULL"
    
        dbACADEMYget.Execute sqlStr
        
        '' 옵션명/ 가격 등이 변경된 경우
        sqlStr = " update db_academy.dbo.tbl_diy_wait_item_option"
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)"
        sqlStr = sqlStr + " , optaddprice=T.optaddprice"
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName ) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_wait_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option_Multiple b"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_wait_item_option.itemid=T.itemid"
        sqlStr = sqlStr + " and db_academy.dbo.tbl_diy_wait_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and ("
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option.optionname<>T.optionname"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_wait_item_option.optaddprice<>T.optaddprice"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_wait_item_option.optaddbuyprice<>T.optaddbuyprice"
        sqlStr = sqlStr + " )"

        dbACADEMYget.Execute sqlStr
    end if

'    --Only 3중옵션
    if (MultiLevel=3) then 
        sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option"
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '복합옵션' as optionTypeName,"
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0,"
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_wait_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option_Multiple b,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option_Multiple c"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and b.itemid=c.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + "     left join db_academy.dbo.tbl_diy_wait_item_option o "
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and T.itemid=o.itemid "
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption "
        sqlStr = sqlStr + " where  o.itemid is NULL"
        
        dbACADEMYget.Execute sqlStr
        
        
        '' 옵션명/ 가격 등이 변경된 경우
        sqlStr = " update db_academy.dbo.tbl_diy_wait_item_option"
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)"
        sqlStr = sqlStr + " , optaddprice=T.optaddprice"
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_wait_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option_Multiple b,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option_Multiple c"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and b.itemid=c.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_wait_item_option.itemid=T.itemid"
        sqlStr = sqlStr + " and db_academy.dbo.tbl_diy_wait_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and ("
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_wait_item_option.optionname<>T.optionname"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_wait_item_option.optaddprice<>T.optaddprice"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_wait_item_option.optaddbuyprice<>T.optaddbuyprice"
        sqlStr = sqlStr + " )"

        dbACADEMYget.Execute sqlStr
    end if
    
end function


'' 옵션수정 - 이중옵션
if (mode="editOptionMultiple") then
    ''TypeCnt, OptCnt
    TypeCnt = request("optionTypename").count 
 
    for i=1 to TypeCnt
        optionTypename  = Trim(request("optionTypename")(i))
        TypeSeq         = Trim(request("TypeSeqTmp")(i))
        
        sqlStr = "update db_academy.dbo.tbl_diy_wait_item_option_Multiple"
        sqlStr = sqlStr + " set optionTypeName='" + html2Db(optionTypename) + "'"
        sqlStr = sqlStr + " where itemid=" & CStr(itemid)
        sqlStr = sqlStr + " and TypeSeq=" & CStr(TypeSeq)
        sqlStr = sqlStr + " and optionTypeName<>'" + html2Db(optionTypename) + "'"
        
        dbACADEMYget.Execute sqlStr
        
    next
    
    OptCnt  = request("KindSeq").count
    for i=1 to OptCnt
        TypeSeq     = Trim(request("TypeSeq")(i))
        KindSeq     = Trim(request("KindSeq")(i))
        optionName  = Trim(request("optionName")(i))
        optaddprice = Trim(request("optaddprice")(i))
        optaddBuyprice = Trim(request("optaddBuyprice")(i))
        
        sqlStr = "update db_academy.dbo.tbl_diy_wait_item_option_Multiple"
        sqlStr = sqlStr + " set optionKindName='" + html2Db(optionName) + "'"
        sqlStr = sqlStr + " ,optaddprice=" & CStr(optaddprice)
        sqlStr = sqlStr + " ,optaddBuyprice=" & CStr(optaddBuyprice)
        sqlStr = sqlStr + " where itemid=" & CStr(itemid)
        sqlStr = sqlStr + " and TypeSeq=" & CStr(TypeSeq)
        sqlStr = sqlStr + " and KindSeq='" & CStr(KindSeq) &"'"
        sqlStr = sqlStr + " and ("
        sqlStr = sqlStr + "     (optionKindName<>'" + html2Db(optionName) + "')"
        sqlStr = sqlStr + "     or (optaddprice<>" & CStr(optaddprice) & ")"
        sqlStr = sqlStr + "     or (optaddBuyprice<>" & CStr(optaddBuyprice) & ")"
        sqlStr = sqlStr + " )"
            
        dbACADEMYget.Execute sqlStr
        
    next
    
    Call ReMatchMultiOption(itemid)
    response.write "<script language='javascript'>alert('수정 되었습니다.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbACADEMYget.close()	:	response.End
end if

'' 옵션수정 - 단일옵션
if (mode="editOption") then
    
    sqlStr = "update db_academy.dbo.tbl_diy_wait_item_option"
    sqlStr = sqlStr + " set optionTypeName=convert(varchar(32),'" + html2Db(optionTypeName) + "')"
    sqlStr = sqlStr + " where itemid=" & CStr(itemid)
    sqlStr = sqlStr + " and optionTypeName<>'" + html2Db(optionTypeName) + "'"
    dbACADEMYget.Execute sqlStr
    
    OptCnt = request("itemoption").count
    
    for i=1 to OptCnt
        itemoption = Trim(request("itemoption")(i))
        optionName = Trim(request("optionName")(i))
        optaddprice = Trim(request("optaddprice")(i))
        optaddBuyprice = Trim(request("optaddBuyprice")(i))
        
        if (Len(itemoption)=4) and (Len(optionName)>0) then
            sqlStr = "update db_academy.dbo.tbl_diy_wait_item_option"
            sqlStr = sqlStr + " set optionName=convert(varchar(96),'" + html2Db(optionName) + "')"
            sqlStr = sqlStr + " ,optaddprice=" & CStr(optaddprice)
            sqlStr = sqlStr + " ,optaddBuyprice=" & CStr(optaddBuyprice)
            sqlStr = sqlStr + " where itemid=" & CStr(itemid)
            sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
            sqlStr = sqlStr + " and ("
            sqlStr = sqlStr + "     (optionName<>'" + html2Db(optionName) + "')"
            sqlStr = sqlStr + "     or (optaddprice<>" & CStr(optaddprice) & ")"
            sqlStr = sqlStr + "     or (optaddBuyprice<>" & CStr(optaddBuyprice) & ")"
            sqlStr = sqlStr + " )"
            
            dbACADEMYget.Execute sqlStr
            
        end if
    next
    
    response.write "<script language='javascript'>alert('수정 되었습니다.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbACADEMYget.close()	:	response.End
end if

    
'''옵션 삭제. - 단일옵션
if (mode = "deleteoption") then
	'' 등록 대기 이므로 삭제.


	
	sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option" + VbCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	sqlStr = sqlStr + " and itemoption='" + CStr(Trim(itemoption)) + "'" + VbCrlf
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	''상품옵션수량
	sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + VBCrlf
	sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
	sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_wait_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_wait_item.itemid=" + CStr(itemid) + VBCrlf
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	''상품한정수량
	sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_wait_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_wait_item.itemid=" + CStr(itemid) + VBCrlf
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	response.write "<script language='javascript'>alert('삭제되었습니다.'); </script>"
	response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
	dbACADEMYget.close()	:	response.End
end if

'' 옵션삭제 - 이중옵션
if (mode = "deleteMultipleOption") then 
    'TypeSeq
    'KindSeq
    
	    sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option_Multiple" + VbCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	    sqlStr = sqlStr + " and TypeSeq=" + CStr(TypeSeq)
	    sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq) + "'"
	    
	    dbACADEMYget.Execute sqlStr
	    
		sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
		if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,2)='Z" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=2) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and LEFT(RIGHT(itemoption,3),1)='" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=3) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and RIGHT(itemoption,1)='" + CStr(KindSeq) + "'"
    	else
    	    sqlStr = sqlStr + " and 1=0"
    	end if

    	dbACADEMYget.Execute sqlStr
    	
    	'' 3차옵션->2차로 변경 or 2차옵션 ->1차로 변경 등..
    	Call ReMatchMultiOption(itemid)

		''상품옵션수량
		sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + VBCrlf
		sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
		sqlStr = sqlStr + " from (" + VBCrlf
		sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
		sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_wait_item_option" + VBCrlf
		sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
		sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
		sqlStr = sqlStr + " ) T" + VBCrlf
		sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_wait_item.itemid=" + CStr(itemid) + VBCrlf
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		''상품한정수량
		sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + VBCrlf
		sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
		sqlStr = sqlStr + " from (" + VBCrlf
		sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
		sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_wait_item_option" + VBCrlf
		sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
		sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
		sqlStr = sqlStr + " ) T" + VBCrlf
		sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_wait_item.itemid=" + CStr(itemid) + VBCrlf
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		response.write "<script language='javascript'>alert('삭제되었습니다.'); </script>"
		response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
		dbACADEMYget.close()	:	response.End
end if


'' 단일 옵션 추가
if (mode = "addoptionCustom") then
    foundcount = 0
    
    for i = 0 to ubound(arritemoption)
        if (Trim(arritemoption(i)) <> "") then
            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option "
            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
            sqlStr = sqlStr + " and ((itemoption = '" + CStr(Trim(arritemoption(i))) + "') or (optionname='" + html2db(arritemoptionname(i)) + "'))"
            
            rsACADEMYget.Open sqlStr,dbACADEMYget,1

            if not rsACADEMYget.EOF then
                found = true
                foundcount = foundcount + 1
            else
                found = false
            end if
            rsACADEMYget.close
            
            ''한정 구분은 상품 한정 구분과 동일
            if (found = false) then
                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(arritemoption(i)) + "', convert(varchar(32),'" + html2db(optionTypename) + "'), convert(varchar(96),'" + CStr(html2db(arritemoptionname(i))) + "'), 'Y', 'Y', '" + itemLimitYn + "', 0, 0) "
                
                dbACADEMYget.Execute sqlStr

            end if
        end if
    next
    
    ''옵션 구분명은 동일
    
    sqlStr = " update db_academy.dbo.tbl_diy_wait_item_option "
    sqlStr = sqlStr + " set optionTypeName=convert(varchar(32),'" + html2db(optionTypename) + "')"
    sqlStr = sqlStr + " where itemid=" + cStr(itemid)
    sqlStr = sqlStr + " and optionTypeName<>'" + html2db(optionTypename) + "'"
    
    dbACADEMYget.Execute sqlStr
    
    sqlStr = " update db_academy.dbo.tbl_diy_wait_item "
    sqlStr = sqlStr + " set optioncnt = IsNULL(T.cnt,0) "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " ( "
    sqlStr = sqlStr + "     select count(itemid) as cnt "
    sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_wait_item_option where itemid = " + CStr(itemid) + " and isusing = 'Y' "
    sqlStr = sqlStr + " ) T "
    sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
    
    dbACADEMYget.Execute sqlStr

    if (foundcount > 0) then
        response.write "<script>alert('일부 옵션은 기존에 있는 옵션과 중복되어 무시되었습니다.');</script>"
    end if

    response.write "<script>alert('옵션이 추가되었습니다.'); opener.location.reload(); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbACADEMYget.close()	:	response.End
end if




''이중옵션 추가
if (mode="addDoubleOption") then
    
    dim optionTypename1, optionTypename2, optionTypename3
    dim itemoption1, itemoption2, itemoption3
    dim optionName1, optionName2, optionName3
    
    optionTypename1 = Trim(request("optionTypename1"))
    optionTypename2 = Trim(request("optionTypename2"))
    optionTypename3 = Trim(request("optionTypename3"))
    itemoption1     = request("itemoption1")
    itemoption2     = request("itemoption2")
    itemoption3     = request("itemoption3")
    optionName1     = request("optionName1")
    optionName2     = request("optionName2")
    optionName3     = request("optionName3")
    
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
    
    for i=1 to Lv1cnt
        buf = Trim(request("optionName1")(i))
        if Len(buf)>0 then Val1cnt = Val1cnt + 1
    next
    
    for i=1 to Lv2cnt
        buf = Trim(request("optionName2")(i))
        if Len(buf)>0 then Val2cnt = Val2cnt + 1
    next
    
    for i=1 to Lv3cnt
        buf = Trim(request("optionName3")(i))
        if Len(buf)>0 then Val3cnt = Val3cnt + 1
    next
    
    if (optionTypename1=optionTypename2) or (optionTypename1=optionTypename3) or (optionTypename2=optionTypename3) then
        ErrMsg = "옵션구분명이 동일할 수 없습니다.\n"
    end if
    
    if (Len(optionTypename1)<1) and (Len(optionTypename2)<1) and (Len(optionTypename3)<1) then
        ErrMsg = "옵션구분명이 입력되지 않았습니다.\n"
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
        response.write "<script>alert('" + ErrMsg + "'); history.back();</script>"
        dbACADEMYget.close()	:	response.End
    end if
    
    foundcount=0


    ''0번은 입력 없음. N까지
    for i=0 to Lv1cnt-1
        for j=0 to Lv2cnt-1
            for k=0 to Lv3cnt-1
                optName1 = Trim(request("optionName1")(i+1))
                optName2 = Trim(request("optionName2")(j+1)) 
                optName3 = Trim(request("optionName3")(k+1))
                
                option1  = Trim(request("itemoption1")(i+1))
                option2  = Trim(request("itemoption2")(j+1))
                option3  = Trim(request("itemoption3")(k+1))
                
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
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and ((optionTypeName=convert(varchar(32),'" + html2db(optionTypename1) + "') and optionKindName=convert(varchar(32),'" + html2db(optName1) + "')))"

                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,1"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(i+1)) & "'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename1) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName1) & "')"
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName1 + "," + optionTypename1 + "," + CStr(1) + "," + CStr(i+1) + "<br>"
                            end if
                        end if
                        
                        found = false
                        if (Len(option2)<1) and (Len(optName2)>0) and (Len(optionTypename2)>0) then
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and ((optionTypeName=convert(varchar(32),'" + html2db(optionTypename2) + "') and optionKindName=convert(varchar(32),'" + html2db(optName2) + "')))"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,2"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(j+1))&"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename2) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName2) & "')"
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName2 + "," + optionTypename2 + "," + CStr(2) + "," + CStr(j+1) + "<br>"
                            end if
                        end if
                        
                        found = false
                        if (Len(option3)<1) and (Len(optName3)>0) and (Len(optionTypename3)>0) then
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and ((optionTypeName=convert(varchar(32),'" + html2db(optionTypename3) + "') and optionKindName=convert(varchar(32),'" + html2db(optName3) + "')))"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,3"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(k+1))&"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename3) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName3) & "')"
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName3 + "," + optionTypename3 + "," + CStr(3) + "," + CStr(k+1) + "<br>"
                            end if
                        end if
                        
                        found = false
                        sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option "
                        sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                        sqlStr = sqlStr + " and ((itemoption = '" + CStr(AssignedOption) + "') or (optionTypeName='복합옵션' and optionname=convert(varchar(96),'" + html2db(optionName) + "')))"
                        
                        rsACADEMYget.Open sqlStr,dbACADEMYget,1
                        if Not rsACADEMYget.Eof then
                            found = true
                        end if
                        rsACADEMYget.Close
                        
                        if (Not found) then 
                            sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                            sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(AssignedOption) + "', '복합옵션', convert(varchar(96),'" + CStr(html2db(optionName)) + "'), 'Y', 'Y', '" + itemLimitYn + "', 0, 0) "
                            
                            dbACADEMYget.Execute sqlStr
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
    sqlStr = " update db_academy.dbo.tbl_diy_wait_item "
    sqlStr = sqlStr + " set optioncnt = IsNULL(T.cnt,0) "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " ( "
    sqlStr = sqlStr + "     select count(itemid) as cnt "
    sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_wait_item_option where itemid = " + CStr(itemid) + " and isusing = 'Y' "
    sqlStr = sqlStr + " ) T "
    sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
    
    dbACADEMYget.Execute sqlStr
     
    response.write "<script>alert('옵션이 추가되었습니다.'); opener.location.reload(); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbACADEMYget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
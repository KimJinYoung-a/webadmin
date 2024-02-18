<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
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

function checkIsUpcheItemEditValid(itemid)
    dim sqlStr
    checkIsUpcheItemEditValid = false
    if (C_ADMIN_USER) then
        checkIsUpcheItemEditValid = True
        Exit function
    end if

    if (C_IS_Maker_Upche) and (request.Cookies("partner")("userid")<>"") then
        sqlStr = " select top 1 itemid from db_academy.dbo.tbl_diy_item " &VBCRLF
        sqlStr = sqlStr & " where itemid="&itemid&VBCRLF
        sqlStr = sqlStr & " and makerid='"&request.Cookies("partner")("userid")&"'"&VBCRLF

        rsACADEMYget.Open sqlStr,dbACADEMYget,1
        if Not rsACADEMYget.Eof then
    	    checkIsUpcheItemEditValid = (rsACADEMYget.RecordCount>0)
    	end if
    	rsACADEMYget.Close
    end if

end function

Dim itemid, mode, itemoption, optionTypename
Dim arritemoption
Dim sqlStr, ErrStr, foundcount, i, j, k, found
dim C_IS_Maker_Upche, C_ADMIN_USER
dim TypeSeq, KindSeq
dim optionTypename1, optionTypename2, optionTypename3
dim Lv1cnt, Lv2cnt, Lv3cnt
dim Val1cnt, Val2cnt, Val3cnt
dim optName1, optName2, optName3
dim Valid1, Valid2, Valid3
dim buf, ErrMsg, AssignedOption, ArrCnt
Dim optprice1, optprice2, optprice3, makerid
Dim optbuyprice1, optbuyprice2, optbuyprice3, optionName
Dim TypeSeq1, KindSeq1, TypeSeq2, KindSeq2, TypeSeq3, KindSeq3

C_IS_Maker_Upche = (request.Cookies("partner")("userdiv") = "9999")
C_ADMIN_USER = request.Cookies("partner")("userdiv")

itemid = Request("itemid")
makerid = request.cookies("partner")("userid")

'Response.write makerid
'Response.end

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If

mode = Request.form("mode")
itemoption = Request("itemoption")
optionTypename = Request("optionTypename")
arritemoption       = request("arritemoption")
arritemoption = Split(arritemoption, "|")

redim ArroptionName(Request.Form("optionName").Count)
redim Arroptaddprice(Request.Form("optaddprice").Count)
redim Arroptaddbuyprice(Request.Form("optaddbuyprice").Count)
ArrCnt = Request.Form("optionName").Count

TypeSeq = request("TypeSeq")
KindSeq = request("KindSeq")

dim IsUpchebeasong, itemLimitYn
IsUpchebeasong =false
itemLimitYn = "N"

''업체 수정인경우 브랜드ID 체크// 배치로 던지는 CASE 있는듯.
if (Not checkIsUpcheItemEditValid(itemid)) then
    response.write "<script>alert('권한이 없습니다. 해당 브랜드 상품이 아닙니다.');</script>"
    dbget.Close(): response.end
end if

''업체배송인경우 입출/판매 관계없이 삭제
sqlStr = " select limityn, deliverytype "
sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item"
sqlStr = sqlStr & " where itemid=" & CStr(itemid)

rsACADEMYget.Open sqlStr,dbACADEMYget,1
if not rsACADEMYget.EOF then
    itemLimitYn = rsACADEMYget("limityn")
    IsUpchebeasong = (rsACADEMYget("deliverytype") = "2") or (rsACADEMYget("deliverytype") = "5") or (rsACADEMYget("deliverytype") = "9") or (rsACADEMYget("deliverytype") = "7")
end if
rsACADEMYget.Close




function ReCalcuItemOption(itemid)
    dim sqlStr
    ''상품옵션수량
	sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
	sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
	sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_item.itemid=" + CStr(itemid) + VBCrlf
	dbACADEMYget.Execute sqlStr

	''상품한정수량
	sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_item.itemid=" + CStr(itemid) + VBCrlf
	dbACADEMYget.Execute sqlStr
	
	''한정이 꺽이면 일시 품절처리 // 2013/09/02 추가
    sqlStr = "update [db_academy].[dbo].tbl_diy_item" + VBCrlf
    sqlStr = sqlStr + " set sellyn='S'" + VBCrlf
    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + VBCrlf
    sqlStr = sqlStr + " and sellyn='Y'" + VBCrlf
    sqlStr = sqlStr + " and limityn='Y'" + VBCrlf
    sqlStr = sqlStr + " and (limitno-limitsold)<1" + VBCrlf
    dbACADEMYget.Execute sqlStr
end function

function ReMatchMultiOption(itemid)
    dim sqlStr
    dim MultiLevel
    
    MultiLevel = 0
    
    sqlStr = " select TypeSeq, Count(KindSeq) as KindCnt "
    sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_item_option_Multiple "
    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    sqlStr = sqlStr + " group by TypeSeq"
    sqlStr = sqlStr + " order by TypeSeq"
    
    rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	    MultiLevel = rsACADEMYget.RecordCount
	rsACADEMYget.close
    	
    ''기존 2차 옵션인 경우 삭제.
    if (MultiLevel=3) then 
        sqlStr = " delete from db_academy.dbo.tbl_diy_item_option"
        sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'"
        sqlStr = sqlStr + " and Right(itemoption,1)='0'"
        
        dbACADEMYget.Execute sqlStr
    end if
    
    if (MultiLevel=2) then 
        sqlStr = " delete from db_academy.dbo.tbl_diy_item_option"
        sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'"
        sqlStr = sqlStr + " and Right(itemoption,1)='00'"
        
        dbACADEMYget.Execute sqlStr
    end if 
    
''response.write  MultiLevel   
    ''옵션 재작성.
'   --Only 1중옵션.
    if (MultiLevel=1) then 
        ''-- 전 옵션 삭제;
        sqlStr = " delete from db_academy.dbo.tbl_diy_item_option_Multiple" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        dbACADEMYget.Execute sqlStr
        
        sqlStr = " delete from db_academy.dbo.tbl_diy_item_option" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        sqlStr = sqlStr & " and Left(itemoption,1)='Z'"
        dbACADEMYget.Execute sqlStr
    end if
    
'   --Only 2중옵션.
    if (MultiLevel=2) then 
        sqlStr = " insert into db_academy.dbo.tbl_diy_item_option"
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '복합옵션' as optionTypeName,"
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0,"
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option_Multiple b"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + " ) T"
        sqlStr = sqlStr + "     left join db_academy.dbo.tbl_diy_item_option o "
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and T.itemid=o.itemid "
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption "
        sqlStr = sqlStr + " where  o.itemid is NULL"
    
        dbACADEMYget.Execute sqlStr
        
        '' 옵션명/ 가격 등이 변경된 경우
        sqlStr = " update db_academy.dbo.tbl_diy_item_option"
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)"
        sqlStr = sqlStr + " , optaddprice=T.optaddprice"
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName ) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option_Multiple b"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_item_option.itemid=T.itemid"
        sqlStr = sqlStr + " and db_academy.dbo.tbl_diy_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and ("
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option.optionname<>T.optionname"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_item_option.optaddprice<>T.optaddprice"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_item_option.optaddbuyprice<>T.optaddbuyprice"
        sqlStr = sqlStr + " )"
'response.write sqlStr
        dbACADEMYget.Execute sqlStr
    end if

'    --Only 3중옵션
    if (MultiLevel=3) then 
        sqlStr = " insert into db_academy.dbo.tbl_diy_item_option"
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '복합옵션' as optionTypeName,"
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0,"
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option_Multiple b,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option_Multiple c"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and b.itemid=c.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + "     left join db_academy.dbo.tbl_diy_item_option o "
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and T.itemid=o.itemid "
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption "
        sqlStr = sqlStr + " where  o.itemid is NULL"
        
        dbACADEMYget.Execute sqlStr
        
        
        '' 옵션명/ 가격 등이 변경된 경우
        sqlStr = " update db_academy.dbo.tbl_diy_item_option"
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)"
        sqlStr = sqlStr + " , optaddprice=T.optaddprice"
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice"
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ,"
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname,"
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice"
        sqlStr = sqlStr + "     from db_academy.dbo.tbl_diy_item_option_Multiple a,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option_Multiple b,"
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option_Multiple c"
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid)
        sqlStr = sqlStr + "     and a.itemid=b.itemid"
        sqlStr = sqlStr + "     and b.itemid=c.itemid"
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq"
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq"
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq "
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq "
        sqlStr = sqlStr + " ) T "
        sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_item_option.itemid=T.itemid"
        sqlStr = sqlStr + " and db_academy.dbo.tbl_diy_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and ("
        sqlStr = sqlStr + "     db_academy.dbo.tbl_diy_item_option.optionname<>T.optionname"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_item_option.optaddprice<>T.optaddprice"
        sqlStr = sqlStr + "     or db_academy.dbo.tbl_diy_item_option.optaddbuyprice<>T.optaddbuyprice"
        sqlStr = sqlStr + " )"

        dbACADEMYget.Execute sqlStr
    end if
    
end function
Function CheckMultiOptionDelYN(IsUpchebeasong,itemid,TypeSeq,KindSeq)
	Dim CheckStr
    if (Not IsUpchebeasong) then
        sqlStr = "select top 1 * from "
    	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    	if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=2) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=3) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
    	end if

    	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    	if Not rsACADEMYget.Eof then
    		CheckStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    	end if
    	rsACADEMYget.close

    	''6개월 이전 판매내역
    	if CheckStr="" then
    		sqlStr = "select top 1 * from "
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    		sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    		if (TypeSeq=1) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=2) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=3) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
        	end if

    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			CheckStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if

    	''입출고내역
    	if CheckStr="" then
    		sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
    		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
    		sqlStr = sqlStr + " where m.code=d.mastercode"
    		sqlStr = sqlStr + " and m.deldt is NULL"
    		sqlStr = sqlStr + " and d.iitemgubun='10'"
    		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
            sqlStr = sqlStr + " and d.deldt is NULL"
            if (TypeSeq=1) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=2) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=3) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
        	end if
            
    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			CheckStr = "삭제하려는 옵션으로 입출고 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if
    end If
    CheckMultiOptionDelYN=CheckStr
End function
Function CheckSingleOptionDelYN(IsUpchebeasong,itemid,itemoption)
	Dim CheckStr
	if (Not IsUpchebeasong) Then
    	''최근 판매내역
    	sqlStr = "select top 1 * from "
    	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    	sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"

    	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    	if Not rsACADEMYget.Eof then
    		CheckStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    	end if
    	rsACADEMYget.close

    	''6개월 이전 판매내역
    	if CheckStr="" then
    		sqlStr = "select top 1 * from "
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    		sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    		sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"

    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			CheckStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if


    	''입출고내역
    	if CheckStr="" then
    		sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
    		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
    		sqlStr = sqlStr + " where m.code=d.mastercode"
    		sqlStr = sqlStr + " and m.deldt is NULL"
    		sqlStr = sqlStr + " and d.iitemgubun='10'"
    		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
    		sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"
            sqlStr = sqlStr + " and d.deldt is NULL"
            
    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			CheckStr = "삭제하려는 옵션으로 입출고 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end If
	end If
	CheckSingleOptionDelYN=CheckStr
End Function

if (mode = "deleteoption") then
	''삭제 가능한 옵션인지 체크
	if (Not IsUpchebeasong) then
    	''최근 판매내역
    	sqlStr = "select top 1 * from "
    	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    	sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"

    	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    	if Not rsACADEMYget.Eof then
    		ErrStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    	end if
    	rsACADEMYget.close

    	''6개월 이전 판매내역
    	if ErrStr="" then
    		sqlStr = "select top 1 * from "
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    		sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    		sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"

    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			ErrStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if

    	''입출고내역
    	if ErrStr="" then
    		sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
    		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
    		sqlStr = sqlStr + " where m.code=d.mastercode"
    		sqlStr = sqlStr + " and m.deldt is NULL"
    		sqlStr = sqlStr + " and d.iitemgubun='10'"
    		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
    		sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"
            sqlStr = sqlStr + " and d.deldt is NULL"
            
    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			ErrStr = "삭제하려는 옵션으로 입출고 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if
	end if

	if (ErrStr="") then
		sqlStr = "delete from db_academy.dbo.tbl_diy_item_option" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + CStr(Trim(itemoption)) + "'" + VbCrlf
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		Call ReCalcuItemOption(itemid)
	end if
end If

'' 옵션삭제 - 이중옵션
if (mode = "deleteMultipleOption") then 
    'TypeSeq
    'KindSeq
    
    if (Not IsUpchebeasong) then
        sqlStr = "select top 1 * from "
    	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    	if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=2) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=3) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
    	end if

    	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    	if Not rsACADEMYget.Eof then
    		ErrStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    	end if
    	rsACADEMYget.close

    	''6개월 이전 판매내역
    	if ErrStr="" then
    		sqlStr = "select top 1 * from "
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    		sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    		if (TypeSeq=1) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=2) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=3) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
        	end if

    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			ErrStr = "삭제하려는 옵션으로 판매된 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if

    	''입출고내역
    	if ErrStr="" then
    		sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
    		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
    		sqlStr = sqlStr + " where m.code=d.mastercode"
    		sqlStr = sqlStr + " and m.deldt is NULL"
    		sqlStr = sqlStr + " and d.iitemgubun='10'"
    		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
            sqlStr = sqlStr + " and d.deldt is NULL"
            if (TypeSeq=1) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=2) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=3) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
        	end if
            
    		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
    		if Not rsACADEMYget.Eof then
    			ErrStr = "삭제하려는 옵션으로 입출고 내역이 있습니다. 삭제하실 수 없습니다. - 관리자 문의요망"
    		end if
    		rsACADEMYget.close
    	end if
    end if
    
    If (ErrStr<>"") Then

	Else
	    sqlStr = "delete from db_academy.dbo.tbl_diy_item_option_Multiple" + VbCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	    sqlStr = sqlStr + " and TypeSeq=" + CStr(TypeSeq)
	    sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq) + "'"
'Response.write sqlStr
'Response.end
	    dbACADEMYget.Execute sqlStr
	    
		sqlStr = "delete from db_academy.dbo.tbl_diy_item_option" + VbCrlf
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

		Call ReCalcuItemOption(itemid)
	end if
end if

'' 단일 옵션 추가
if (mode = "addoptionCustom") then
    foundcount = 0
    
    for i = 1 to ArrCnt
        if (Trim(arritemoption(i-1)) <> "") Then
			ArroptionName(i) = Request.Form("optionName")(i)
			Arroptaddprice(i) = Request.Form("optaddprice")(i)
			Arroptaddbuyprice(i) = Request.Form("optaddbuyprice")(i)
            sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option "
            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
            sqlStr = sqlStr + " and ((itemoption = '" + CStr(Trim(arritemoption(i-1))) + "') or (optionname='" + html2db(ArroptionName(i)) + "'))"
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
                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
                sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(arritemoption(i-1)) + "', convert(varchar(32),'" + html2db(optionTypename) + "'), convert(varchar(96),'" + CStr(html2db(ArroptionName(i))) + "'), 'Y', 'Y', '" + itemLimitYn + "', 0, 0," + CStr(Arroptaddprice(i)) + "," + CStr(Arroptaddbuyprice(i)) + ") "
                dbACADEMYget.Execute sqlStr
			Else
				sqlStr = " update db_academy.dbo.tbl_diy_item_option"
                sqlStr = sqlStr + " set optionTypeName=convert(varchar(32),'" + html2db(optionTypename) + "')"
				sqlStr = sqlStr + " ,optionname=convert(varchar(96),'" + CStr(html2db(ArroptionName(i))) + "')"
				sqlStr = sqlStr + " ,optaddprice='" + CStr(arroptaddprice(i)) + "'"
				sqlStr = sqlStr + " ,optaddbuyprice='" + CStr(arroptaddbuyprice(i)) + "'"
				sqlStr = sqlStr + " where itemid='" + CStr(itemid) + "'"
				sqlStr = sqlStr + " and itemoption='" + CStr(Trim(arritemoption(i-1))) + "'"
                dbACADEMYget.Execute sqlStr
            end if
        end if
    next
    
    ''옵션 구분명은 동일
    
    sqlStr = " update db_academy.dbo.tbl_diy_item_option "
    sqlStr = sqlStr + " set optionTypeName=convert(varchar(32),'" + html2db(optionTypename) + "')"
    sqlStr = sqlStr + " where itemid=" + cStr(itemid)
    sqlStr = sqlStr + " and optionTypeName<>'" + html2db(optionTypename) + "'"
    
    dbACADEMYget.Execute sqlStr
    
    Call ReCalcuItemOption(itemid)

    if (foundcount > 0) then
       'ErrStr = "일부 옵션은 기존에 있는 옵션과 중복되어 무시되었습니다."
    end if
end If

'' 단일 옵션 추가
if (mode = "ResetSingleOptionCustom") then
    foundcount = 0

	'################### 기존 등록 옵션 삭제 #######################################################
	sqlStr = "delete from db_academy.dbo.tbl_diy_item_option"
	sqlStr = sqlStr & " where itemid=" & itemid
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	'################### 기존 등록 옵션 삭제 #######################################################
	sqlStr = "delete from db_academy.dbo.tbl_diy_item_option_Multiple"
	sqlStr = sqlStr & " where itemid=" & itemid
	rsACADEMYget.Open sqlStr,dbACADEMYget,1

    For i = 1 To ArrCnt
		ArroptionName(i) = Request.Form("optionName")(i)
		Arroptaddprice(i) = Request.Form("optaddprice")(i)
		Arroptaddbuyprice(i) = Request.Form("optaddbuyprice")(i)
		if (Trim(arritemoption(i-1)) <> "") Then
			sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option "
            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
            sqlStr = sqlStr + " and ((itemoption = '" + CStr(Trim(arritemoption(i-1))) + "') or (optionname='" + html2db(ArroptionName(i)) + "'))"
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
                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
                sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(arritemoption(i-1)) + "', convert(varchar(32),'" + html2db(optionTypename) + "'), convert(varchar(96),'" + CStr(html2db(ArroptionName(i))) + "'), 'Y', 'Y', '" + itemLimitYn + "', 0, 0," + CStr(Arroptaddprice(i)) + "," + CStr(Arroptaddbuyprice(i)) + ")"
                dbACADEMYget.Execute sqlStr
            end if
        end if
    next
    
    ''옵션 구분명은 동일
    
    sqlStr = " update db_academy.dbo.tbl_diy_item_option "
    sqlStr = sqlStr + " set optionTypeName=convert(varchar(32),'" + html2db(optionTypename) + "')"
    sqlStr = sqlStr + " where itemid=" + cStr(itemid)
    sqlStr = sqlStr + " and optionTypeName<>'" + html2db(optionTypename) + "'"
    
    dbACADEMYget.Execute sqlStr
    
    Call ReCalcuItemOption(itemid)

    if (foundcount > 0) then
       ErrStr "일부 옵션은 기존에 있는 옵션과 중복되어 무시되었습니다."
    end if
end If

if (mode="addDoubleOption") then
    
    optionTypename1 = Trim(Request.Form("opttypename1"))
    optionTypename2 = Trim(Request.Form("opttypename2"))
    optionTypename3 = Trim(Request.Form("opttypename3"))
	redim arroptionName1(Request.Form("optname1").Count)
	redim arroptaddprice1(Request.Form("optaddprice1").Count)
	redim arroptaddbuyprice1(Request.Form("optbuyprice1").Count)
	redim arroptionName2(Request.Form("optname2").Count)
	redim arroptaddprice2(Request.Form("optaddprice2").Count)
	redim arroptaddbuyprice2(Request.Form("optbuyprice2").Count)
	redim arroptionName3(Request.Form("optname3").Count)
	redim arroptaddprice3(Request.Form("optaddprice3").Count)
	redim arroptaddbuyprice3(Request.Form("optbuyprice3").Count)

    Lv1cnt = Request.Form("optname1").Count
	Lv2cnt = Request.Form("optname2").Count
	Lv3cnt = Request.Form("optname3").Count
    
    Val1cnt = 0
    Val2cnt = 0
    Val3cnt = 0
    
    for i=1 to Lv1cnt
		arroptionName1(i) = Request.Form("optname1")(i)
		arroptaddprice1(i) = Request.Form("optaddprice1")(i)
		arroptaddbuyprice1(i) = Request.Form("optbuyprice1")(i)
        buf = Trim(arroptionName1(i))
        if Len(buf)>0 then Val1cnt = Val1cnt + 1
    next
    
    for i=1 to Lv2cnt
		arroptionName2(i) = Request.Form("optname2")(i)
		arroptaddprice2(i) = Request.Form("optaddprice2")(i)
		arroptaddbuyprice2(i) = Request.Form("optbuyprice2")(i)
        buf = Trim(arroptionName2(i))
        if Len(buf)>0 then Val2cnt = Val2cnt + 1
    next
    
    for i=1 to Lv3cnt
		arroptionName3(i) = Request.Form("optname3")(i)
		arroptaddprice3(i) = Request.Form("optaddprice3")(i)
		arroptaddbuyprice3(i) = Request.Form("optbuyprice3")(i)
        buf = Trim(arroptionName3(i))
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
        WaitRegDoubleOptionProc =ErrMsg 
        ''Exit function
    end if

    foundcount=0
	If Lv3cnt=0 Then Lv3cnt=1
    ''0번은 입력 없음. N까지
    for i=0 to Lv1cnt-1
        for j=0 to Lv2cnt-1
            for k=0 to Lv3cnt-1
                optName1 = arroptionName1(i+1)
				optprice1  = arroptaddprice1(i+1)
				optbuyprice1  = arroptaddbuyprice1(i+1)

                If (Lv2cnt) Then
					optName2 = arroptionName2(j+1)
					optprice2  = arroptaddprice2(j+1)
					optbuyprice2  = arroptaddbuyprice2(j+1)
				Else
					optName2 = ""
					optprice2  = ""
					optbuyprice2  = ""
				End If
				If (Lv3cnt>1) Then
					optName3 = arroptionName3(k+1)
					optprice3  = arroptaddprice3(k+1)
					optbuyprice3  = arroptaddbuyprice3(k+1)
				Else
					optName3 = ""
					optprice3  = ""
					optbuyprice3  = ""
				End If

                Valid1  = (Len(optionTypename1)>0) and (Len(optName1)>0)
                Valid2  = (Len(optionTypename2)>0) and (Len(optName2)>0)
                Valid3  = (Len(optionTypename3)>0) and (Len(optName3)>0)
                
                AssignedOption = "Z"  '''변경해야함.. =>9
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

                        if (Len(optName1)>0) and (Len(optionTypename1)>0) then
                            sqlStr = " select itemid, TypeSeq, KindSeq  from db_academy.dbo.tbl_diy_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and replace(optionTypeName,' ','')='" + html2db(Replace(optionTypename1," ","")) + "' and replace(optionKindName,' ','')='" + html2db(Replace(optName1," ","")) + "'"
							'Response.write sqlstr
							'Response.end
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
								TypeSeq1 = rsACADEMYget("TypeSeq")
								KindSeq1 = rsACADEMYget("KindSeq")
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice, optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,1"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(i+1)) &"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename1) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName1) & "')"
								sqlStr = sqlStr + " ," & Cstr(optprice1)
								sqlStr = sqlStr + " ," & Cstr(optbuyprice1)
                                sqlStr = sqlStr + " )"
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName1 + "," + optionTypename1 + "," + CStr(1) + "," + CStr(i+1) + "<br>"
							Else
								sqlStr = " update db_academy.dbo.tbl_diy_item_option_Multiple"
								sqlStr = sqlStr + " set optaddprice='" + CStr(optprice1) + "'"
								sqlStr = sqlStr + " ,optaddbuyprice='" + CStr(optbuyprice1) + "'"
								sqlStr = sqlStr + " where itemid='" + CStr(itemid) + "'"
								sqlStr = sqlStr + " and TypeSeq='" + CStr(TypeSeq1) + "'"
								sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq1) + "'"
								dbACADEMYget.Execute sqlStr
                            end if
                        end if
                        
                        found = false
                        if (Len(optName2)>0) and (Len(optionTypename2)>0) then
                            sqlStr = " select itemid, TypeSeq, KindSeq from db_academy.dbo.tbl_diy_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and replace(optionTypeName,' ','')='" + html2db(Replace(optionTypename2," ","")) + "' and replace(optionKindName,' ','')='" + html2db(Replace(optName2," ","")) + "'"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
								TypeSeq2 = rsACADEMYget("TypeSeq")
								KindSeq2 = rsACADEMYget("KindSeq")
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice,optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,2"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(J+1)) &"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename2) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName2) & "')"
								sqlStr = sqlStr + " ," & Cstr(optprice2)
								sqlStr = sqlStr + " ," & Cstr(optbuyprice2)
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName2 + "," + optionTypename2 + "," + CStr(2) + "," + CStr(j+1) + "<br>"
							Else
								sqlStr = " update db_academy.dbo.tbl_diy_item_option_Multiple"
								sqlStr = sqlStr + " set optaddprice='" + CStr(optprice2) + "'"
								sqlStr = sqlStr + " ,optaddbuyprice='" + CStr(optbuyprice2) + "'"
								sqlStr = sqlStr + " where itemid='" + CStr(itemid) + "'"
								sqlStr = sqlStr + " and TypeSeq='" + CStr(TypeSeq2) + "'"
								sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq2) + "'"
								dbACADEMYget.Execute sqlStr
                            end if
                        end if
                        
                        found = false
                        if (Len(optName3)>0) and (Len(optionTypename3)>0) then
                            sqlStr = " select itemid, TypeSeq, KindSeq from db_academy.dbo.tbl_diy_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and replace(optionTypeName,' ','')='" + html2db(Replace(optionTypename3," ","")) + "' and replace(optionKindName,' ','')='" + html2db(Replace(optName3," ","")) + "'"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
								TypeSeq3 = rsACADEMYget("TypeSeq")
								KindSeq3 = rsACADEMYget("KindSeq")
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice,optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,3"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(K+1)) &"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename3) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName3) & "')"
								sqlStr = sqlStr + " ," & Cstr(optprice3)
								sqlStr = sqlStr + " ," & Cstr(optbuyprice3)
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName3 + "," + optionTypename3 + "," + CStr(3) + "," + CStr(k+1) + "<br>"
							Else
								sqlStr = " update db_academy.dbo.tbl_diy_item_option_Multiple"
								sqlStr = sqlStr + " set optaddprice='" + CStr(optprice3) + "'"
								sqlStr = sqlStr + " ,optaddbuyprice='" + CStr(optbuyprice3) + "'"
								sqlStr = sqlStr + " where itemid='" + CStr(itemid) + "'"
								sqlStr = sqlStr + " and TypeSeq='" + CStr(TypeSeq3) + "'"
								sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq3) + "'"
								dbACADEMYget.Execute sqlStr
                            end if
                        end if
                        
                        found = false
                        sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option "
                        sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                        sqlStr = sqlStr + " and ((itemoption = '" + CStr(AssignedOption) + "') or (optionTypeName='복합옵션' and optionname='" + html2db(optionName) + "'))"
                        
                        rsACADEMYget.Open sqlStr,dbACADEMYget,1
                        if Not rsACADEMYget.Eof then
                            found = true
                        end if
                        rsACADEMYget.Close
                        
                        if (Not found) then 
                            sqlStr = " insert into db_academy.dbo.tbl_diy_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                            sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(AssignedOption) + "', '복합옵션', '" + CStr(html2db(optionName)) + "', 'Y', 'Y', '" + Request.Form("limityn") + "', " & Request.Form("limitno") & ", 0) "
                            
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
    Call ReCalcuItemOption(itemid)

end if
'' 단일옵션 --> 이중옵션으로 변경
if (mode="ResetMultiOptionCustom") then
    
    optionTypename1 = Trim(Request.Form("opttypename1"))
    optionTypename2 = Trim(Request.Form("opttypename2"))
    optionTypename3 = Trim(Request.Form("opttypename3"))
	redim arroptionName1(Request.Form("optname1").Count)
	redim arroptaddprice1(Request.Form("optaddprice1").Count)
	redim arroptaddbuyprice1(Request.Form("optbuyprice1").Count)
	redim arroptionName2(Request.Form("optname2").Count)
	redim arroptaddprice2(Request.Form("optaddprice2").Count)
	redim arroptaddbuyprice2(Request.Form("optbuyprice2").Count)
	redim arroptionName3(Request.Form("optname3").Count)
	redim arroptaddprice3(Request.Form("optaddprice3").Count)
	redim arroptaddbuyprice3(Request.Form("optbuyprice3").Count)

    Lv1cnt = Request.Form("optname1").Count
	Lv2cnt = Request.Form("optname2").Count
	Lv3cnt = Request.Form("optname3").Count
    
    Val1cnt = 0
    Val2cnt = 0
    Val3cnt = 0
    
    for i=1 to Lv1cnt
		arroptionName1(i) = Request.Form("optname1")(i)
		arroptaddprice1(i) = Request.Form("optaddprice1")(i)
		arroptaddbuyprice1(i) = Request.Form("optbuyprice1")(i)
        buf = Trim(arroptionName1(i))
        if Len(buf)>0 then Val1cnt = Val1cnt + 1
    next
    
    for i=1 to Lv2cnt
		arroptionName2(i) = Request.Form("optname2")(i)
		arroptaddprice2(i) = Request.Form("optaddprice2")(i)
		arroptaddbuyprice2(i) = Request.Form("optbuyprice2")(i)
        buf = Trim(arroptionName2(i))
        if Len(buf)>0 then Val2cnt = Val2cnt + 1
    next
    
    for i=1 to Lv3cnt
		arroptionName3(i) = Request.Form("optname3")(i)
		arroptaddprice3(i) = Request.Form("optaddprice3")(i)
		arroptaddbuyprice3(i) = Request.Form("optbuyprice3")(i)
        buf = Trim(arroptionName3(i))
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
        WaitRegDoubleOptionProc =ErrMsg 
        ''Exit function
    end if

    foundcount=0

	If Lv1cnt > 0 Then
		'################### 기존 등록 옵션 삭제 #######################################################
		sqlStr = "delete from db_academy.dbo.tbl_diy_item_option"
		sqlStr = sqlStr & " where itemid=" & itemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		'################### 기존 등록 옵션 삭제 #######################################################
		sqlStr = "delete from db_academy.dbo.tbl_diy_item_option_Multiple"
		sqlStr = sqlStr & " where itemid=" & itemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
	End If
	If Lv3cnt=0 Then Lv3cnt=1
    ''0번은 입력 없음. N까지
    for i=0 to Lv1cnt-1
        for j=0 to Lv2cnt-1
            for k=0 to Lv3cnt-1
                optName1 = Trim(arroptionName1(i+1))
				optprice1  = Trim(arroptaddprice1(i+1))
				optbuyprice1  = Trim(arroptaddbuyprice1(i+1))

                If (Lv2cnt) Then
					optName2 = Trim(arroptionName2(j+1))
					optprice2  = Trim(arroptaddprice2(j+1))
					optbuyprice2  = Trim(arroptaddbuyprice2(j+1))
				Else
					optName2 = ""
					optprice2  = ""
					optbuyprice2  = ""
				End If
				If (Lv3cnt>1) Then
					optName3 = Trim(arroptionName3(k+1))
					optprice3  = Trim(arroptaddprice3(k+1))
					optbuyprice3  = Trim(arroptaddbuyprice3(k+1))
				Else
					optName3 = ""
					optprice3  = ""
					optbuyprice3  = ""
				End If

                Valid1  = (Len(optionTypename1)>0) and (Len(optName1)>0)
                Valid2  = (Len(optionTypename2)>0) and (Len(optName2)>0)
                Valid3  = (Len(optionTypename3)>0) and (Len(optName3)>0)
                
                 AssignedOption = "Z"  '''변경해야함.. =>9
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

                        if (Len(optName1)>0) and (Len(optionTypename1)>0) then
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename1) + "' and optionKindName='" + html2db(optName1) + "'))"

                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice, optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,1"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(i+1)) &"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename1) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName1) & "')"
								sqlStr = sqlStr + " ," & Cstr(optprice1)
								sqlStr = sqlStr + " ," & Cstr(optbuyprice1)
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName1 + "," + optionTypename1 + "," + CStr(1) + "," + CStr(i+1) + "<br>"
                            end if
                        end if
                        
                        found = false
                        if (Len(optName2)>0) and (Len(optionTypename2)>0) then
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename2) + "' and optionKindName='" + html2db(optName2) + "'))"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice,optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,2"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(J+1)) &"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename2) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName2) & "')"
								sqlStr = sqlStr + " ," & Cstr(optprice2)
								sqlStr = sqlStr + " ," & Cstr(optbuyprice2)
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName2 + "," + optionTypename2 + "," + CStr(2) + "," + CStr(j+1) + "<br>"
                            end if
                        end if
                        
                        found = false
                        if (Len(optName3)>0) and (Len(optionTypename3)>0) then
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename3) + "' and optionKindName='" + html2db(optName3) + "'))"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice,optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & itemid 
                                sqlStr = sqlStr + " ,3"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(K+1)) &"'"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optionTypename3) & "')"
                                sqlStr = sqlStr + " ,convert(varchar(32),'" & html2db(optName3) & "')"
								sqlStr = sqlStr + " ," & Cstr(optprice3)
								sqlStr = sqlStr + " ," & Cstr(optbuyprice3)
                                sqlStr = sqlStr + " )"
                                
                                dbACADEMYget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName3 + "," + optionTypename3 + "," + CStr(3) + "," + CStr(k+1) + "<br>"
                            end if
                        end if
                        
                        found = false
                        sqlStr = " select itemid from db_academy.dbo.tbl_diy_item_option "
                        sqlStr = sqlStr + " where itemid = " + CStr(itemid)  
                        sqlStr = sqlStr + " and ((itemoption = '" + CStr(AssignedOption) + "') or (optionTypeName='복합옵션' and optionname='" + html2db(optionName) + "'))"
                        
                        rsACADEMYget.Open sqlStr,dbACADEMYget,1
                        if Not rsACADEMYget.Eof then
                            found = true
                        end if
                        rsACADEMYget.Close
                        
                        if (Not found) then 
                            sqlStr = " insert into db_academy.dbo.tbl_diy_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                            sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(AssignedOption) + "', '복합옵션', '" + CStr(html2db(optionName)) + "', 'Y', 'Y', '" + Request.Form("limityn") + "', " & Request.Form("limitno") & ", 0) "
                            
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
    Call ReCalcuItemOption(itemid)

end if

'' 이중에서 단일 옵션으로 변경
if (mode = "CheckResetSingleOption") then
	Dim MultiCheck
	i=0
	sqlStr = " select itemoption, optionTypeName from db_academy.dbo.tbl_diy_item_option "
	sqlStr = sqlStr + " where itemid = " + CStr(itemid)
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	If Not rsACADEMYget.EOF Then
		ReDim Arritemoption(i)
		Arritemoption(i)=rsACADEMYget("itemoption")
		If rsACADEMYget("optionTypeName") = "복합옵션" Then
			MultiCheck="Y"
		End If
		i=i+1
	End If
	rsACADEMYget.close

	For k=0 To k > i
		If MultiCheck<>"Y" Then
			ErrStr = CheckSingleOptionDelYN(IsUpchebeasong,itemid, Arritemoption(k))
			If ErrStr <> "" Then
				%>
				<script>parent.fnOptionDelCheckEnd("<%=ErrStr%>");</script>
				<%
			rsACADEMYget.close : Response.End
			End If
		End If
	Next

	If MultiCheck="Y" Then
		i=0
		sqlStr = " select TypeSeq, KindSeq from db_academy.dbo.tbl_diy_item_option_Multiple "
		sqlStr = sqlStr + " where itemid = " + CStr(itemid)
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If Not rsACADEMYget.EOF Then
			ReDim ArrTypeSeq(i)
			ReDim ArrKindSeq(i)
			ArrTypeSeq(i)=rsACADEMYget("TypeSeq")
			ArrKindSeq(i)=rsACADEMYget("KindSeq")
			i=i+1
		End If
		For j=0 To j > i
			ErrStr = CheckMultiOptionDelYN(IsUpchebeasong,itemid,ArrTypeSeq(j),ArrKindSeq(j))
			If ErrStr <> "" Then
				%>
				<script>parent.fnOptionDelCheckEnd("<%=ErrStr%>");</script>
				<%
			rsACADEMYget.close : Response.End
			End If
		Next
	End If
End If

%>
<% If mode="deleteoption" Then %>
<script>parent.fnOptionDelCheckEnd("<%=ErrStr%>");</script>
<% ElseIf mode="addoptionCustom" Then %>
<script>parent.fnOptionAddEnd("<%=ErrStr%>");</script>
<% ElseIf mode="deleteMultipleOption" Then %>
<script>parent.fnOptionDelCheckEnd("<%=ErrStr%>");</script>
<% ElseIf mode="addDoubleOption" Then %>
<script>parent.fnMultipleOptionEditEnd("<%=ErrStr%>");</script>
<% ElseIf mode="CheckResetSingleOption" Then %>
<script>parent.fnOptionDelCheckEnd("<%=ErrStr%>");</script>
<% ElseIf mode="ResetSingleOptionCustom" Then %>
<script>parent.fnOptionAddEnd("<%=ErrStr%>");</script>
<% ElseIf mode="ResetMultiOptionCustom" Then %>
<script>parent.fnMultipleOptionEditEnd("<%=ErrStr%>");</script>
<% End If %>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim itemid, limityn, limitno, limitsold
dim sellyn, isusing, reqstring
dim pojangok
dim itemoptionarr, optlimitnoarr, optlimitsoldarr, optisusingarr
dim index

itemid      = requestCheckvar(request("itemid"),10)
limityn     = requestCheckvar(request("limityn"),1)
limitno     = requestCheckvar(request("limitno"),10)
limitsold   = requestCheckvar(request("limitsold"),10)
sellyn      = requestCheckvar(request("sellyn"),1)
isusing     = requestCheckvar(request("isusing"),1)
reqstring   = request("reqstring")


itemoptionarr 	= request("itemoptionarr")
optlimitnoarr	= request("optlimitnoarr")
optlimitsoldarr = request("optlimitsoldarr")
optisusingarr	= request("optisusingarr")

'response.write "itemid=" & itemid & "<br>"
'response.write "limityn=" & limityn & "<br>"
'response.write "limitno=" & limitno & "<br>"
'response.write "limitsold=" & limitsold & "<br>"
'response.write "sellyn=" & sellyn & "<br>"
'response.write "isusing=" & isusing & "<br>"
'response.write "reqstring=" & reqstring & "<br>"
'
'response.write "itemoptionarr=" & itemoptionarr & "<br>"
'response.write "optlimitnoarr=" & optlimitnoarr & "<br>"
'response.write "optlimitsoldarr=" & optlimitsoldarr & "<br>"
'response.write "optisusingarr=" & optisusingarr & "<br>"

if itemid = "" then
	response.write "<script>alert('상품코드가 없습니다'); history.back();</script>"
	dbget.close()	:	response.End
end if

''2014/03/18 추가
if (limityn <> "Y") and (limityn <> "N") then
	response.write "<script>alert('파라메터 오류 limityn'); history.back();</script>"
	dbget.close()	:	response.End
end if

if (sellyn <> "Y") and (sellyn <> "N") and (sellyn <> "S") then
	response.write "<script>alert('파라메터 오류 sellyn'); history.back();</script>"
	dbget.close()	:	response.End
end if

if (isusing <> "Y") and (isusing <> "N") then
	response.write "<script>alert('파라메터 오류 isusing'); history.back();</script>"
	dbget.close()	:	response.End
end if
''2014/03/18

itemoptionarr 	= split(itemoptionarr,",")
optlimitnoarr	= split(optlimitnoarr,",")
optlimitsoldarr = split(optlimitsoldarr,",")
optisusingarr 	= split(optisusingarr,",")


dim OptionExists
OptionExists = (UBound(itemoptionarr)>1)
dim BufOptionKey

'response.write "OptionExists=" & OptionExists
'dbget.close()	:	response.End
'==============================================================================
'기존 상품 정보
dim mwdiv, orgsellyn, orglimityn, orglimitno, orglimitsold, orgoptioncnt,orgsellSTDate
    
sqlStr = "select top 1 sellyn, limityn, limitno, limitsold, mwdiv, optioncnt, sellSTDate " + VbCrlf
sqlStr = sqlStr + " from [db_item].[dbo].tbl_item" + VbCrlf
sqlStr = sqlStr + " where 1 = 1 "
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + " " + VbCrlf
sqlStr = sqlStr + " and makerid = '" + CStr(session("ssBctID")) + "' "
rsget.Open sqlStr,dbget,1

if  not rsget.EOF  then
    orgsellyn = rsget("sellyn")
    orglimityn = rsget("limityn")
    orglimitno = rsget("limitno")
    orglimitsold = rsget("limitsold")
    mwdiv = rsget("mwdiv")
    orgoptioncnt = rsget("optioncnt") 
    orgsellSTDate   = rsget("sellSTDate") 
else
    rsget.Close

    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if
rsget.Close


'==============================================================================
'기존 옵션 정보
dim orgarritemoption, orgarritemoptionname, orgarrisusing, orgarroptlimityn, orgarroptlimitno, orgarroptlimitsold
''orgarroptsellyn - 사용안함(optisusing 과 일치)

sqlStr = " select top 100 o.itemoption, isnull(o.optionname,'') as itemoptionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold "
sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option o "
sqlStr = sqlStr + " where 1 = 1 "
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + " "
sqlStr = sqlStr + " and itemoption <> '0000' "
sqlStr = sqlStr + " order by itemoption "
rsget.Open sqlStr,dbget,1

if  not rsget.EOF  then
    do until rsget.Eof
        orgarritemoption        = orgarritemoption + rsget("itemoption") + "|"
        orgarritemoptionname    = orgarritemoptionname + db2html(rsget("itemoptionname")) + "|"
        orgarrisusing           = orgarrisusing + rsget("isusing") + "|"
        orgarroptlimityn        = orgarroptlimityn + rsget("optlimityn") + "|"
        orgarroptlimitno        = orgarroptlimitno + CStr(rsget("optlimitno")) + "|"
        orgarroptlimitsold      = orgarroptlimitsold + CStr(rsget("optlimitsold")) + "|"
        ''orgarroptsellyn         = orgarroptsellyn + rsget("optsellyn") + "|"
		rsget.movenext
	loop
end if
rsget.Close




dim refer, AssignedCnt, iAssignedRow
refer = request.ServerVariables("HTTP_REFERER")
AssignedCnt = 0

dim sqlStr, i, j 
 

''업체인경우 직접 수정.
if (mwdiv = "U") then
    if (limityn="Y") then
    	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
    	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
    	sqlStr = sqlStr + " , lastupdate=getdate()"+ VBCrlf      
    	    if orgSellyn <>"Y" and sellyn ="Y" and isNull(orgsellSTDate) then
	    sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf        
	        end if
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	rsget.Open sqlStr, dbget, 1

    	''옵션한정여부한정
    	sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	rsget.Open sqlStr, dbget, 1

    	for i=0 to UBound(itemoptionarr)
    		if (Len(Trim(itemoptionarr(i)))=4) then
    			if (itemoptionarr(i)="0000") then
    				sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    				sqlStr = sqlStr + " set limitno=" + optlimitnoarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " , limitsold=" + optlimitsoldarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    				rsget.Open sqlStr, dbget, 1
    			else
    				sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    				sqlStr = sqlStr + " set isusing='" + optisusingarr(i) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optsellyn='" + optisusingarr(i) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optlimitno=" + optlimitnoarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " , optlimitsold=" + optlimitsoldarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    				sqlStr = sqlStr + " and itemoption='" + Trim(itemoptionarr(i)) + "'"

    				rsget.Open sqlStr, dbget, 1
    			end if
    		end if
    	next
    else
    	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
    	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
    	sqlStr = sqlStr + " , lastupdate=getdate()"+ VBCrlf      
    	if orgSellyn <>"Y" and sellyn ="Y" and isNull(orgsellSTDate) then
	    sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf        
	    end if
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	rsget.Open sqlStr, dbget, 1


    	''옵션한정여부한정
    	sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	rsget.Open sqlStr, dbget, 1

    	for i=0 to UBound(itemoptionarr)
    		if (Len(Trim(itemoptionarr(i)))=4) then
    			if (itemoptionarr(i)="0000") then
    				sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    				sqlStr = sqlStr + " set limitno=" + optlimitnoarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " , limitsold=" + optlimitsoldarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    				rsget.Open sqlStr, dbget, 1
    			else
    				sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    				sqlStr = sqlStr + " set isusing='" + optisusingarr(i) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optsellyn='" + optisusingarr(i) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optlimitno=" + optlimitnoarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " , optlimitsold=" + optlimitsoldarr(i) + "" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    				sqlStr = sqlStr + " and itemoption='" + Trim(itemoptionarr(i)) + "'"

    				rsget.Open sqlStr, dbget, 1
    			end if
    		end if
    	next
    end if


    ''상품옵션수량
    sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
    sqlStr = sqlStr + " from (" + VBCrlf
    sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
    sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
    sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
    sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
    sqlStr = sqlStr + " ) T" + VBCrlf
    sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf

    rsget.Open sqlStr, dbget, 1

'임시 주석처리 2014-08-18 정윤정
'    sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
'    sqlStr = sqlStr + " set optionname=v.codeview" + VBCrlf
'    sqlStr = sqlStr + " from [db_item].[dbo].vw_all_option v" + VBCrlf
'    sqlStr = sqlStr + " where  [db_item].[dbo].tbl_item_option.itemid=" + CStr(itemid) + VBCrlf
'    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=v.optioncode" + VBCrlf
'
'    rsget.Open sqlStr, dbget, 1


    if (orgoptioncnt > 0) then
    	''상품한정수량
    	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
    	sqlStr = sqlStr + " ,sellyn=(CASE WHEN sellyn='Y' and IsNULL(T.usingoptionCnt,0)=0 THEN 'S' ELSE sellyn END)" + VBCrlf ''옵션 전부 사용안함으로 설정 하는 경우 //2013/09/02 추가
    	sqlStr = sqlStr + " from (" + VBCrlf
    	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold, count(*) as usingoptionCnt" + VBCrlf
    	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
    	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
    	sqlStr = sqlStr + " 	and isusing='Y' " + VBCrlf
    	sqlStr = sqlStr + " ) T" + VBCrlf
    	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf
    	rsget.Open sqlStr, dbget, 1


        sqlStr = " update [db_item].[dbo].tbl_item_option "
        sqlStr = sqlStr + " set optlimityn = [db_item].[dbo].tbl_item.limityn "   ''optsellyn = T.sellyn,
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid = " + CStr(itemid) + " "
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemid = [db_item].[dbo].tbl_item.itemid"
        rsget.Open sqlStr,dbget,1
    end if

    ''한정이 꺽이면 일시 품절처리 // 2013/09/02 추가
    sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    sqlStr = sqlStr + " set sellyn='S'" + VBCrlf
    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + VBCrlf
    sqlStr = sqlStr + " and sellyn='Y'" + VBCrlf
    sqlStr = sqlStr + " and limityn='Y'" + VBCrlf
    sqlStr = sqlStr + " and (limitno-limitsold)<1" + VBCrlf
    dbget.Execute sqlStr


''텐배인경우 수정요청
else
	limitno = 0
	limitsold = 0
	for i=0 to UBound(itemoptionarr)
		if (Len(Trim(itemoptionarr(i)))=4) then
			limitno = limitno + optlimitnoarr(i)
			limitsold = limitsold + optlimitsoldarr(i)
		end if
	next

    if (OptionExists) then
        BufOptionKey = "XXXX"
    else
        BufOptionKey = "0000"
    end if

    ' 상품정보 = 따로 수정 옵션이 없을 경우 0000 있을경우 XXXX XXXX 인경우
    if (orgsellyn<>sellyn) or (orglimityn<>limityn)  or (CStr(orglimitno)<>CStr(limitno)) or (CStr(orglimitsold)<>CStr(limitsold)) then
    	sqlStr = "insert into [db_temp].[dbo].tbl_upche_itemedit" + VbCrlf
    	sqlStr = sqlStr + " (itemid,itemoption, "
    	sqlStr = sqlStr + " oldsellyn, oldlimityn, oldlimitno, oldlimitsold, " + VbCrlf
    	sqlStr = sqlStr + " sellyn, limityn,limitno, limitsold, " + VbCrlf
    	sqlStr = sqlStr + " isupchebeasong, isfinish, etcstr,edittype)" + VbCrlf
    	sqlStr = sqlStr + " values(" + itemid + ",'" & BufOptionKey & "'," + VbCrlf
    	sqlStr = sqlStr + " '" + orgsellyn + "',"  + VbCrlf
    	sqlStr = sqlStr + " '" + orglimityn + "',"  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(orglimitno) + ","  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(orglimitsold) + ","  + VbCrlf
    	sqlStr = sqlStr + " '" + sellyn + "',"  + VbCrlf
    	sqlStr = sqlStr + " '" + limityn + "',"  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(limitno) + ","  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(limitsold) + ","  + VbCrlf
    	sqlStr = sqlStr + " 'N',"  + VbCrlf
    	sqlStr = sqlStr + " 'N','" + html2db(reqstring) + "','E')"  + VbCrlf
    	'response.write sqlStr

    	dbget.Execute sqlStr, iAssignedRow

    	AssignedCnt = AssignedCnt + iAssignedRow
    end if


    if (OptionExists) then
        '옵션정보???
        'TODO : dispyn 필드를 isusing 으로 사용한다.
        'TODO : 옵션명도 변경할수 있도록 수정하면 좋다.
        orgarritemoption        = Split(orgarritemoption, "|")
        orgarritemoptionname    = Split(orgarritemoptionname, "|")
        orgarrisusing           = Split(orgarrisusing, "|")
        orgarroptlimityn        = Split(orgarroptlimityn, "|")
        orgarroptlimitno        = Split(orgarroptlimitno, "|")
        orgarroptlimitsold      = Split(orgarroptlimitsold, "|")

        ''orgarroptsellyn         = Split(orgarroptsellyn, "|")

        for i = 0 to UBound(orgarritemoption) - 1
            index = -1
            iAssignedRow = 0
            if (Trim(orgarritemoption(i)) <> "") then
                for j = 0 to UBound(itemoptionarr) - 1
                    if ((Trim(orgarritemoption(i)) = Trim(itemoptionarr(j))) and (Trim(itemoptionarr(j)) <> "0000")) then
                        index = j
                        exit for
                    end if
                next

                if (index <> -1) then
                    ''변경내역이 있는경우만. 옵션 만
                    if (orgarrisusing(i)<>optisusingarr(index)) or (orglimityn<>limityn) or (CStr(orgarroptlimitno(i))<>CStr(optlimitnoarr(index))) or (CStr(orgarroptlimitsold(i))<>CStr(optlimitsoldarr(index))) then
                    	'response.write "1" & orgarrisusing(i) & "," & optisusingarr(index) & "<br>"
                    	'response.write "1" & orglimitno & "," & limityn & "<br>"
                    	'response.write "1" & orgarroptlimitno(i) & "," & optlimitnoarr(index) & "<br>"
                    	'response.write "1" & orgarroptlimitsold(i) & "," & optlimitsoldarr(index) & "<br>"


                    	sqlStr = "insert into [db_temp].[dbo].tbl_upche_itemedit" + VbCrlf
                    	sqlStr = sqlStr + " (itemid, itemoption, itemoptionname, oldsellyn, olddispyn, " + VbCrlf
                    	sqlStr = sqlStr + " oldlimityn, oldlimitno, oldlimitsold, sellyn, dispyn, limityn," + VbCrlf
                    	sqlStr = sqlStr + " limitno, limitsold, isupchebeasong, isfinish, etcstr,edittype)" + VbCrlf
                    	sqlStr = sqlStr + " values(" + itemid + ",'" + Trim(orgarritemoption(i)) + "','" + Trim(html2db(orgarritemoptionname(i))) + "'," + VbCrlf
                    	sqlStr = sqlStr + " '" + orgsellyn + "',"  + VbCrlf         ''상품판매여부
                    	sqlStr = sqlStr + " '" + orgarrisusing(i) + "',"  + VbCrlf
                    	sqlStr = sqlStr + " '" + orglimityn + "',"  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(orgarroptlimitno(i)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(orgarroptlimitsold(i)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " '" + sellyn + "',"  + VbCrlf
                    	sqlStr = sqlStr + " '" + optisusingarr(index) + "',"  + VbCrlf
                    	sqlStr = sqlStr + " '" + limityn + "',"  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(optlimitnoarr(index)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(optlimitsoldarr(index)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " 'N',"  + VbCrlf
                    	sqlStr = sqlStr + " 'N','" + html2db(reqstring) + "','E')"  + VbCrlf
                    	dbget.Execute sqlStr, iAssignedRow
                        AssignedCnt = AssignedCnt + iAssignedRow
                    end if
                end if
            end if
        next
    End IF


    if (AssignedCnt<1) then
        response.write "<script>alert('Err - 변경된 내역이 없습니다.\n\n 판매가나 기타 변경사항을 요청 하실 경우\n\n 담당MD에게 직접 문의해 주세요.');</script>"
        response.write "<script>location.replace('" + refer + "');</script>"
    else
        response.write "<script>alert('수정요청되었습니다.\n\n 실제 수정은 다음날 텐바이텐에서 확인후 반영됩니다.');</script>"
        response.write "<script>location.replace('" + refer + "');</script>"
    end if
    dbget.close()	:	response.End
end if

%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
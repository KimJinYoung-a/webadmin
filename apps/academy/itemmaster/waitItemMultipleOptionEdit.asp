<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Session.codepage="65001"
Response.codepage="65001"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
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
end Function

function ReMatchMultiOption(itemid)
    dim sqlStr
    dim MultiLevel, itemLimitYn
    
    MultiLevel = 0

	''업체배송인경우 입출/판매 관계없이 삭제
	sqlStr = " select limityn, deliverytype "
	sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item"
	sqlStr = sqlStr & " where itemid=" & CStr(itemid)

	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if not rsACADEMYget.EOF then
		itemLimitYn = rsACADEMYget("limityn")
	end if
	rsACADEMYget.Close

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

function WaitRegDoubleOptionProc(waititemid)
    dim found, foundcount, optionName, i,j,k
    
    dim optionTypename1, optionTypename2, optionTypename3
    optionTypename1 = Trim(Request.Form("opttypename1"))
    optionTypename2 = Trim(Request.Form("opttypename2"))
    optionTypename3 = Trim(Request.Form("opttypename3"))

	dim Lv1cnt, Lv2cnt, Lv3cnt 
	Lv1cnt = Request.Form("optname1").Count
	Lv2cnt = Request.Form("optname2").Count
	Lv3cnt = Request.Form("optname3").Count

    dim Val1cnt, Val2cnt, Val3cnt
    dim optName1, optName2, optName3
    dim Valid1, Valid2, Valid3
    dim buf, ErrMsg, AssignedOption
	Dim optprice1, optprice2, optprice3
	Dim optbuyprice1, optbuyprice2, optbuyprice3
    
    Val1cnt = 0
    Val2cnt = 0
    Val3cnt = 0

	redim arroptionName1(Request.Form("optname1").Count)
	redim arroptaddprice1(Request.Form("optaddprice1").Count)
	redim arroptaddbuyprice1(Request.Form("optbuyprice1").Count)
	redim arroptionName2(Request.Form("optname2").Count)
	redim arroptaddprice2(Request.Form("optaddprice2").Count)
	redim arroptaddbuyprice2(Request.Form("optbuyprice2").Count)
	redim arroptionName3(Request.Form("optname3").Count)
	redim arroptaddprice3(Request.Form("optaddprice3").Count)
	redim arroptaddbuyprice3(Request.Form("optbuyprice3").Count)

    For i=1 To Lv1cnt
		arroptionName1(i) = Request.Form("optname1")(i)
		arroptaddprice1(i) = Request.Form("optaddprice1")(i)
		arroptaddbuyprice1(i) = Request.Form("optbuyprice1")(i)
        buf = Trim(arroptionName1(i))
        if Len(buf)>0 then Val1cnt = Val1cnt + 1
    Next

    For i=1 To Lv2cnt
		arroptionName2(i) = Request.Form("optname2")(i)
		arroptaddprice2(i) = Request.Form("optaddprice2")(i)
		arroptaddbuyprice2(i) = Request.Form("optbuyprice2")(i)
        buf = Trim(arroptionName2(i))
        if Len(buf)>0 then Val2cnt = Val2cnt + 1
    Next

    For i=1 To Lv3cnt
		arroptionName3(i) = Request.Form("optname3")(i)
		arroptaddprice3(i) = Request.Form("optaddprice3")(i)
		arroptaddbuyprice3(i) = Request.Form("optbuyprice3")(i)
        buf = Trim(arroptionName3(i))
        if Len(buf)>0 then Val3cnt = Val3cnt + 1
    Next

    If (optionTypename1=optionTypename2) or (optionTypename1=optionTypename3) or (optionTypename2=optionTypename3) then
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
    
    Dim sqlStr
    foundcount=0

	If Lv1cnt > 0 Then
		'################### 기존 등록 옵션 삭제 #######################################################
		sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option"
		sqlStr = sqlStr & " where itemid=" & waititemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		'################### 기존 등록 옵션 삭제 #######################################################
		sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option_Multiple"
		sqlStr = sqlStr & " where itemid=" & waititemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
	End If
	If Lv3cnt=0 Then Lv3cnt=1
    ''0번은 입력 없음. N까지
    For i=0 To Lv1cnt-1
        For j=0 To Lv2cnt-1
            For k=0 To Lv3cnt-1
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
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(waititemid)  
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename1) + "' and optionKindName='" + html2db(optName1) + "'))"

                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice, optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & waititemid 
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
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(waititemid)  
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename2) + "' and optionKindName='" + html2db(optName2) + "'))"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice,optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & waititemid 
                                sqlStr = sqlStr + " ,2"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(j+1)) &"'"
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
                            sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(waititemid)  
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename3) + "' and optionKindName='" + html2db(optName3) + "'))"
                            
                            rsACADEMYget.Open sqlStr,dbACADEMYget,1
                            if Not rsACADEMYget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsACADEMYget.Close
                            
                            if (Not found) then
                                sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option_Multiple"
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName,optaddprice,optaddbuyprice)"
                                sqlStr = sqlStr + " values("
                                sqlStr = sqlStr + " " & waititemid 
                                sqlStr = sqlStr + " ,3"
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(k+1)) &"'"
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
                        sqlStr = " select itemid from db_academy.dbo.tbl_diy_wait_item_option "
                        sqlStr = sqlStr + " where itemid = " + CStr(waititemid)  
                        sqlStr = sqlStr + " and ((itemoption = '" + CStr(AssignedOption) + "') or (optionTypeName='복합옵션' and optionname='" + html2db(optionName) + "'))"
                        
                        rsACADEMYget.Open sqlStr,dbACADEMYget,1
                        if Not rsACADEMYget.Eof then
                            found = true
                        end if
                        rsACADEMYget.Close
                        
                        if (Not found) then 
                            sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                            sqlStr = sqlStr + " values(" + CStr(waititemid) + ", '" + CStr(AssignedOption) + "', '복합옵션', '" + CStr(html2db(optionName)) + "', 'Y', 'Y', '" + Request.Form("limityn") + "', " & Request.Form("limitno") & ", 0) "
                            
                            dbACADEMYget.Execute sqlStr
                            ''response.write AssignedOption + ":" +  optName1 + "," + optName2 + "," + optName3 + "<BR>"
                        end if
                    end if
                    
                end if
            Next
        Next
    Next
	'' 2차옵션->3차로 변경  등..
    Call ReMatchMultiOption(waititemid)
End Function

Dim waititemid, iErrMsg, makerid
waititemid = requestCheckVar(request("waititemid"),10)
makerid = request.cookies("partner")("userid")

If waititemid = "" Then
	dim DesignerID, sqlStr
	DesignerID = Request.Form("designerid")
	'###########################################################################
	'상품 데이터 입력
	'###########################################################################
	sqlStr = "insert into db_academy.dbo.tbl_diy_wait_item" + vbCrlf
	sqlStr = sqlStr & " (itemdiv,makerid,itemname,regdate,buycash, sellcash, mileage, sellyn, deliverytype,limityn,currstate)" + vbCrlf
	sqlStr = sqlStr & " values(" + vbCrlf
	sqlStr = sqlStr & "'01'" + vbCrlf
	sqlStr = sqlStr & ",'" & DesignerID & "'" + vbCrlf
	sqlStr = sqlStr & ",'tempitem'" + vbCrlf
	sqlStr = sqlStr & ",getdate()" + vbCrlf
	sqlStr = sqlStr & ",0" + vbCrlf
	sqlStr = sqlStr & ",0" + vbCrlf
	sqlStr = sqlStr & ",0" + vbCrlf
	sqlStr = sqlStr & ",'N'" + vbCrlf
	sqlStr = sqlStr & ",'9'" + vbCrlf
	sqlStr = sqlStr & ",'N'" + vbCrlf
	sqlStr = sqlStr & ",3)" + vbCrlf
	'Response.write sqlStr
	'Response.end
	dbACADEMYget.Execute sqlStr
	'###########################################################################
	'상품 아이디 가져오기
	'###########################################################################
	sqlStr = "Select IDENT_CURRENT('db_academy.dbo.tbl_diy_wait_item') as maxitemid "
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		waititemid = rsACADEMYget("maxitemid")
	rsACADEMYget.close
End If

If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If

iErrMsg = WaitRegDoubleOptionProc(waititemid)
if (iErrMsg="") then
%>
<script type="text/javascript">
<!--
	parent.fnMultipleOptionEditEnd("<%=waititemid%>");
//-->
</script>
<%
Else
%>
<script type="text/javascript">
<!--
	alert('<%=iErrMsg%>');
//-->
</script>
<%
end if
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
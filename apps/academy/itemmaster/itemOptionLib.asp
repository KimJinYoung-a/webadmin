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

function WaitRegDoubleOptionProc(waititemid)
    dim found, foundcount, optionName, i,j,k
    
    dim optionTypename1, optionTypename2, optionTypename3
    dim itemoption1, itemoption2, itemoption3
    dim optionName1, optionName2, optionName3
	Dim optaddprice1, optaddprice2,optaddprice3
	Dim optaddbuyprice1,optaddbuyprice2,optaddbuyprice3
    
    optionTypename1 = Trim(Request.Form("optionTypename1"))
    optionTypename2 = Trim(Request.Form("optionTypename2"))
    optionTypename3 = Trim(Request.Form("optionTypename3"))
    optionName1     = Request.Form("optionName1") & ","
    optionName2     = Request.Form("optionName2") & ","
    optionName3     = Request.Form("optionName3") & ","
	optaddprice1     = Request.Form("optaddprice1") & ","
	optaddprice2     = Request.Form("optaddprice2") & ","
	optaddprice3     = Request.Form("optaddprice3") & ","
	optaddbuyprice1     = Request.Form("optaddbuyprice1") & ","
	optaddbuyprice2     = Request.Form("optaddbuyprice2") & ","
	optaddbuyprice3     = Request.Form("optaddbuyprice3") & ","

	Dim arroptionName1,arroptionName2,arroptionName3
	Dim arroptaddprice1,arroptaddprice2,arroptaddprice3
	Dim arroptaddbuyprice1,arroptaddbuyprice2,arroptaddbuyprice3

	arroptionName1 = split(optionName1,",")
	arroptionName2 = split(optionName2,",")
	arroptionName3 = split(optionName3,",")
	arroptaddprice1 = split(optaddprice1,",")
	arroptaddprice2 = split(optaddprice2,",")
	arroptaddprice3 = split(optaddprice3,",")
	arroptaddbuyprice1 = split(optaddbuyprice1,",")
	arroptaddbuyprice2 = split(optaddbuyprice2,",")
	arroptaddbuyprice3 = split(optaddbuyprice3,",")

    dim Lv1cnt, Lv2cnt, Lv3cnt
    dim Val1cnt, Val2cnt, Val3cnt
    dim option1, option2, option3
    dim optName1, optName2, optName3
    dim Valid1, Valid2, Valid3
    dim buf, ErrMsg, AssignedOption
	Dim optprice1, optprice2, optprice3
	Dim optbuyprice1, optbuyprice2, optbuyprice3
    
    Lv1cnt = ubound(arroptionName1)
    Lv2cnt = ubound(arroptionName2)
    Lv3cnt = ubound(arroptionName3)
    
    Val1cnt = 0
    Val2cnt = 0
    Val3cnt = 0
    
    for i=0 to Lv1cnt-1
        buf = Trim(arroptionName1(i))
        if Len(buf)>0 then Val1cnt = Val1cnt + 1
    next
    
    for i=0 to Lv2cnt-1
        buf = Trim(arroptionName2(i))
        if Len(buf)>0 then Val2cnt = Val2cnt + 1
    next
    
    for i=0 to Lv3cnt-1
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
		sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option"
		sqlStr = sqlStr & " where itemid=" & waititemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		'################### 기존 등록 옵션 삭제 #######################################################
		sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option_Multiple"
		sqlStr = sqlStr & " where itemid=" & waititemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
	End If

    ''0번은 입력 없음. N까지
    for i=0 to Lv1cnt-1
        for j=0 to Lv2cnt-1
            for k=0 to Lv3cnt-1
                optName1 = Trim(arroptionName1(i))
				optprice1  = Trim(arroptaddprice1(i))
				optbuyprice1  = Trim(arroptaddbuyprice1(i))

                If (Lv2cnt) Then
					optName2 = Trim(arroptionName2(j))
					optprice2  = Trim(arroptaddprice2(j))
					optbuyprice2  = Trim(arroptaddbuyprice2(j))
				Else
					optName2 = ""
					optprice2  = ""
					optbuyprice2  = ""
				End If
				If (Lv3cnt) Then
					optName3 = Trim(arroptionName3(k))
					optprice3  = Trim(arroptaddprice3(k))
					optbuyprice3  = Trim(arroptaddbuyprice3(k))
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
                            sqlStr = sqlStr + " values(" + CStr(waititemid) + ", '" + CStr(AssignedOption) + "', '복합옵션', '" + CStr(html2db(optionName)) + "', 'Y', 'Y', '" + Request.Form("limityn") + "', " & limitno & ", 0) "
                            
                            dbACADEMYget.Execute sqlStr
                            ''response.write AssignedOption + ":" +  optName1 + "," + optName2 + "," + optName3 + "<BR>"
                        end if
                    end if
                    
                end if
            next
        next
    next
end function
%>

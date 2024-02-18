<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/upchejungsan/upchejungsan_function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,checkone,id, ipkumregdate, checkoneEx, jgubun
id          = request("id")
mode        = request("mode")
checkone    = request.form("checkone")
ipkumregdate = request("ipkumregdate")
checkoneEx = request.form("checkoneEx")
jgubun      = request("jgubun")

if (checkone="") then checkone="0"
if (checkoneEx="") then checkoneEx="0"
Dim isMixedFile : isMixedFile= (request.form("ck_Mibus")="CX")

Dim reqIcheDate   : reqIcheDate   = requestCheckVar(request("reqIcheDate"),10)

Dim UseUpFile : UseUpFile = requestCheckVar(request("UseUpFile"),10)
Dim ipFileNo : ipFileNo = requestCheckVar(request("ipFileNo"),10)

Dim firstSel  : firstSel = requestCheckVar(request("firstSel"),10)
Dim secondSel : secondSel = requestCheckVar(request("secondSel"),10)
Dim thirdSel : thirdSel = requestCheckVar(request("thirdSel"),10)
Dim ipFileDIdx: ipFileDIdx = requestCheckVar(request("ipFileDIdx"),10)

IF (ipFileNo="") then ipFileNo=0

Dim ipFileState, ipFileGbn

dim sqlstr, AssignedRow, AssignedRow2
Dim targetGbn : targetGbn = requestCheckVar(request("targetGbn"),10)
Dim targetGbnEx : targetGbnEx = "OF"
Dim ipFileName

Dim retMakerId, retGroupid
Dim retipFileNo, rettargetIdx
Dim NotReqgroupIdList0,NotReqgroupIdList1,NotReqgroupIdList2

Dim IsForce : IsForce=FALSE

Dim IsExtJExists : IsExtJExists = (checkoneEx<>"0")
Dim retCnt

'rw checkone
'rw checkoneEx
'rw mode
'response.end

if mode="bankingupload" then ''신규방식
    IF (UseUpFile<>"") THEN

        IF (ipFileNo>0) then
            sqlstr = "select ipFileState, jgubun from db_jungsan.dbo.tbl_jungsan_ipkumFile_MASTER "
            sqlstr = sqlstr + " where ipFileNo="&ipFileNo
            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
                ipFileState = rsget("ipFileState")
                jgubun      = rsget("jgubun")
            end if
    	    rsget.Close

    	    if (ipFileState>1) then
    	        response.write "수정 불가 - 파일 진행 상태 오류 [FileNo:"&ipFileNo&" : State :"&ipFileState&"]"
                response.end
    	    end if
        ENd IF


        retipFileNo = 0
        sqlstr = " select top 1 S.ipFileNo, S.targetIdx, m.designerid, m.groupid From db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
        sqlstr = sqlstr + " Join  [db_jungsan].[dbo].tbl_designer_jungsan_master M"
        sqlstr = sqlstr + " on S.targetIdx=M.id"
        sqlstr = sqlstr + " where S.targetGbn='"&targetGbn&"'"
        sqlstr = sqlstr + " and S.targetIdx in (" + checkone + ")"

        rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			retipFileNo = rsget("ipFileNo")
			rettargetIdx = rsget("targetIdx")
			retMakerId = rsget("designerid")
			retGroupid = rsget("groupid")
		end if
		rsget.Close

		IF (retipFileNo<>0) then
		    response.write "브랜드ID "&retMakerId&"["&retGroupid&"] 이미 업로드 된 내역을 다시 올릴 수 없음."
            response.end
		End IF

		''' 오프. 기 업로드 내역 Check
		sqlstr = " select top 1 S.ipFileNo, S.targetIdx, m.makerid, m.groupid From db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
        sqlstr = sqlstr + " Join  [db_jungsan].[dbo].tbl_off_jungsan_master M"
        sqlstr = sqlstr + " on S.targetIdx=M.idx"
        sqlstr = sqlstr + " where S.targetGbn='"&targetGbnEx&"'"
        sqlstr = sqlstr + " and S.targetIdx in (" + checkoneEx + ")"

        rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			retipFileNo = rsget("ipFileNo")
			rettargetIdx = rsget("targetIdx")
			retMakerId = rsget("makerid")
			retGroupid = rsget("groupid")
		end if
		rsget.Close

		IF (retipFileNo<>0) then
		    response.write "브랜드ID "&retMakerId&"["&retGroupid&"] 이미 업로드 된 내역을 다시 올릴 수 없음."
            response.end
		End IF

		''정산구분 check-----------------------------------------------------------------------------------------
		if (jgubun<>"") then
    		retipFileNo = 0
            sqlstr = " select top 1 M.id, m.designerid, m.groupid From [db_jungsan].[dbo].tbl_designer_jungsan_master M"
            sqlstr = sqlstr + " where M.id in (" + checkone + ")"
            sqlstr = sqlstr + " and M.jgubun<>'"&jgubun&"'"

            rsget.Open sqlStr,dbget,1
    		if Not rsget.Eof then
    			retipFileNo = rsget("id")
    			retMakerId = rsget("designerid")
    			retGroupid = rsget("groupid")
    		end if
    		rsget.Close

    		IF (retipFileNo<>0) then
    		    response.write "브랜드ID "&retMakerId&"["&retGroupid&"] 정산구분이 올바르지 않음."
                response.end
    		End IF

    		''' 오프. 기 업로드 내역 Check
    		sqlstr = " select top 1 M.idx, m.makerid, m.groupid From [db_jungsan].[dbo].tbl_off_jungsan_master M"
            sqlstr = sqlstr + " where M.idx in (" + checkoneEx + ")"
            sqlstr = sqlstr + " and M.jgubun<>'"&jgubun&"'"

            rsget.Open sqlStr,dbget,1
    		if Not rsget.Eof then
    			retipFileNo = rsget("idx")
    			retMakerId = rsget("makerid")
    			retGroupid = rsget("groupid")
    		end if
    		rsget.Close

    		IF (retipFileNo<>0) then
    		    response.write "브랜드ID "&retMakerId&"["&retGroupid&"] 정산구분이 올바르지 않음."
                response.end
    		End IF
    	end if
		''----------------------------------------------------------------------------------------------------------
		''마이너스 체크 // 같은거래처(groupid)에 마이너스 금액이 남아 있을경우 올라가지 못하게.
        sqlstr = " select Top 100 g.Groupid from [db_jungsan].[dbo].tbl_designer_jungsan_master g"
        sqlstr = sqlstr + " where g.id in (" + checkone + ")"
        sqlstr = sqlstr + " and g.Groupid in ("
        sqlstr = sqlstr + "     select M.Groupid from [db_jungsan].[dbo].tbl_designer_jungsan_master M"
        sqlstr = sqlstr + "     where M.id not in (" + checkone + ")"
        sqlstr = sqlstr + "     and M.bankingupflag ='N'"
        sqlstr = sqlstr + "     and M.finishFlag=3"
        sqlstr = sqlstr + "     and (ub_totalsuplycash+ me_totalsuplycash+wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash)<0"
        sqlstr = sqlstr + " ) and ( g.id in (" + checkone + ") or (IsNULL(g.ipkum_acctno,'')='')  )"

        rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			Do Until rsget.Eof
    			NotReqgroupIdList0 = NotReqgroupIdList0 + rsget("Groupid") +","
    			rsget.movenext
			loop
		end if
		rsget.Close

		IF (NotReqgroupIdList0<>"") and (Not IsForce) then
		    response.write " 정산 확정내역중 마이너스 금액 / 또는 계좌 오류 있음. : " & NotReqgroupIdList0 & "<br>" &"["&checkone&"]"
            response.end  '' 주석제거 (2017/01/19)
		End IF


		''업로드 총금액이 그룹별로 마이너스가 될 수 없음 // 기 업로드 내역 포함 조사
		sqlstr = "select Groupid, Sum(jungsanSum) from ("
        sqlstr = sqlstr + "     select groupid, (ub_totalsuplycash+ me_totalsuplycash+wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash) as jungsanSum"
        sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlstr = sqlstr + "     where  id  in (" + checkone + ")"
        sqlstr = sqlstr + " Union ALL "
        sqlstr = sqlstr + "     select m2.groupid, m2.tot_jungsanprice as jungsanSum "
        sqlstr = sqlstr + "     from  [db_jungsan].[dbo].tbl_off_jungsan_master m2"
        sqlstr = sqlstr + "     where m2.idx in (" + checkoneEx + ")"
        sqlstr = sqlstr + " Union ALL "
        sqlstr = sqlstr + "     select J.groupid, (J.ub_totalsuplycash+ J.me_totalsuplycash+J.wi_totalsuplycash+J.sh_totalsuplycash+J.et_totalsuplycash+J.dlv_totalsuplycash) as jungsanSum"
        sqlstr = sqlstr + "     from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail P1"
	    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_designer_jungsan_master J"
	    sqlstr = sqlstr + "     on P1.ipFileNo="&ipFileNo
	    sqlstr = sqlstr + "     and P1.targetGbn='ON'"
	    sqlstr = sqlstr + "     and P1.targetIdx=J.id"
        sqlstr = sqlstr + " Union ALL "
        sqlstr = sqlstr + "     select F.groupid, tot_jungsanprice as jungsanSum"
        sqlstr = sqlstr + "     from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail P2"
	    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_off_jungsan_master F"
	    sqlstr = sqlstr + "     on P2.ipFileNo="&ipFileNo
	    sqlstr = sqlstr + "     and P2.targetGbn='OF'"
	    sqlstr = sqlstr + "     and P2.targetIdx=F.idx"
        sqlstr = sqlstr + " ) T group by T.groupid"
        sqlstr = sqlstr + " having Sum(jungsanSum)<0"
''rw sqlstr
        rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			Do Until rsget.Eof
    			NotReqgroupIdList1 = NotReqgroupIdList1 + rsget("Groupid") +","
    			rsget.movenext
			loop
		end if
		rsget.Close

		IF (NotReqgroupIdList1<>"") and (Not IsForce) then
		    response.write "업로드 총금액이 마이너스가 될 수 없음 "&NotReqgroupIdList1&"  합계금액 마이너스 있음."
            response.end '' 주석제거 (2017/01/19)
		End IF

        '''일단 사용안함..//// 원복해야함..
        ''if (FALSE) and (Not IsExtJExists) THEN
        if (Not IsExtJExists) THEN
    		''오프 체크 // 온라인만 올릴경우
    		sqlstr = " select Top 100 designerId, Groupid from [db_jungsan].[dbo].tbl_designer_jungsan_master"
            sqlstr = sqlstr + " where id in (" + checkone + ")"
            sqlstr = sqlstr + " and Groupid in ("
            sqlstr = sqlstr + "     select M.Groupid from [db_jungsan].[dbo].tbl_off_jungsan_master M"
            sqlstr = sqlstr + "     where 1=1"
            sqlstr = sqlstr + "     and M.bankingupflag ='N'"
            sqlstr = sqlstr + "     and M.finishFlag=3"
            sqlstr = sqlstr + "     and (tot_jungsanprice)<0"
     sqlstr = sqlstr + "     and M.Groupid not in ('G05289')"  ''2018/01/05 예외처리 ''닌텐도 사입 선정산
            sqlstr = sqlstr + " )"

            rsget.Open sqlStr,dbget,1
    		if Not rsget.Eof then
    			Do Until rsget.Eof
        			NotReqgroupIdList2 = NotReqgroupIdList2 + rsget("Groupid") +","
        			rsget.movenext
    			loop
    		end if
    		rsget.Close

    		IF (NotReqgroupIdList2<>"") and (Not IsForce) then
    		    response.write "브랜드ID "&NotReqgroupIdList2&" <b>오프라인 정산</b> 확정내역중 마이너스 금액 있음. - 업체 제외 되고 업로드"
                ''response.end  '' 주석제거 (2017/01/19) 주석처리 (2017/02/22)
    		End IF
		end if

        IF (ipFileNo=0) then
            IF (targetGbn="ON") then
                ipFileName = "온라인 "
            ELSEIF (targetGbn="OF") then
                ipFileName = "오프라인 "
            ELSEIF (targetGbn="WN") then
                ipFileName = "아카데미(원천징수) "
            ELSE
                ipFileName = targetGbn
            END IF

            ipFileName = ipFileName & CHKIIF(jgubun="CC"," 수수료 정산","")
            ipFileName = ipFileName & CHKIIF(jgubun="MM"," 매입 정산","")

            ipFileName = ipFileName & " " & reqIcheDate & " 정기결제"
            if (isMixedFile) then ipFileName=ipFileName& " (상계처리)"

            sqlStr = "Insert into db_jungsan.dbo.tbl_jungsan_ipkumFile_Master"
            sqlStr = sqlStr & " (ipFileName,ipFileRegdate,ipFileState, ipfilegbn,ReqDate,jgubun)"
            sqlStr = sqlStr & " values ('"&ipFileName&"',getdate(),0,'"&targetGbn&"','"&reqIcheDate&"','"&jgubun&"')"
            dbget.Execute sqlStr

            sqlStr = "select IDENT_CURRENT('db_jungsan.dbo.tbl_jungsan_ipkumFile_Master') as ipFileNo"
			rsget.Open sqlStr,dbget,1
            IF Not rsget.Eof THEN
                ipFileNo = rsget("ipFileNo")
            ENd IF
            rsget.Close
        ENd IF

        NotReqgroupIdList0=Trim(NotReqgroupIdList0)
        NotReqgroupIdList1=Trim(NotReqgroupIdList1)
        NotReqgroupIdList2=Trim(NotReqgroupIdList2)

        if (Right(NotReqgroupIdList0,1)=",") then NotReqgroupIdList0=Left(NotReqgroupIdList0,Len(NotReqgroupIdList0)-1)
        if (Right(NotReqgroupIdList1,1)=",") then NotReqgroupIdList1=Left(NotReqgroupIdList1,Len(NotReqgroupIdList1)-1)
        if (Right(NotReqgroupIdList2,1)=",") then NotReqgroupIdList2=Left(NotReqgroupIdList2,Len(NotReqgroupIdList2)-1)

        NotReqgroupIdList0 = Replace(NotReqgroupIdList0,",","','")
        NotReqgroupIdList1 = Replace(NotReqgroupIdList1,",","','")
        NotReqgroupIdList2 = Replace(NotReqgroupIdList2,",","','")

        NotReqgroupIdList0 = "'"&NotReqgroupIdList0&"'"
        NotReqgroupIdList1 = "'"&NotReqgroupIdList1&"'"
        NotReqgroupIdList2 = "'"&NotReqgroupIdList2&"'"


        sqlstr = " insert into db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
        sqlstr = sqlstr & " (ipFileNo,targetGbn,targetIdx,regdate)"
        sqlstr = sqlstr & " select "&ipFileNo
        IF (targetGbn="WN") THEN
            sqlstr = sqlstr & " ,'ON'"
        ELSE
            sqlstr = sqlstr & " ,'"&targetGbn&"'"
        END IF
        sqlstr = sqlstr & " ,M.id"
        sqlstr = sqlstr & " ,getdate()"
        sqlstr = sqlstr & " From [db_jungsan].[dbo].tbl_designer_jungsan_master M"
        sqlstr = sqlstr & "     left Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
        sqlstr = sqlstr & "     on M.id=S.targetIdx"
        sqlstr = sqlstr & "     and S.targetGbn='"&targetGbn&"'"
        sqlstr = sqlstr + " where M.id in (" + checkone + ")"
        sqlstr = sqlstr + " and M.bankingupflag ='N'"
        sqlstr = sqlstr + " and S.ipFileNo Is NULL"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"
        dbget.Execute sqlStr, AssignedRow
''rw sqlstr
        sqlstr = "update M "
    	sqlstr = sqlstr + " set bankingupflag='Y'"
    	sqlstr = sqlstr + " ,ipkum_bank=isNULL(A.jungsan_bank,G.jungsan_bank)"          ''2016/12/15 수정
    	sqlstr = sqlstr + " ,ipkum_acctno=isNULL(A.jungsan_acctno,G.jungsan_acctno)"    ''2016/12/15 수정
    	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master M"
    	sqlstr = sqlstr + "     Join [db_partner].[dbo].tbl_partner_group G"
    	sqlstr = sqlstr + "     on M.groupid=G.groupid"
    	sqlstr = sqlstr + "     left join db_partner.dbo.tbl_partner_addJungsanInfo A" ''2016/12/15 추가
    	sqlstr = sqlstr + "     on M.designerid=A.partnerid"
    	sqlstr = sqlstr + " where M.id in (" + checkone + ")"
    	sqlstr = sqlstr + " and M.bankingupflag ='N'"
    	sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"

    	dbget.Execute sqlStr, AssignedRow2

        IF (IsExtJExists) THEN
            Dim PAssignedRow, PAssignedRow2
            PAssignedRow = AssignedRow
            PAssignedRow2 = AssignedRow2

            sqlstr = " insert into db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
            sqlstr = sqlstr & " (ipFileNo,targetGbn,targetIdx,regdate)"
            sqlstr = sqlstr & " select "&ipFileNo
            sqlstr = sqlstr & " ,'"&targetGbnEx&"'"
            sqlstr = sqlstr & " ,M.idx"
            sqlstr = sqlstr & " ,getdate()"
            sqlstr = sqlstr & " From [db_jungsan].[dbo].tbl_off_jungsan_master M"
            sqlstr = sqlstr & "     left Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
            sqlstr = sqlstr & "     on M.idx=S.targetIdx"
            sqlstr = sqlstr & "     and S.targetGbn='"&targetGbnEx&"'"
            sqlstr = sqlstr + " where M.idx in (" + checkoneEx + ")"
            sqlstr = sqlstr + " and M.bankingupflag ='N'"
            sqlstr = sqlstr + " and S.ipFileNo Is NULL"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"
            dbget.Execute sqlStr, AssignedRow

            sqlstr = "update M "
        	sqlstr = sqlstr + " set bankingupflag='Y'"
        	sqlstr = sqlstr + " ,ipkum_bank=isNULL(A.jungsan_bank,G.jungsan_bank)"
        	sqlstr = sqlstr + " ,ipkum_acctno=isNULL(A.jungsan_acctno,G.jungsan_acctno)"
        	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_off_jungsan_master M"
        	sqlstr = sqlstr + "     Join [db_partner].[dbo].tbl_partner_group G"
        	sqlstr = sqlstr + "     on M.groupid=G.groupid"
        	sqlstr = sqlstr + "     left join db_partner.dbo.tbl_partner_addJungsanInfo A"      ''2016/12/15 추가
        	sqlstr = sqlstr + "     on M.makerid=A.partnerid"
        	sqlstr = sqlstr + " where M.idx in (" + checkoneEx + ")"
        	sqlstr = sqlstr + " and M.bankingupflag ='N'"
        	sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"

        	dbget.Execute sqlStr, AssignedRow2

        	AssignedRow = AssignedRow + PAssignedRow
        	AssignedRow2 = AssignedRow2 + PAssignedRow2
        end if

''response.end

        response.write "<script language='javascript'>"
        response.write "alert('저장 되었습니다. - "& AssignedRow &"/"&AssignedRow2&" 건...');"
        IF (NotReqgroupIdList0&NotReqgroupIdList1&NotReqgroupIdList2="''''''") then
            response.write "location.replace('"&refer&"');"
        End IF
        response.write "</script>"

        response.end
    ENd If

    rw "오류 / 기존방식 사용불가?"
    response.end
    '''기존방식.
	sqlstr = "update M "
	sqlstr = sqlstr + " set bankingupflag='Y'"
	sqlstr = sqlstr + " ,ipkum_bank=G.jungsan_bank"
	sqlstr = sqlstr + " ,ipkum_acctno=G.jungsan_acctno"
	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master M"
	sqlstr = sqlstr + "     Join [db_partner].[dbo].tbl_partner_group G"
	sqlstr = sqlstr + "     on M.groupid=G.groupid"
	sqlstr = sqlstr + " where M.id in (" + checkone + ")"
	sqlstr = sqlstr + " and M.bankingupflag ='N'"

	dbget.Execute sqlStr, AssignedRow

elseif mode="delflagWF" then
    ''ipFileDIdx
    retCnt=0

    ''이곳에서 db_partner.dbo.tbl_eAppPayRequest_SubList 내역 수정해야함..
    sqlstr = " select d.ipFileNo, d.targetGbn "
    sqlstr = sqlstr & ", (select M.ipFileState from db_jungsan.dbo.tbl_jungsan_ipkumFile_MASTER M where M.ipFileNo=D.ipFileNo) as ipFileState"
    sqlstr = sqlstr & " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D"
    sqlstr = sqlstr & " where D.targetIdx=" & id
    sqlstr = sqlstr & " and D.ipFileDetailIDx=" & ipFileDIdx

    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    retCnt = rsget.RecordCount
	    ipFileNo = rsget("ipFileNo")
		ipFileState      = rsget("ipFileState")
		targetGbn = rsget("targetGbn")
	end if
	rsget.Close

    if (retCnt>1) then
        response.write "관리자문의요망 - 삭제 불가"
        response.end
    end if
    ''Check Valid Del
    if (ipFileNo<>0) then
        if not(C_ADMIN_AUTH or C_MngPowerUser) then
            if (ipFileState>3) then  ''0=>3
                response.write "삭제 불가 - 기결제 내역 또는 상태오류"
                response.end
            end if
        end if

	    sqlstr = " delete from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
        sqlstr = sqlstr & " where targetIdx=" + id
        sqlstr = sqlstr & " and ipFileDetailIDx=" & ipFileDIdx

        dbget.Execute sqlStr

        IF (targetGbn="ON") THEN
            sqlstr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
        	sqlstr = sqlstr + " set bankingupflag='N'"
        	sqlstr = sqlstr + " ,ipkum_bank=NULL"
        	sqlstr = sqlstr + " ,ipkum_acctno=NULL"
        	sqlstr = sqlstr + " where id=" + id
        	dbget.Execute sqlStr, AssignedRow
	    ELSEIF (targetGbn="OF") THEN
	        sqlstr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
        	sqlstr = sqlstr + " set bankingupflag='N'"
        	sqlstr = sqlstr + " ,ipkum_bank=NULL"
        	sqlstr = sqlstr + " ,ipkum_acctno=NULL"
        	sqlstr = sqlstr + " where idx=" + id
        	dbget.Execute sqlStr, AssignedRow
	    END IF
    end if

elseif mode="delmast" then
	sqlstr = " IF Not Exists(select * from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D where ipFileNo="&ipFileNo&")"
    sqlstr = sqlstr + " BEGIN"
	sqlstr = sqlstr + "     delete from db_jungsan.dbo.tbl_jungsan_ipkumFile_master"
	sqlstr = sqlstr + "     where ipFileNo="&ipFileNo&""
	sqlstr = sqlstr + "     and ipFileState=0"
    sqlstr = sqlstr + " END"

	dbget.Execute sqlstr, AssignedRow

	if (AssignedRow<1) then
        response.write "관리자문의요망 - 삭제 불가"
        response.end
    end if

elseif mode="delflag" then
    retCnt=0

    ''이곳에서 db_partner.dbo.tbl_eAppPayRequest_SubList 내역 수정해야함..
    sqlstr = " select d.ipFileNo "
    sqlstr = sqlstr & ", (select M.ipFileState from db_jungsan.dbo.tbl_jungsan_ipkumFile_MASTER M where M.ipFileNo=D.ipFileNo) as ipFileState"
    sqlstr = sqlstr & " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D"
    sqlstr = sqlstr & " where D.targetIdx=" + id
    sqlstr = sqlstr & " and D.targetGbn='"&targetGbn&"'"

    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    retCnt = rsget.RecordCount
	    ipFileNo = rsget("ipFileNo")
		ipFileState      = rsget("ipFileState")
	end if
	rsget.Close

    if (retCnt>1) then
        response.write "관리자문의요망 - 삭제 불가"
        response.end
    end if
    ''Check Valid Del
    if (ipFileNo<>0) then
        if (ipFileState>0) then
            response.write "삭제 불가 - 기결제 내역 또는 상태오류"
            response.end
        end if

	    sqlstr = " delete from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
        sqlstr = sqlstr & " where targetIdx=" + id
        sqlstr = sqlstr & " and targetGbn='"&targetGbn&"'"

        dbget.Execute sqlStr

    end if

	sqlstr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlstr = sqlstr + " set bankingupflag='N'"
	sqlstr = sqlstr + " ,ipkum_bank=NULL"
	sqlstr = sqlstr + " ,ipkum_acctno=NULL"
	sqlstr = sqlstr + " where id=" + id
	dbget.Execute sqlStr, AssignedRow

elseif mode="ipkumGroup" then
    ''그룹아이디 및 입금계좌가 같아야함. // 일단 script에서 check
    sqlstr = " update F"
    sqlstr = sqlstr + " set F.refipFileDetailiDx="&secondSel
    sqlstr = sqlstr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F"
    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
    sqlstr = sqlstr + "     on F.ipFileNo=M.ipFileNo"
    sqlstr = sqlstr + " where F.ipFileNo="&ipFileNo
    sqlstr = sqlstr + " and F.ipFileDetailiDx="&firstSel
    sqlstr = sqlstr + " and F.ipFileDetailState=0"
    sqlstr = sqlstr + " and M.ipFileState=0"

    dbget.Execute sqlStr, AssignedRow

    if (thirdSel<>"") then
        sqlstr = " update F"
        sqlstr = sqlstr + " set F.refipFileDetailiDx="&secondSel
        sqlstr = sqlstr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F"
        sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
        sqlstr = sqlstr + "     on F.ipFileNo=M.ipFileNo"
        sqlstr = sqlstr + " where F.ipFileNo="&ipFileNo
        sqlstr = sqlstr + " and F.ipFileDetailiDx="&thirdSel
        sqlstr = sqlstr + " and F.ipFileDetailState=0"
        sqlstr = sqlstr + " and M.ipFileState=0"
        dbget.Execute sqlStr, AssignedRow
    end if
elseif mode="ipkumGroupMulti" then
    Dim ipFileDetailiDxArr : ipFileDetailiDxArr = request("itemidarr")
    ''그룹아이디 및 입금계좌가 같아야함. // 일단 script에서 check
    sqlstr = " update F"
    sqlstr = sqlstr + " set F.refipFileDetailiDx="&firstSel
    sqlstr = sqlstr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F"
    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
    sqlstr = sqlstr + "     on F.ipFileNo=M.ipFileNo"
    sqlstr = sqlstr + " where F.ipFileNo="&ipFileNo
    sqlstr = sqlstr + " and F.ipFileDetailiDx in ("&ipFileDetailiDxArr&")"
    sqlstr = sqlstr + " and F.ipFileDetailState=0"
    sqlstr = sqlstr + " and M.ipFileState=0"
    dbget.Execute sqlStr, AssignedRow

    response.write "<script>alert('저장 되었습니다. - " & AssignedRow & "건');</script>"
    response.write "<script>opener.location.reload();self.close();</script>"
	dbget.close()	:	response.End
elseif mode="delGroup" then
	dim grpidx, grpcnt
	grpidx  = requestCheckVar(request("grpidx"),10)
	grpcnt = 0
	 sqlStr = "select count(refipFileDetailiDx)   "
	  sqlstr = sqlstr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F"
    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
    sqlstr = sqlstr + "     on F.ipFileNo=M.ipFileNo"
    sqlstr = sqlstr + " where F.ipFileNo="&ipFileNo
    sqlStr = sqlStr + " and F.refipFileDetailiDx="&grpidx
    sqlstr = sqlstr + " and F.ipFileDetailState=0"

    if not(C_ADMIN_AUTH or C_MngPart) then
        sqlstr = sqlstr + " and M.ipFileState=0"    ' 대기 상태
    end if

     rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		grpcnt = rsget(0)
	end if
	rsget.close

	  sqlstr = " update F"
    sqlstr = sqlstr + " set F.refipFileDetailiDx=NULL "
    sqlstr = sqlstr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail  F"
    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
    sqlstr = sqlstr + "     on F.ipFileNo=M.ipFileNo"
    sqlstr = sqlstr + " where F.ipFileNo="&ipFileNo
    if grpcnt >1 then 'idx가 1개 이상 묶인 그룹만 한 idx 삭제처리 , 1개로 묶인 그룹은 그룹모두 삭제
    sqlstr = sqlstr + " and F.ipFileDetailiDx="&firstSel
  end if
    sqlStr = sqlStr + " and F.refipFileDetailiDx="&grpidx
    sqlstr = sqlstr + " and F.ipFileDetailState=0"

    if not(C_ADMIN_AUTH or C_MngPart) then
        sqlstr = sqlstr + " and M.ipFileState=0"    ' 대기 상태
    end if

    dbget.Execute sqlStr, AssignedRow


elseif mode="ipkumfinishWF" then
    AssignedRow2 = 0

    sqlstr = " update M"
    sqlstr = sqlstr + " set ipkumdate='" + ipkumregdate + "'"
    sqlStr = sqlStr + " , finishflag='7'"
    sqlStr = sqlStr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D"
    sqlStr = sqlStr + "     Join db_jungsan.dbo.tbl_designer_jungsan_master M"
    sqlStr = sqlStr + "     On M.id=D.targetIdx"
    sqlStr = sqlStr + "     and D.targetGbn='ON'"
    sqlstr = sqlstr + " where D.ipFileNo="&ipFileNo
    dbget.Execute sqlStr, AssignedRow
    AssignedRow2 = AssignedRow2 + AssignedRow
    sqlstr = " update M"
    sqlstr = sqlstr + " set ipkumdate='" + ipkumregdate + "'"
    sqlStr = sqlStr + " , finishflag='7'"
    sqlStr = sqlStr + " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D"
    sqlStr = sqlStr + "     Join db_jungsan.dbo.tbl_Off_jungsan_master M"
    sqlStr = sqlStr + "     On M.idx=D.targetIdx"
    sqlStr = sqlStr + "     and D.targetGbn='OF'"
    sqlstr = sqlstr + " where D.ipFileNo="&ipFileNo
    dbget.Execute sqlStr, AssignedRow
    AssignedRow2 = AssignedRow2 + AssignedRow

    sqlstr = " update db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
    sqlstr = sqlstr + " set ipFileDetailState=7"
    sqlstr = sqlstr + " where ipFileNo="&ipFileNo
    sqlstr = sqlstr + " and ipFileDetailState<7"
    dbget.Execute sqlStr

    sqlstr = " update db_jungsan.dbo.tbl_jungsan_ipkumFile_MASTER"
    sqlstr = sqlstr + " set ipFileState=7"
    sqlstr = sqlstr + " ,IcheDate='"&ipkumregdate&"'"
    sqlstr = sqlstr + " where ipFileNo="&ipFileNo
    sqlstr = sqlstr + " and ipFileState<7"
    dbget.Execute sqlStr

    AssignedRow = AssignedRow2
elseif mode="makeItemBuyingErpData" then
    ''수정중

elseif mode="makeItemBuyingErpData_OLD" then
    '''구 방식 사용안함 2012/02/17
    response.end
    Dim payRequestIdx, eapppartIdx

    sqlstr = " select ipFileGbn,ipFileState, IcheDate from db_jungsan.dbo.tbl_jungsan_ipkumFile_Master"
    sqlstr = sqlstr + " where ipfileno="&ipFileNo

    rsget.Open sqlStr,dbget,1
    IF Not rsget.Eof THEN
        ipFileGbn = rsget("ipFileGbn")
        ipFileState = rsget("ipFileState")
        reqIcheDate = rsget("IcheDate")
    ENd IF
    rsget.Close

    if (ipFileState<7) then
        response.write "상태값 오류 "&ipFileState
        response.end
    end if

    IF (ipFileGbn="ON") then eapppartIdx="0000000101"
    IF (ipFileGbn="OF") then eapppartIdx="0000000201"

    payRequestIdx = MakePayReq(reqIcheDate, 0, eapppartIdx)

    sqlstr = " insert into db_partner.dbo.tbl_eAppPayRequest_SubList"
    sqlstr = sqlstr & " (payRequestIdx,refType,refKey,payState,erpKey)"
    sqlstr = sqlstr & " select "&payRequestIdx
    sqlstr = sqlstr & " ,CASE WHEN D.targetGbn='ON' THEN 1 WHEN D.targetGbn='OF' THEN 2 ELSE 0 END"
    sqlstr = sqlstr & " ,D.targetIdx"
    sqlstr = sqlstr & " ,0"
    sqlstr = sqlstr & " ,NULL"
    sqlstr = sqlstr & " from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D"
    sqlstr = sqlstr & " 	Left Join db_jungsan.dbo.tbl_designer_jungsan_master M"
    sqlstr = sqlstr & " 	On D.targetGbn='ON' and D.targetIdx=M.id"
    sqlstr = sqlstr & " 	Left Join db_jungsan.dbo.tbl_off_jungsan_master F"
    sqlstr = sqlstr & " 	On D.targetGbn='OF' and D.targetIdx=F.idx"
    sqlstr = sqlstr & " where D.ipFileNo="&ipFileNo
    sqlstr = sqlstr & " and D.ipFileDetailState=7"

    dbget.Execute sqlstr, AssignedRow

    CALL RecalcuPayRequestPrice(payRequestIdx)

    sqlstr = " update db_jungsan.dbo.tbl_jungsan_ipkumFile_Master"
    sqlstr = sqlstr & " set ipFileState=8"
    sqlstr = sqlstr & " where ipFileNo="&ipFileNo
    dbget.Execute sqlstr

    sqlstr = " update db_partner.dbo.tbl_eAppPayRequest"
    sqlstr = sqlstr & " set payrequestState=7"
    sqlstr = sqlstr & " where payrequestIdx="&payRequestIdx

    dbget.Execute sqlstr

elseif mode="ipkumfinish" then
    response.write "사용불가 메뉴"
    response.end
'	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'	sqlStr = sqlStr + " set ipkumdate='" + ipkumregdate + "'"
'	sqlStr = sqlStr + " , finishflag='7'"
'	sqlstr = sqlstr + " where id in (" + checkone + ")"
'	dbget.Execute sqlStr, AssignedRow
ElseIf mode = "deleteFileNo" Then
    Dim tmpFileGbn
    sqlstr = ""
	sqlstr = sqlstr & " SELECT TOP 1 ipFileGbn FROM db_jungsan.dbo.tbl_jungsan_ipkumFile_master "
	sqlstr = sqlstr & " WHERE ipFileNo="&ipFileNo&""
	sqlstr = sqlstr & " and ipFileState=0"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
    If Not rsget.Eof Then
        tmpFileGbn = rsget("ipFileGbn")
    End If
    rsget.close

    If tmpFileGbn = "ON" Then
        sqlstr = ""
        sqlstr = sqlstr & " UPDATE m "
        sqlstr = sqlstr & " SET m.bankingupflag = 'N' "
        sqlstr = sqlstr & " FROM db_jungsan.dbo.tbl_designer_jungsan_master as m "
        sqlstr = sqlstr & " JOIN db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail d on m.id = d.targetIdx "
        sqlstr = sqlstr & " WHERE d.ipFileNo = "&ipFileNo&""
        dbget.Execute sqlstr
    Else
        sqlstr = ""
        sqlstr = sqlstr & " UPDATE m "
        sqlstr = sqlstr & " SET m.bankingupflag = 'N' "
        sqlstr = sqlstr & " FROM [db_jungsan].[dbo].tbl_off_jungsan_master as m "
        sqlstr = sqlstr & " JOIN db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail d on m.idx = d.targetIdx "
        sqlstr = sqlstr & " WHERE d.ipFileNo = "&ipFileNo&""
        dbget.Execute sqlstr
    End If

    sqlstr = ""
    sqlstr = sqlstr & " DELETE FROM db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail "
    sqlstr = sqlstr & " WHERE ipFileNo="&ipFileNo&""
    dbget.Execute sqlstr

    sqlstr = ""
    sqlstr = sqlstr & " DELETE FROM db_jungsan.dbo.tbl_jungsan_ipkumFile_master "
    sqlstr = sqlstr & " WHERE ipFileNo="&ipFileNo&""
    sqlstr = sqlstr & " and ipFileState=0"
    dbget.Execute sqlstr, AssignedRow

end if
%>

<script language="javascript">
<% if mode="delflag" or mode="delflagWF" or mode="delmast" or mode="deleteFileNo" then %>
alert('삭제 되었습니다. - <%= AssignedRow %>건');
//opener.location.reload();
window.close();
<% else %>
alert('저장 되었습니다. - <%= AssignedRow %>건');
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
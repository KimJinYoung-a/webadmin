<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,checkone,id, ipkumregdate, checkoneEx, jgubun
id = request("id")
mode = request("mode")
checkone = request.form("checkone")
ipkumregdate = request("ipkumregdate")
checkoneEx = request.form("checkoneEx")
jgubun      = request("jgubun")

if (checkone="") then checkone="0"
if (checkoneEx="") then checkoneEx="0"
Dim isMixedFile : isMixedFile= (request.form("ck_Mibus")="CX")

Dim reqIcheDate   : reqIcheDate   = requestCheckVar(request("reqIcheDate"),10)

Dim UseUpFile : UseUpFile = requestCheckVar(request("UseUpFile"),10)
Dim ipFileNo : ipFileNo = requestCheckVar(request("ipFileNo"),10)
IF (ipFileNo="") then ipFileNo=0

Dim ipFileState


dim sqlstr, AssignedRow, AssignedRow2
Dim targetGbn : targetGbn = "OF"
Dim targetGbnEx : targetGbnEx = "ON"
Dim ipFileName

Dim retMakerId, retGroupid
Dim retipFileNo, rettargetIdx
Dim NotReqgroupIdList0,NotReqgroupIdList1,NotReqgroupIdList2

Dim IsForce : IsForce=FALSE

Dim IsExtJExists : IsExtJExists = (checkoneEx<>"0")

if mode="bankingupload" then
    IF (UseUpFile<>"") THEN
        IF (ipFileNo>0) then
        	rw checkone
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
        sqlstr = " select top 1 S.ipFileNo, S.targetIdx, m.makerid, m.groupid From db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
        sqlstr = sqlstr + " Join  [db_jungsan].[dbo].tbl_off_jungsan_master M"
        sqlstr = sqlstr + " on S.targetIdx=M.idx"
        sqlstr = sqlstr + " where S.targetGbn='"&targetGbn&"'"
        sqlstr = sqlstr + " and S.targetIdx in (" + checkone + ")"

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

		'''온라인 기업로드 check
		sqlstr = " select top 1 S.ipFileNo, S.targetIdx, m.designerid, m.groupid From db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
        sqlstr = sqlstr + " Join  [db_jungsan].[dbo].tbl_designer_jungsan_master M"
        sqlstr = sqlstr + " on S.targetIdx=M.id"
        sqlstr = sqlstr + " where S.targetGbn='"&targetGbnEx&"'"
        sqlstr = sqlstr + " and S.targetIdx in (" + checkoneEx + ")"

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

        ''정산구분 check-----------------------------------------------------------------------------------------
		if (jgubun<>"") then
    		retipFileNo = 0
            sqlstr = " select top 1 M.idx, m.makerid, m.groupid From [db_jungsan].[dbo].tbl_off_jungsan_master M"
            sqlstr = sqlstr + " where M.idx in (" + checkone + ")"
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

    		'''온라인 기업로드 check
    		sqlstr = " select top 1 M.id, m.designerid, m.groupid From  [db_jungsan].[dbo].tbl_designer_jungsan_master M"
            sqlstr = sqlstr + " where M.id in (" + checkoneEx + ")"
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
    	end if


		''----------------------------------------------------------------------------------------------------------
		''마이너스 체크 // 같은거래처(groupid)에 마이너스 금액이 남아 있을경우 올라가지 못하게.
        sqlstr = " select top 1000 g.Groupid from [db_jungsan].[dbo].tbl_off_jungsan_master g"
        sqlstr = sqlstr + " where g.idx in (" + checkone + ")"
        sqlstr = sqlstr + " and Groupid in ("
        sqlstr = sqlstr + "     select M.Groupid from [db_jungsan].[dbo].tbl_off_jungsan_master M"
        sqlstr = sqlstr + "     where M.idx not in (" + checkone + ")"
        sqlstr = sqlstr + "     and M.bankingupflag ='N'"
        sqlstr = sqlstr + "     and M.finishFlag=3"
        sqlstr = sqlstr + "     and (tot_jungsanprice)<0"
        sqlstr = sqlstr + " ) and ( g.idx in (" + checkone + ") or (IsNULL(g.ipkum_acctno,'')='')  )"


        rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			Do Until rsget.Eof
    			NotReqgroupIdList0 = NotReqgroupIdList0 + rsget("Groupid") +","
    			rsget.movenext
			loop
		end if
		rsget.Close

		IF (NotReqgroupIdList0<>"") and (Not IsForce) then
		    response.write " 정산 확정내역중 마이너스  / 또는 계좌 오류 있음. : " & NotReqgroupIdList0 & "<br>"
            ''response.end
		End IF


		''업로드 총금액이 그룹별로 마이너스가 될 수 없음 // 기 업로드 내역 포함 조사
		sqlstr = "select Groupid, Sum(jungsanSum) from ("
        sqlstr = sqlstr + "     select groupid, (ub_totalsuplycash+ me_totalsuplycash+wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash) as jungsanSum"
        sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlstr = sqlstr + "     where  id  in (" + checkoneEx + ")"
        sqlstr = sqlstr + " Union "
        sqlstr = sqlstr + "     select m2.groupid, m2.tot_jungsanprice as jungsanSum "
        sqlstr = sqlstr + "     from  [db_jungsan].[dbo].tbl_off_jungsan_master m2"
        sqlstr = sqlstr + "     where m2.idx in (" + checkone + ")"
        sqlstr = sqlstr + " Union "
        sqlstr = sqlstr + "     select J.groupid, (J.ub_totalsuplycash+ J.me_totalsuplycash+J.wi_totalsuplycash+J.sh_totalsuplycash+J.et_totalsuplycash+J.dlv_totalsuplycash) as jungsanSum"
        sqlstr = sqlstr + "     from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail P1"
	    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_designer_jungsan_master J"
	    sqlstr = sqlstr + "     on P1.ipFileNo="&ipFileNo
	    sqlstr = sqlstr + "     and P1.targetGbn='ON'"
	    sqlstr = sqlstr + "     and P1.targetIdx=J.id"
        sqlstr = sqlstr + " Union "
        sqlstr = sqlstr + "     select F.groupid, tot_jungsanprice as jungsanSum"
        sqlstr = sqlstr + "     from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail P2"
	    sqlstr = sqlstr + "     Join db_jungsan.dbo.tbl_off_jungsan_master F"
	    sqlstr = sqlstr + "     on P2.ipFileNo="&ipFileNo
	    sqlstr = sqlstr + "     and P2.targetGbn='OF'"
	    sqlstr = sqlstr + "     and P2.targetIdx=F.idx"
        sqlstr = sqlstr + " ) T group by T.groupid"
        sqlstr = sqlstr + " having Sum(jungsanSum)<0"

        rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			Do Until rsget.Eof
    			NotReqgroupIdList1 = NotReqgroupIdList1 + rsget("Groupid") +","
    			rsget.movenext
			loop
		end if
		rsget.Close

		IF (NotReqgroupIdList1<>"") and (Not IsForce) then
		    response.write "업로드 총금액이 마이너스가 될 수 없음.  "&NotReqgroupIdList1&"  "
            ''response.end
		End IF

		'''일단 사용안함..
        if (FALSE) and (Not IsExtJExists) THEN
    		''온라인 체크
    		sqlstr = " select Top 100 makerid, Groupid from [db_jungsan].[dbo].tbl_off_jungsan_master"
            sqlstr = sqlstr + " where idx in (" + checkone + ")"
            sqlstr = sqlstr + " and Groupid in ("
            sqlstr = sqlstr + "     select M.Groupid from [db_jungsan].[dbo].tbl_designer_jungsan_master M"
            sqlstr = sqlstr + "     where 1=1"
            sqlstr = sqlstr + "     and M.bankingupflag ='N'"
            sqlstr = sqlstr + "     and M.finishFlag=3"
            sqlstr = sqlstr + "     and (ub_totalsuplycash+ me_totalsuplycash+wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash)<0"
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
    		    response.write "그룹 ID "&NotReqgroupIdList2&" <b>온라인 정산</b> 확정내역중 마이너스 금액 있음. - 업체 제외 되고 업로드"
                ''response.end
    		End IF
		end if


        IF (ipFileNo=0) then
            IF (targetGbn="ON") then
                ipFileName = "온라인 "
            ELSEIF (targetGbn="OF") then
                ipFileName = "오프라인 "
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
        sqlstr = sqlstr & " ,'"&targetGbn&"'"
        sqlstr = sqlstr & " ,M.idx"
        sqlstr = sqlstr & " ,getdate()"
        sqlstr = sqlstr & " From [db_jungsan].[dbo].tbl_off_jungsan_master M"
        sqlstr = sqlstr & "     left Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
        sqlstr = sqlstr & "     on M.idx=S.targetIdx"
        sqlstr = sqlstr & "     and S.targetGbn='"&targetGbn&"'"
        sqlstr = sqlstr + " where M.idx in (" + checkone + ")"
        sqlstr = sqlstr + " and M.bankingupflag ='N'"
        sqlstr = sqlstr + " and S.ipFileNo Is NULL"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
        sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"
        dbget.Execute sqlStr, AssignedRow
'rw sqlstr
        sqlstr = "update M "
    	sqlstr = sqlstr + " set bankingupflag='Y'"
    	sqlstr = sqlstr + " ,ipkum_bank=isNULL(A.jungsan_bank,G.jungsan_bank)"
        sqlstr = sqlstr + " ,ipkum_acctno=isNULL(A.jungsan_acctno,G.jungsan_acctno)"
    	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_off_jungsan_master M"
    	sqlstr = sqlstr + "     Join [db_partner].[dbo].tbl_partner_group G"
    	sqlstr = sqlstr + "     on M.groupid=G.groupid"
    	sqlstr = sqlstr + "     left join db_partner.dbo.tbl_partner_addJungsanInfo A"      ''2016/12/15 추가
        sqlstr = sqlstr + "     on M.makerid=A.partnerid"
    	sqlstr = sqlstr + " where M.idx in (" + checkone + ")"
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
            sqlstr = sqlstr & " ,M.id"
            sqlstr = sqlstr & " ,getdate()"
            sqlstr = sqlstr & " From [db_jungsan].[dbo].tbl_designer_jungsan_master M"
            sqlstr = sqlstr & "     left Join db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail S"
            sqlstr = sqlstr & "     on M.id=S.targetIdx"
            sqlstr = sqlstr & "     and S.targetGbn='"&targetGbnEx&"'"
            sqlstr = sqlstr + " where M.id in (" + checkoneEx + ")"
            sqlstr = sqlstr + " and M.bankingupflag ='N'"
            sqlstr = sqlstr + " and S.ipFileNo Is NULL"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"
            dbget.Execute sqlStr, AssignedRow
'rw sqlstr
            sqlstr = "update M "
        	sqlstr = sqlstr + " set bankingupflag='Y'"
        	sqlstr = sqlstr + " ,ipkum_bank=isNULL(A.jungsan_bank,G.jungsan_bank)"
        	sqlstr = sqlstr + " ,ipkum_acctno=isNULL(A.jungsan_acctno,G.jungsan_acctno)"
        	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master M"
        	sqlstr = sqlstr + "     Join [db_partner].[dbo].tbl_partner_group G"
        	sqlstr = sqlstr + "     on M.groupid=G.groupid"
        	sqlstr = sqlstr + "     left join db_partner.dbo.tbl_partner_addJungsanInfo A"      ''2016/12/15 추가
        	sqlstr = sqlstr + "     on M.designerid=A.partnerid"
        	sqlstr = sqlstr + " where M.id in (" + checkoneEx + ")"
        	sqlstr = sqlstr + " and M.bankingupflag ='N'"
        	sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList0&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList1&")"
            sqlstr = sqlstr + " and M.groupid not in ("&NotReqgroupIdList2&")"

        	dbget.Execute sqlStr, AssignedRow2

            AssignedRow = AssignedRow + PAssignedRow
        	AssignedRow2 = AssignedRow2 + PAssignedRow2
        end if
        response.write "<script language='javascript'>"
        response.write "alert('저장 되었습니다. - "& AssignedRow &"/"&AssignedRow2&" 건');"
        IF (NotReqgroupIdList0&NotReqgroupIdList1&NotReqgroupIdList2="''''''") then
            response.write "location.replace('"&refer&"');"
        End IF
        response.write "</script>"

        response.end
    ENd If

    sqlstr = "update M "
	sqlstr = sqlstr + " set bankingupflag='Y'"
	sqlstr = sqlstr + " ,ipkum_bank=isNULL(A.jungsan_bank,G.jungsan_bank)"
    sqlstr = sqlstr + " ,ipkum_acctno=isNULL(A.jungsan_acctno,G.jungsan_acctno)"
	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_off_jungsan_master M"
	sqlstr = sqlstr + "     Join [db_partner].[dbo].tbl_partner_group G"
	sqlstr = sqlstr + "     on M.groupid=G.groupid"
	sqlstr = sqlstr + "     left join db_partner.dbo.tbl_partner_addJungsanInfo A"      ''2016/12/15 추가
    sqlstr = sqlstr + "     on M.makerid=A.partnerid"
	sqlstr = sqlstr + " where M.idx in (" + checkone + ")"
	sqlstr = sqlstr + " and M.bankingupflag ='N'"

	dbget.Execute sqlStr, AssignedRow

elseif mode="delflag" then
    Dim retCnt : retCnt=0
    Dim payState

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

	sqlstr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
	sqlstr = sqlstr + " set bankingupflag='N'"
	sqlstr = sqlstr + " ,ipkum_bank=NULL"
	sqlstr = sqlstr + " ,ipkum_acctno=NULL"
	sqlstr = sqlstr + " where idx=" + id
	dbget.Execute sqlStr, AssignedRow

elseif mode="ipkumfinish" then
    response.write "사용불가 메뉴"
    response.end
'	sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
'	sqlStr = sqlStr + " set ipkumdate='" + ipkumregdate + "'"
'	sqlStr = sqlStr + " , finishflag='7'"
'	sqlstr = sqlstr + " where idx in (" + checkone + ")"
'	dbget.Execute sqlStr, AssignedRow
end if
%>

<script language="javascript">
<% if mode="delflag" then %>
alert('삭제 되었습니다. - <%= AssignedRow %>건');
opener.location.reload();
window.close();
<% else %>
alert('저장 되었습니다. - <%= AssignedRow %>건');
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
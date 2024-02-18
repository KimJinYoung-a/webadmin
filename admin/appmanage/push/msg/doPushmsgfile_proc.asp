<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 푸시 csv 타게팅
' Hieditor : 2018.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<%
dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
    sNow = now()
    sY= Year(sNow)
    sM = Format00(2,Month(sNow))
    sD = Format00(2,Day(sNow))
    sH = Format00(2,Hour(sNow))
    sMi = Format00(2,Minute(sNow))
    sS = Format00(2,Second(sNow))
    sDateName = sY&sM&sD&sH&sMi&sS

dim uploadform, fs, sDefaultPath, sFile, mode, iMaxLen, strFileType, sFilePath, objFile, content, contarr, i
dim useridarr, sqlStr, idx, targetKey, baseIdx, TargetsubCNT, NotTargetCNT, searchStr, push_targetcnt, delFile
Set uploadform = Server.CreateObject("TABSUpload4.Upload")
Set fs		= Server.CreateObject("Scripting.FileSystemObject")
    i=0
    sDefaultPath = Server.MapPath("\admin\appmanage\push\msg")
    uploadform.Start sDefaultPath '업로드경로
    iMaxLen 		= "1"	'파일크기
    iMaxLen = iMaxLen * 1024 * 1024 '최대용량(MB)
    strFileType = "csv"
    mode = requestCheckVar(uploadform("mode"),32)
    idx = RequestCheckVar(uploadform("idx"),10)
	TargetsubCNT = 0
	NotTargetCNT = 0
    push_targetcnt = 0

Select Case mode
	Case "csvtarget"
        IF (fnChkFile(uploadform.Form("sFile"), iMaxLen, strFileType)) THEN	'파일체크	(용량, 확장자)
            ' 푸시 상세내역 가져오기
            sqlStr = "select top 1" + VbCrlf
            sqlStr = sqlStr & " targetKey, baseIdx" + VbCrlf
            sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve" + VbCrlf
            sqlStr = sqlStr & " where idx = "& idx

            'response.write sqlStr & "<Br>"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            If not rsget.EOF Then
                targetKey = rsget("targetKey")
                    if isnull(targetKey) then targetKey = ""
                baseIdx = rsget("baseIdx")
                    if isnull(baseIdx) then baseIdx = ""
            end if
            rsget.close

            '파일저장
            sFile = sDateName & "." & uploadform("sFile").FileType
            sFilePath = sDefaultPath & "\" & sFile
            sFilePath = uploadform("sFile").SaveAs(sFilePath, True)

            Set objFile = fs.OpenTextFile(sFilePath,1)
            i=0
            Do While objFile.AtEndOfStream <> True
                content=trim(objFile.ReadLine)    ' 파일을 라인단위로 읽음 한줄씩 처리. 다음줄을 읽으라는 명령을 안줘도 됨

                ' 첫줄은 무시
                if i > 0 then
                    if content <> "" then
                        useridarr = useridarr & content & ","
                    end if
                else
                    ' 텐바이텐 고객번호일경우
                    if targetKey="10" then
                        if content<>"텐바이텐고객번호" then
                            response.write "<script type='text/javascript'>"
                            response.write "	alert('CSV파일이 텐바이텐고객번호 가 아닙니다.\n다시한번 확인하세요.');"
                            response.write "</script>"
                            session.codePage = 949
                            dbget.close()	:	response.End
                        end if

                    ' 텐바이텐 고객아이디일경우
                    elseif targetKey="11" then
                        if content<>"텐바이텐고객아이디" then
                            response.write "<script type='text/javascript'>"
                            response.write "	alert('CSV파일이 텐바이텐고객아이디 가 아닙니다.\n다시한번 확인하세요.');"
                            response.write "</script>"
                            session.codePage = 949
                            dbget.close()	:	response.End
                        end if
                    end if
                end if
                i = i + 1
            Loop

            if trim(replace(useridarr,",","")) = "" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('텐바이텐 고객번호 or 텐바이텐 고객아이디를 입력해 주세요.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if

            useridarr = trim(left(useridarr,len(useridarr)-1))
            Set objFile = Nothing

            if ubound(split(useridarr,","))+1 > 30000 then
                response.write "<script type='text/javascript'>"
                response.write "	alert('3만명씩 짤라서 입력해 주세요.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if

            ' 텐바이텐 고객번호일경우
            if targetKey="10" then
                sqlStr = "select userid" + VbCrlf
                sqlStr = sqlStr & " into #push_target" + VbCrlf
                sqlStr = sqlStr & " from db_user.dbo.tbl_logindata with (nolock)" + VbCrlf
                sqlStr = sqlStr & " where useq in ("& useridarr &")" + VbCrlf

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr

            ' 텐바이텐 고객아이디일경우
            elseif targetKey="11" then
                useridarr =  "'" & replace(useridarr,",","','") & "'"

                sqlStr = "select userid" + VbCrlf
                sqlStr = sqlStr & " into #push_target" + VbCrlf
                sqlStr = sqlStr & " from db_user.dbo.tbl_user_n with (nolock)" + VbCrlf
                sqlStr = sqlStr & " where userid in ("& useridarr &")" + VbCrlf

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if

            sqlStr = "select count(*) as push_targetcnt from #push_target" + VbCrlf

            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            If not rsget.EOF Then
                push_targetcnt = rsget("push_targetcnt")
            end if
            rsget.close

            if push_targetcnt < 1 then
                response.write "<script type='text/javascript'>"
                response.write "	alert('CSV파일에 있는 대상자중에 텐바이텐회원이 없습니다.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if

            ' 나머지타켓이 아닌데 멀티타겟으로 메인타겟이 될경우
            if targetKey <>"1" and baseIdx="" then
                ' 현재idx가 다른 멀티타겟에 메인타겟으로 쓰인 카운트 체크
                sqlStr = "select count(*) as TargetsubCNT" + VbCrlf
                sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve" + VbCrlf
                sqlStr = sqlStr & " where idx <> "& idx &"" + VbCrlf
                sqlStr = sqlStr & " and baseIdx = "& idx &"" + VbCrlf

                rsget.CursorLocation = adUseClient
                rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                If not rsget.EOF Then
                    TargetsubCNT = rsget("TargetsubCNT")
                end if
                rsget.close

                if TargetsubCNT > 0 then
                    sqlStr = "select count(*) as NotTargetCNT" + VbCrlf
                    sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_reserve p" + VbCrlf
                    sqlStr = sqlStr & " left Join (select rsvIdx,count(*) as CNT from db_contents.dbo.tbl_app_push_TargetTemp group by rsvIdx) T" + VbCrlf
                    sqlStr = sqlStr & " 	on p.idx=T.rsvIdx" + VbCrlf
                    sqlStr = sqlStr & " where p.idx<>"& idx &"" + VbCrlf
                    sqlStr = sqlStr & " and p.baseIdx="& idx &"" + VbCrlf
                    sqlStr = sqlStr & " and T.rsvIdx is NULL" + VbCrlf	 ' 타게팅 값이 없음
                    sqlStr = sqlStr & " and p.isusing='Y'" + VbCrlf

            		rsget.CursorLocation = adUseClient
		            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                    If not rsget.EOF Then
                        NotTargetCNT = rsget("NotTargetCNT")
                    end if
                    rsget.close

                    if NotTargetCNT > 0 then
                        response.write "<script type='text/javascript'>"
                        response.write "	alert('타게팅 값이 없음.');"
                        response.write "</script>"
                        session.codePage = 949
                        dbget.close()	:	response.End
                    end if

                    searchStr = searchStr & " and R.regIdx not in (" & VbCrlf
                    searchStr = searchStr & "   select T.psregIdx from db_contents.dbo.tbl_app_push_reserve p" & VbCrlf
                    searchStr = searchStr & "   Join db_contents.dbo.tbl_app_push_TargetTemp T with (nolock)" & VbCrlf
                    searchStr = searchStr & "   on T.rsvIdx=P.idx" & VbCrlf
                    searchStr = searchStr & "   and P.baseIdx="& idx &"" & VbCrlf
                    searchStr = searchStr & " )" & VbCrlf
                end if
            end if

            ' 푸시 대상자 입력
            sqlStr = "insert into db_contents.dbo.tbl_app_push_TargetTemp (rsvIdx, psregIdx, userid)" + VbCrlf
            sqlStr = sqlStr & "	select "& idx &",t.regIdx,t.userid" + VbCrlf
            sqlStr = sqlStr & "	from (" + VbCrlf
            sqlStr = sqlStr & "		select R.regIdx , RANK() Over (partition by r.userid, appkey order by lastupdate desc) as LastRank" + VbCrlf
            sqlStr = sqlStr & "		, r.userid"
            sqlStr = sqlStr & "		from db_contents.dbo.tbl_app_regInfo R with (nolock)" + VbCrlf
            sqlStr = sqlStr & "		join #push_target t" + VbCrlf
            sqlStr = sqlStr & "			on r.userid = t.userid" + VbCrlf
            sqlStr = sqlStr & "		where R.isusing='Y'" + VbCrlf
            sqlStr = sqlStr & "		and ((R.appkey=6 and R.appVer>='36')" + VbCrlf
            sqlStr = sqlStr & "			or (R.appkey=5 and R.appVer>='1')" + VbCrlf
            sqlStr = sqlStr & "		) " & searchStr
            sqlStr = sqlStr & "	) T" + VbCrlf
            sqlStr = sqlStr & "	left join db_contents.dbo.tbl_app_push_TargetTemp pt" + VbCrlf
            sqlStr = sqlStr & "	    on t.regIdx = pt.psregidx" + VbCrlf
            sqlStr = sqlStr & "	    and pt.rsvidx = "& idx &"" + VbCrlf
            sqlStr = sqlStr & "	where pt.psregidx is null" + VbCrlf
            'sqlStr = sqlStr & " and T.LastRank=1" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            ' 연속 입력으로 인한 중복제거
            sqlStr = "delete pt" + VbCrlf
            sqlStr = sqlStr & "	from db_contents.dbo.tbl_app_push_TargetTemp pt" + VbCrlf
            sqlStr = sqlStr & "	join (" + VbCrlf
            sqlStr = sqlStr & "		select psregIdx, count(psregIdx) as cnt" + VbCrlf
            sqlStr = sqlStr & "		from db_contents.dbo.tbl_app_push_TargetTemp" + VbCrlf
            sqlStr = sqlStr & "		where rsvidx = "& idx &"" + VbCrlf
            sqlStr = sqlStr & "		group by psregIdx" + VbCrlf
            sqlStr = sqlStr & "		having count(psregIdx) > 1" + VbCrlf
            sqlStr = sqlStr & "	) as t" + VbCrlf
            sqlStr = sqlStr & "		on pt.psregIdx = t.psregIdx" + VbCrlf
            sqlStr = sqlStr & "	where pt.rsvidx = "& idx &"" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            if targetKey >1 and baseIdx<>"" then
                ' 현재idx가 다른 멀티타겟 쓰인경우 중복제거
                sqlStr = "DELETE T" + VbCrlf
                sqlStr = sqlStr & "	from db_contents.dbo.tbl_app_push_TargetTemp T" + VbCrlf
                sqlStr = sqlStr & "	Join db_contents.dbo.tbl_app_push_TargetTemp p" + VbCrlf
                sqlStr = sqlStr & "	    on T.psregIdx=p.psregIdx and P.rsvIdx<>"& idx &"" + VbCrlf
                sqlStr = sqlStr & "	Join db_contents.dbo.tbl_app_push_reserve R" + VbCrlf
                sqlStr = sqlStr & "	    on P.rsvIdx=R.idx and R.baseIdx="& baseIdx &"" + VbCrlf
                sqlStr = sqlStr & "	where T.rsvIdx="& idx &"" + VbCrlf

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if

            ' 수신거부자 제외
            sqlStr = "DELETE T" + VbCrlf
            sqlStr = sqlStr & " from db_contents.dbo.tbl_app_push_TargetTemp T" + VbCrlf
            sqlStr = sqlStr & " Join db_contents.dbo.tbl_app_regInfo R" + VbCrlf
            sqlStr = sqlStr & "    on T.rsvIdx in ("& idx &")" + VbCrlf
            sqlStr = sqlStr & "    and T.psregIdx=R.regIdx" + VbCrlf
            sqlStr = sqlStr & "    and isNULL(R.pushyn,'')='N'" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            ' 최종 타게팅 업데이트
            sqlStr = "DECLARE @mayTargetCnt int" + VbCrlf
            sqlStr = sqlStr & " select @mayTargetCnt=count(*) from db_contents.dbo.tbl_app_push_TargetTemp where" + VbCrlf
            sqlStr = sqlStr & " rsvIdx="& idx &"" + VbCrlf
            sqlStr = sqlStr & " update  db_contents.dbo.tbl_app_push_reserve" + VbCrlf
            sqlStr = sqlStr & " set targetState=7" + VbCrlf
            sqlStr = sqlStr & " ,mayTargetCnt=@mayTargetCnt where" + VbCrlf
            sqlStr = sqlStr & " idx="& idx &"" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            sqlStr = "drop table #push_target" + VbCrlf
            dbget.Execute sqlStr

            ' 저장한 파일삭제
            Set delFile = fs.GetFile(sFilePath)
                delFile.Delete 
            set delFile = Nothing

			response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
			Response.write "<script type='text/javascript'>opener.location.reload();opener.opener.location.reload();self.close();</script>"
			session.codePage = 949
			dbget.close()	:	response.End
        END IF

    CASE ELSE
        response.write "<script type='text/javascript'>alert('정의되지 않았음 "&mode&"');</script>"
		session.codePage = 949
		dbget.close()	:	response.End
End Select

Set uploadform = Nothing
Set fs = Nothing

Function fnChkFile(ByVal sfile, ByVal smaxlen, ByVal fileType) 
	Dim  strFileSize, strFileType ,strMimeType,strFileName, arrfileType, chkReturn, i
	IF  sfile = "" THEN  
		fnChkFile = FALSE
		Exit Function
	END IF	
	strFileSize = sfile.FileSize
	strFileName = sfile.FileName  
	strFileType = LCase(sfile.FileType)  
	strMimeType = sfile.ContentType
	if strFileSize  > smaxlen then	'용량 체크
		response.write "파일크기는 " & smaxlen & "MB 이하만 가능합니다."
		dbget.close() : response.end
	end if
	arrfileType = split(filetype,"^")
	chkReturn = 0
	for i = 0 to ubound(arrfiletype)
		if strFileType = arrfiletype(i) THEN chkReturn = 1 
	next
	if chkReturn = 0 then
		response.write fileType & " 형식의 파일만 가능합니다."
		dbget.close() : response.end
	end if
	fnChkFile = TRUE 
End Function
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>
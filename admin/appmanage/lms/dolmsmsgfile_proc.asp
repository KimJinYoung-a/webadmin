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
' Description : LMS발송관리
' Hieditor : 2020.03.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
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

dim uploadform, fs, sDefaultPath, sFile, mode, iMaxLen, strFileType, sFilePath, objFile, content, contarr, i, lastadminid
dim useridarr, sqlStr, ridx, targetkey, TargetsubCNT, NotTargetCNT, searchStr, lms_targetcnt, delFile
dim exception7dayyn, member_pushyn_checkyn, itemid, useridfield, itemidfield, contentline, olms, sendmethod, member_smsok_checkyn
dim title, contents, failed_subject, failed_msg, itemidexists, itemidexistscnt, exceptionlogin, exceptionuserlevelarr
dim member_kakaoalrimyn_checkyn, mileageexists, mileagedateexists, mileagefield, mileagedatefield
Set uploadform = Server.CreateObject("TABSUpload4.Upload")
Set fs		= Server.CreateObject("Scripting.FileSystemObject")
    i=0
    itemidexists=false
    mileageexists=false
    mileagedateexists=false
    itemidexistscnt=0
    sDefaultPath = Server.MapPath("\admin\appmanage\lms")
    uploadform.Start sDefaultPath '업로드경로
    iMaxLen 		= "1"	'파일크기
    iMaxLen = iMaxLen * 1024 * 1024 '최대용량(MB)
    strFileType = "csv"
    mode = requestCheckVar(uploadform("mode"),32)
    ridx = requestcheckvar(getNumeric(uploadform("ridx")),10)
	TargetsubCNT = 0
	NotTargetCNT = 0
    lms_targetcnt = 0

lastadminid = session("ssBctId")

Select Case mode
	Case "csvtarget"
        IF (fnChkFile(uploadform.Form("sFile"), iMaxLen, strFileType)) THEN	'파일체크	(용량, 확장자)
            sqlStr = "select top 1" & VbCrlf
            sqlStr = sqlStr & " targetkey, exception7dayyn, member_pushyn_checkyn, isnull(exceptionlogin,'') as exceptionlogin" & VbCrlf
            sqlStr = sqlStr & " , isnull(exceptionuserlevelarr,'') as exceptionuserlevelarr, member_kakaoalrimyn_checkyn" & VbCrlf
            sqlStr = sqlStr & " from db_contents.dbo.tbl_lms_reserve with (readuncommitted)" + VbCrlf
            sqlStr = sqlStr & " where ridx = "& ridx

            'response.write sqlStr & "<Br>"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            If not rsget.EOF Then
                targetkey = rsget("targetkey")
                if isnull(targetkey) then targetkey = ""
                exception7dayyn = rsget("exception7dayyn")
                member_pushyn_checkyn = rsget("member_pushyn_checkyn")
                exceptionlogin = rsget("exceptionlogin")
                exceptionuserlevelarr = rsget("exceptionuserlevelarr")
                member_kakaoalrimyn_checkyn = rsget("member_kakaoalrimyn_checkyn")
            end if
            rsget.close

            if targetkey="" or isnull(targetkey) then
                response.write "<script type='text/javascript'>"
                response.write "	alert('발송타켓이 없습니다.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if

            set olms = new clms_msg_list
                olms.FRectrIdx = ridx
                olms.lmsmsg_getrow()

            if olms.FResultCount > 0 then
                sendmethod			= olms.FOneItem.fsendmethod
                title			= olms.FOneItem.ftitle
                contents			= olms.FOneItem.fcontents
                failed_subject = olms.FOneItem.ffailed_subject
                failed_msg = olms.FOneItem.ffailed_msg
            end if
            if sendmethod="KAKAOALRIM" then
                member_smsok_checkyn=""
            else
                member_smsok_checkyn="Y"
	            member_kakaoalrimyn_checkyn=""
            end if
            if instr(title,"${PRODUCTNAME}")>0 or instr(contents,"${PRODUCTNAME}")>0 or instr(failed_subject,"${PRODUCTNAME}")>0 or instr(failed_subject,"${PRODUCTNAME}")>0 then
                itemidexists=true
            end if
            if instr(title,"${MILEAGE}")>0 or instr(contents,"${MILEAGE}")>0 or instr(failed_subject,"${MILEAGE}")>0 or instr(failed_subject,"${MILEAGE}")>0 then
                mileageexists=true
            end if
            if instr(title,"${MILEAGEDATE}")>0 or instr(contents,"${MILEAGEDATE}")>0 or instr(failed_subject,"${MILEAGEDATE}")>0 or instr(failed_subject,"${MILEAGEDATE}")>0 then
                mileagedateexists=true
            end if

            '파일저장
            sFile = sDateName & "." & uploadform("sFile").FileType
            sFilePath = sDefaultPath & "\" & sFile
            sFilePath = uploadform("sFile").SaveAs(sFilePath, True)

            ' 엑셀 모수를 담을 임시테이블 작성
            sqlStr = "create table #lms_target(" & VbCrlf
            sqlStr = sqlStr & " hpno nvarchar(16) NULL" & VbCrlf
            sqlStr = sqlStr & " , useq int NULL" & VbCrlf
            sqlStr = sqlStr & " , userid nvarchar(32) NULL" & VbCrlf
            sqlStr = sqlStr & " , pushyn nvarchar(1) NULL" & VbCrlf
            sqlStr = sqlStr & " , smsok nvarchar(1) NULL" & VbCrlf
            sqlStr = sqlStr & " , itemid int NULL" & VbCrlf
            sqlStr = sqlStr & " , mileage money NULL DEFAULT (0)" & VbCrlf
            sqlStr = sqlStr & " , mileagedate nvarchar(32) NULL" & VbCrlf
            sqlStr = sqlStr & " )" & VbCrlf
            sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_hpno ON #lms_target(hpno ASC)" & VbCrlf
            sqlStr = sqlStr & " CREATE NONCLUSTERED INDEX IX_userid ON #lms_target(userid ASC)" & VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            Set objFile = fs.OpenTextFile(sFilePath,1)
            i=0
            Do While objFile.AtEndOfStream <> True
                content=trim(objFile.ReadLine)    ' 파일을 라인단위로 읽음 한줄씩 처리. 다음줄을 읽으라는 명령을 안줘도 됨
                contentline = Split(content, ",")
                if isarray(contentline) then
                    useridfield = trim(contentline(0))
                    if ubound(contentline)>0 then
                        itemidfield = trim(contentline(1))
                    end if
                    if ubound(contentline)>1 then
                        mileagefield = trim(contentline(2))
                    end if
                    if ubound(contentline)>2 then
                        mileagedatefield = trim(contentline(3))
                    end if
                end if
                if itemidfield="" or isnull(itemidfield) then itemidfield = "NULL"
                if mileagefield="" or isnull(mileagefield) then mileagefield = "NULL"

                ' 두번째 줄부터
                if i > 0 then
                    if useridfield <> "" and not(isnull(useridfield)) then
                        ' 휴대폰번호
                        if targetkey="1" then
                            ' CSV파일을 엑셀로 열경우 앞자리가 0 서식때문에 기본으로 짤려서 없어짐. 답없음. 강제로 맨앞에 0추가 한다.
                            if left(useridfield,1)="1" then useridfield="0"&useridfield

                            sqlStr = "insert into #lms_target (hpno,useq,userid,pushyn,smsok,itemid) values (" & VbCrlf
                            sqlStr = sqlStr & " N'"& useridfield &"', NULL, NULL, N'N', N'N', NULL" & VbCrlf
                            sqlStr = sqlStr & " )" & VbCrlf

                            'response.write sqlStr & "<br>"
                            dbget.Execute sqlStr

                        ' 텐바이텐 고객번호일경우
                        elseif targetkey="2" then
                            if itemidexists then
                                if isnull(itemidfield) or itemidfield="" or itemidfield=0 then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('상품번호에 빈칸이 있습니다.');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                                itemidexistscnt=0
                                sqlStr = "SELECT" & vbcrlf
                                sqlStr = sqlStr & " count(itemid) as cnt" & vbcrlf
                                sqlStr = sqlStr & " From db_item.dbo.tbl_item with (readuncommitted)" & vbcrlf
                                sqlStr = sqlStr & " WHERE itemid = "& itemidfield &"" & vbcrlf

                                'response.write sqlStr &"<br>"
                                rsget.CursorLocation = adUseClient
                                rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                                IF not rsget.EOF THEN
                                    itemidexistscnt 	= rsget("cnt")
                                End IF			
                                rsget.Close	
                                if itemidexistscnt<1 then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('입력하신 상품번호 중에 존재하지 않는 상품이 있습니다.(상품번호 : "& itemidfield &")');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                            end if
                            if mileageexists then
                                if isnull(mileagefield) or mileagefield="" or mileagefield="0" then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('마일리지(지급/소멸)에 빈칸이 있습니다.');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                            end if
                            if mileagedateexists then
                                if isnull(mileagedatefield) or mileagedatefield="" then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('마일리지처리일(지급/소멸)에 빈칸이 있습니다.');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                            end if

                            sqlStr = "insert into #lms_target (hpno,useq,userid,pushyn,smsok,itemid,mileage,mileagedate) values (" & VbCrlf
                            sqlStr = sqlStr & " NULL, N'"& useridfield &"', NULL, N'N', N'N', "& itemidfield &", "& mileagefield &", '"& mileagedatefield &"'" & VbCrlf
                            sqlStr = sqlStr & " )" & VbCrlf

                            'response.write sqlStr & "<br>"
                            dbget.Execute sqlStr

                        ' 텐바이텐 고객아이디일경우
                        elseif targetkey="3" then
                            if itemidexists then
                                if isnull(itemidfield) or itemidfield="" or itemidfield=0 then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('상품번호에 빈칸이 있습니다.');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                                itemidexistscnt=0
                                sqlStr = "SELECT" & vbcrlf
                                sqlStr = sqlStr & " count(itemid) as cnt" & vbcrlf
                                sqlStr = sqlStr & " From db_item.dbo.tbl_item with (readuncommitted)" & vbcrlf
                                sqlStr = sqlStr & " WHERE itemid = "& itemidfield &"" & vbcrlf

                                'response.write sqlStr &"<br>"
                                rsget.CursorLocation = adUseClient
                                rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                                IF not rsget.EOF THEN
                                    itemidexistscnt 	= rsget("cnt")
                                End IF			
                                rsget.Close	
                                if itemidexistscnt<1 then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('입력하신 상품번호 중에 존재하지 않는 상품이 있습니다.(상품번호 : "& itemidfield &")');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                            end if
                            if mileageexists then
                                if isnull(mileagefield) or mileagefield="" or mileagefield="0" then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('마일리지(지급/소멸)에 빈칸이 있습니다.');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                            end if
                            if mileagedateexists then
                                if isnull(mileagedatefield) or mileagedatefield="" then
                                    response.write "<script type='text/javascript'>"
                                    response.write "	alert('마일리지처리일(지급/소멸)에 빈칸이 있습니다.');"
                                    response.write "</script>"
                                    session.codePage = 949
                                    dbget.close()	:	response.End
                                end if
                            end if

                            sqlStr = "insert into #lms_target (hpno,useq,userid,pushyn,smsok,itemid,mileage,mileagedate) values (" & VbCrlf
                            sqlStr = sqlStr & " NULL, NULL, N'"& useridfield &"', N'N', N'N', "& itemidfield &", "& mileagefield &", '"& mileagedatefield &"'" & VbCrlf
                            sqlStr = sqlStr & " )" & VbCrlf

                            'response.write sqlStr & "<br>"
                            'response.end
                            dbget.Execute sqlStr

                        end if

                        useridarr = useridarr & useridfield & ","
                    else
                        response.write "<script type='text/javascript'>"
                        response.write "	alert('텐바이텐고객 입력정보중에 빈칸이 있습니다.');"
                        response.write "</script>"
                        session.codePage = 949
                        dbget.close()	:	response.End
                    end if
                
                ' 첫줄
                else
                    ' 엑셀파일 첫번째 라인 구분자 체크
                    ' 휴대폰번호
                    if targetkey="1" then
                        if useridfield<>"휴대폰번호" then
                            response.write "<script type='text/javascript'>"
                            response.write "	alert('CSV파일이 휴대폰번호 가 아닙니다.\n다시한번 확인하세요.');"
                            response.write "</script>"
                            session.codePage = 949
                            dbget.close()	:	response.End
                        end if

                    ' 텐바이텐 고객번호일경우
                    elseif targetkey="2" then
                        if instr(useridfield,"텐바이텐고객번호") < 1 then
                            response.write "<script type='text/javascript'>"
                            response.write "	alert('CSV파일이 텐바이텐고객번호 가 아닙니다.\n다시한번 확인하세요.');"
                            response.write "</script>"
                            session.codePage = 949
                            dbget.close()	:	response.End
                        end if

                    ' 텐바이텐 고객아이디일경우
                    elseif targetkey="3" then
                        if instr(useridfield,"텐바이텐고객아이디") < 1 then
                            response.write "<script type='text/javascript'>"
                            response.write "	alert('CSV파일이 텐바이텐고객아이디 가 아닙니다.\n다시한번 확인하세요.');"
                            response.write "</script>"
                            session.codePage = 949
                            dbget.close()	:	response.End
                        end if
                    else
                        response.write "<script type='text/javascript'>"
                        response.write "	alert('CSV파일 구분자가 아닙니다.');"
                        response.write "</script>"
                        session.codePage = 949
                        dbget.close()	:	response.End
                    end if
                end if
                i = i + 1
            Loop

            useridarr = trim(left(useridarr,len(useridarr)-1))
            Set objFile = Nothing

            if ubound(split(useridarr,","))+1 > 30000 then
                response.write "<script type='text/javascript'>"
                response.write "	alert('3만명씩 업로드해 주세요.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if

            ' 고객정보 업데이트. 텐바이텐고객번호 일경우
            if targetkey="2" then
                sqlStr = "update t set t.hpno=u.usercell, t.userid=l.userid, t.smsok=u.smsok" & VbCrlf
                sqlStr = sqlStr & " from #lms_target t" & VbCrlf
                sqlStr = sqlStr & " join db_user.dbo.tbl_logindata l with (nolock)" & VbCrlf
                sqlStr = sqlStr & "     on t.useq=l.useq" & VbCrlf
                sqlStr = sqlStr & " join db_user.dbo.tbl_user_n u with (nolock)" & VbCrlf
                sqlStr = sqlStr & "     on l.userid=u.userid" & VbCrlf

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr

            ' 고객정보 업데이트. 텐바이텐고객아이디 일경우
            elseif targetkey="3" then
                sqlStr = "update t set t.hpno=u.usercell, t.smsok=u.smsok" & VbCrlf
                sqlStr = sqlStr & " from #lms_target t" & VbCrlf
                sqlStr = sqlStr & " join db_user.dbo.tbl_user_n u with (nolock)" & VbCrlf
                sqlStr = sqlStr & "     on t.userid=u.userid" & VbCrlf

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if

            ' 실제 테이블에 휴대폰번호 사이에 - 들어간거 전부 치환
            sqlStr = "update #lms_target set hpno=replace(isnull(hpno,''),'-','')" & VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            ' 고객정보중에 휴대폰번호가 입력 안된건 모두 삭제
            sqlStr = "delete from #lms_target where hpno=''" & VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            ' 2:텐바이텐고객번호 / 3:텐바이텐고객아이디
            if targetkey="2" or targetkey="3" then
                if member_smsok_checkyn="Y" then
                    ' 휴대폰수신거부자 제외
                    sqlStr = "delete from #lms_target where isnull(smsok,'N')='N'" & VbCrlf

                    'response.write sqlStr & "<br>"
                    dbget.Execute sqlStr
                end if
                if sendmethod="KAKAOALRIM" then
                    ' 알림톡수신거부자제외
                    if member_kakaoalrimyn_checkyn="Y" then
                        ' 휴대폰수신거부자 제외
                        sqlStr = "delete t from #lms_target t join db_contents.dbo.tbl_lms_agree a with (nolock) on t.userid=a.userid and a.kakaoalrimyn='N'" & VbCrlf

                        'response.write sqlStr & "<br>"
                        dbget.Execute sqlStr
                    end if
                end if

                ' 푸시수신정보
                if member_pushyn_checkyn<>"" then
                    sqlStr = "update t" & VbCrlf
                    sqlStr = sqlStr & " set t.pushyn='Y'" & VbCrlf
                    sqlStr = sqlStr & " from #lms_target as t" & VbCrlf
                    sqlStr = sqlStr & " join db_contents.dbo.tbl_app_regInfo as B with (noLock)" & VbCrlf
                    sqlStr = sqlStr & " 	on t.userid=B.userid" & VbCrlf
                    sqlStr = sqlStr & "     and B.pushyn='Y'" & VbCrlf
                    sqlStr = sqlStr & "     and B.isusing='Y'" & VbCrlf
                    sqlStr = sqlStr & "     and ((B.appkey=6 and B.appVer>='36')" & VbCrlf
                    sqlStr = sqlStr & "     or (B.appkey=5 and B.appVer>='1'))" & VbCrlf

                    'response.write sqlStr & "<br>"
                    dbget.Execute sqlStr
                end if

                ' 푸시수신여부 처리
                if member_pushyn_checkyn="Y" then
                    sqlStr = "delete from #lms_target where pushyn='N'" & VbCrlf

                    'response.write sqlStr & "<br>"
                    dbget.Execute sqlStr
                elseif member_pushyn_checkyn="N" then
                    sqlStr = "delete from #lms_target where pushyn='Y'" & VbCrlf

                    'response.write sqlStr & "<br>"
                    dbget.Execute sqlStr
                end if
            end if

            ' 최근발송데이터삭제
            if exception7dayyn<>"" then
                sqlStr = "delete s from #lms_target s join db_contents.dbo.tbl_lms_TargetTemp ss with (noLock) on s.hpno=ss.hpno and ss.ridx <>"& ridx &""

                if exception7dayyn="currentday" then    ' 오늘 
                    sqlStr = sqlStr & " and convert(nvarchar(10),ss.regdate,121) = convert(nvarchar(10),getdate(),121)"
                elseif exception7dayyn="before1day" then		' 1일전 
                    sqlStr = sqlStr & " and ss.regdate >= dateadd(day,-1,getdate())"
                elseif exception7dayyn="before2day" then		' 2일전 
                    sqlStr = sqlStr & " and ss.regdate >= dateadd(day,-2,getdate())"
                elseif exception7dayyn="before3day" then		' 3일전 
                    sqlStr = sqlStr & " and ss.regdate >= dateadd(day,-3,getdate())"
                elseif exception7dayyn="before7day" or exception7dayyn="Y" then		' 7일전 
                    sqlStr = sqlStr & " and ss.regdate >= dateadd(day,-7,getdate())"
                elseif exception7dayyn="before14day" then		' 14일전 
                    sqlStr = sqlStr & " and ss.regdate >= dateadd(day,-14,getdate())"
                elseif exception7dayyn="before21day" then		' 21일전 
                    sqlStr = sqlStr & " and ss.regdate >= dateadd(day,-21,getdate())"
                elseif exception7dayyn="currentmonth" then		' 당월 
                    sqlStr = sqlStr & " and ss.regdate>=convert(nvarchar(7),getdate(),121)+'-01'"
                elseif exception7dayyn="before1_1month" then		' 1달전(1일기준) 
                    sqlStr = sqlStr & " and ss.regdate>=convert(nvarchar(7),dateadd(month,-1,getdate()),121)+'-01'"
                elseif exception7dayyn="before2_1month" then		' 2달전(1일기준) 
                    sqlStr = sqlStr & " and ss.regdate>=convert(nvarchar(7),dateadd(month,-2,getdate()),121)+'-01'"
                elseif exception7dayyn="before3_1month" then		' 3달전(1일기준) 
                    sqlStr = sqlStr & " and ss.regdate>=convert(nvarchar(7),dateadd(month,-3,getdate()),121)+'-01'"
                end if

                sqlStr = sqlStr & " join db_contents.dbo.tbl_lms_reserve r with (noLock) on ss.ridx=r.ridx and r.isusing='Y' and r.sendmethod=N'"& sendmethod &"'"

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if

            if exceptionlogin<>"" then
                ' 로그인한사람제외
                if exceptionlogin="currentday" then		' 오늘 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),getdate(),121)"
                elseif exceptionlogin="before1day" then		' 1일전 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),dateadd(day,-1,getdate()),121)"
                elseif exceptionlogin="before2day" then		' 2일전 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),dateadd(day,-2,getdate()),121)"
                elseif exceptionlogin="before3day" then		' 3일전 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),dateadd(day,-3,getdate()),121)"
                elseif exceptionlogin="before7day" then		' 7일전 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),dateadd(day,-7,getdate()),121)"
                elseif exceptionlogin="before14day" then		' 14일전 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),dateadd(day,-14,getdate()),121)"
                elseif exceptionlogin="before21day" then		' 21일전 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(10),dateadd(day,-21,getdate()),121)"
                elseif exceptionlogin="currentmonth" then		' 당월 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(7),getdate(),121)+'-01'"
                elseif exceptionlogin="before1_1month" then		' 1달전(1일기준) 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(7),dateadd(month,-1,getdate()),121)+'-01'"
                elseif exceptionlogin="before2_1month" then		' 2달전(1일기준) 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(7),dateadd(month,-2,getdate()),121)+'-01'"
                elseif exceptionlogin="before3_1month" then		' 3달전(1일기준) 
                    sqlStr = "delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.lastlogin>=convert(nvarchar(7),dateadd(month,-3,getdate()),121)+'-01'"
                end if

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if

            ' 해당회원등급제외
            if exceptionuserlevelarr<>"" then
            	sqlStr="delete s from #lms_target s join db_user.dbo.tbl_logindata l with (nolock) on s.userid=l.userid and l.userlevel in ("& exceptionuserlevelarr &")"

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if

            sqlStr = "select count(hpno) as lms_targetcnt from #lms_target" + VbCrlf

            'response.write sqlStr & "<Br>"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
            If not rsget.EOF Then
                lms_targetcnt = rsget("lms_targetcnt")
            end if
            rsget.close

            if lms_targetcnt < 1 then
                response.write "<script type='text/javascript'>"
                response.write "	alert('CSV파일에 대상자가 없습니다.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if

            ' 타켓중으로 변경
            sqlStr = "update db_contents.dbo.tbl_lms_reserve set targetState=3 where ridx="& ridx &"" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            ' 기존에 타켓팅된게 있다면 삭제. CSV의 경우 파일을 타켓팅후 추가로 타켓 파일을 또 넣을수 있다.
            'sqlStr = "DELETE From db_contents.dbo.tbl_lms_targettemp where ridx="& ridx &"" + VbCrlf

            'response.write sqlStr & "<br>"
            'dbget.Execute sqlStr

            ' 대상자 입력
            sqlStr = "insert into db_contents.dbo.tbl_lms_targettemp(ridx,hpno,useq,userid,regdate,itemid,mileage,mileagedate)" + VbCrlf
            sqlStr = sqlStr & "	select distinct "& ridx &", t.hpno, t.useq, t.userid, getdate(), t.itemid, t.mileage, t.mileagedate" + VbCrlf
            sqlStr = sqlStr & "	from #lms_target T" + VbCrlf
            sqlStr = sqlStr & "	left join db_contents.dbo.tbl_lms_targettemp pt" + VbCrlf
            sqlStr = sqlStr & "	    on t.hpno = pt.hpno" + VbCrlf
            sqlStr = sqlStr & "	    and pt.ridx = "& ridx &"" + VbCrlf
            sqlStr = sqlStr & "	where pt.ridx is null" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            ' 최종 타게팅 업데이트
            sqlStr = "DECLARE @targetcnt int" + VbCrlf
            sqlStr = sqlStr & " select @targetcnt=count(*) from db_contents.dbo.tbl_lms_TargetTemp where" + VbCrlf
            sqlStr = sqlStr & " ridx="& ridx &"" + VbCrlf
            sqlStr = sqlStr & " update db_contents.dbo.tbl_lms_reserve" + VbCrlf
            sqlStr = sqlStr & " set targetState=7" + VbCrlf
            sqlStr = sqlStr & " ,state=1" + VbCrlf      ' 타켓시 발송예약으로 바로변경. 마케팅요청
            sqlStr = sqlStr & " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr & " ,lastadminid=N'"& lastadminid &"'" + VbCrlf
            sqlStr = sqlStr & " ,targetcnt=@targetcnt where" + VbCrlf
            sqlStr = sqlStr & " ridx="& ridx &"" + VbCrlf

            'response.write sqlStr & "<br>"
            dbget.Execute sqlStr

            sqlStr = "drop table #lms_target" + VbCrlf
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
        session.codePage = 949
		dbget.close() : response.end
	end if
	arrfileType = split(filetype,"^")
	chkReturn = 0
	for i = 0 to ubound(arrfiletype)
		if strFileType = arrfiletype(i) THEN chkReturn = 1 
	next
	if chkReturn = 0 then
		response.write fileType & " 형식의 파일만 가능합니다."
        session.codePage = 949
		dbget.close() : response.end
	end if
	fnChkFile = TRUE 
End Function
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>
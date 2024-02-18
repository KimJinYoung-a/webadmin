
<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : copyitem_process.asp
' Description :  이벤트 아이템 복사
' History : 2019.02.28 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim eCode, CCode, vChangeContents, vSCMChangeSQL, strSql

eCode  	= requestCheckVar(Request.Form("eC"),10)	'이벤트코드
CCode	= requestCheckVar(Request.Form("cC"),10)	'아이템 복사를 위한 이벤트 코드

Dim cnt , gcnt , tempi , tempii, eTemplate, eTemplate_mo

'//그룹개수
strSql = "select count(*) as gcnt " & VbCrlf
strSql = strSql & " from db_event.dbo.tbl_eventitem_group " & VbCrlf
strSql = strSql & " where evtgroup_using = 'Y' and evt_code = " & Ccode

rsget.Open strSql, dbget
IF not (rsget.EOF or rsget.BOF) THEN
    gcnt = rsget("gcnt")
END IF
rsget.close

'//화면템플릿 업데이트
strSql = " select evt_Template, case  when (evt_kind = 25 or evt_kind = 19 or evt_kind = 26) then evt_Template else evt_Template_mo end as evt_template_mo  from  db_event.dbo.tbl_event_display as d inner join db_event.dbo.tbl_event as e on d.evt_code = e.evt_code where d.evt_code = "&CCode&""
    rsget.Open strSql, dbget
IF not (rsget.EOF or rsget.BOF) THEN
    eTemplate = rsget("evt_Template")
    if eTemplate = "" or isNull(eTemplate) then eTemplate = "NULL"
    eTemplate_mo = rsget("evt_Template_mo")
    if eTemplate_mo = "" or isNull(eTemplate_mo) then eTemplate_mo = "NULL"
END IF
rsget.close

If gcnt > 0 Then '// 그룹이 있을 경우
    dbget.beginTrans '//트렌젝션  

    strSql = "update db_event.dbo.tbl_event_display set " & VbCrlf
    strSql = strSql &" evt_template =  "&eTemplate&"  , evt_template_mo=  "&eTemplate_mo &" where evt_code= " & eCode   
    dbget.execute strSql

    IF Err.Number = 0 Then
        '//그룹 일단 다 복사
        strSql = " insert into db_event.dbo.tbl_eventitem_group " & VbCrlf 
        strSql = strSql & " (evt_code , evtgroup_desc , evtgroup_sort " & VbCrlf
        strSql = strSql & " , evtgroup_pcode , evtgroup_depth , evtgroup_using, evtgroup_desc_mo, evtgroup_sort_mo, evtgroup_pcode_mo, evtgroup_depth_mo , evtgroup_isDisp, evtgroup_isDisp_mo) " & VbCrlf
        strSql = strSql & " select '"& eCode &"', t.evtgroup_desc  , t.evtgroup_sort " & VbCrlf
        strSql = strSql & " , t.evtgroup_pcode , t.evtgroup_depth , t.evtgroup_using, isNull(t.evtgroup_desc_mo,evtgroup_desc), isNull(t.evtgroup_sort_mo,t.evtgroup_sort) " & VbCrlf
        strSql = strSql & " , isNull(t.evtgroup_pcode_mo,t.evtgroup_pcode), isNull(t.evtgroup_depth_mo,t.evtgroup_depth) , isNull(t.evtgroup_isDisp, 1) , isNull(t.evtgroup_isDisp_mo,1)" & VbCrlf
        strSql = strSql & " From db_event.dbo.tbl_eventitem_group as t " & VbCrlf
        strSql = strSql & " where t.evt_code = '"& CCode &"' and t.evtgroup_using ='Y' " 
            
        dbget.execute strSql

        IF Err.Number = 0 Then
            '//후에 그룹코드 변경 업데이트
            strSql = " update b set " & VbCrlf
            strSql = strSql & " b.evtgroup_pcode = (select c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c where c.evt_code = b.evt_code and c.evtgroup_depth = a.evtgroup_depth and c.evtgroup_using ='Y' ) " & VbCrlf
            strSql = strSql & " from db_event.dbo.tbl_eventitem_group as a " & VbCrlf
            strSql = strSql & " inner join " & VbCrlf
            strSql = strSql & " db_event.dbo.tbl_eventitem_group as b " & VbCrlf
            strSql = strSql & " on a.evtgroup_code = b.evtgroup_pcode " & VbCrlf
            strSql = strSql & " where b.evt_code = '"& eCode &"' and b.evtgroup_using='Y' and a.evtgroup_using='Y' " 
            dbget.execute strSql

            '//모바일 그룹코드 변경 업데이트 
            strSql = " update b set " & VbCrlf
            strSql = strSql & " b.evtgroup_pcode_mo = (select distinct c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c where c.evt_code =  b.evt_code and c.evtgroup_depth_mo =  isNull(a.evtgroup_depth_mo,a.evtgroup_depth)  and c.evtgroup_using ='Y') " & VbCrlf
            strSql = strSql & " from db_event.dbo.tbl_eventitem_group as a " & VbCrlf
            strSql = strSql & " inner join " & VbCrlf
            strSql = strSql & " db_event.dbo.tbl_eventitem_group as b " & VbCrlf
            strSql = strSql & " on a.evtgroup_code = b.evtgroup_pcode_mo " & VbCrlf
            strSql = strSql & " where b.evt_code = '"& eCode &"'  and b.evtgroup_using='Y' and a.evtgroup_using='Y' "
                dbget.execute strSql

            strSql = " update g set " & VbCrlf
            strSql = strSql & "  evtgroup_code_mo =  (select min(evtgroup_code) from db_event.dbo.tbl_Eventitem_Group " & VbCrlf  
            strSql = strSql & "        where evt_code = g.evt_code and evtgroup_depth_mo = g.evtgroup_depth_mo and evtgroup_using ='Y' group by evtgroup_depth_mo) " & VbCrlf
            strSql = strSql & " from db_event.dbo.tbl_Eventitem_Group  as g " & VbCrlf
            strSql = strSql & " where evt_code =  '"& eCode &"' and evtgroup_using='Y'" & VbCrlf 
            dbget.execute strSql
            
            IF Err.Number = 0 Then
                '//상품 그룹복사 전체
                strSql = " insert into [db_event].[dbo].tbl_eventitem " & VbCrlf
                strSql = strSql & " (evt_code,itemid,evtgroup_code,evtitem_sort , evtitem_imgsize,evtitem_sort_mo, evtitem_isDisp, evtitem_isDisp_mo) " & VbCrlf
                strSql = strSql & " select '"& eCode &"', i.itemid, i.evtgroup_code ,i.evtitem_sort ,i.evtitem_imgsize, isNull(i.evtitem_sort_mo,i.evtitem_sort), isNull(i.evtitem_isDisp,1), isNull(i.evtitem_isDisp_mo,1) " & VbCrlf
                strSql = strSql & " from [db_event].[dbo].tbl_eventitem i " & VbCrlf
                strSql = strSql & " where evt_code= '"& CCode &"' and evtitem_isusing ='1' " & VbCrlf
                strSql = strSql & " and itemid not in " & VbCrlf
                strSql = strSql & " (select itemid from [db_event].[dbo].tbl_eventitem " & VbCrlf
                strSql = strSql & " where evt_code= '"& eCode &"' and evtitem_isusing ='1' " & VbCrlf
                strSql = strSql & " ) "

                dbget.execute strSql
                
                IF Err.Number = 0 Then
                    '//상품 그룹복사 - 그룹코드 전체 변경
                    strSql = " update i Set " & VbCrlf
                    strSql = strSql & " i.evtgroup_code =  " & VbCrlf
                    strSql = strSql & " (select c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c  " & VbCrlf
                    strSql = strSql & " 	where c.evt_code = '"& eCode &"'  " & VbCrlf
                    strSql = strSql & " 	and c.evtgroup_depth = a.evtgroup_depth  and c.evtgroup_using='Y' " & VbCrlf
                    strSql = strSql & " ) " & VbCrlf
                    strSql = strSql & " from [db_event].[dbo].tbl_eventitem as i " & VbCrlf
                    strSql = strSql & " inner Join " & VbCrlf
                    strSql = strSql & " db_event.dbo.tbl_eventitem_group as a " & VbCrlf
                    strSql = strSql & " on i.evtgroup_code = a.evtgroup_code " & VbCrlf
                    strSql = strSql & " where i.evt_code = '"& eCode &"' and a.evtgroup_using='Y' and i.evtitem_isusing ='1'"
                    dbget.execute strSql

                    IF Err.Number = 0 Then
                        dbget.CommitTrans
                        
                        vChangeContents = vChangeContents & "- 이벤트 상품 복사. " & CCode & " 상품을 " & eCode & " 로 복사" & vbCrLf
                        '### 수정 로그 저장(event)
                        vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
                        vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & CCode & "', '" & menupos & "', "
                        vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
                        dbget.execute(vSCMChangeSQL)
                        
                        Response.write "<script>alert('상품이 복사 되었습니다.');</script>"
                        Response.write "<script>parent.opener.location.reload();</script>"
                        Response.write "<script>parent.self.close();</script>"
                        dbget.close()	:	response.End
                    Else
                        dbget.RollBackTrans
                        Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
                    END IF
                Else
                    dbget.RollBackTrans
                    Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
                END IF
            Else
                dbget.RollBackTrans
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
            END IF
        Else 
            dbget.RollBackTrans
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
        END IF
    Else
        dbget.RollBackTrans
        Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
    END IF

Else '// 그룹이 없을경우 상품만 복사
    '//상품개수
    strSql = "select count(*) as cnt " & VbCrlf
    strSql = strSql & " from [db_event].[dbo].tbl_eventitem i "  & VbCrlf
    strSql = strSql & " where evt_code= " & CCode 
    strSql = strSql & " and itemid not in " & VbCrlf
    strSql = strSql & " (select itemid from [db_event].[dbo].tbl_eventitem " & VbCrlf
    strSql = strSql & " where evt_code= " & eCode & " and evtitem_isusing ='1' "&VbCrlf
    strSql = strSql & " ) and evtitem_isusing ='1' " 

    rsget.Open strSql, dbget
    IF not (rsget.EOF or rsget.BOF) THEN
        cnt = rsget("cnt")
    END IF
    rsget.close

'	Response.write cnt
'	Response.end
    
    If cnt > 0 Then 
    dbget.beginTrans '//트렌젝션

        strSql = " insert into [db_event].[dbo].tbl_eventitem " & VbCrlf
        strSql = strSql & " (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_imgsize, evtitem_sort_mo) " & VbCrlf
        strSql = strSql & " select " & CStr(eCode) & ", i.itemid, '0' ,evtitem_sort,i.evtitem_imgsize, isNull(i.evtitem_sort_mo, i.evtitem_sort)  " & VbCrlf
        strSql = strSql & " from [db_event].[dbo].tbl_eventitem i "  & VbCrlf
        strSql = strSql & " where evt_code= " & CCode 
        strSql = strSql & " and itemid not in " & VbCrlf
        strSql = strSql & " (select itemid from [db_event].[dbo].tbl_eventitem " & VbCrlf
        strSql = strSql & " where evt_code= " & eCode 
        strSql = strSql & "  and evtitem_isusing ='1' )  and evtitem_isusing ='1' " 

        dbget.execute strSql

        IF Err.Number = 0 Then
            dbget.CommitTrans
            
            vChangeContents = vChangeContents & "- 이벤트 상품 복사. " & CCode & " 상품을 " & eCode & " 로 복사" & vbCrLf
            '### 수정 로그 저장(event)
            vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
            vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & CCode & "', '" & menupos & "', "
            vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
            dbget.execute(vSCMChangeSQL)
            
            Response.write "<script>alert('상품이 복사 되었습니다.');</script>"
            response.write "<script type='text/javascript'>"
            response.write "	opener.document.location.reload();self.close();"
            response.write "</script>"
            dbget.close()	:	response.End
        Else
            dbget.RollBackTrans
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
        END IF
    Else
        Call sbAlertMsg ("이미 상품이 복사 되었습니다.", "back", "")
    End If 

End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
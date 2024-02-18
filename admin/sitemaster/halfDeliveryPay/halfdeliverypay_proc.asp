<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 배송비 반반 부담 설정 컨텐츠 등록 처리 페이지
' Hieditor : 2020.08.28 원승현 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
    dim startdate, enddate, idx
    dim sqlstr, i, mode
    dim starttime, endtime, halfdeliverypay, isusing, iid, adminid, loginUserId
    dim tmpiid, halfdeliverypayidx
    dim defaultDeliveryType, defaultFreeBeasongLimit, defaultDeliveryPay, makerid
    dim isusingtype, itemisusingarr, returncurrpage, returnitemname, returnresearch, returnitemid, returnstartdate, returnenddate
    dim returnisusing, returnbrandid, tmpidx, pageParam, returnregusertext, returnregusertype, itemdeliveryType, overlapalerttext

    mode                        =	requestCheckvar(Request("mode"),10)                     '// 처리 구분
    idx                         =	requestCheckvar(Request("idx"),20)                      '// 수정시 필요한 idx 값
    startdate                   =	requestCheckvar(Request("startdate"),20)                '// 시작일자
    enddate                     =	requestCheckvar(Request("enddate"),20)                  '// 종료일자
    starttime                   =	requestCheckvar(Request("starttime"),30)                '// 시작일자의 시간
    endtime                     =	requestCheckvar(Request("endtime"),30)                  '// 종료일자의 시간
    halfdeliverypay             =	requestCheckvar(Request("halfdeliverypay"),100)         '// 배송비 부담금액
    isusing                     =	requestCheckvar(Request("isusing"),10)                  '// 사용구분(y/n)
    iid                         =	requestCheckvar(Request("iid"),2048)                    '// 상품등록값(array)
    menupos                     =	requestCheckvar(Request("menupos"),50)                  '// 메뉴pos값
    adminid                     =	requestCheckvar(Request("adminid"),10)                  '// 관리자 아이디값
    loginUserId                 =   session("ssBctId")                                      '// 현재 로그인한 사용자의 아이디
    defaultDeliveryType         =	requestCheckvar(Request("defaultdeliveryType"),10)      '// 수정시 필요한 조건배송여부
    defaultFreeBeasongLimit     =	requestCheckvar(Request("defaultFreeBeasongLimit"),10)  '// 수정시 필요한 무료배송기준금액
    defaultDeliveryPay          =	requestCheckvar(Request("defaultDeliverPay"),10)        '// 수정시 필요한 배송비
    
    isusingtype                 =	requestCheckvar(Request("isusingtype"),10)              '// 사용여부 전체 수정시 필요한 값
    itemisusingarr              =	requestCheckvar(Request("itemisusingarr"),2048)         '// 사용여부 전체 수정시 수정될 상품 idx값
    returncurrpage              =	requestCheckvar(Request("returncurrpage"),10)           '// 사용여부 전체 수정시 처리 완료 후 돌아갈 페이지 값
    returnitemname              =	requestCheckvar(Request("returnitemname"),200)          '// 사용여부 전체 수정시 처리 완료 후 돌아갈 상품명 값
    returnresearch              =	requestCheckvar(Request("returnresearch"),10)           '// 사용여부 전체 수정시 처리 완료 후 돌아갈 검색여부 값
    returnitemid                =	requestCheckvar(Request("returnitemid"),2048)           '// 사용여부 전체 수정시 처리 완료 후 돌아갈 상품코드 값
    returnstartdate             =	requestCheckvar(Request("returnstartdate"),20)          '// 사용여부 전체 수정시 처리 완료 후 돌아갈 시작일 값
    returnenddate               =	requestCheckvar(Request("returnenddate"),20)            '// 사용여부 전체 수정시 처리 완료 후 돌아갈 종료일 값
    returnisusing               =	requestCheckvar(Request("returnisusing"),10)            '// 사용여부 전체 수정시 처리 완료 후 돌아갈 사용여부 값
    returnbrandid               =	requestCheckvar(Request("returnbrandid"),100)           '// 사용여부 전체 수정시 처리 완료 후 돌아갈 브랜드 아이디 값
    returnregusertype           =	requestCheckvar(Request("returnregusertype"),100)       '// 사용여부 전체 수정시 처리 완료 후 돌아갈 작성자 검색 구분 값
    returnregusertext           =	requestCheckvar(Request("returnregusertext"),100)       '// 사용여부 전체 수정시 처리 완료 후 돌아갈 작성자 검색 값
    overlapalerttext            =   ""

    if mode = "add" then
        if startdate="" or isNull(startdate) then
            response.write "<script>alert('시작일을 입력해주세요.');history.back();</script>"
            response.end
        end if	

        if enddate="" or isNull(enddate) then
            response.write "<script>alert('종료일을 입력해주세요.');history.back();</script>"
            response.end
        end if

        startdate = startdate&" "&starttime
        enddate = enddate&" "&endtime    

        if halfdeliverypay="" or isNull(halfdeliverypay) then
            response.write "<script>alert('배송비 부담금액을 입력해주세요.');history.back();</script>"
            response.end
        end If 

        if iid="" or isNull(iid) then
            response.write "<script>alert('상품을 입력해주세요.');history.back();</script>"
            response.end
        end If

        tmpiid = split(iid,",")

        for i=0 to ubound(tmpiid)
            sqlstr = "SELECT i.itemid, i.makerid, c.userid, c.defaultDeliveryType, c.defaultFreeBeasongLimit, c.defaultDeliverPay, p.idx, i.deliverytype "
            sqlstr = sqlstr & " FROM db_item.dbo.tbl_item i WITH(NOLOCK)"
            sqlstr = sqlstr & " INNER JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON i.makerid = c.userid "
            sqlstr = sqlstr & " LEFT JOIN db_sitemaster.dbo.tbl_HalfDeliveryPay p WITH(NOLOCK) ON i.itemid = p.itemid AND p.isusing='Y' "
            sqlstr = sqlstr & " WHERE i.itemid = '"&tmpiid(i)&"' "
            rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
            If not(rsget.bof or rsget.eof) Then
                halfdeliverypayidx = rsget("idx")
                defaultDeliveryType = rsget("defaultDeliveryType")
                defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
                defaultDeliveryPay = rsget("defaultDeliverPay")
                makerid = rsget("makerid")
                itemdeliveryType = rsget("deliverytype")
            Else
                response.write "<script>alert('선택하신 상품중에 정보가 없는 상품이 있습니다\n상품코드:"&tmpiid(i)&"');history.back();</script>"
                rsget.close
                response.end
            End If
            rsget.close

            If trim(halfdeliverypayidx) = "" or isNull(halfdeliverypayidx) Then
                sqlstr = "INSERT INTO db_sitemaster.dbo.tbl_halfdeliverypay (itemid, brandid, startdate, enddate, defaultdeliveryType, defaultFreeBeasongLimit, defaultDeliverPay, halfDeliveryPay, isusing, regdate, lastupdate, adminid, lastupdateadminid, itemdeliveryType)"
                sqlstr = sqlstr & " values ('"&tmpiid(i)&"','" & makerid & "','" & startdate & "' , '" & enddate & "', '" & defaultDeliveryType & "' , '" & defaultFreeBeasongLimit & "' , '" & defaultDeliveryPay & "' , '" & halfdeliverypay & "', '" & isusing & "', getdate(), getdate(), '" & loginUserId & "', '" & loginUserId & "', '" & itemdeliveryType & "')"
                dbget.execute sqlstr
            Else
                '// 기존에 등록된 상품이면 overlapalerttext 변수에 담아서 등록 alert 보여줄때 표시 해준다.)
                If Trim(overlapalerttext) = "" Then
                    overlapalerttext = tmpiid(i)
                Else
                    overlapalerttext = overlapalerttext&","&tmpiid(i)
                End If

                '// 사고 위험때문에 등록 시 기존에 등록된 상품은 업데이트를 하지 않는다.
                'sqlstr = " UPDATE db_sitemaster.dbo.tbl_halfdeliverypay SET "
                'sqlstr = sqlstr & " startdate = '"& startdate &"' "
                'sqlstr = sqlstr & " ,enddate = '"& enddate &"' "
                'sqlstr = sqlstr & " ,defaultdeliveryType = '"& defaultDeliveryType &"' "
                'sqlstr = sqlstr & " ,defaultFreeBeasongLimit = '"& defaultFreeBeasongLimit &"' "
                'sqlstr = sqlstr & " ,defaultDeliverPay = '"& defaultDeliveryPay &"' "
                'sqlstr = sqlstr & " ,halfDeliveryPay = '"& halfdeliverypay &"' "
                'sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
                'sqlstr = sqlstr & " ,lastupdate = getdate() "
                'sqlstr = sqlstr & " ,lastupdateadminid = '"& adminid &"' "
                'sqlstr = sqlstr & " where idx = "& halfdeliverypayidx &" "
                'response.write sqlstr
                'dbget.execute sqlstr
            End If
        next

        '// 기존에 등록된 상품을 다시 등록하면 alert을 띄워줌
        If Trim(overlapalerttext) <> "" Then
            response.write "<script>alert('기존에 등록된 상품인 "&overlapalerttext&"\n코드들을 제외한 상품이 등록되었습니다.');opener.location.href='index.asp';window.close();</script>"
        Else
            response.write "<script>alert('등록되었습니다.');opener.location.href='index.asp';window.close();</script>"
        End If
        response.end



	elseif mode = "edit" then
        if idx="" or isNull(idx) then
            response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
            response.end
        end If

        if startdate="" or isNull(startdate) then
            response.write "<script>alert('시작일을 입력해주세요.');history.back();</script>"
            response.end
        end if	

        if enddate="" or isNull(enddate) then
            response.write "<script>alert('종료일을 입력해주세요.');history.back();</script>"
            response.end
        end if

        startdate = startdate&" "&starttime
        enddate = enddate&" "&endtime    

        if halfdeliverypay="" or isNull(halfdeliverypay) then
            response.write "<script>alert('배송비 부담금액을 입력해주세요.');history.back();</script>"
            response.end
        end If

        sqlstr = " UPDATE db_sitemaster.dbo.tbl_halfdeliverypay SET "
        sqlstr = sqlstr & " startdate = '"& startdate &"' "
        sqlstr = sqlstr & " ,enddate = '"& enddate &"' "
        sqlstr = sqlstr & " ,defaultdeliveryType = '"& defaultDeliveryType &"' "
        sqlstr = sqlstr & " ,defaultFreeBeasongLimit = '"& defaultFreeBeasongLimit &"' "
        sqlstr = sqlstr & " ,defaultDeliverPay = '"& defaultDeliveryPay &"' "
        sqlstr = sqlstr & " ,halfDeliveryPay = '"& halfdeliverypay &"' "
        sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
        sqlstr = sqlstr & " ,lastupdate = getdate() "
        sqlstr = sqlstr & " ,lastupdateadminid = '"& adminid &"' "
        sqlstr = sqlstr & " where idx = "& idx &" "
        'response.write sqlstr
        dbget.execute sqlstr

        response.write "<script>document.domain='10x10.co.kr';alert('수정되었습니다.');opener.location.reload();window.close();</script>"
        response.end        

    elseif mode = "isusingall" Then
        tmpidx = split(itemisusingarr,",")

        for i=0 to ubound(tmpidx)
            sqlstr = " UPDATE db_sitemaster.dbo.tbl_halfdeliverypay SET "
            sqlstr = sqlstr & " isusing = '"& isusingtype &"' "
            sqlstr = sqlstr & " ,lastupdate = getdate() "
            sqlstr = sqlstr & " ,lastupdateadminid = '"& session("ssBctId") &"' "
            sqlstr = sqlstr & " where idx = "& tmpidx(i) &" "
            'response.write sqlstr
            dbget.execute sqlstr
        next

		If returncurrpage = "" Then returncurrpage = 1
		pageParam = "?page="&returncurrpage&"&itemname="& returnitemname &"&research="& returnresearch &"&itemid="& returnitemid &"&startdate="&returnstartdate &"&enddate="&returnenddate&"&isusing="&returnisusing&"&brandid="&returnbrandid&"&regusertype="&returnregusertype&"&regusertext="&returnregusertext


        response.write "<script>alert('수정되었습니다.');location.href='index.asp"&pageParam&"';</script>"
        response.end
	end If
%>

<script language = "javascript">
/*
    alert('저장되었습니다.');
    opener.location.href="index.asp<%=pageParam%>";
    window.close();
*/
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
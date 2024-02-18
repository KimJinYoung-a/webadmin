<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 보너스 쿠폰 적용 제외 상품or브랜드 등록 처리 페이지
' Hieditor : 2021.02.02 원승현 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
    dim idx
    dim sqlstr, i, mode
    dim isusing, itemid, adminid, loginUserId
    dim makerid, excludingCouponIdx, excludingCouponType
    dim isusingtype, itemisusingarr, returncurrpage, returnitemname, returnresearch, returnitemid, returnstartdate, returnenddate
    dim returnisusing, returnbrandid, tmpidx, pageParam, returnregusertext, returnregusertype, itemdeliveryType

    mode                        =	requestCheckvar(Request("mode"),10)                     '// 처리 구분
    idx                         =	requestCheckvar(Request("idx"),20)                      '// 수정시 필요한 idx 값
    isusing                     =	requestCheckvar(Request("isusing"),10)                  '// 사용구분(y/n)
    itemid                      =	requestCheckvar(Request("itemid"),2048)                 '// 상품등록값
    makerid                      =	requestCheckvar(Request("makerid"),2048)                '// 브랜드등록값    
    excludingCouponType         =	requestCheckvar(Request("excludingCouponType"),128)     '// 등록타입(I-상품, B-브랜드)
    menupos                     =	requestCheckvar(Request("menupos"),50)                  '// 메뉴pos값
    adminid                     =	requestCheckvar(Request("adminid"),10)                  '// 관리자 아이디값
    loginUserId                 =   session("ssBctId")                                      '// 현재 로그인한 사용자의 아이디
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

    if mode = "additem" then

        if itemid="" or isNull(itemid) then
            response.write "<script>alert('상품을 입력해주세요.');history.back();</script>"
            response.end
        end If

        sqlstr = "SELECT i.itemid, i.makerid, c.userid, p.idx, i.deliverytype "
        sqlstr = sqlstr & " FROM db_item.dbo.tbl_item i WITH(NOLOCK)"
        sqlstr = sqlstr & " INNER JOIN db_user.dbo.tbl_user_c c WITH(NOLOCK) ON i.makerid = c.userid "
        sqlstr = sqlstr & " LEFT JOIN db_order.dbo.tbl_ExcludingCouponData p WITH(NOLOCK) ON i.itemid = p.itemid "
        sqlstr = sqlstr & " WHERE i.itemid = '"&itemid&"' "
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
        If not(rsget.bof or rsget.eof) Then
            excludingCouponIdx = rsget("idx")
            makerid = rsget("makerid")
        Else
            response.write "<script>alert('상품 정보가 없습니다.\n상품코드:"&itemid&"');history.back();</script>"
            rsget.close
            response.end
        End If
        rsget.close

        If trim(excludingCouponIdx) = "" or isNull(excludingCouponIdx) Then
            sqlstr = "INSERT INTO db_order.dbo.tbl_ExcludingCouponData (type, itemid, isusing, regdate, lastupdate, adminid, lastupdateadminid)"
            sqlstr = sqlstr & " values ('I', '"&itemid&"', '" & isusing & "', getdate(), getdate(), '" & loginUserId & "', '" & loginUserId & "')"
            dbget.execute sqlstr
        Else
            '// 사고 위험때문에 등록 시 기존에 등록된 상품은 업데이트를 하지 않는다.
            'sqlstr = " UPDATE db_order.dbo.tbl_ExcludingCouponData SET "
            'sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
            'sqlstr = sqlstr & " ,lastupdate = getdate() "
            'sqlstr = sqlstr & " ,lastupdateadminid = '"& adminid &"' "
            'sqlstr = sqlstr & " where idx = "& excludingCouponIdx &" "
            'response.write sqlstr
            'dbget.execute sqlstr
        End If

        response.write "<script>alert('등록되었습니다.');opener.location.href='index.asp';window.close();</script>"
        response.end

    elseif mode = "addbrand" then
        if makerid="" or isNull(makerid) then
            response.write "<script>alert('브랜드를 입력해주세요.');history.back();</script>"
            response.end
        end If

        sqlstr = "SELECT c.userid, p.idx "
        sqlstr = sqlstr & " FROM db_user.dbo.tbl_user_c c WITH(NOLOCK) "
        sqlstr = sqlstr & " LEFT JOIN db_order.dbo.tbl_ExcludingCouponData p WITH(NOLOCK) ON c.userid = p.brandid "
        sqlstr = sqlstr & " WHERE c.userid = '"&makerid&"' "
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
        If not(rsget.bof or rsget.eof) Then
            excludingCouponIdx = rsget("idx")
        Else
            response.write "<script>alert('브랜드 정보가 없습니다.\n브랜드아이디:"&makerid&"');history.back();</script>"
            rsget.close
            response.end
        End If
        rsget.close

        If trim(excludingCouponIdx) = "" or isNull(excludingCouponIdx) Then
            sqlstr = "INSERT INTO db_order.dbo.tbl_ExcludingCouponData (type, brandid, isusing, regdate, lastupdate, adminid, lastupdateadminid)"
            sqlstr = sqlstr & " values ('B', '"&makerid&"', '" & isusing & "', getdate(), getdate(), '" & loginUserId & "', '" & loginUserId & "')"
            dbget.execute sqlstr
        Else
            '// 사고 위험때문에 등록 시 기존에 등록된 브랜드는 업데이트를 하지 않는다.
            'sqlstr = " UPDATE db_order.dbo.tbl_ExcludingCouponData SET "
            'sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
            'sqlstr = sqlstr & " ,lastupdate = getdate() "
            'sqlstr = sqlstr & " ,lastupdateadminid = '"& adminid &"' "
            'sqlstr = sqlstr & " where idx = "& excludingCouponIdx &" "
            'response.write sqlstr
            'dbget.execute sqlstr
        End If

        response.write "<script>alert('등록되었습니다.');opener.location.href='index.asp';window.close();</script>"
        response.end        

	elseif mode = "edit" then
        if idx="" or isNull(idx) then
            response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
            response.end
        end If

        sqlstr = " UPDATE db_order.dbo.tbl_ExcludingCouponData SET "
        sqlstr = sqlstr & " type = '"& excludingCouponType &"' "
        sqlstr = sqlstr & " ,itemid = '"& itemid &"' "
        sqlstr = sqlstr & " ,brandid = '"& makerid &"' "        
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
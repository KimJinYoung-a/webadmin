<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ʽ� ���� ���� ���� ��ǰor�귣�� ��� ó�� ������
' Hieditor : 2021.02.02 ������ ����
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

    mode                        =	requestCheckvar(Request("mode"),10)                     '// ó�� ����
    idx                         =	requestCheckvar(Request("idx"),20)                      '// ������ �ʿ��� idx ��
    isusing                     =	requestCheckvar(Request("isusing"),10)                  '// ��뱸��(y/n)
    itemid                      =	requestCheckvar(Request("itemid"),2048)                 '// ��ǰ��ϰ�
    makerid                      =	requestCheckvar(Request("makerid"),2048)                '// �귣���ϰ�    
    excludingCouponType         =	requestCheckvar(Request("excludingCouponType"),128)     '// ���Ÿ��(I-��ǰ, B-�귣��)
    menupos                     =	requestCheckvar(Request("menupos"),50)                  '// �޴�pos��
    adminid                     =	requestCheckvar(Request("adminid"),10)                  '// ������ ���̵�
    loginUserId                 =   session("ssBctId")                                      '// ���� �α����� ������� ���̵�
    isusingtype                 =	requestCheckvar(Request("isusingtype"),10)              '// ��뿩�� ��ü ������ �ʿ��� ��
    itemisusingarr              =	requestCheckvar(Request("itemisusingarr"),2048)         '// ��뿩�� ��ü ������ ������ ��ǰ idx��
    returncurrpage              =	requestCheckvar(Request("returncurrpage"),10)           '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� ������ ��
    returnitemname              =	requestCheckvar(Request("returnitemname"),200)          '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� ��ǰ�� ��
    returnresearch              =	requestCheckvar(Request("returnresearch"),10)           '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� �˻����� ��
    returnitemid                =	requestCheckvar(Request("returnitemid"),2048)           '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� ��ǰ�ڵ� ��
    returnstartdate             =	requestCheckvar(Request("returnstartdate"),20)          '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� ������ ��
    returnenddate               =	requestCheckvar(Request("returnenddate"),20)            '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� ������ ��
    returnisusing               =	requestCheckvar(Request("returnisusing"),10)            '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� ��뿩�� ��
    returnbrandid               =	requestCheckvar(Request("returnbrandid"),100)           '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� �귣�� ���̵� ��
    returnregusertype           =	requestCheckvar(Request("returnregusertype"),100)       '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� �ۼ��� �˻� ���� ��
    returnregusertext           =	requestCheckvar(Request("returnregusertext"),100)       '// ��뿩�� ��ü ������ ó�� �Ϸ� �� ���ư� �ۼ��� �˻� ��

    if mode = "additem" then

        if itemid="" or isNull(itemid) then
            response.write "<script>alert('��ǰ�� �Է����ּ���.');history.back();</script>"
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
            response.write "<script>alert('��ǰ ������ �����ϴ�.\n��ǰ�ڵ�:"&itemid&"');history.back();</script>"
            rsget.close
            response.end
        End If
        rsget.close

        If trim(excludingCouponIdx) = "" or isNull(excludingCouponIdx) Then
            sqlstr = "INSERT INTO db_order.dbo.tbl_ExcludingCouponData (type, itemid, isusing, regdate, lastupdate, adminid, lastupdateadminid)"
            sqlstr = sqlstr & " values ('I', '"&itemid&"', '" & isusing & "', getdate(), getdate(), '" & loginUserId & "', '" & loginUserId & "')"
            dbget.execute sqlstr
        Else
            '// ��� ���趧���� ��� �� ������ ��ϵ� ��ǰ�� ������Ʈ�� ���� �ʴ´�.
            'sqlstr = " UPDATE db_order.dbo.tbl_ExcludingCouponData SET "
            'sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
            'sqlstr = sqlstr & " ,lastupdate = getdate() "
            'sqlstr = sqlstr & " ,lastupdateadminid = '"& adminid &"' "
            'sqlstr = sqlstr & " where idx = "& excludingCouponIdx &" "
            'response.write sqlstr
            'dbget.execute sqlstr
        End If

        response.write "<script>alert('��ϵǾ����ϴ�.');opener.location.href='index.asp';window.close();</script>"
        response.end

    elseif mode = "addbrand" then
        if makerid="" or isNull(makerid) then
            response.write "<script>alert('�귣�带 �Է����ּ���.');history.back();</script>"
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
            response.write "<script>alert('�귣�� ������ �����ϴ�.\n�귣����̵�:"&makerid&"');history.back();</script>"
            rsget.close
            response.end
        End If
        rsget.close

        If trim(excludingCouponIdx) = "" or isNull(excludingCouponIdx) Then
            sqlstr = "INSERT INTO db_order.dbo.tbl_ExcludingCouponData (type, brandid, isusing, regdate, lastupdate, adminid, lastupdateadminid)"
            sqlstr = sqlstr & " values ('B', '"&makerid&"', '" & isusing & "', getdate(), getdate(), '" & loginUserId & "', '" & loginUserId & "')"
            dbget.execute sqlstr
        Else
            '// ��� ���趧���� ��� �� ������ ��ϵ� �귣��� ������Ʈ�� ���� �ʴ´�.
            'sqlstr = " UPDATE db_order.dbo.tbl_ExcludingCouponData SET "
            'sqlstr = sqlstr & " ,isusing = '"& isusing &"' "
            'sqlstr = sqlstr & " ,lastupdate = getdate() "
            'sqlstr = sqlstr & " ,lastupdateadminid = '"& adminid &"' "
            'sqlstr = sqlstr & " where idx = "& excludingCouponIdx &" "
            'response.write sqlstr
            'dbget.execute sqlstr
        End If

        response.write "<script>alert('��ϵǾ����ϴ�.');opener.location.href='index.asp';window.close();</script>"
        response.end        

	elseif mode = "edit" then
        if idx="" or isNull(idx) then
            response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
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

        response.write "<script>document.domain='10x10.co.kr';alert('�����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
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


        response.write "<script>alert('�����Ǿ����ϴ�.');location.href='index.asp"&pageParam&"';</script>"
        response.end
	end If
%>

<script language = "javascript">
/*
    alert('����Ǿ����ϴ�.');
    opener.location.href="index.asp<%=pageParam%>";
    window.close();
*/
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
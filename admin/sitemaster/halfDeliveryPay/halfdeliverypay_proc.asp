<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ۺ� �ݹ� �δ� ���� ������ ��� ó�� ������
' Hieditor : 2020.08.28 ������ ����
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

    mode                        =	requestCheckvar(Request("mode"),10)                     '// ó�� ����
    idx                         =	requestCheckvar(Request("idx"),20)                      '// ������ �ʿ��� idx ��
    startdate                   =	requestCheckvar(Request("startdate"),20)                '// ��������
    enddate                     =	requestCheckvar(Request("enddate"),20)                  '// ��������
    starttime                   =	requestCheckvar(Request("starttime"),30)                '// ���������� �ð�
    endtime                     =	requestCheckvar(Request("endtime"),30)                  '// ���������� �ð�
    halfdeliverypay             =	requestCheckvar(Request("halfdeliverypay"),100)         '// ��ۺ� �δ�ݾ�
    isusing                     =	requestCheckvar(Request("isusing"),10)                  '// ��뱸��(y/n)
    iid                         =	requestCheckvar(Request("iid"),2048)                    '// ��ǰ��ϰ�(array)
    menupos                     =	requestCheckvar(Request("menupos"),50)                  '// �޴�pos��
    adminid                     =	requestCheckvar(Request("adminid"),10)                  '// ������ ���̵�
    loginUserId                 =   session("ssBctId")                                      '// ���� �α����� ������� ���̵�
    defaultDeliveryType         =	requestCheckvar(Request("defaultdeliveryType"),10)      '// ������ �ʿ��� ���ǹ�ۿ���
    defaultFreeBeasongLimit     =	requestCheckvar(Request("defaultFreeBeasongLimit"),10)  '// ������ �ʿ��� �����۱��رݾ�
    defaultDeliveryPay          =	requestCheckvar(Request("defaultDeliverPay"),10)        '// ������ �ʿ��� ��ۺ�
    
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
    overlapalerttext            =   ""

    if mode = "add" then
        if startdate="" or isNull(startdate) then
            response.write "<script>alert('�������� �Է����ּ���.');history.back();</script>"
            response.end
        end if	

        if enddate="" or isNull(enddate) then
            response.write "<script>alert('�������� �Է����ּ���.');history.back();</script>"
            response.end
        end if

        startdate = startdate&" "&starttime
        enddate = enddate&" "&endtime    

        if halfdeliverypay="" or isNull(halfdeliverypay) then
            response.write "<script>alert('��ۺ� �δ�ݾ��� �Է����ּ���.');history.back();</script>"
            response.end
        end If 

        if iid="" or isNull(iid) then
            response.write "<script>alert('��ǰ�� �Է����ּ���.');history.back();</script>"
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
                response.write "<script>alert('�����Ͻ� ��ǰ�߿� ������ ���� ��ǰ�� �ֽ��ϴ�\n��ǰ�ڵ�:"&tmpiid(i)&"');history.back();</script>"
                rsget.close
                response.end
            End If
            rsget.close

            If trim(halfdeliverypayidx) = "" or isNull(halfdeliverypayidx) Then
                sqlstr = "INSERT INTO db_sitemaster.dbo.tbl_halfdeliverypay (itemid, brandid, startdate, enddate, defaultdeliveryType, defaultFreeBeasongLimit, defaultDeliverPay, halfDeliveryPay, isusing, regdate, lastupdate, adminid, lastupdateadminid, itemdeliveryType)"
                sqlstr = sqlstr & " values ('"&tmpiid(i)&"','" & makerid & "','" & startdate & "' , '" & enddate & "', '" & defaultDeliveryType & "' , '" & defaultFreeBeasongLimit & "' , '" & defaultDeliveryPay & "' , '" & halfdeliverypay & "', '" & isusing & "', getdate(), getdate(), '" & loginUserId & "', '" & loginUserId & "', '" & itemdeliveryType & "')"
                dbget.execute sqlstr
            Else
                '// ������ ��ϵ� ��ǰ�̸� overlapalerttext ������ ��Ƽ� ��� alert �����ٶ� ǥ�� ���ش�.)
                If Trim(overlapalerttext) = "" Then
                    overlapalerttext = tmpiid(i)
                Else
                    overlapalerttext = overlapalerttext&","&tmpiid(i)
                End If

                '// ��� ���趧���� ��� �� ������ ��ϵ� ��ǰ�� ������Ʈ�� ���� �ʴ´�.
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

        '// ������ ��ϵ� ��ǰ�� �ٽ� ����ϸ� alert�� �����
        If Trim(overlapalerttext) <> "" Then
            response.write "<script>alert('������ ��ϵ� ��ǰ�� "&overlapalerttext&"\n�ڵ���� ������ ��ǰ�� ��ϵǾ����ϴ�.');opener.location.href='index.asp';window.close();</script>"
        Else
            response.write "<script>alert('��ϵǾ����ϴ�.');opener.location.href='index.asp';window.close();</script>"
        End If
        response.end



	elseif mode = "edit" then
        if idx="" or isNull(idx) then
            response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
            response.end
        end If

        if startdate="" or isNull(startdate) then
            response.write "<script>alert('�������� �Է����ּ���.');history.back();</script>"
            response.end
        end if	

        if enddate="" or isNull(enddate) then
            response.write "<script>alert('�������� �Է����ּ���.');history.back();</script>"
            response.end
        end if

        startdate = startdate&" "&starttime
        enddate = enddate&" "&endtime    

        if halfdeliverypay="" or isNull(halfdeliverypay) then
            response.write "<script>alert('��ۺ� �δ�ݾ��� �Է����ּ���.');history.back();</script>"
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
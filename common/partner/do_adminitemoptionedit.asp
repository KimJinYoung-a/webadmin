<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ���
' History : ���ʻ����ڸ�
'			2017.04.10 �ѿ�� ����(���Ȱ���ó��)
'           2019.04.23 ������ �ɼ� ���� ���ϵ��� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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

function checkIsUpcheItemEditValid(itemid)
    dim sqlStr
    checkIsUpcheItemEditValid = false
    if (C_ADMIN_USER) then
        checkIsUpcheItemEditValid = True
        Exit function
    end if

    if (C_IS_Maker_Upche) and (session("ssBctId")<>"") then
        sqlStr = " select top 1 itemid from db_item.dbo.tbl_item " &VBCRLF
        sqlStr = sqlStr & " where itemid="&itemid&VBCRLF
        sqlStr = sqlStr & " and makerid='"&session("ssBctId")&"'"&VBCRLF

        rsget.Open sqlStr, dbget, 1
        if Not rsget.Eof then
    	    checkIsUpcheItemEditValid = (rsget.RecordCount>0)
    	end if
    	rsget.close
    end if

end function

dim refer, vChangeContents, vSCMChangeSQL
refer = request.ServerVariables("HTTP_REFERER")

dim itemid, itemoption
dim mode
dim arritemoption, arritemoptionname
dim optionTypename, optionName
dim optaddprice, optaddBuyprice

dim i, j, k, index, sqlStr, foundcount, found, ErrStr
dim TypeSeq, KindSeq

dim TypeCnt, OptCnt
 dim sRetValue

itemid              = requestCheckvar(request("itemid"),10)
itemoption          = requestCheckvar(request("itemoption"),4)
mode                = requestCheckVar(request("mode"),32)
optionTypename      = requestCheckVar(request("optionTypename"),32)
arritemoption       = request("arritemoption")
arritemoptionname   = request("arritemoptionname")

TypeSeq             = requestCheckvar(request("TypeSeq"),10)
KindSeq             = requestCheckvar(request("KindSeq"),10)

arritemoption = Split(arritemoption, "|")
arritemoptionname = Split(arritemoptionname, "|")

vChangeContents = vChangeContents & "��ǰ �ɼ� " & vbCrLf
vChangeContents = vChangeContents & "- ��ǰ�ڵ� : itemid = " & itemid & vbCrLf
vChangeContents = vChangeContents & "- mode ���簪 : mode = " & mode & vbCrLf

if itemid="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�ڵ尡 �����ϴ�.');"
	response.write "</script>"
	dbget.close()	:	response.End
end if

dim IsUpchebeasong, itemLimitYn
IsUpchebeasong =false
itemLimitYn = "N"

''��ü �����ΰ�� �귣��ID üũ// ��ġ�� ������ CASE �ִµ�.
if (Not checkIsUpcheItemEditValid(itemid)) then
    response.write "<script>alert('������ �����ϴ�. �ش� �귣�� ��ǰ�� �ƴմϴ�.');</script>"
    dbget.Close(): response.end
end if

''��ü����ΰ�� ����/�Ǹ� ������� ����
sqlStr = " select limityn, deliverytype "
sqlStr = sqlStr & " from [db_item].[dbo].tbl_item"
sqlStr = sqlStr & " where itemid=" & CStr(itemid)

rsget.Open sqlStr,dbget,1
if not rsget.EOF then
    itemLimitYn = rsget("limityn")
    IsUpchebeasong = (rsget("deliverytype") = "2") or (rsget("deliverytype") = "5") or (rsget("deliverytype") = "9") or (rsget("deliverytype") = "7")
end if
rsget.Close

''response.write mode
function ReCalcuItemOption(itemid)
    dim sqlStr
    ''��ǰ�ɼǼ���
	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf
	dbget.Execute sqlStr

	''��ǰ��������
	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf
	dbget.Execute sqlStr

	''������ ���̸� �Ͻ� ǰ��ó�� // 2013/09/02 �߰�
    sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    sqlStr = sqlStr + " set sellyn='S'" + VBCrlf
    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + VBCrlf
    sqlStr = sqlStr + " and sellyn='Y'" + VBCrlf
    sqlStr = sqlStr + " and limityn='Y'" + VBCrlf
    sqlStr = sqlStr + " and (limitno-limitsold)<1" + VBCrlf
    dbget.Execute sqlStr
end function

function ReMatchMultiOption(itemid)
    dim sqlStr
    dim MultiLevel

    MultiLevel = 0

    sqlStr = " select TypeSeq, Count(KindSeq) as KindCnt " &VbCRLF
    sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_Multiple " &VbCRLF
    sqlStr = sqlStr + " where itemid=" + CStr(itemid) &VbCRLF
    sqlStr = sqlStr + " group by TypeSeq" &VbCRLF
    sqlStr = sqlStr + " order by TypeSeq" &VbCRLF

    rsget.Open sqlStr, dbget, 1
	    MultiLevel = rsget.RecordCount
	rsget.close

    ''���� 2�� �ɼ��� ��� ����.
    if (MultiLevel=3) then
        sqlStr = " delete from [db_item].[dbo].tbl_item_option" &VbCRLF
        sqlStr = sqlStr + " where itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and Left(itemoption,1)='Z'" &VbCRLF
        sqlStr = sqlStr + " and Right(itemoption,1)='0'" &VbCRLF

        dbget.Execute sqlStr
    end if

'    if (MultiLevel=2) then
'        sqlStr = " delete from [db_item].[dbo].tbl_item_option"
'        sqlStr = sqlStr + " where itemid=" + CStr(itemid)
'        sqlStr = sqlStr + " and Left(itemoption,1)='Z'"
'        sqlStr = sqlStr + " and Right(itemoption,1)='00'"
'
'        dbget.Execute sqlStr
'    end if

''response.write  MultiLevel
    ''�ɼ� ���ۼ�.
'   --Only 1�߿ɼ�.
    if (MultiLevel=1) then
        ''-- �� �ɼ� ����;
        sqlStr = " delete from [db_item].[dbo].tbl_item_option_Multiple" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        dbget.Execute sqlStr

        sqlStr = " delete from [db_item].[dbo].tbl_item_option" & VbCrlf
        sqlStr = sqlStr & " where itemid=" + CStr(itemid)
        sqlStr = sqlStr & " and Left(itemoption,1)='Z'"
        dbget.Execute sqlStr

''        sqlStr = " insert into [db_item].[dbo].tbl_item_option"
''        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) "
''        sqlStr = sqlStr + " select T.itemid, ('Z' + convert(varchar(1),T.KindSeq) + '0' + '0') as itemoption,"
''        sqlStr = sqlStr + " T.optionTypeName, T.optionKindName, 'Y','Y','" + itemLimitYn + "', 0, 0,"
''        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice"
''        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_Multiple T"
''        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o "
''        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid)
''        sqlStr = sqlStr + "     and T.itemid=o.itemid "
''        sqlStr = sqlStr + "     and ('Z' + convert(varchar(1),T.KindSeq) + '0' + '0')=o.itemoption "
''        sqlStr = sqlStr + " where  o.itemid is NULL"
''
''        dbget.Execute sqlStr
''
''        '' �ɼǸ�/ ���� ���� ����� ���
''        sqlStr = " update [db_item].[dbo].tbl_item_option"
''        sqlStr = sqlStr + " set optionname=T.optionname"
''        sqlStr = sqlStr + " , optaddprice=T.optaddprice"
''        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice"
''        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_Multiple T "
''        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.itemid"
''        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
''        sqlStr = sqlStr + " and ("
''        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option.optionname<>T.optionname"
''        sqlStr = sqlStr + "     or [db_item].[dbo].tbl_item_option.optaddprice<>T.optaddprice"
''        sqlStr = sqlStr + "     or [db_item].[dbo].tbl_item_option.optaddbuyprice<>T.optaddbuyprice"
''        sqlStr = sqlStr + " )"
''
''        dbget.Execute sqlStr
    end if

'   --Only 2�߿ɼ�.
    if (MultiLevel=2) then
        sqlStr = " insert into [db_item].[dbo].tbl_item_option" &VbCRLF
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) " &VbCRLF
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '���տɼ�' as optionTypeName," &VbCRLF
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0," &VbCRLF
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_item].[dbo].tbl_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option_Multiple b" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T" &VbCRLF
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o " &VbCRLF
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and T.itemid=o.itemid " &VbCRLF
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption " &VbCRLF
        sqlStr = sqlStr + " where  o.itemid is NULL"

        dbget.Execute sqlStr

        '' �ɼǸ�/ ���� ���� ����� ���
        sqlStr = " update [db_item].[dbo].tbl_item_option" &VbCRLF
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)" &VbCRLF
        sqlStr = sqlStr + " , optaddprice=T.optaddprice" &VbCRLF
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + '0') as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName ) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_item].[dbo].tbl_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option_Multiple b" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T " &VbCRLF
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.itemid" &VbCRLF
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption" &VbCRLF
        sqlStr = sqlStr + " and (" &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option.optionname<>T.optionname" &VbCRLF
        sqlStr = sqlStr + "     or [db_item].[dbo].tbl_item_option.optaddprice<>T.optaddprice" &VbCRLF
        sqlStr = sqlStr + "     or [db_item].[dbo].tbl_item_option.optaddbuyprice<>T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " )"
'response.write sqlStr
        dbget.Execute sqlStr
    end if

'    --Only 3�߿ɼ�
    if (MultiLevel=3) then
        sqlStr = " insert into [db_item].[dbo].tbl_item_option" &VbCRLF
        sqlStr = sqlStr + " (itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, optaddprice, optaddbuyprice) " &VbCRLF
        sqlStr = sqlStr + " select T.itemid, T.itemoption, '���տɼ�' as optionTypeName," &VbCRLF
        sqlStr = sqlStr + " convert(varchar(96),T.optionname), 'Y','Y','" + itemLimitYn + "', 0, 0," &VbCRLF
        sqlStr = sqlStr + " T.optaddprice, T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_item].[dbo].tbl_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option_Multiple b," &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option_Multiple c" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and b.itemid=c.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T " &VbCRLF
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option o " &VbCRLF
        sqlStr = sqlStr + "     on o.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and T.itemid=o.itemid " &VbCRLF
        sqlStr = sqlStr + "     and T.itemoption=o.itemoption " &VbCRLF
        sqlStr = sqlStr + " where  o.itemid is NULL"

        dbget.Execute sqlStr


        '' �ɼǸ�/ ���� ���� ����� ���
        sqlStr = " update [db_item].[dbo].tbl_item_option" &VbCRLF
        sqlStr = sqlStr + " set optionname=convert(varchar(96),T.optionname)" &VbCRLF
        sqlStr = sqlStr + " , optaddprice=T.optaddprice" &VbCRLF
        sqlStr = sqlStr + " , optaddbuyprice=T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " from (" &VbCRLF
        sqlStr = sqlStr + "     select  a.itemid, ('Z' + convert(varchar(1),a.KindSeq) + convert(varchar(1),b.KindSeq) + convert(varchar(1),c.KindSeq)) as itemoption ," &VbCRLF
        sqlStr = sqlStr + "     (A.optionKindName + ',' + B.optionKindName + ',' + C.optionKindName) as optionname," &VbCRLF
        sqlStr = sqlStr + "     (A.optAddprice+B.optaddprice+C.optaddprice) as optaddprice , (A.optAddbuyprice+B.optaddbuyprice+C.optaddbuyprice) as optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + "     from [db_item].[dbo].tbl_item_option_Multiple a," &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option_Multiple b," &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option_Multiple c" &VbCRLF
        sqlStr = sqlStr + "     where a.itemid=" + CStr(itemid) &VbCRLF
        sqlStr = sqlStr + "     and a.itemid=b.itemid" &VbCRLF
        sqlStr = sqlStr + "     and b.itemid=c.itemid" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<>b.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<>c.TypeSeq" &VbCRLF
        sqlStr = sqlStr + "     and a.TypeSeq<b.TypeSeq " &VbCRLF
        sqlStr = sqlStr + "     and b.TypeSeq<c.TypeSeq " &VbCRLF
        sqlStr = sqlStr + " ) T " &VbCRLF
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.itemid" &VbCRLF
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption" &VbCRLF
        sqlStr = sqlStr + " and (" &VbCRLF
        sqlStr = sqlStr + "     [db_item].[dbo].tbl_item_option.optionname<>T.optionname" &VbCRLF
        sqlStr = sqlStr + "     or [db_item].[dbo].tbl_item_option.optaddprice<>T.optaddprice" &VbCRLF
        sqlStr = sqlStr + "     or [db_item].[dbo].tbl_item_option.optaddbuyprice<>T.optaddbuyprice" &VbCRLF
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr
    end if

    ''������ �ɼ� ���� 2013/09/30 �߰�
    if (MultiLevel=2) or (MultiLevel=3) then
        sqlStr = "exec [db_item].[dbo].sp_TEN_MultiOptionNotMatchDEL "& CStr(itemid)
        dbget.Execute sqlStr
    end if
end function


'' �ɼǼ��� - ���߿ɼ�
if (mode="editOptionMultiple") then
    ''TypeCnt, OptCnt
    TypeCnt = request("optionTypename").count

    for i=1 to TypeCnt
        optionTypename  = requestCheckVar(Trim(request("optionTypename")(i)),32)
        TypeSeq         = requestCheckVar(Trim(request("TypeSeqTmp")(i)),10)

        sqlStr = "update [db_item].[dbo].tbl_item_option_Multiple" &VbCRLF
        sqlStr = sqlStr + " set optionTypeName='" + html2Db(optionTypename) + "'" &VbCRLF
        sqlStr = sqlStr + " where itemid=" & CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and TypeSeq=" & CStr(TypeSeq) &VbCRLF
        sqlStr = sqlStr + " and optionTypeName<>'" + html2Db(optionTypename) + "'" &VbCRLF

        dbget.Execute sqlStr

    next

    OptCnt  = request("KindSeq").count
    for i=1 to OptCnt
        TypeSeq     = requestCheckVar(Trim(request("TypeSeq")(i)),10)
        KindSeq     = requestCheckVar(Trim(request("KindSeq")(i)),10)
        optionName  = requestCheckVar(Trim(request("optionName")(i)),96)
        optaddprice = requestCheckVar(Trim(request("optaddprice")(i)),20)
        optaddBuyprice = requestCheckVar(Trim(request("optaddBuyprice")(i)),20)
        if (optaddprice="") then optaddprice="0"  ''�߰� 2013/06/18
        if optaddBuyprice = "" then optaddBuyprice = 0
        IF optaddprice > 0 and optaddBuyprice = 0 then '�߰����� �ִ� ��� �߰� ���ް� �ԷµǾ�� �Ѵ�. 2015-07-21
            response.write "<script language='javascript'>alert('�߰��ݾ��� ���ް��� �����Ǿ����� �ʽ��ϴ�.Ȯ�� �� �ٽ� ������ּ��� '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
              dbget.close()	:	response.End
        end if

        if optaddprice < 0  then
            response.write "<script language='javascript'>alert('�߰��ݾ� ���ް��� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�1. (�߰��ݾ��� ������ 0) '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
        if optaddBuyprice < 0  then
            response.write "<script language='javascript'>alert('�߰��ݾ� ���ް��� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�1. (�߰��ݾ��� ������ 0) '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if

        sqlStr = "update [db_item].[dbo].tbl_item_option_Multiple" &VbCRLF
        sqlStr = sqlStr + " set optionKindName='" + html2Db(optionName) + "'" &VbCRLF
        sqlStr = sqlStr + " ,optaddprice=" & CStr(optaddprice) &VbCRLF
        sqlStr = sqlStr + " ,optaddBuyprice=" & CStr(optaddBuyprice) &VbCRLF
        sqlStr = sqlStr + " where itemid=" & CStr(itemid) &VbCRLF
        sqlStr = sqlStr + " and TypeSeq=" & CStr(TypeSeq) &VbCRLF
        sqlStr = sqlStr + " and KindSeq='" & CStr(KindSeq) & "'" &VbCRLF
        sqlStr = sqlStr + " and (" &VbCRLF
        sqlStr = sqlStr + "     (optionKindName<>'" + html2Db(optionName) + "')" &VbCRLF
        sqlStr = sqlStr + "     or (optaddprice<>" & CStr(optaddprice) & ")" &VbCRLF
        sqlStr = sqlStr + "     or (optaddBuyprice<>" & CStr(optaddBuyprice) & ")" &VbCRLF
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr

    next

    vChangeContents = vChangeContents & "- �ɼǼ��� - ���߿ɼ�" & vbCrLf
    vChangeContents = vChangeContents & "- TypeSeq = " & CStr(TypeSeq) & vbCrLf
    vChangeContents = vChangeContents & "- KindSeq = " & CStr(KindSeq) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǱ��и� : optionTypename = " & html2Db(request("optionTypename")) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǻ󼼸� : optionKindName = " & html2Db(request("optionName")) & vbCrLf
    vChangeContents = vChangeContents & "- �߰����� : optaddprice = " & request("optaddprice") & vbCrLf
    vChangeContents = vChangeContents & "- ���ް� : optaddBuyprice = " & request("optaddBuyprice") & vbCrLf

   	'### ���� �α� ����(item option)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'itemoption', '" & itemid & "', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

    Call ReMatchMultiOption(itemid)

    Call ReCalcuItemOption(itemid)

    response.write "<script language='javascript'>alert('���� �Ǿ����ϴ�.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if

'' �ɼǼ��� - ���Ͽɼ�
if (mode="editOption") then

    sqlStr = "update [db_item].[dbo].tbl_item_option"
    sqlStr = sqlStr + " set optionTypeName='" + html2Db(optionTypeName) + "'"
    sqlStr = sqlStr + " where itemid=" & CStr(itemid)
    sqlStr = sqlStr + " and optionTypeName<>'" + html2Db(optionTypeName) + "'"
    dbget.Execute sqlStr

    OptCnt = request("itemoption").count

    for i=1 to OptCnt
        itemoption = requestCheckVar(Trim(request("itemoption")(i)),4)
        optionName = requestCheckVar(Trim(request("optionName")(i)),96)
        optaddprice = requestCheckVar(Trim(request("optaddprice")(i)),20)
        optaddBuyprice = requestCheckVar(Trim(request("optaddBuyprice")(i)),20)

        if (optaddprice="") then optaddprice="0"  ''�߰� 2013/06/18
        if optaddBuyprice = "" then optaddBuyprice = 0
        IF optaddprice > 0 and optaddBuyprice = 0 then '�߰����� �ִ� ��� �߰� ���ް� �ԷµǾ�� �Ѵ�. 2015-07-21
            response.write "<script language='javascript'>alert('�߰��ݾ��� ���ް��� �����Ǿ����� �ʽ��ϴ�.Ȯ�� �� �ٽ� ������ּ��� '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
              dbget.close()	:	response.End
        end if

        if optaddprice < 0  then
            response.write "<script language='javascript'>alert('�߰��ݾ� ���ް��� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�1. (�߰��ݾ��� ������ 0) '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
        if optaddBuyprice < 0  then
            response.write "<script language='javascript'>alert('�߰��ݾ� ���ް��� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�1. (�߰��ݾ��� ������ 0) '); </script>"
            response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if

        if (Len(itemoption)=4) and (Len(optionName)>0) then
            sqlStr = "update [db_item].[dbo].tbl_item_option"&VbCRLF
            sqlStr = sqlStr + " set optionName='" + html2Db(optionName) + "'"&VbCRLF
            sqlStr = sqlStr + " ,optaddprice=" & CStr(optaddprice)&VbCRLF
            sqlStr = sqlStr + " ,optaddBuyprice=" & CStr(optaddBuyprice)&VbCRLF
            sqlStr = sqlStr + " where itemid=" & CStr(itemid)&VbCRLF
            sqlStr = sqlStr + " and itemoption='" & itemoption & "'"&VbCRLF
            sqlStr = sqlStr + " and ("
            sqlStr = sqlStr + "     (optionName<>'" + html2Db(optionName) + "')"&VbCRLF
            sqlStr = sqlStr + "     or (optaddprice<>" & CStr(optaddprice) & ")"&VbCRLF
            sqlStr = sqlStr + "     or (optaddBuyprice<>" & CStr(optaddBuyprice) & ")"&VbCRLF
            sqlStr = sqlStr + " )"

            dbget.Execute sqlStr

        end if
    next

    vChangeContents = vChangeContents & "- �ɼǼ��� - ���Ͽɼ�" & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ��ڵ� : itemoption = " & request("itemoption") & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǱ��и� : optionTypename = " & html2Db(request("optionTypename")) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǻ󼼸� : optionName = " & html2Db(request("optionName")) & vbCrLf
    vChangeContents = vChangeContents & "- �߰����� : optaddprice = " & request("optaddprice") & vbCrLf
    vChangeContents = vChangeContents & "- ���ް� : optaddBuyprice = " & request("optaddBuyprice") & vbCrLf

   	'### ���� �α� ����(item option)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'itemoption', '" & itemid & "', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

    Call ReCalcuItemOption(itemid)

    response.write "<script language='javascript'>alert('���� �Ǿ����ϴ�.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if


'''�ɼ� ����. - ���Ͽɼ�
if (mode = "deleteoption") then
	''���� ������ �ɼ����� üũ

    response.write "<script language='javascript'>alert('�ɼ��� ������ �� �����ϴ�. ������ ���� �����ϼ���.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End

	if (Not IsUpchebeasong) then
    	''�ֱ� �Ǹų���
    	sqlStr = "select top 1 * from "
    	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    	sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"

    	rsget.Open sqlStr, dbget, 1
    	if Not rsget.Eof then
    		ErrStr = "�����Ϸ��� �ɼ����� �Ǹŵ� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    	end if
    	rsget.close

    	''6���� ���� �Ǹų���
    	if ErrStr="" then
    		sqlStr = "select top 1 * from "
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    		sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    		sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"

    		rsget.Open sqlStr, dbget, 1
    		if Not rsget.Eof then
    			ErrStr = "�����Ϸ��� �ɼ����� �Ǹŵ� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    		end if
    		rsget.close
    	end if

    	''�������
    	if ErrStr="" then
    		sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
    		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
    		sqlStr = sqlStr + " where m.code=d.mastercode"
    		sqlStr = sqlStr + " and m.deldt is NULL"
    		sqlStr = sqlStr + " and d.iitemgubun='10'"
    		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
    		sqlStr = sqlStr + " and d.itemoption='" + Trim(itemoption) + "'"
            sqlStr = sqlStr + " and d.deldt is NULL"

    		rsget.Open sqlStr, dbget, 1
    		if Not rsget.Eof then
    			ErrStr = "�����Ϸ��� �ɼ����� ����� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    		end if
    		rsget.close
    	end if
	end if

    ''��Ǫ��ǰ��� ����(��ü+�ٹ� ����) //2016.05.19 ������ �߰�(�̹��� �̻�� ��û)
    sRetValue = ""
    sqlStr = " if exists(select shopitemid from db_shop.dbo.[tbl_shop_item] where itemgubun='10' and shopitemid = "& CStr(itemid) & " and itemoption ='" + Trim(itemoption) + "' and isusing ='Y' and onofflinkyn ='Y' )  select 'Y'   "
    sqlStr = sqlStr & " else select 'N' "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    	if Not rsget.Eof then
    	    sRetValue = rsget(0)
    	    if sRetValue ="Y" then
    		''ErrStr = "�����Ϸ��� �ɼ����� ��Ǫ��ǰ��� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"  ''2016/10/04 �ּ�ó��.
    	    end if
    	end if
    rsget.close

    ''2016/10/04 �߰�.
    if (sRetValue="Y") then
        ''��� ������ �ִ°��� �Դ� üũ
        sRetValue = ""
        sqlStr = " if exists(select itemid from db_summary.dbo.tbl_current_shopstock_summary where itemgubun='10' and itemid = "& CStr(itemid) &" and itemoption ='" + Trim(itemoption) + "' "
        sqlStr = sqlStr & " and shopid in ('streetshop011','streetshop018','streetshop103','streetshop809','streetshop810') and (sysstockno<>0 or realstockno<>0))  select 'Y'   "
        sqlStr = sqlStr & " else select 'N' "
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        	if Not rsget.Eof then
        	    sRetValue = rsget(0)
        	    if sRetValue ="Y" then
        		ErrStr = "�����Ϸ��� �ɼ����� ��Ǫ��ǰ��� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
        	    end if
        	end if
        rsget.close
    end if

	if (ErrStr<>"") then
		response.write "<script language='javascript'>alert('" + ErrStr + "'); </script>"
		response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End
	else
		sqlStr = "delete from [db_item].[dbo].tbl_item_option" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + CStr(Trim(itemoption)) + "'" + VbCrlf
		'rsget.Open sqlStr, dbget, 1

		'Call ReCalcuItemOption(itemid)

		vChangeContents = vChangeContents & "- �ɼ� ���� - ���Ͽɼ�" & vbCrLf
	    vChangeContents = vChangeContents & "- �ɼ��ڵ� : itemoption = " & request("itemoption") & vbCrLf

	   	'### ���� �α� ����(item option)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'itemoption', '" & itemid & "', '" & Request("menupos") & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		'dbget.execute(vSCMChangeSQL)

		'response.write "<script language='javascript'>alert('�����Ǿ����ϴ�.'); </script>"
		response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End
	end if
end if

'' �ɼǻ��� - ���߿ɼ�
if (mode = "deleteMultipleOption") then
    'TypeSeq
    'KindSeq
    response.write "<script language='javascript'>alert('�ɼ��� ������ �� �����ϴ�. ������ ���� �����ϼ���.'); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End

    if (Not IsUpchebeasong) then
        sqlStr = "select top 1 * from "
    	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
    	sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    	if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=2) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=3) then
    	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
    	end if

    	rsget.Open sqlStr, dbget, 1
    	if Not rsget.Eof then
    		ErrStr = "�����Ϸ��� �ɼ����� �Ǹŵ� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    	end if
    	rsget.close

    	''6���� ���� �Ǹų���
    	if ErrStr="" then
    		sqlStr = "select top 1 * from "
    		sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d"
    		sqlStr = sqlStr + " where d.itemid=" + CStr(itemid)
    		if (TypeSeq=1) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=2) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=3) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
        	end if

    		rsget.Open sqlStr, dbget, 1
    		if Not rsget.Eof then
    			ErrStr = "�����Ϸ��� �ɼ����� �Ǹŵ� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    		end if
    		rsget.close
    	end if

    	''�������
    	if ErrStr="" then
    		sqlStr = "select top 1 * from [db_storage].[dbo].tbl_acount_storage_detail d,"
    		sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m"
    		sqlStr = sqlStr + " where m.code=d.mastercode"
    		sqlStr = sqlStr + " and m.deldt is NULL"
    		sqlStr = sqlStr + " and d.iitemgubun='10'"
    		sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
            sqlStr = sqlStr + " and d.deldt is NULL"
            if (TypeSeq=1) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,2)='Z" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=2) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and LEFT(RIGHT(d.itemoption,3),1)='" + CStr(KindSeq) + "'"
        	elseif (TypeSeq=3) then
        	    sqlStr = sqlStr + " and LEFT(d.itemoption,1)='Z'"
        	    sqlStr = sqlStr + " and RIGHT(d.itemoption,1)='" + CStr(KindSeq) + "'"
        	end if

    		rsget.Open sqlStr, dbget, 1
    		if Not rsget.Eof then
    			ErrStr = "�����Ϸ��� �ɼ����� ����� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    		end if
    		rsget.close
    	end if
    end if

 ''��Ǫ��ǰ��� ����(��ü+�ٹ� ����) //2016.05.19 ������ �߰�(�̹��� �̻�� ��û)
    sRetValue =""
    sqlStr = " if exists(select shopitemid from db_shop.dbo.[tbl_shop_item] where itemgubun='10' and shopitemid = "& CStr(itemid)
        if (TypeSeq=1) then
            sqlStr = sqlStr & " and LEFT(itemoption,2)='Z" & CStr(KindSeq) & "'"
        elseif (TypeSeq=2) then
        	sqlStr = sqlStr & " and LEFT(itemoption,1)='Z'"
        	sqlStr = sqlStr & " and LEFT(RIGHT(itemoption,3),1)='" & CStr(KindSeq) & "'"
        elseif (TypeSeq=3) then
        	    sqlStr = sqlStr & " and LEFT(itemoption,1)='Z'"
        	    sqlStr = sqlStr & " and RIGHT(itemoption,1)='" & CStr(KindSeq) & "'"
       end if
    sqlStr = sqlStr & " and isusing ='Y' and onofflinkyn ='Y' )  select 'Y'   "
    sqlStr = sqlStr & " else select 'N' "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    	if Not rsget.Eof then
    	    sRetValue = rsget(0)
    	    if sRetValue ="Y" then
    		''ErrStr = "�����Ϸ��� �ɼ����� ��Ǫ��ǰ��� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
    	    end if
    	end if
    rsget.close

    ''2016/10/04 �߰�.
    if (sRetValue="Y") then
        ''��� ������ �ִ°��� �Դ� üũ
        sRetValue = ""
        sqlStr = " if exists(select itemid from db_summary.dbo.tbl_current_shopstock_summary where itemgubun='10' and itemid = "& CStr(itemid)
            if (TypeSeq=1) then
                sqlStr = sqlStr & " and LEFT(itemoption,2)='Z" & CStr(KindSeq) & "'"
            elseif (TypeSeq=2) then
            	sqlStr = sqlStr & " and LEFT(itemoption,1)='Z'"
            	sqlStr = sqlStr & " and LEFT(RIGHT(itemoption,3),1)='" & CStr(KindSeq) & "'"
            elseif (TypeSeq=3) then
            	    sqlStr = sqlStr & " and LEFT(itemoption,1)='Z'"
            	    sqlStr = sqlStr & " and RIGHT(itemoption,1)='" & CStr(KindSeq) & "'"
           end if
        sqlStr = sqlStr & " and shopid in ('streetshop011','streetshop018','streetshop103','streetshop809','streetshop810')"
        sqlStr = sqlStr & " and (sysstockno<>0 or realstockno<>0))  select 'Y'   "
        sqlStr = sqlStr & " else select 'N' "

        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        	if Not rsget.Eof then
        	    sRetValue = rsget(0)
        	    if sRetValue ="Y" then
        		ErrStr = "�����Ϸ��� �ɼ����� ��Ǫ��ǰ��� ������ �ֽ��ϴ�. �����Ͻ� �� �����ϴ�. - ������ ���ǿ��"
        	    end if
        	end if
        rsget.close
    end if

    if (ErrStr<>"") then
		response.write "<script language='javascript'>alert('" + ErrStr + "'); history.back();</script>"
		dbget.close()	:	response.End
	else
	    sqlStr = "delete from [db_item].[dbo].tbl_item_option_Multiple" + VbCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	    sqlStr = sqlStr + " and TypeSeq=" + CStr(TypeSeq)
	    sqlStr = sqlStr + " and KindSeq='" + CStr(KindSeq) + "'"

	    'dbget.Execute sqlStr

		sqlStr = "delete from [db_item].[dbo].tbl_item_option" + VbCrlf
		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
		if (TypeSeq=1) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,2)='Z" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=2) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and LEFT(RIGHT(itemoption,3),1)='" + CStr(KindSeq) + "'"
    	elseif (TypeSeq=3) then
    	    sqlStr = sqlStr + " and LEFT(itemoption,1)='Z'"
    	    sqlStr = sqlStr + " and RIGHT(itemoption,1)='" + CStr(KindSeq) + "'"
    	else
    	    sqlStr = sqlStr + " and 1=0"
    	end if

    	'dbget.Execute sqlStr

    	'' 3���ɼ�->2���� ���� or 2���ɼ� ->1���� ���� ��..
    	'Call ReMatchMultiOption(itemid)

		'Call ReCalcuItemOption(itemid)

		vChangeContents = vChangeContents & "- �ɼǻ��� - ���߿ɼ�" & vbCrLf
	    vChangeContents = vChangeContents & "- TypeSeq = " & CStr(TypeSeq) & vbCrLf
	    vChangeContents = vChangeContents & "- KindSeq = " & CStr(KindSeq) & vbCrLf

	   	'### ���� �α� ����(item option)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'itemoption', '" & itemid & "', '" & Request("menupos") & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		'dbget.execute(vSCMChangeSQL)

		'response.write "<script language='javascript'>alert('�����Ǿ����ϴ�.'); </script>"
		response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
		dbget.close()	:	response.End
	end if
end if


'' ���� �ɼ� �߰�
if (mode = "addoptionCustom") then
    foundcount = 0

    for i = 0 to ubound(arritemoption)
        if (Trim(arritemoption(i)) <> "") then
            sqlStr = " select itemid from [db_item].[dbo].tbl_item_option "
            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and ((itemoption = '" + CStr(requestCheckVar(Trim(arritemoption(i)),4)) + "') or (optionname='" + requestCheckVar(html2db(arritemoptionname(i)),96) + "'))"

            rsget.Open sqlStr,dbget,1

            if not rsget.EOF then
                found = true
                foundcount = foundcount + 1
            else
                found = false
            end if
            rsget.close

            ''���� ������ ��ǰ ���� ���а� ����
            if (found = false) then
                sqlStr = " insert into [db_item].[dbo].tbl_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(requestCheckVar(arritemoption(i),4)) + "', '" + html2db(optionTypename) + "', '" + CStr(requestCheckVar(html2db(arritemoptionname(i)),96)) + "', 'Y', 'Y', '" + itemLimitYn + "', 0, 0) "

                dbget.Execute sqlStr

            end if
        end if
    next

    ''�ɼ� ���и��� ����

    sqlStr = " update [db_item].[dbo].tbl_item_option " &VbCRLF
    sqlStr = sqlStr + " set optionTypeName='" + html2db(optionTypename) + "'" &VbCRLF
    sqlStr = sqlStr + " where itemid=" + cStr(itemid) &VbCRLF
    sqlStr = sqlStr + " and optionTypeName<>'" + html2db(optionTypename) + "'" &VbCRLF

    dbget.Execute sqlStr

    Call ReCalcuItemOption(itemid)

    vChangeContents = vChangeContents & "- �ɼ� �߰� - ���Ͽɼ�" & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ��ڵ� : itemoption = " & request("arritemoption") & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǱ��и� : optionTypename = " & html2db(optionTypename) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǻ󼼸� : optionName = " & html2Db(request("arritemoptionname")) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ����� : optlimityn = " & itemLimitYn & vbCrLf

   	'### ���� �α� ����(item option)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'itemoption', '" & itemid & "', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

    if (foundcount > 0) then
        response.write "<script>alert('�Ϻ� �ɼ��� ������ �ִ� �ɼǰ� �ߺ��Ǿ� ���õǾ����ϴ�.');</script>"
    end if

    response.write "<script>alert('�ɼ��� �߰��Ǿ����ϴ�.'); opener.location.reload(); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if




''���߿ɼ� �߰�
if (mode="addDoubleOption") then

    dim optionTypename1, optionTypename2, optionTypename3
    dim itemoption1, itemoption2, itemoption3
    dim optionName1, optionName2, optionName3

    optionTypename1 = requestCheckVar(Trim(request("optionTypename1")),32)
    optionTypename2 = requestCheckVar(Trim(request("optionTypename2")),32)
    optionTypename3 = requestCheckVar(Trim(request("optionTypename3")),32)
    itemoption1     = requestCheckVar(request("itemoption1"),4)
    itemoption2     = requestCheckVar(request("itemoption2"),4)
    itemoption3     = requestCheckVar(request("itemoption3"),4)
    optionName1     = requestCheckVar(request("optionName1"),96)
    optionName2     = requestCheckVar(request("optionName2"),96)
    optionName3     = requestCheckVar(request("optionName3"),96)

    dim Lv1cnt, Lv2cnt, Lv3cnt
    dim Val1cnt, Val2cnt, Val3cnt
    dim option1, option2, option3
    dim optName1, optName2, optName3
    dim Valid1, Valid2, Valid3
    dim buf, ErrMsg, AssignedOption

    Lv1cnt = request("itemoption1").count
    Lv2cnt = request("itemoption2").count
    Lv3cnt = request("itemoption3").count

    Val1cnt = 0
    Val2cnt = 0
    Val3cnt = 0

    '' üũ �����߰� //2016/04/19
    if ((Lv1cnt>35) or (Lv2cnt>35) or (Lv3cnt>35)) then
        ErrMsg = ErrMsg & "���߿ɼ��� �ɼǱ��д� �ִ� 35������ �����մϴ�.\n"
    end if

    for i=1 to Lv1cnt
        buf = requestCheckVar(Trim(request("optionName1")(i)),96)
        if Len(buf)>0 then Val1cnt = Val1cnt + 1
    next

    for i=1 to Lv2cnt
        buf = requestCheckVar(Trim(request("optionName2")(i)),96)
        if Len(buf)>0 then Val2cnt = Val2cnt + 1
    next

    for i=1 to Lv3cnt
        buf = requestCheckVar(Trim(request("optionName3")(i)),96)
        if Len(buf)>0 then Val3cnt = Val3cnt + 1
    next

    if (optionTypename1=optionTypename2) or (optionTypename1=optionTypename3) or (optionTypename2=optionTypename3) then
        ErrMsg = "�ɼǱ��и��� ������ �� �����ϴ�.\n"
    end if

    if (Len(optionTypename1)<1) and (Len(optionTypename2)<1) and (Len(optionTypename3)<1) then
        ErrMsg = "�ɼǱ��и��� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if (Val1cnt>0) and (Len(optionTypename1)<1) then
        ErrMsg = ErrMsg & "�ɼǱ��и�1�� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if (Val2cnt>0) and (Len(optionTypename2)<1) then
        ErrMsg = ErrMsg & "�ɼǱ��и�2�� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if (Val3cnt>0) and (Len(optionTypename3)<1) then
        ErrMsg = ErrMsg & "�ɼǱ��и�3�� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if (Val1cnt<1) and (Len(optionTypename1)>0) then
        ErrMsg = ErrMsg & "�ɼǱ��и�1�� ���� �ɼ��� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if (Val2cnt<1) and (Len(optionTypename2)>0) then
        ErrMsg = ErrMsg & "�ɼǱ��и�2�� ���� �ɼ��� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if (Val3cnt<1) and (Len(optionTypename3)>0) then
        ErrMsg = ErrMsg & "�ɼǱ��и�3�� ���� �ɼ��� �Էµ��� �ʾҽ��ϴ�.\n"
    end if

    if ((Val1cnt<1) and (Val2cnt<1)) or ((Val2cnt<1) and (Val3cnt<1)) or ((Val1cnt<1) and (Val3cnt<1)) then
        ErrMsg = ErrMsg & "���߿ɼ����� ��� �Ͻ÷��� �ɼǱ����� 2�� �̻� ����ϼž� �մϴ�.\n"
    end if

    ''������� �Է��ؾ� ����
    if ((Val1cnt<1) or (Val2cnt<1)) then
        ErrMsg = ErrMsg & "���߿ɼ����� ��� �Ͻ÷��� �ɼǱ��� 1 ���� ������ 2�� �̻� ����ϼž� �մϴ�.\n"
    end if

    if (Len(ErrMsg)>0) then
        response.write "<script>alert('" + ErrMsg + "'); history.back();</script>"
        dbget.close()	:	response.End
    end if

    foundcount=0


    ''0���� �Է� ����. N����
    for i=0 to Lv1cnt-1
        for j=0 to Lv2cnt-1
            for k=0 to Lv3cnt-1
                optName1 = requestCheckVar(Trim(request("optionName1")(i+1)),96)
                optName2 = requestCheckVar(Trim(request("optionName2")(j+1)),96)
                optName3 = requestCheckVar(Trim(request("optionName3")(k+1)),96)

                option1  = requestCheckVar(Trim(request("itemoption1")(i+1)),4)
                option2  = requestCheckVar(Trim(request("itemoption2")(j+1)),4)
                option3  = requestCheckVar(Trim(request("itemoption3")(k+1)),4)

                Valid1  = (Len(optionTypename1)>0) and (Len(optName1)>0)
                Valid2  = (Len(optionTypename2)>0) and (Len(optName2)>0)
                Valid3  = (Len(optionTypename3)>0) and (Len(optName3)>0)

                AssignedOption = "Z"  '''�����ؾ���.. =>9
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
                ''�޸�����
                optionName = Replace(optionName,",,",",")
                if Right(optionName,1)="," then optionName=Left(optionName,Len(optionName)-1)

                if (Valid1 and Valid2) or (Valid1 and Valid3) or (Valid2 and Valid3) then
                    if ((i=0) or (Valid1)) and  ((j=0) or (Valid2)) and ((k=0) or (Valid3))  then
                        ''���� �ɼ��� �����ϴ��� Check.

                        found = false

                        if (Len(option1)<1) and (Len(optName1)>0) and (Len(optionTypename1)>0) then
                            sqlStr = " select itemid from [db_item].[dbo].tbl_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename1) + "' and optionKindName='" + html2db(optName1) + "'))"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsget.Close

                            if (Not found) then
                                sqlStr = " insert into [db_item].[dbo].tbl_item_option_Multiple" &VbCRLF
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)" &VbCRLF
                                sqlStr = sqlStr + " values(" &VbCRLF
                                sqlStr = sqlStr + " " & itemid &VbCRLF
                                sqlStr = sqlStr + " ,1" &VbCRLF
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(i+1)) &"'" &VbCRLF
                                sqlStr = sqlStr + " ,'" & html2db(optionTypename1) & "'" &VbCRLF
                                sqlStr = sqlStr + " ,'" & html2db(optName1) & "'" &VbCRLF
                                sqlStr = sqlStr + " )"

                                dbget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName1 + "," + optionTypename1 + "," + CStr(1) + "," + CStr(i+1) + "<br>"
                            end if
                        end if

                        found = false
                        if (Len(option2)<1) and (Len(optName2)>0) and (Len(optionTypename2)>0) then
                            sqlStr = " select itemid from [db_item].[dbo].tbl_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename2) + "' and optionKindName='" + html2db(optName2) + "'))"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsget.Close

                            if (Not found) then
                                found = false
                                sqlStr = " select itemid from [db_item].[dbo].tbl_item_option_Multiple "
                                sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                                sqlStr = sqlStr + " and TypeSeq = 2 and KindSeq = " & optKindSeq2Code(CStr(j+1))

                                rsget.Open sqlStr,dbget,1
                                if Not rsget.Eof then
                                    found = true
                                end if
                                rsget.Close

                                if found then
                                    ErrMsg = ErrMsg & "�߰��Ϸ��� �ɼ��� ������ ������ ���޵Ǿ�� �մϴ�.\n"
                                    response.write "<script>alert('" + ErrMsg + "'); history.back();</script>"
                                    dbget.close()	:	response.End
                                end if

                                sqlStr = " insert into [db_item].[dbo].tbl_item_option_Multiple" &VbCRLF
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)" &VbCRLF
                                sqlStr = sqlStr + " values(" &VbCRLF
                                sqlStr = sqlStr + " " & itemid &VbCRLF
                                sqlStr = sqlStr + " ,2" &VbCRLF
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(j+1)) & "'" &VbCRLF
                                sqlStr = sqlStr + " ,'" & html2db(optionTypename2) & "'" &VbCRLF
                                sqlStr = sqlStr + " ,'" & html2db(optName2) & "'" &VbCRLF
                                sqlStr = sqlStr + " )"

                                dbget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName2 + "," + optionTypename2 + "," + CStr(2) + "," + CStr(j+1) + "<br>"
                            end if
                        end if

                        found = false
                        if (Len(option3)<1) and (Len(optName3)>0) and (Len(optionTypename3)>0) then
                            sqlStr = " select itemid from [db_item].[dbo].tbl_item_option_Multiple "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                            sqlStr = sqlStr + " and ((optionTypeName='" + html2db(optionTypename3) + "' and optionKindName='" + html2db(optName3) + "'))"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                found = true
                                foundcount = foundcount + 1
                            end if
                            rsget.Close

                            if (Not found) then
                                sqlStr = " insert into [db_item].[dbo].tbl_item_option_Multiple" &VbCRLF
                                sqlStr = sqlStr + " (itemid, TypeSeq, KindSeq, optionTypeName, optionKindName)" &VbCRLF
                                sqlStr = sqlStr + " values(" &VbCRLF
                                sqlStr = sqlStr + " " & itemid &VbCRLF
                                sqlStr = sqlStr + " ,3" &VbCRLF
                                sqlStr = sqlStr + " ,'" & optKindSeq2Code(CStr(k+1)) & "'" &VbCRLF
                                sqlStr = sqlStr + " ,'" & html2db(optionTypename3) & "'" &VbCRLF
                                sqlStr = sqlStr + " ,'" & html2db(optName3) & "'" &VbCRLF
                                sqlStr = sqlStr + " )"

                                dbget.Execute sqlStr
                                ''response.write AssignedOption + "," + optName3 + "," + optionTypename3 + "," + CStr(3) + "," + CStr(k+1) + "<br>"
                            end if
                        end if

                        found = false
                        sqlStr = " select itemid from [db_item].[dbo].tbl_item_option "
                        sqlStr = sqlStr + " where itemid = " + CStr(itemid)
                        sqlStr = sqlStr + " and ((itemoption = '" + CStr(AssignedOption) + "') or (optionTypeName='���տɼ�' and optionname='" + html2db(optionName) + "'))"

                        rsget.Open sqlStr,dbget,1
                        if Not rsget.Eof then
                            found = true
                        end if
                        rsget.Close

                        if (Not found) then
                            sqlStr = " insert into [db_item].[dbo].tbl_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold) "
                            sqlStr = sqlStr + " values(" + CStr(itemid) + ", '" + CStr(AssignedOption) + "', '���տɼ�', '" + CStr(html2db(optionName)) + "', 'Y', 'Y', '" + itemLimitYn + "', 0, 0) "

                            dbget.Execute sqlStr
                            ''response.write AssignedOption + ":" +  optName1 + "," + optName2 + "," + optName3 + "<BR>"
                        end if
                    end if

                end if
            next
        next
    next

    '' 2���ɼ�->3���� ����  ��..
    Call ReMatchMultiOption(itemid)


    Call ReCalcuItemOption(itemid)

    vChangeContents = vChangeContents & "- �ɼ� �߰� - ���߿ɼ�" & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ��ڵ�1 : itemoption1 = " & request("itemoption1") & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ��ڵ�2 : itemoption2 = " & request("itemoption2") & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ��ڵ�3 : itemoption3 = " & request("itemoption3") & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǱ��и�1 : optionTypename1 = " & html2Db(Trim(request("optionTypename1"))) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǱ��и�2 : optionTypename2 = " & html2Db(Trim(request("optionTypename2"))) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǱ��и�3 : optionTypename3 = " & html2Db(Trim(request("optionTypename3"))) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǻ󼼸�1 : optionName1 = " & html2Db(request("optionName1")) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǻ󼼸�2 : optionName2 = " & html2Db(request("optionName2")) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼǻ󼼸�3 : optionName3 = " & html2Db(request("optionName3")) & vbCrLf
    vChangeContents = vChangeContents & "- �ɼ����� : optlimityn = " & itemLimitYn & vbCrLf

   	'### ���� �α� ����(item option)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'itemoption', '" & itemid & "', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

    response.write "<script>alert('�ɼ��� �߰��Ǿ����ϴ�.'); opener.location.reload(); </script>"
    response.write "<script language='javascript'>location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

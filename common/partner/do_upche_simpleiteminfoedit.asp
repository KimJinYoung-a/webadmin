<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ����
' History : ������ ����
'			2017.04.14 �ѿ�� ����(���Ȱ���ó��)

'####################################################
%>
<!-- #include virtual="/partner/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim itemid, limityn, limitno, limitsold
dim sellyn, isusing, reqstring
dim pojangok
dim itemoptionarr, optlimitnoarr, optlimitsoldarr, optisusingarr
dim index

itemid      = getNumeric(requestCheckvar(request("itemid"),10))
limityn     = requestCheckvar(request("limityn"),1)
limitno     = getNumeric(requestCheckvar(request("limitno"),10))
limitsold   = getNumeric(requestCheckvar(request("limitsold"),10))
sellyn      = requestCheckvar(request("sellyn"),1)
isusing     = requestCheckvar(request("isusing"),1)
reqstring   = requestCheckvar(request("reqstring"),128)


itemoptionarr 	= request("itemoptionarr")
optlimitnoarr	= request("optlimitnoarr")
optlimitsoldarr = request("optlimitsoldarr")
optisusingarr	= request("optisusingarr")

'response.write "itemid=" & itemid & "<br>"
'response.write "limityn=" & limityn & "<br>"
'response.write "limitno=" & limitno & "<br>"
'response.write "limitsold=" & limitsold & "<br>"
'response.write "sellyn=" & sellyn & "<br>"
'response.write "isusing=" & isusing & "<br>"
'response.write "reqstring=" & reqstring & "<br>"
'
'response.write "itemoptionarr=" & itemoptionarr & "<br>"
'response.write "optlimitnoarr=" & optlimitnoarr & "<br>"
'response.write "optlimitsoldarr=" & optlimitsoldarr & "<br>"
'response.write "optisusingarr=" & optisusingarr & "<br>"

if itemid = "" then
	response.write "<script type='text/javascript'>alert('��ǰ�ڵ尡 �����ϴ�'); history.back();</script>"
	dbget.close()	:	response.End
end if

''2014/03/18 �߰�
if (limityn <> "Y") and (limityn <> "N") then
	response.write "<script type='text/javascript'>alert('�Ķ���� ���� limityn'); history.back();</script>"
	dbget.close()	:	response.End
end if

if (sellyn <> "Y") and (sellyn <> "N") and (sellyn <> "S") then
	response.write "<script type='text/javascript'>alert('�Ķ���� ���� sellyn'); history.back();</script>"
	dbget.close()	:	response.End
end if

if (isusing <> "Y") and (isusing <> "N") then
	response.write "<script type='text/javascript'>alert('�Ķ���� ���� isusing'); history.back();</script>"
	dbget.close()	:	response.End
end if
''2014/03/18

'rw itemoptionarr
'rw optlimitnoarr
'rw optlimitsoldarr
'rw optisusingarr

itemoptionarr 	= split(itemoptionarr,",")
optlimitnoarr	= split(optlimitnoarr,",")
optlimitsoldarr = split(optlimitsoldarr,",")
optisusingarr 	= split(optisusingarr,",")


dim OptionExists
OptionExists = (UBound(itemoptionarr)>1)
dim BufOptionKey

'response.write "OptionExists=" & OptionExists
'dbget.close()	:	response.End
'==============================================================================
'���� ��ǰ ����
dim mwdiv, orgsellyn, orglimityn, orglimitno, orglimitsold, orgoptioncnt,orgsellSTDate
dim orgisusing
    
sqlStr = "select top 1 sellyn, limityn, limitno, limitsold, mwdiv, optioncnt, sellSTDate, isusing " + VbCrlf
sqlStr = sqlStr + " from [db_item].[dbo].tbl_item" + VbCrlf
sqlStr = sqlStr + " where 1 = 1 "
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + " " + VbCrlf
sqlStr = sqlStr + " and makerid = '" + CStr(session("ssBctID")) + "' "
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

if  not rsget.EOF  then
    orgsellyn = rsget("sellyn")
    orglimityn = rsget("limityn")
    orglimitno = rsget("limitno")
    orglimitsold = rsget("limitsold")
    mwdiv = rsget("mwdiv")
    orgoptioncnt = rsget("optioncnt") 
    orgsellSTDate   = rsget("sellSTDate") 
    orgisusing = rsget("isusing")
else
    rsget.Close

    response.write "<script type='text/javascript'>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if
rsget.Close

''2017/06/02 �߰�
if (orgisusing="N") then
    response.write "<script type='text/javascript'>alert('�Ǹ� ������ ��ǰ�Դϴ�. �����Ұ�'); history.back();</script>"
    dbget.close()	:	response.End
end if

'==============================================================================
'���� �ɼ� ����
dim orgarritemoption, orgarritemoptionname, orgarrisusing, orgarroptlimityn, orgarroptlimitno, orgarroptlimitsold
''orgarroptsellyn - ������(optisusing �� ��ġ)

sqlStr = " select top 100 o.itemoption, isnull(o.optionname,'') as itemoptionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold "
sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option o "
sqlStr = sqlStr + " where 1 = 1 "
sqlStr = sqlStr + " and itemid=" + CStr(itemid) + " "
sqlStr = sqlStr + " and itemoption <> '0000' "
sqlStr = sqlStr + " order by itemoption "
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

if  not rsget.EOF  then
    do until rsget.Eof
        orgarritemoption        = orgarritemoption + rsget("itemoption") + "|"
        orgarritemoptionname    = orgarritemoptionname + db2html(rsget("itemoptionname")) + "|"
        orgarrisusing           = orgarrisusing + rsget("isusing") + "|"
        orgarroptlimityn        = orgarroptlimityn + rsget("optlimityn") + "|"
        orgarroptlimitno        = orgarroptlimitno + CStr(rsget("optlimitno")) + "|"
        orgarroptlimitsold      = orgarroptlimitsold + CStr(rsget("optlimitsold")) + "|"
        ''orgarroptsellyn         = orgarroptsellyn + rsget("optsellyn") + "|"
		rsget.movenext
	loop
end if
rsget.Close




dim refer, AssignedCnt, iAssignedRow
refer = request.ServerVariables("HTTP_REFERER")
AssignedCnt = 0

dim sqlStr, i, j 
 

''��ü�ΰ�� ���� ����.
if (mwdiv = "U") then
    if (limityn="Y") then
    	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
    	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
    	sqlStr = sqlStr + " , lastupdate=getdate()"+ VBCrlf      
    	    if orgSellyn <>"Y" and sellyn ="Y" and isNull(orgsellSTDate) then
	    sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf        
	        end if
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	dbget.Execute sqlStr

    	''�ɼ�������������
    	sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	dbget.Execute sqlStr

    	for i=0 to UBound(itemoptionarr)
    		if (Len(requestCheckVar(Trim(itemoptionarr(i)),4))=4) then
    			if (itemoptionarr(i)="0000") then
    				sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    				sqlStr = sqlStr + " set limitno=" + requestCheckVar(optlimitnoarr(i),10) + "" + VBCrlf
    				sqlStr = sqlStr + " , limitsold=" + requestCheckVar(optlimitsoldarr(i),10) + "" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    				dbget.Execute sqlStr
    			else
    				sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    				sqlStr = sqlStr + " set isusing='" + requestCheckVar(optisusingarr(i),1) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optsellyn='" + requestCheckVar(optisusingarr(i),1) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optlimitno=" + requestCheckVar(optlimitnoarr(i),10) + "" + VBCrlf
    				sqlStr = sqlStr + " , optlimitsold=" + requestCheckVar(optlimitsoldarr(i),10) + "" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    				sqlStr = sqlStr + " and itemoption='" + requestCheckVar(Trim(itemoptionarr(i)),4) + "'"

    				dbget.Execute sqlStr
    			end if
    		end if
    	next
    else
    	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
    	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
    	sqlStr = sqlStr + " , lastupdate=getdate()"+ VBCrlf      
    	if orgSellyn <>"Y" and sellyn ="Y" and isNull(orgsellSTDate) then
	    sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf        
	    end if
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	dbget.Execute sqlStr


    	''�ɼ�������������
    	sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    	dbget.Execute sqlStr

    	for i=0 to UBound(itemoptionarr)
    		if (Len(Trim(itemoptionarr(i)))=4) then
    			if (itemoptionarr(i)="0000") then
    			    '' �������Ǹ��ΰ�� �ʿ� ����.
    				'sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    				'sqlStr = sqlStr + " set limitno=" + requestCheckVar(optlimitnoarr(i),10) + "" + VBCrlf
    				'sqlStr = sqlStr + " , limitsold=" + requestCheckVar(optlimitsoldarr(i),10) + "" + VBCrlf
    				'sqlStr = sqlStr + " where itemid=" + CStr(itemid)

    				'dbget.Execute sqlStr
    			else
    			    ' �������Ǹ��ΰ�� �ʿ� ����. �ּ�ó��
    			    ''sqlStr = sqlStr + " , optlimitno=" + requestCheckVar(optlimitnoarr(i),10) + "" + VBCrlf
    				''sqlStr = sqlStr + " , optlimitsold=" + requestCheckVar(optlimitsoldarr(i),10) + "" + VBCrlf
    				
    				sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
    				sqlStr = sqlStr + " set isusing='" + requestCheckVar(optisusingarr(i),1) + "'" + VBCrlf
    				sqlStr = sqlStr + " , optsellyn='" + requestCheckVar(optisusingarr(i),1) + "'" + VBCrlf
    				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
    				sqlStr = sqlStr + " and itemoption='" + requestCheckVar(Trim(itemoptionarr(i)),4) + "'"

    				dbget.Execute sqlStr
    			end if
    		end if
    	next
    end if


    ''��ǰ�ɼǼ���
    sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
    sqlStr = sqlStr + " from (" + VBCrlf
    sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
    sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
    sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
    sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
    sqlStr = sqlStr + " ) T" + VBCrlf
    sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf

    dbget.Execute sqlStr

'�ӽ� �ּ�ó�� 2014-08-18 ������
'    sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
'    sqlStr = sqlStr + " set optionname=v.codeview" + VBCrlf
'    sqlStr = sqlStr + " from [db_item].[dbo].vw_all_option v" + VBCrlf
'    sqlStr = sqlStr + " where  [db_item].[dbo].tbl_item_option.itemid=" + CStr(itemid) + VBCrlf
'    sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=v.optioncode" + VBCrlf
'
'    rsget.Open sqlStr, dbget, 1


    if (orgoptioncnt > 0) then
    	''��ǰ��������
    	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
    	sqlStr = sqlStr + " ,sellyn=(CASE WHEN sellyn='Y' and IsNULL(T.usingoptionCnt,0)=0 THEN 'S' ELSE sellyn END)" + VBCrlf ''�ɼ� ���� ���������� ���� �ϴ� ��� //2013/09/02 �߰�
    	sqlStr = sqlStr + " from (" + VBCrlf
    	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold, count(*) as usingoptionCnt" + VBCrlf
    	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
    	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
    	sqlStr = sqlStr + " 	and isusing='Y' " + VBCrlf
    	sqlStr = sqlStr + " ) T" + VBCrlf
    	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf
    	dbget.Execute sqlStr


        sqlStr = " update [db_item].[dbo].tbl_item_option "
        sqlStr = sqlStr + " set optlimityn = [db_item].[dbo].tbl_item.limityn "   ''optsellyn = T.sellyn,
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid = " + CStr(itemid) + " "
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemid = [db_item].[dbo].tbl_item.itemid"
        dbget.Execute sqlStr
    end if

    ''������ ���̸� �Ͻ� ǰ��ó�� // 2013/09/02 �߰�
    sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
    sqlStr = sqlStr + " set sellyn='S'" + VBCrlf
    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + VBCrlf
    sqlStr = sqlStr + " and sellyn='Y'" + VBCrlf
    sqlStr = sqlStr + " and limityn='Y'" + VBCrlf
    sqlStr = sqlStr + " and (limitno-limitsold)<1" + VBCrlf
    dbget.Execute sqlStr


''�ٹ��ΰ�� ������û
else
	limitno = 0
	limitsold = 0
	for i=0 to UBound(itemoptionarr)
		if (Len(requestCheckVar(Trim(itemoptionarr(i)),4))=4) then
			limitno = limitno + requestCheckVar(optlimitnoarr(i),10)
			limitsold = limitsold + requestCheckVar(optlimitsoldarr(i),10)
		end if
	next

    if (OptionExists) then
        BufOptionKey = "XXXX"
    else
        BufOptionKey = "0000"
    end if

    ' ��ǰ���� = ���� ���� �ɼ��� ���� ��� 0000 ������� XXXX XXXX �ΰ��
    if (orgsellyn<>sellyn) or (orglimityn<>limityn)  or (CStr(orglimitno)<>CStr(limitno)) or (CStr(orglimitsold)<>CStr(limitsold)) then
    	sqlStr = "insert into [db_temp].[dbo].tbl_upche_itemedit" + VbCrlf
    	sqlStr = sqlStr + " (itemid,itemoption, "
    	sqlStr = sqlStr + " oldsellyn, oldlimityn, oldlimitno, oldlimitsold, " + VbCrlf
    	sqlStr = sqlStr + " sellyn, limityn,limitno, limitsold, " + VbCrlf
    	sqlStr = sqlStr + " isupchebeasong, isfinish, etcstr,edittype)" + VbCrlf
    	sqlStr = sqlStr + " values(" + itemid + ",'" & BufOptionKey & "'," + VbCrlf
    	sqlStr = sqlStr + " '" + orgsellyn + "',"  + VbCrlf
    	sqlStr = sqlStr + " '" + orglimityn + "',"  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(orglimitno) + ","  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(orglimitsold) + ","  + VbCrlf
    	sqlStr = sqlStr + " '" + sellyn + "',"  + VbCrlf
    	sqlStr = sqlStr + " '" + limityn + "',"  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(limitno) + ","  + VbCrlf
    	sqlStr = sqlStr + " " + CStr(limitsold) + ","  + VbCrlf
    	sqlStr = sqlStr + " 'N',"  + VbCrlf
    	sqlStr = sqlStr + " 'N','" + html2db(reqstring) + "','E')"  + VbCrlf
    	'response.write sqlStr

    	dbget.Execute sqlStr, iAssignedRow

    	AssignedCnt = AssignedCnt + iAssignedRow
    end if


    if (OptionExists) then
        '�ɼ�����???
        'TODO : dispyn �ʵ带 isusing ���� ����Ѵ�.
        'TODO : �ɼǸ� �����Ҽ� �ֵ��� �����ϸ� ����.
        orgarritemoption        = Split(orgarritemoption, "|")
        orgarritemoptionname    = Split(orgarritemoptionname, "|")
        orgarrisusing           = Split(orgarrisusing, "|")
        orgarroptlimityn        = Split(orgarroptlimityn, "|")
        orgarroptlimitno        = Split(orgarroptlimitno, "|")
        orgarroptlimitsold      = Split(orgarroptlimitsold, "|")

        ''orgarroptsellyn         = Split(orgarroptsellyn, "|")

        for i = 0 to UBound(orgarritemoption) - 1
            index = -1
            iAssignedRow = 0
            if (Trim(orgarritemoption(i)) <> "") then
                for j = 0 to UBound(itemoptionarr) - 1
                    if ((Trim(orgarritemoption(i)) = Trim(itemoptionarr(j))) and (Trim(itemoptionarr(j)) <> "0000")) then
                        index = j
                        exit for
                    end if
                next

                if (index <> -1) then
                    ''���泻���� �ִ°�츸. �ɼ� ��
                    if (orgarrisusing(i)<>optisusingarr(index)) or (orglimityn<>limityn) or (CStr(orgarroptlimitno(i))<>CStr(optlimitnoarr(index))) or (CStr(orgarroptlimitsold(i))<>CStr(optlimitsoldarr(index))) then
                    	'response.write "1" & orgarrisusing(i) & "," & optisusingarr(index) & "<br>"
                    	'response.write "1" & orglimitno & "," & limityn & "<br>"
                    	'response.write "1" & orgarroptlimitno(i) & "," & optlimitnoarr(index) & "<br>"
                    	'response.write "1" & orgarroptlimitsold(i) & "," & optlimitsoldarr(index) & "<br>"


                    	sqlStr = "insert into [db_temp].[dbo].tbl_upche_itemedit" + VbCrlf
                    	sqlStr = sqlStr + " (itemid, itemoption, itemoptionname, oldsellyn, olddispyn, " + VbCrlf
                    	sqlStr = sqlStr + " oldlimityn, oldlimitno, oldlimitsold, sellyn, dispyn, limityn," + VbCrlf
                    	sqlStr = sqlStr + " limitno, limitsold, isupchebeasong, isfinish, etcstr,edittype)" + VbCrlf
                    	sqlStr = sqlStr + " values(" + itemid + ",'" + requestCheckVar(Trim(orgarritemoption(i)),4) + "','" + requestCheckVar(Trim(html2db(orgarritemoptionname(i))),96) + "'," + VbCrlf
                    	sqlStr = sqlStr + " '" + orgsellyn + "',"  + VbCrlf         ''��ǰ�Ǹſ���
                    	sqlStr = sqlStr + " '" + requestCheckVar(orgarrisusing(i),1) + "',"  + VbCrlf
                    	sqlStr = sqlStr + " '" + orglimityn + "',"  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(requestCheckVar(orgarroptlimitno(i),10)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(requestCheckVar(orgarroptlimitsold(i),10)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " '" + sellyn + "',"  + VbCrlf
                    	sqlStr = sqlStr + " '" + requestCheckVar(optisusingarr(index),1) + "',"  + VbCrlf
                    	sqlStr = sqlStr + " '" + limityn + "',"  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(requestCheckVar(optlimitnoarr(index),10)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " " + CStr(requestCheckVar(optlimitsoldarr(index),10)) + ","  + VbCrlf
                    	sqlStr = sqlStr + " 'N',"  + VbCrlf
                    	sqlStr = sqlStr + " 'N','" + html2db(reqstring) + "','E')"  + VbCrlf
 
                    	dbget.Execute sqlStr, iAssignedRow
                        AssignedCnt = AssignedCnt + iAssignedRow
                    end if
                end if
            end if
        next
    End IF


    if (AssignedCnt<1) then
        response.write "<script type='text/javascript'>alert('Err - ����� ������ �����ϴ�.\n\n �ǸŰ��� ��Ÿ ��������� ��û �Ͻ� ���\n\n ���MD���� ���� ������ �ּ���.');</script>"
        response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
    else
        response.write "<script type='text/javascript'>alert('������û�Ǿ����ϴ�.\n\n ���� ������ ������ �ٹ����ٿ��� Ȯ���� �ݿ��˴ϴ�.');</script>"
        response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
    end if
    dbget.close()	:	response.End
end if

%>
<script type='text/javascript'>
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
'// ��ȹ�� ��ǰ ����Ʈ ����
dim cnt, i
dim EXitemid , mode , isUsing , itemid , idx , MDPick , pickitem
dim strSQL,msg
dim poscode : poscode = request("poscode")
dim page : page = request("page")
dim evt_code , startdate , enddate , evtsorting , evtisusing

'// ��ǰ ���
mode = request.Form("mode")
EXitemid = request.Form("eid")
MDPick = request.Form("mdpick")

itemid= requestCheckvar(request("iid"),10)
idx = requestCheckvar(request("eidx"),10)
pickitem = requestCheckvar(request("pickitem"),1)

evt_code = requestCheckvar(request("evt_code"),10)
startdate = requestCheckvar(request("StartDate"),10)
enddate = requestCheckvar(request("EndDate"),10)
evtsorting = requestCheckvar(request("evtsorting"),10)
evtisusing = requestCheckvar(request("evtisusing"),1)

'// �׷��ڵ� 
dim gidx , gubuncode , mastercode , detailcode , catetitle , groupisusing
gidx = requestCheckvar(request("gidx"),10)
gubuncode = requestCheckvar(request("gubuncode"),1)
mastercode = requestCheckvar(request("mastercode"),10)
detailcode = requestCheckvar(request("detailcode"),10)
catetitle = requestCheckvar(request("title"),50)
groupisusing = requestCheckvar(request("isusing"),1)
	
IF mode="add" Then
    dim totcnt
    strSQL ="SELECT count(itemid) FROM db_event.dbo.tbl_exhibition_items where mastercode = '"& mastercode &"' and itemid = '"& itemid &"' "
    rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		totcnt = rsget(0)
	End IF
	rsget.close

    if totcnt > 0 then 
        Alert_return "��ǰ�ڵ� ["& itemid &"] �̹� ��� �Ǿ� �ֽ��ϴ�."
        dbget.close()	:	response.End
    end if 

	strSQL = " INSERT INTO db_event.dbo.tbl_division_items "  & vbcrlf
    strSQL = strSQL & " (mastercode , detailcode , itemid , pickitem , adminid) " & vbcrlf
    strSQL = strSQL & " values " & vbcrlf
    strSQL = strSQL & " ('"& mastercode &"' , '"& detailcode &"' , '"& itemid &"' , "& pickitem &" , '"& session("ssBctId") &"') " & vbcrlf

	dbget.execute(strSQL)

    msg = "���� �Ǿ����ϴ�"

    strSQL ="select SCOPE_IDENTITY() "

	rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		idx = rsget(0)
	End IF
	rsget.close

	Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_reg_item.asp?mode=edit&idx="& idx

ELSEIF mode="edit" Then

    strSQL = " UPDATE db_event.dbo.tbl_division_items SET "  & vbcrlf
    strSQL = strSQL & " mastercode = '"& mastercode &"' " & vbcrlf
    strSQL = strSQL & " ,detailcode = '"& detailcode &"' " & vbcrlf
    strSQL = strSQL & " ,itemid = '"& itemid &"' " & vbcrlf
    strSQL = strSQL & " ,pickitem = "& pickitem &" " & vbcrlf
    strSQL = strSQL & " ,lastupdate = getdate() " & vbcrlf
    strSQL = strSQL & " ,lastadminid = '"& session("ssBctId") &"' " & vbcrlf
    strSQL = strSQL & " where idx = "& idx &" " & vbcrlf	

	dbget.execute(strSQL)

    msg = "���� �Ǿ����ϴ�"

	Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_reg_item.asp?mode=edit&idx="& idx

Elseif mode = "pickreg" Then
    dim tmpItemIdx
    tmpItemIdx = EXitemid

	EXitemid = split(EXitemid,",")
	MDPick = split(MDPick,",")
	cnt = ubound(EXitemid)    

    if cnt > 0 then  
        For i = 0 to cnt    
            strSQL =" UPDATE db_event.dbo.tbl_exhibition_items "&_
                    " SET pickitem = '" & MDPick(i) & "' "
                    strSQL = strSQL & " WHERE idx = "& EXitemid(i)
            dbget.execute(strSQL)
        Next
    else
        strSQL =" UPDATE db_event.dbo.tbl_exhibition_items "&_
                " SET pickitem = '1' "
                strSQL = strSQL & " WHERE idx = "& tmpItemIdx
        dbget.execute(strSQL)    
    end if

       
	msg = "���� �Ǿ����ϴ�"

	Alert_move msg,"/admin/eventmanage/event/v5/popup/index.asp?menupos="&poscode&"&page="&page&"&mastercode="&mastercode

elseif mode = "gubunAdd" then '// gubuncode �Է�
    dim query '// ī�װ� ���� ���� ��Ͻ� mastercode ���� �� detailcode 0 �ڵ�����
    if mastercode="" then
        query = " select isnull(max(mastercode),0) as lastidx from db_event.dbo.tbl_division_groupcode "
        rsget.Open query,dbget,1
        if not rsget.EOF  then
            mastercode = rsget("lastidx")
        end if
        rsget.close

        mastercode = mastercode + 1
        detailcode = -1
    end if

    strSQL = " INSERT INTO db_event.dbo.tbl_division_groupcode "  & vbcrlf
    strSQL = strSQL & " (gubuncode , mastercode , detailcode , title , isusing) " & vbcrlf
    strSQL = strSQL & " values " & vbcrlf
    strSQL = strSQL & " ('"& gubuncode &"' , '"& mastercode &"' , '"& detailcode &"' , '"& catetitle &"' , "& groupisusing &") " & vbcrlf

    dbget.execute(strSQL)

    msg = "���� �Ǿ����ϴ�"

    Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_division_manage.asp"

elseif mode = "gubunModify" then '// gubuncode �Է�
    
    strSQL = " UPDATE db_event.dbo.tbl_division_groupcode SET "  & vbcrlf
    strSQL = strSQL & " mastercode = '"& mastercode &"' " & vbcrlf
    strSQL = strSQL & " ,detailcode = '"& detailcode &"' " & vbcrlf
    strSQL = strSQL & " ,title = '"& catetitle &"' " & vbcrlf
    strSQL = strSQL & " ,isusing = "& groupisusing &" " & vbcrlf
    strSQL = strSQL & " where gidx = "& gidx &" " & vbcrlf

    dbget.execute(strSQL)

    msg = "���� �Ǿ����ϴ�"

    Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_division_manage.asp"

elseif mode = "mdpicksortingedit" then '// pick ���� ����
    dim sIdx , sSortNo , sIsUsing

	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
        if detailcode > 0 then 
        sIsUsing = request.form("pickitem"&sIdx)
        else
        sIsUsing = request.form("isusing"&sIdx)
        end if		

		strSQL = strSQL & " UPDATE db_event.dbo.tbl_division_items SET "  & VBCRLF
		strSQL = strSQL & " pickitem = '"&sIsUsing&"'" & VBCRLF
        if detailcode > 0 then 
            strSQL = strSQL & " ,categorysorting = '"&sSortNo&"'" & VBCRLF
        else
            strSQL = strSQL & " ,picksorting = '"&sSortNo&"'" & VBCRLF
        end if 

		strSQL = strSQL & " WHERE idx =" & sIdx &";" & VBCRLF
	Next

'response.write strSQL
'response.end

	if strSQL <> "" then 
		dbget.execute strSQL
	end if 

    msg = "���� �Ǿ����ϴ�"

	Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_pickitems.asp?mastercode="& mastercode &"&detailcode="& detailcode &""

elseif mode = "delitem" then  

    strSQL = " DELETE FROM db_event.dbo.tbl_division_items where idx="& idx

    dbget.execute(strSQL)

    msg = "���� �Ǿ����ϴ�"

    Alert_move msg,"/admin/eventmanage/event/v5/popup/index.asp?menupos="&poscode&"&page="&page

elseif mode = "evtadd" then 

    strSQL = " INSERT INTO db_event.dbo.tbl_division_eventgroup "  & vbcrlf
    strSQL = strSQL & " (evt_code , mastercode , detailcode , startdate , enddate , isusing , evtsorting) " & vbcrlf
    strSQL = strSQL & " values " & vbcrlf
    strSQL = strSQL & " ('"& evt_code &"' , '"& mastercode &"' , '"& detailcode &"' , '"& startdate &"' , '"& enddate &"', "& evtisusing &", "& evtsorting &") " & vbcrlf

    dbget.execute(strSQL)

    strSQL ="select SCOPE_IDENTITY() "
    
	rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		idx = rsget(0)
	End IF
	rsget.close

    msg = "���� �Ǿ����ϴ�"

    Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_reg_event.asp?idx="& idx &"&mastercode="& mastercode

elseif mode = "evtmodify" then 

    strSQL = " UPDATE db_event.dbo.tbl_division_eventgroup SET "  & vbcrlf
    strSQL = strSQL & " evt_code = '"& evt_code &"' " & vbcrlf
    strSQL = strSQL & " ,startdate = '"& startdate &"' " & vbcrlf
    strSQL = strSQL & " ,enddate = '"& enddate &"' " & vbcrlf
    strSQL = strSQL & " ,isusing = "& evtisusing &" " & vbcrlf
    strSQL = strSQL & " ,evtsorting = "& evtsorting &" " & vbcrlf
    strSQL = strSQL & " where idx = "& idx &" " & vbcrlf

    dbget.execute(strSQL)

    msg = "���� �Ǿ����ϴ�"

    Alert_move msg,"/admin/eventmanage/event/v5/popup/pop_reg_event.asp?idx="& idx &"&mastercode="& mastercode

End IF
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
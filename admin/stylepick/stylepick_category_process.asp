<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.05 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<%  
dim mode ,menupos ,sqlstr ,cnt , msg, i , catetypeorg , catecodeorg ,isusingorg
dim catetype,catecode,catename,isusing,orderno,lastadminid,cateidx ,totalcount ,tmpitem
	mode = request("mode")
	catetype= request("catetype")
	catecode= request("catecode")
	catename= request("catename")
	isusing= request("isusing")
	orderno= request("orderno")
	cateidx= request("cateidx")					
	catetypeorg= request("catetypeorg")
	catecodeorg= request("catecodeorg")
	isusingorg= request("isusingorg")		
	lastadminid = session("ssBctId")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

'/�űԵ��
if mode = "itemadd" then

	if catetype = "" or catecode = "" or catename = "" or orderno = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('�ڵ忡 ������ �ֽ��ϴ�.������ ���� �ϼ���');"
		response.write "	location.replace('" + referer + "');"
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if

	sqlstr = "SELECT count(*) as count"
	
	if catetype = "CD1" then
		sqlstr = sqlstr & " FROM [db_giftplus].dbo.tbl_stylepick_cate_cd1"
	elseif catetype = "CD2" then
		sqlstr = sqlstr & " FROM [db_giftplus].dbo.tbl_stylepick_cate_cd2"
	elseif catetype = "CD3" then
		sqlstr = sqlstr & " FROM [db_giftplus].dbo.tbl_stylepick_cate_cd3"
	end if
	
	sqlstr = sqlstr & " where isusing = 'Y'"

	if catetype = "CD1" then
		sqlstr = sqlstr & " and cd1 = '"&catecode&"'"
	elseif catetype = "CD2" then
		sqlstr = sqlstr & " and cd2 = '"&catecode&"'"
	elseif catetype = "CD3" then
		sqlstr = sqlstr & " and cd3 = '"&catecode&"'"
	end if
	
	'response.write sqlstr &"<br>"
	rsget.open sqlstr ,dbget,1
	
	if not rsget.eof then
		cnt = rsget("count")
	end if

	rsget.close

	if cnt >0 then
		response.write "<script language='javascript'>"
		response.write	"	alert('�����Ǿ��ų� ,�̹� ������� ī�װ� �Դϴ�.\nī�װ� �ڵ带 Ȯ���� �ٽ� �Է����ּ���');"
		response.write "	location.replace('" + referer + "');"
		response.write "</script>"	
		dbget.close()	:	response.End
	end if

	if catetype = "CD1" then
		sqlStr = " insert into [db_giftplus].[dbo].[tbl_stylepick_cate_cd1]" + VbCrlf
    	sqlStr = sqlStr + " (cd1 ,catename ,isusing ,orderno ,lastadminid)"+ VbCrlf
	elseif catetype = "CD2" then
		sqlStr = " insert into [db_giftplus].[dbo].[tbl_stylepick_cate_cd2]" + VbCrlf
		sqlStr = sqlStr + " (cd2 ,catename ,isusing ,orderno ,lastadminid)"+ VbCrlf
	elseif catetype = "CD3" then
		sqlStr = " insert into [db_giftplus].[dbo].[tbl_stylepick_cate_cd3]" + VbCrlf
		sqlStr = sqlStr + " (cd3 ,catename ,isusing ,orderno ,lastadminid)"+ VbCrlf
	end if

    sqlStr = sqlStr + " values("	    
    sqlStr = sqlStr + " '" + catecode + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + db2html(catename) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ," + orderno + "" + VbCrlf
    sqlStr = sqlStr + " ,'" + lastadminid + "'" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    'response.write sqlStr    
    dbget.Execute sqlStr

	msg = "�ű����� �Ǿ����ϴ�"
	response.write "<script language='javascript'>"
	response.write "	alert('" & msg & "');"
	response.write "	location.replace('" + referer + "');"
	response.write "</script>"
	dbget.close()	:	response.End

'/����
elseif mode = "itemedit" then

	if catetype = "" or catecode = "" or catename = "" or orderno = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('�ڵ忡 ������ �ֽ��ϴ�.������ ���� �ϼ���');"
		response.write "	location.replace('" + referer + "');"
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if

	'//ī�װ� �ڵ� , ��뿩�θ� ���� �Ұ�� ī�װ��� �ش� ��ǰ�� ���� �ϴ��� üũ �Ѵ�.	'/10�Ǳ��� ����
	if catecode <> catecodeorg or isusing <> isusingorg then

		sqlstr = "SELECT top 10"
		sqlstr = sqlstr & " si.itemid"
		sqlstr = sqlstr & " , (c1.catename) as c1name ,(c2.catename) as c2name ,(c3.catename) as c3name"
		sqlstr = sqlstr & " FROM [db_giftplus].dbo.tbl_stylepick_item si"
	    sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
	    sqlStr = sqlStr & " 	on si.cd1 = c1.cd1"
	    sqlStr = sqlStr & " 	and c1.isusing='Y'"
	    sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd2 c2"
	    sqlStr = sqlStr & " 	on si.cd1 = c2.cd2"
	    sqlStr = sqlStr & " 	and c2.isusing='Y'"
	    sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd3 c3"
	    sqlStr = sqlStr & " 	on si.cd1 = c3.cd3"
	    sqlStr = sqlStr & " 	and c3.isusing='Y'"	    	    
		sqlstr = sqlstr & " where si.isusing='Y'"

		if catetype = "CD1" then
			sqlstr = sqlstr & " and si.cd1='"&catecodeorg&"'"
		elseif catetype = "CD2" then
			sqlstr = sqlstr & " and si.cd2='"&catecodeorg&"'"
		elseif catetype = "CD3" then
			sqlstr = sqlstr & " and si.cd3='"&catecodeorg&"'"
		end if
				
		'response.write sqlstr &"<Br>"
		rsget.open sqlstr ,dbget,1
		
		totalcount = rsget.recordcount
		
		if not rsget.EOF then
			do until rsget.EOF			
			i = i + 1						

			if catetype = "CD1" then
				tmpitem = tmpitem & "[��Ÿ�� / " & rsget("c1name") & "] ��ǰ�ڵ�:" & rsget("itemid") & "\n"
			elseif catetype = "CD2" then
				tmpitem = tmpitem & "[�з� / " & rsget("c2name") & "] ��ǰ�ڵ�:" & rsget("itemid") & "\n"
			elseif catetype = "CD3" then
				tmpitem = tmpitem & "[ī�װ�3 / " & rsget("c3name") & "] ��ǰ�ڵ�:" & rsget("itemid") & "\n"
			end if
									
			rsget.movenext
			loop
		end if
		
		rsget.Close
		
		'//�̺�Ʈ ��ǰ üũ
		totalcount = 0
	    sqlStr = "select top 10"
	    sqlStr = sqlStr & " ei.evtitemidx ,ei.evtidx ,ei.itemid ,ei.regdate ,ei.isusing"
	    sqlStr = sqlStr & " ,e.evtidx,e.title,e.subcopy,e.state,e.banner_img,e.startdate,e.enddate"
	    sqlStr = sqlStr & " ,e.isusing,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate,c1.catename"
	    sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event_item ei"
	    sqlStr = sqlStr & " join db_giftplus.dbo.tbl_stylepick_event e"
	    sqlStr = sqlStr & " 	on ei.evtidx = e.evtidx"
	    sqlStr = sqlStr & " 	and e.state <> 9 and (getdate() <= e.startdate or getdate() between e.startdate and e.enddate)"
	    sqlStr = sqlStr & " 	and e.isusing='Y'"
	    sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
	    sqlStr = sqlStr & " 	on e.cd1 = c1.cd1"
	    sqlStr = sqlStr & " 	and c1.isusing='Y'"
	    sqlStr = sqlStr & " where ei.isusing='Y'"

		if catetype = "CD1" then
			sqlstr = sqlstr & " and e.cd1='"&catecodeorg&"'"
		end if

		'response.write sqlstr &"<Br>"
		rsget.open sqlstr ,dbget,1
		
		totalcount = rsget.recordcount
		
		if not rsget.EOF then
			do until rsget.EOF			
			i = i + 1						
								
			if catetype = "CD1" then
				tmpitem = tmpitem & "[" & rsget("catename") & " / ��ȹ���ڵ�:" & rsget("evtidx") & "] ��ǰ�ڵ�:" & rsget("itemid") & "\n"
			end if
									
			rsget.movenext
			loop
		end if
		
		rsget.Close
	
		'if tmpitem <> "" then
		'	tmpitem = "���Ϲݻ�ǰ�� ��ȹ����ǰ ���� 10�Ǳ��� ����˴ϴ�\n\n" & tmpitem
		'	response.write	"<script language='javascript'>"
		'	response.write	"	alert('��ǰ�� �����ִ� ī�װ��� ī�װ��ڵ�� ��뿩�δ� ���� �Ҽ� �����ϴ�.\nȮ���� �ٽ� �Է����ּ���\n\n"&tmpitem&"');"
		''	response.write "	location.replace('" + referer + "');"				
		'	response.write	"</script>"
		'	dbget.close()	:	response.End
		'end if	
	end if

	if catetype = "CD1" then
		sqlStr = " update [db_giftplus].[dbo].[tbl_stylepick_cate_cd1] set" + VbCrlf
		sqlStr = sqlStr + " cd1='" + catecode + "'" + VbCrlf
	elseif catetype = "CD2" then
		sqlStr = " update [db_giftplus].[dbo].[tbl_stylepick_cate_cd2] set" + VbCrlf
		sqlStr = sqlStr + " cd2='" + catecode + "'" + VbCrlf
	elseif catetype = "CD3" then
		sqlStr = " update [db_giftplus].[dbo].[tbl_stylepick_cate_cd3] set" + VbCrlf
		sqlStr = sqlStr + " cd3='" + catecode + "'" + VbCrlf
	end if

    sqlStr = sqlStr + " ,catename='" + db2html(catename) + "'" + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,orderno=" + orderno + "" + VbCrlf
    sqlStr = sqlStr + " ,lastadminid='" + lastadminid + "'" + VbCrlf

	if catetype = "CD1" then
		sqlStr = sqlStr + " where cd1="&catecodeorg&""
	elseif catetype = "CD2" then
		sqlStr = sqlStr + " where cd2="&catecodeorg&""
	elseif catetype = "CD3" then
		sqlStr = sqlStr + " where cd3="&catecodeorg&""
	end if
    
    'response.write sqlStr
    dbget.Execute sqlStr

	msg = "���� �Ǿ����ϴ�"
	response.write "<script language='javascript'>"
	response.write "	alert('" & msg & "');"
	response.write "	location.replace('" + referer + "');"
	response.write "</script>"
	dbget.close()	:	response.End	

end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
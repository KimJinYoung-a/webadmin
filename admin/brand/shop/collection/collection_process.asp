<%@ language=vbscript %>

<% option explicit %>

<%

'###########################################################

' Description :  �귣�彺Ʈ��Ʈ

' History : 2013.08.29 �ѿ�� ����

'###########################################################

%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->

<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/offshop_function.asp"-->

<!-- #include virtual="/lib/classes/street/shopcls.asp"-->

<%

dim idx, makerid, title, subtitle, state, mainimg, isusing, sortNo, regdate, tmpstate

dim lastupdate, regadminid, lastadminid, comment, sqlStr, mode, adminid, detailitemcnt, itemidarr

dim i, existsitemcnt, existsbrandcnt

	idx 	= request("idx")

	makerid 	= request("makerid")

	title 	= request("title")

	subtitle 	= request("subtitle")

	state 	= request("state")

	mainimg 	= request("mainimg")

	isusing 	= request("isusing")

	sortNo 	= request("sortNo")

	comment 	= request("comment")

	mode 	= request("mode")

	menupos 	= request("menupos")

	itemidarr 	= request("itemidarr")

	

adminid = session("ssBctId")

detailitemcnt = 0

existsitemcnt = 0

existsbrandcnt = 0



If mode = "I" Then



	sqlStr = "SELECT count(*) as cnt"

	sqlStr = sqlStr & " from db_user.dbo.tbl_user_c"

	sqlStr = sqlStr & " WHERE userid='"&makerid&"'"

	

	'response.write sqlStr & "<BR>"

	rsget.Open sqlStr, dbget, 1

    If Not rsget.Eof then

    	existsbrandcnt = rsget("cnt")

	End If

    rsget.Close



	If existsbrandcnt = 0 Then

		Response.Write  "<script language='javascript'>"

		Response.Write  "	alert('�ش�Ǵ� �귣�尡 �����ϴ�.');"

		Response.Write  "	location.replace('/admin/brand/shop/collection/collectionModify.asp?menupos="&menupos&"');"

		Response.Write  "</script>"

		dbget.close()	:	response.End

	End If



	if checkNotValidHTML(comment) or checkNotValidHTML(title) or checkNotValidHTML(subtitle) then

	%>

		<script>

			alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');

			history.go(-1);

		</script>		

	<%		

		dbget.close()	:	response.End

	end if

	

	sqlStr = "INSERT INTO db_brand.dbo.tbl_street_shop_collection (" + vbcrlf

	sqlStr = sqlStr & " makerid, title, subtitle, state, mainimg, isusing, regadminid, lastadminid, comment)" + vbcrlf

	sqlStr = sqlStr & " 	select" + vbcrlf

	sqlStr = sqlStr & " 	c.userid as makerid, '"&html2db(title)&"', '"&html2db(subtitle)&"', 3, '"&mainimg&"', '"&isusing&"','"&adminid&"'" + vbcrlf

	sqlStr = sqlStr & " 	,'"&adminid&"', '"&html2db(comment)&"'" + vbcrlf

	sqlStr = sqlStr & " 	from db_user.dbo.tbl_user_c c" + vbcrlf

	sqlStr = sqlStr & " 	where userid='"& makerid &"'"

	

	'response.write sqlStr & "<BR>"

	dbget.execute sqlStr



	sqlStr = "select IDENT_CURRENT('db_brand.dbo.tbl_street_shop_collection') as idx"

	rsget.Open sqlStr, dbget, 1



	If Not rsget.Eof then

		idx = rsget("idx")

	End If

	rsget.close



	Response.Write "<script language='javascript'>alert('����Ǿ����ϴ�');location.replace('/admin/brand/shop/collection/collectionModify.asp?idx="&idx&"&menupos="&menupos&"');</script>"



ElseIf mode = "U" Then

	if idx="" then

	%>

		<script>

			alert('IDX�� �����ϴ�.');

			history.go(-1);

		</script>		

	<%		

		dbget.close()	:	response.End

	end if
	


	if checkNotValidHTML(comment) or checkNotValidHTML(title) or checkNotValidHTML(subtitle) then

	%>

		<script>

			alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');

			history.go(-1);

		</script>		

	<%		

		dbget.close()	:	response.End

	end if

		

	sqlStr = "UPDATE db_brand.dbo.tbl_street_shop_collection SET" + VBCRLF

	sqlStr = sqlStr & " makerid = '"& makerid &"'" + VBCRLF

	sqlStr = sqlStr & " , title = '"& html2db(title) &"'" + VBCRLF

	sqlStr = sqlStr & " , subtitle = '"& html2db(subtitle) &"'" + VBCRLF

	sqlStr = sqlStr & " ,state = "&state&"" + VBCRLF

	sqlStr = sqlStr & " ,mainimg = '"&mainimg&"'" + VBCRLF

	sqlStr = sqlStr & " , isusing = '"&isusing&"'" + VBCRLF

	sqlStr = sqlStr & " ,sortNo = '"&sortNo&"'" + VBCRLF

	sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf

	sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf

	sqlStr = sqlStr & " , comment = '"& html2db(comment) &"'" + VBCRLF

	sqlStr = sqlStr & " where idx ='" & Cstr(idx) & "'"



	'response.write sqlStr & "<BR>"	

	dbget.execute sqlStr



	response.write "<script language='javascript'>"

	response.write "	alert('����Ǿ����ϴ�');"

	response.write "	location.replace('/admin/brand/shop/collection/collectionModify.asp?idx="&idx&"&menupos="&menupos&"');"

	response.write "</script>"	



'/���� ����

elseif mode="chstate" then

	if idx="" then

	%>

		<script>

			alert('���� �����ϴ�.');

			history.go(-1);

		</script>		

	<%		

		dbget.close()	:	response.End

	end if

	

	sqlStr = "SELECT count(*) as cnt"

	sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection as M"

	sqlStr = sqlStr & " JOIN db_brand.dbo.tbl_street_shop_collection_item as D"

	sqlStr = sqlStr & " 	on M.idx=D.masteridx"

	sqlStr = sqlStr & " WHERE m.idx="&idx&" and D.isusing='Y' "

	

	'response.write sqlStr & "<BR>"

	rsget.Open sqlStr, dbget, 1

    If Not rsget.Eof then

    	detailitemcnt = rsget("cnt")

	End If

    rsget.Close

	

	if state="3" or state="7" then

		If detailitemcnt = "0" Then

			Response.Write  "<script language='javascript'>"

			Response.Write  "	alert('collection ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�.\n����Ͻð� �ٽ� �õ� �ϼ���.');"

			Response.Write  "	history.go(-1);"

			Response.Write  "</script>"

			dbget.close()	:	response.End		

		End If

	End If	


	sqlStr = "UPDATE db_brand.dbo.tbl_street_shop_collection SET" + VBCRLF

	sqlStr = sqlStr & " state = "&state&"" + VBCRLF

	sqlStr = sqlStr & " where idx ='" & Cstr(idx) & "'"



	'response.write sqlStr & "<BR>"	

	dbget.execute sqlStr



	response.write "<script language='javascript'>"

	response.write "	alert('����Ǿ����ϴ�');"

	response.write "	location.replace('/admin/brand/shop/collection/collectionModify.asp?idx="&idx&"&menupos="&menupos&"');"

	response.write "</script>"



'/��ǰ�߰�

elseif mode="itemreg" then



	If idx = "" Then

		Response.Write  "<script language='javascript'>"

		Response.Write  "	alert('���а��� �����ϴ�');"

		Response.Write  "</script>"

		dbget.close()	:	response.End		

	End If

	If itemidarr = "" Then

		Response.Write  "<script language='javascript'>"

		Response.Write  "	alert('��ǰ�� �����ϼ���.');"

		Response.Write  "</script>"

		dbget.close()	:	response.End		

	End If

	

	itemidarr = split(itemidarr,",")

	for i = 0 to ubound(itemidarr)

		existsitemcnt = 0

		sqlStr = "SELECT count(*) as cnt"

		sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection_item as d"

		sqlStr = sqlStr & " WHERE d.isusing='Y' and d.masteridx="& trim(idx) &""

		

		'response.write sqlStr & "<BR>"

		rsget.Open sqlStr, dbget, 1

	    If Not rsget.Eof then

	    	existsitemcnt = rsget("cnt")

		End If

	    rsget.Close

		if existsitemcnt>39 then

			response.write "<script language='javascript'>"

			response.write "	alert('��ǰ�� 40������ ��ϵʴϴ�.');"

			response.write "</script>"

			exit for

		end if



		sqlStr = "insert into db_brand.dbo.tbl_street_shop_collection_item(" & VBCRLF

		sqlStr = sqlStr & " masteridx, itemid, isusing, regadminid, lastadminid)" & VBCRLF

		sqlStr = sqlStr & " 	select top 500" & VBCRLF

		sqlStr = sqlStr & " 	"& trim(idx) &", i.itemid, 'Y', '"&adminid&"', '"&adminid&"'" & VBCRLF

		sqlStr = sqlStr & " 	from db_item.dbo.tbl_item i" & VBCRLF

		sqlStr = sqlStr & " 	left join db_brand.dbo.tbl_street_shop_collection_item d" & VBCRLF

		sqlStr = sqlStr & " 		on i.itemid=d.itemid" & VBCRLF

		sqlStr = sqlStr & " 		and d.masteridx="& trim(idx) &"" & VBCRLF

		sqlStr = sqlStr & " 		and d.isusing='Y'" & VBCRLF

		sqlStr = sqlStr & " 	where d.detailidx is null" & VBCRLF

		sqlStr = sqlStr & " 	and i.itemid in ("& trim(itemidarr(i)) &")"

		

		'response.write sqlStr & "<Br>"

		dbget.execute sqlStr

	next



	response.write "<script language='javascript'>"

	response.write "	alert('����Ǿ����ϴ�');"

	response.write "	opener.document.location.reload();"

	response.write "	self.close();"

	response.write "</script>"

	

else

	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"

	dbget.close()	:	response.End	

End If

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->

<!-- #include virtual="/lib/db/dbclose.asp" -->
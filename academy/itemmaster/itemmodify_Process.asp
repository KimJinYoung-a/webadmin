<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
	'// ���� ��� ����
	dim mode, i
	mode = RequestCheckvar(Request("mode"),32)

    '// ��ǰ��ȣ�� �޴´� //
    dim realitemid
    realitemid = RequestCheckvar(Request("itemid"),10)
    
    '// ��ۺ� ��å
    dim deliveryType, itemcontent, ordercomment, designercomment, requirecontents, refundpolicy
    deliveryType = RequestCheckvar(Request("deliveryType"),2)
	itemcontent = Request("itemcontent")
	ordercomment = Request("ordercomment")
	designercomment = Request("designercomment")
	requirecontents = Request("requirecontents")
	refundpolicy = Request("refundpolicy")
  	if itemcontent <> "" then
		if checkNotValidHTML(itemcontent) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
	if ordercomment <> "" then
		if checkNotValidHTML(ordercomment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
	if designercomment <> "" then
		if checkNotValidHTML(designercomment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
	if requirecontents <> "" then
		if checkNotValidHTML(requirecontents) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
	if refundpolicy <> "" then
		if checkNotValidHTML(refundpolicy) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
	'###########################################################################
	'��ǰ ���� ����
	'###########################################################################

	'// Ʈ������ ����
	dbACADEMYget.beginTrans

	'// ���� ��� ����(�⺻���� ����, �������� ����)
	dim sqlStr
	Select Case mode
		Case "ItemBasicInfo"
			'###########################################################################
			'��ǰ �⺻���� ����
			'###########################################################################

			sqlStr = "update db_academy.dbo.tbl_diy_item" + vbCrlf
			sqlStr = sqlStr & " set itemdiv='" & Cstr(RequestCheckvar(Request("itemdiv"),10)) & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemname='" & chrbyte(html2db(Request("itemname")),64,"") & "'" & vbCrlf
			sqlStr = sqlStr & " ,lastupdate=getdate()"
			'' ��ü ���� �ڵ� �߰�
    		sqlStr = sqlStr & " ,upchemanagecode='" & html2db(RequestCheckvar(Request("upchemanagecode"),32)) & "'" & vbCrlf
			sqlStr = sqlStr & " ,requireimgchk='" & html2db(RequestCheckvar(Request("requireimgchk"),32)) & "'" & vbCrlf
			sqlStr = sqlStr & " where itemid=" & CStr(realitemid) & "" + vbCrlf

			dbACADEMYget.execute(sqlStr)
			
			
			sqlStr = "update db_academy.dbo.tbl_diy_item_Contents" + vbCrlf
			sqlStr = sqlStr & " set itemcontent='" & html2db(itemcontent) & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemsource='" & chrbyte(html2db(Request("itemsource")),128,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemsize='" & chrbyte(html2db(Request("itemsize")),128,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemWeight='" & chrbyte(html2db(Request("itemWeight")),12,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,sourcearea='" & chrbyte(html2db(Request("sourcearea")),128,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,makername='" & chrbyte(html2db(Request("makername")),64,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,keywords='" & chrbyte(html2db(Request("keywords")),128,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,usinghtml='" & RequestCheckvar(Request("usinghtml"),1) & "'" + vbCrlf
			sqlStr = sqlStr & " ,ordercomment='" & chrbyte(html2db(ordercomment),1024,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,designercomment='" & chrbyte(html2db(designercomment),255,"") & "'" + vbCrlf

			sqlStr = sqlStr & " ,cstodr='" & RequestCheckvar(Request("cstodr"),1) & "'" + vbCrlf
			sqlStr = sqlStr & " ,requireMakeDay='" & RequestCheckvar(Request("requireMakeDay"),10) & "'" + vbCrlf
			sqlStr = sqlStr & " ,requirecontents='" & html2db(requirecontents) & "'" + vbCrlf
			sqlStr = sqlStr & " ,refundpolicy='" & html2db(refundpolicy) & "'" + vbCrlf
			sqlStr = sqlStr & " ,infoDiv='" & RequestCheckvar(Request("infoDiv"),2) & "'" + vbCrlf
			sqlStr = sqlStr & " ,safetyYn='" & RequestCheckvar(Request("safetyYn"),1) & "'" + vbCrlf
			sqlStr = sqlStr & " ,safetyDiv='" & RequestCheckvar(Request("safetyDiv"),9) & "'" + vbCrlf
			sqlStr = sqlStr & " ,safetyNum='" & chrbyte(html2db(Request("safetyNum")),24,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,freight_min='" & getNumeric(RequestCheckvar(Request("freight_min"),10)) & "'" + vbCrlf
			sqlStr = sqlStr & " ,freight_max='" & getNumeric(RequestCheckvar(Request("freight_max"),10)) & "'" + vbCrlf
			sqlStr = sqlStr & " ,requireMakeEmail='" & html2db(RequestCheckvar(Request("requireMakeEmail"),100)) & "'" + vbCrlf

			sqlStr = sqlStr & " where itemid=" & CStr(realitemid) & "" + vbCrlf
	        
	        dbACADEMYget.execute(sqlStr)
			
			
			'// ����ī�װ� �ֱ� //
			sqlStr = "delete from db_academy.dbo.tbl_display_cate_item_Academy Where itemid='" & realitemid & "';" & vbCrLf
			If (Request("catecode").Count>0) Then
				sqlStr = sqlStr & "update db_academy.dbo.tbl_diy_item set dispcate1=null Where itemid='" & realitemid & "';" & vbCrLf
				for i=1 to Request("catecode").Count
					sqlStr = sqlStr & "Insert into db_academy.dbo.tbl_display_cate_item_Academy (catecode, itemid, depth, sortNo, isDefault) values "
					sqlStr = sqlStr & "('" & RequestCheckvar(Request("catecode")(i),10) & "'"
					sqlStr = sqlStr & ",'" & realitemid & "'"
					sqlStr = sqlStr & ",'" & RequestCheckvar(Request("catedepth")(i),10) & "',9999"
					sqlStr = sqlStr & ",'" & RequestCheckvar(Request("isDefault")(i),1) & "');" & vbCrLf
					
					if Request("isDefault")(i)="y" then
						sqlStr = sqlStr & "update db_academy.dbo.tbl_diy_item set dispcate1='" & left(RequestCheckvar(Request("catecode")(i),10),3) & "' Where itemid='" & realitemid & "';" & vbCrLf
					end if
				next
			end if
			dbACADEMYget.execute(sqlStr)

			'###########################################################################
			'��ǰ ǰ�������� ���� 
			'###########################################################################
			if RequestCheckvar(Request("infoDiv"),2)<>"" then
				dim infoCd, infoCont, infoChk
			
				'�迭�� ó��
				redim infoCd(Request("infoCd").Count)
				redim infoCont(Request("infoCont").Count)
				redim infoChk(Request("infoChk").Count)
				for i=1 to Request("infoCd").Count
					infoCd(i) = RequestCheckvar(Request("infoCd")(i),5)
					infoCont(i) = Request("infoCont")(i)
					infoChk(i) = RequestCheckvar(Request("infoChk")(i),1)
					if infoCont(i) <> "" then
						if checkNotValidHTML(infoCont(i)) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
						response.write "</script>"
						response.End
						end if
					end if
				next
			
				'������ ����
				sqlStr = "Delete From db_academy.dbo.tbl_diy_item_infoCont Where itemid='" & CStr(realitemid) & "'"
				dbACADEMYget.execute(sqlStr)
			
				'DB�� ó��
				for i=1 to ubound(infoCd)
					'�Է°��� �ִ� ��츸 ����
					if infoChk(i)<>"" or infoCont(i)<>"" then
						sqlStr = "Insert into db_academy.dbo.tbl_diy_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
						sqlStr = sqlStr & "('" & CStr(realitemid) & "'"
						sqlStr = sqlStr & ",'" & CStr(infoCd(i)) & "'"
						sqlStr = sqlStr & ",'" & CStr(infoChk(i)) & "'"
						sqlStr = sqlStr & ",'" & html2db(infoCont(i)) & "')"
						dbACADEMYget.execute(sqlStr)
					end if
				Next
			end If
			'###########################################################################
			
			'// 2016.2.16 �ű��߰� ��ǰ�󼼼��� ������ �߰� - ������
			'// 2016.07.12 ���� - ����ȭ
			'// ������ ������ �� ���Խ����� src, width, height�� �̾Ƴ�
			If Trim(request("itemvideo")) <> "" Then
				Dim itemvideo, RetStr, RetSrc, RetWidth, RetHeight, regEx, Matches, Match, VideoTempSrc, VideoTempWidth, VideoTempHeight, videoType, dbsql
				itemvideo = request("itemvideo")
				itemvideo = Trim(Replace(itemvideo,"""","'"))
				'// iframe �̿��� �ڵ�� �߶����
				itemvideo = Left(itemvideo, InStrRev(itemvideo, "</iframe>")+9)

				'// ���� Ÿ������(���������� ��޿�����)
				If InStr(itemvideo, "youtube")>0 Then
					videoType = "youtube"
				ElseIf InStr(itemvideo, "vimeo")>0 Then
					videoType = "vimeo"
				Else
					videoType = "etc"
				End If

				Set regEx = New RegExp
				regEx.IgnoreCase = True
				regEx.Global = True

				regEx.pattern = "<iframe [^<>]*>"
				Set Matches = regEx.execute(itemvideo)
				For Each Match In Matches
					VideoTempSrc =  Mid(Match.Value, InStrRev(Match.Value,"src='")+5)
					RetSrc = Left(VideoTempSrc, InStr(VideoTempSrc, "'")-1)
					
					If InStrRev(Match.Value,"width='") > 0 then
					VideoTempWidth =  Mid(Match.Value, InStrRev(Match.Value,"width='")+7)
					RetWidth = Left(VideoTempWidth, InStr(VideoTempWidth, "'")-1)
					End If 
					
					If InStrRev(Match.Value,"height='") > 0 then
					VideoTempHeight =  Mid(Match.Value, InStrRev(Match.Value,"height='")+8)
					RetHeight = Left(VideoTempHeight, InStr(VideoTempHeight, "'")-1)
					End If 
				Next
				Set regEx = Nothing
				Set Matches = Nothing

				sqlStr = " select idx FROM db_academy.dbo.tbl_diy_item_videos WHERE videogubun='video1' And itemid =" + CStr(realitemid)
				rsACADEMYget.Open sqlStr,dbACADEMYget,1
				if Not rsACADEMYget.Eof Then
					If Not(videoType="etc") Then
						'// �����Ͱ� �ִٸ� ������Ʈ ����.
						dbsql = "update db_academy.dbo.tbl_diy_item_videos" + vbCrlf
						dbsql = dbsql & " set videourl='" &RetSrc& "'" + vbCrlf
						dbsql = dbsql & " ,videowidth='" & RetWidth & "'" + vbCrlf
						dbsql = dbsql & " ,videoheight='" & RetHeight & "'" + vbCrlf
						dbsql = dbsql & " ,videotype='" & videoType & "'" + vbCrlf
						dbsql = dbsql & " ,videofullurl='" & chrbyte(html2db(itemvideo),255,"") & "'" + vbCrlf
						dbsql = dbsql & " ,modifydate=getdate()" + vbCrlf
						dbsql = dbsql & " where idx='"&rsACADEMYget("idx")&"' And itemid='" & CStr(realitemid) & "'" + vbCrlf
						dbACADEMYget.execute(dbsql)
					End If
				Else
					If Not(videoType="etc") Then
						'// �����Ͱ� ������ �μ�Ʈ ����.
						dbsql = " insert into db_academy.dbo.tbl_diy_item_videos (itemid, videogubun, videotype, videourl, videowidth, videoheight, videofullurl, regdate) values " + vbCrlf
						dbsql = dbsql & " ('"&CStr(realitemid)&"', 'video1', '"&videoType&"', '"&RetSrc&"', '"&RetWidth&"', '"&RetHeight&"','"&chrbyte(html2db(itemvideo),255,"")&"', getdate()) " + vbCrlf
						dbACADEMYget.execute(dbsql)
					End If
				end if
				rsACADEMYget.close
			Else
				'// �ƹ����� �ȳѾ�Դµ� db�� ���� ������ ������� �Ǵ�. ������.
				sqlStr = " select idx FROM db_academy.dbo.tbl_diy_item_videos WHERE videogubun='video1' And itemid =" + CStr(realitemid)  
				rsACADEMYget.Open sqlStr,dbACADEMYget,1
				if Not rsACADEMYget.Eof Then
					dbsql = " Delete from db_academy.dbo.tbl_diy_item_videos Where videogubun='video1' And itemid=" + CStr(realitemid)
					dbACADEMYget.execute(dbsql)
				End If
				rsACADEMYget.close
			End If

		Case "ItemPriceInfo"
			'###########################################################################
			'��ǰ �Ǹ�/�������� ����
			'###########################################################################

			'// ���� ���� ����
	        dim sailprice, sailsuplycash, orgprice, orgsuplycash, sellcash, buycash
	        
	
			if RequestCheckvar(Request("saleyn"),2)= "Y" then
				sailprice = RequestCheckvar(Request("sailprice"),10)
				sailsuplycash = RequestCheckvar(Request("sailsuplycash"),10)
				orgprice = RequestCheckvar(Request("sellcash"),10)
				orgsuplycash = RequestCheckvar(Request("buycash"),10)
				sellcash = RequestCheckvar(Request("sailprice"),10)
				buycash = RequestCheckvar(Request("sailsuplycash"),10)
			else
				sailprice = RequestCheckvar(Request("sailprice"),10)
				sailsuplycash = RequestCheckvar(Request("sailsuplycash"),10)
				orgprice = RequestCheckvar(Request("sellcash"),10)
				orgsuplycash = RequestCheckvar(Request("buycash"),10)
				sellcash = RequestCheckvar(Request("sellcash"),10)
				buycash = RequestCheckvar(Request("buycash"),10)
			end if
            
            ''//��ۺ� ��å ** 
            if (RequestCheckvar(request("mwdiv"),2)="U") then
                ''��ü ����� ��� ��ü�� ��ۺ� �ΰ��� �ƴϸ� 2 - ����⺻
                if (deliveryType<>"9") and (deliveryType<>"7") then
                    deliveryType = "2"
                end if
            else
                ''��ü ����� �ƴѰ�� �������� �ƴϸ� 1 - �ٹ�⺻
                if (deliveryType<>"4") then
                    deliveryType = "1"
                end if
            end if
        
        
			'// ��ǰ ������ �Է� //
			sqlStr = "update db_academy.dbo.tbl_diy_item" + vbCrlf
			sqlStr = sqlStr & " set sellcash=" & Cstr(sellcash) & "" + vbCrlf
			sqlStr = sqlStr & " ,buycash=" & Cstr(buycash) & "" + vbCrlf
			sqlStr = sqlStr & " ,mileage=" & RequestCheckvar(Request("mileage"),10) & "" + vbCrlf
			sqlStr = sqlStr & " ,vatyn='" & RequestCheckvar(Request("vatyn"),1) & "'" + vbCrlf
			sqlStr = sqlStr & " ,sellyn='" & RequestCheckvar(Request("sellyn"),2) & "'" + vbCrlf
			sqlStr = sqlStr & " ,isusing='" & RequestCheckvar(Request("isusing"),2) & "'" + vbCrlf
			sqlStr = sqlStr & " ,saleyn='" & RequestCheckvar(Request("saleyn"),2) & "'" + vbCrlf
			sqlStr = sqlStr & " ,sailprice=" & Cstr(sailprice) & "" + vbCrlf
			sqlStr = sqlStr & " ,sailsuplycash=" & Cstr(sailsuplycash) & "" + vbCrlf
			sqlStr = sqlStr & " ,orgprice=" & Cstr(orgprice) & "" + vbCrlf
			sqlStr = sqlStr & " ,orgsuplycash=" & Cstr(orgsuplycash) & "" + vbCrlf
			sqlStr = sqlStr & " ,mwdiv='" & RequestCheckvar(Request("mwdiv"),1) & "'" + vbCrlf
			sqlStr = sqlStr & " ,deliverytype='" & deliverytype & "'" + vbCrlf
			sqlStr = sqlStr & " ,availPayType='" & RequestCheckvar(Request("availPayType"),1) & "'" + vbCrlf
			sqlStr = sqlStr & " ,lastupdate=getdate()"
			sqlStr = sqlStr & " where itemid='" & realitemid & "'" + vbCrlf
			dbACADEMYget.execute(sqlStr)
	End Select


	'�귣�� �̸� �ֱ�
	sqlStr =	"Update db_academy.dbo.tbl_diy_item Set " &_
				"	 brandname=U.socname" &_
				"		from [TENDB].[db_user].[dbo].tbl_user_c as U" &_
				"		where db_academy.dbo.tbl_diy_item.itemid=" &  CStr(realitemid) &_
				"			and db_academy.dbo.tbl_diy_item.makerid=U.userid"
	dbACADEMYget.execute(sqlStr)

	'// ��ī�װ� ���� //
	dim NewCd1, NewCd2, NewCd3, NewDiv, ArrTemp, lp
    dim CateArr
	if Request("cate_div")<>"" then
		'�� ����
		ArrTemp = Request("cate_large")
		NewCd1 = Split(ArrTemp,",")
		ArrTemp = Request("cate_mid")
		NewCd2 = Split(ArrTemp,",")
		ArrTemp = Request("cate_small")
		NewCd3 = Split(ArrTemp,",")
		ArrTemp = Request("cate_div")
		NewDiv = Split(ArrTemp,",")
        
        CateArr = ""
        for lp=0 to Ubound(NewDiv)
            CateArr = CateArr + "'" + CStr(realitemid) + Trim(NewDiv(lp)) + Trim(NewCd1(lp)) + Trim(NewCd2(lp)) + Trim(NewCd3(lp)) + "'" + ","
        next
        CateArr = Trim(CateArr)
        
        if right(CateArr,1)="," then CateArr= Left(CateArr,Len(CateArr)-1)
        
        sqlStr = " Delete From db_academy.dbo.tbl_diy_item_category " 
        sqlStr = sqlStr & " Where itemid=" & realitemid
        sqlStr = sqlStr & " and ((Convert(varchar(10),itemid) + code_div + code_large + code_mid + code_small)"
        sqlStr = sqlStr & "      not in (" & CateArr & ") "
        sqlStr = sqlStr & "      )"
        dbACADEMYget.execute(sqlStr)
 
		for lp=0 to Ubound(NewDiv)
			if Trim(NewDiv(lp))="D" then
				sqlStr = " Update db_academy.dbo.tbl_diy_item Set "
				sqlStr = sqlStr & "	 cate_large='" & Trim(NewCd1(lp)) & "' " 
				sqlStr = sqlStr & "	 ,cate_mid='" & Trim(NewCd2(lp)) & "' "
				sqlStr = sqlStr & "	 ,cate_small='" & Trim(NewCd3(lp)) & "' " 
				sqlStr = sqlStr & " Where itemid=" & realitemid 
				sqlStr = sqlStr & " and (cate_large<>'" & Trim(NewCd1(lp)) & "' " 
				sqlStr = sqlStr & "     or cate_mid<>'" & Trim(NewCd2(lp)) & "' " 
				sqlStr = sqlStr & "     or cate_small<>'" & Trim(NewCd3(lp)) & "' " 
				sqlStr = sqlStr & " ) "

				dbACADEMYget.execute(sqlStr)
			end if
			
			''���� ī�װ��� ���°�츸 �Է�
			sqlStr = "Insert into db_academy.dbo.tbl_diy_item_category "
			sqlStr = sqlStr & " (itemid,code_large,code_mid,code_small,code_div)  " 
			sqlStr = sqlStr & " select i.itemid" 
			sqlStr = sqlStr & " ,'" & Trim(NewCd1(lp)) & "'" 
			sqlStr = sqlStr & " ,'" & Trim(NewCd2(lp)) & "'" 
			sqlStr = sqlStr & " ,'" & Trim(NewCd3(lp)) & "'" 
			sqlStr = sqlStr & " ,'" & Trim(NewDiv(lp)) & "'"
			sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
			sqlStr = sqlStr & "     left join db_academy.dbo.tbl_diy_item_category c"
			sqlStr = sqlStr & "     on i.itemid=c.itemid"
			sqlStr = sqlStr & "     and c.code_large='" & Trim(NewCd1(lp)) & "'" 
			sqlStr = sqlStr & "     and c.code_mid='" & Trim(NewCd2(lp)) & "'" 
			sqlStr = sqlStr & "     and c.code_small='" & Trim(NewCd3(lp)) & "'" 
			sqlStr = sqlStr & "     and c.code_div='" & Trim(NewDiv(lp)) & "'" 
			sqlStr = sqlStr & " where i.itemid=" & realitemid 
			sqlStr = sqlStr & " and c.itemid Is NULL"
			
			dbACADEMYget.execute(sqlStr)
		next       
	end if

	'##### DB ���� ó�� #####
    If Err.Number = 0 Then
    	dbACADEMYget.CommitTrans				'Ŀ��(����)
    	Response.Write	"<script language=javascript>" &_
    					"	alert('�����͸� �����Ͽ����ϴ�.');" &_
    					"	opener.location.reload();" &_
    					"	self.close();" &_
    					"</script>"
    Else
        dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)
    	Response.Write	"<script language=javascript>" &_
    					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
    					"	self.history.back();" &_
    					"</script>"
    End If

        
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
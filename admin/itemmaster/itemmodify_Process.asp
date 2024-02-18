<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : ��ǰ����
' History : ������ ����
'			2023.03.03 �ѿ�� ����(��ǰ��� A/S å����/��ȭ��ȣ ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'// ���� ��� ����
	dim mode, i, vChangeContents, vSCMChangeSQL, vChangeContentsCa, vSCMChangeSQLCa, itemname
	mode = Request("mode")
	itemname = requestCheckvar(Request("itemname"),64)
	vChangeContents = "- HTTP_REFERER : " & request.ServerVariables("HTTP_REFERER") & vbCrLf
    
    '// ���԰� ǥ�ÿ��� : ���԰� �����Ϸ� ����
	dim reipgodate
	reipgodate = requestCheckvar(Request("reipgodate"),20)
	
	On Error Resume Next
    reipgodate = CDate(reipgodate)
    If Err then reipgodate=""
    
    Err = 0

    On Error Goto 0
    
    '// ��ǰ��ȣ�� �޴´� //
    dim realitemid
    realitemid = requestCheckvar(Request("itemid"),10)
    
    '// ��ۺ� ��å
    dim deliveryType
    deliveryType = requestCheckvar(Request("deliveryType"),10)
    
    dim reserveItemTp
    reserveItemTp = requestCheckvar(Request("reserveItemTp"),10)
    if (reserveItemTp="") then reserveItemTp=0
    
	'###########################################################################
	'��ǰ ���� ����
	'###########################################################################

	'// Ʈ������ ����
	dbget.beginTrans

	'// ���� ��� ����(�⺻���� ����, �������� ����)
	dim sqlStr
	Select Case mode
		Case "ItemBasicInfo"

			if trim(stripHTML(itemname))="" Then
				dbget.RollBackTrans				'�ѹ�(�����߻���)
				Response.Write	"<script language=javascript>" &_
								"	alert('��ǰ���� ���ų� HTML�±� ���·� �ۼ��Ǿ����ϴ�. ���� �� �ٽ� ������ּ���.');" &_
								"	self.history.back();" &_
								"</script>"
				dbget.close() : response.end
			end if

			'###########################################################################
			'��ǰ üũ
			'###########################################################################
			'/2016.07.06 �ѿ�� �߰�
			if requestCheckvar(Request("makerid"),32)<>"" then
				sqlStr = "select top 1 userid" & vbcrlf
				sqlStr = sqlStr & " from db_user.dbo.tbl_user_c" & vbcrlf
				sqlStr = sqlStr & " where userid= '"& requestCheckvar(Request("makerid"),32) &"'" & vbcrlf
	
				'response.write sqlStr & "<Br>"			
				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof Then
				Else
					dbget.RollBackTrans				'�ѹ�(�����߻���)
			    	Response.Write	"<script language=javascript>" &_
			    					"	alert('�Է��Ͻ� �귣��ID �� �������� �ʽ��ϴ�.');" &_
			    					"	self.history.back();" &_
			    					"</script>"
			    	dbget.close() : response.end
				end if
				rsget.close
			end if

			'/2016.07.06 �ѿ�� �߰�
			if requestCheckvar(Request("frontMakerid"),32)<>"" then
				sqlStr = "select top 1 userid" & vbcrlf
				sqlStr = sqlStr & " from db_user.dbo.tbl_user_c" & vbcrlf
				sqlStr = sqlStr & " where userid= '" & requestCheckvar(Request("frontMakerid"),32) & "'" & vbcrlf
	
				'response.write sqlStr & "<Br>"			
				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof Then
				Else
					dbget.RollBackTrans				'�ѹ�(�����߻���)
			    	Response.Write	"<script language=javascript>" &_
			    					"	alert('�Է��Ͻ� ǥ�� �귣�� �� �������� �ʽ��ϴ�.');" &_
			    					"	self.history.back();" &_
			    					"</script>"
			    	dbget.close() : response.end
				end if
				rsget.close
			end if

			'###########################################################################
			'��ǰ �⺻���� ����
			'###########################################################################

			sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
			sqlStr = sqlStr & " set itemdiv='" & Cstr(Request("itemdiv")) & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemname='" & chrbyte(html2db(itemname),64,"") & "'" & vbCrlf
			sqlStr = sqlStr & " ,lastupdate=getdate()"
			'' ��ü ���� �ڵ� �߰�
    		sqlStr = sqlStr & " ,upchemanagecode='" & html2db(Request("upchemanagecode")) & "'" & vbCrlf
			'' �ܵ���ǰ ���� �߰�
    		sqlStr = sqlStr & " ,tenOnlyYn='" & Request("tenOnlyYn") & "'" & vbCrlf
			sqlStr = sqlStr & " ,adultType=isnull('" & Request("adultType") & "', 0)" & vbCrlf
			'' ��ǰ���� �߰�(2009.04.03;������ �߰�)
    		sqlStr = sqlStr & " ,itemWeight='" & chrbyte(html2db(Request("itemWeight")),64,"") & "'" & vbCrlf
    		'' �귣�� �߰� (2015.09.15 ������ �߰�)
    		sqlStr = sqlStr & ", makerid ='"& requestCheckvar(Request("makerid"),32)&"'"&vbCrlf
			'' ǥ�ú귣�� �߰�(2012.03.05;������ �߰�)
			sqlStr = sqlStr & " ,frontMakerid='" & chkIIF(requestCheckvar(Request("frontMakerid"),32)<>"",requestCheckvar(Request("frontMakerid"),32),requestCheckvar(Request("makerid"),32)) & "'" & vbCrlf
    		'' �ܵ�(����)���Ż�ǰ (2012.03.26;������ �߰�)
    		sqlStr = sqlStr & " ,reserveItemTp='" & reserveItemTp & "'" & vbCrlf

			sqlStr = sqlStr & " where itemid=" & CStr(realitemid) & "" + vbCrlf
 
			dbget.execute(sqlStr)
			vChangeContents = vChangeContents & "- ��ǰ�� : itemname = " & chrbyte(html2db(Request("itemname")),64,"") & vbCrLf
			vChangeContents = vChangeContents & "- �ܵ�(����)���� : reserveItemTp = " & reserveItemTp & vbCrLf

			sqlStr = "update [db_item].[dbo].tbl_item_Contents" + vbCrlf
			sqlStr = sqlStr & " set itemcontent='" & html2db(Request("itemcontent")) & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemsource='" & chrbyte(html2db(Request("itemsource")),128,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,itemsize='" & chrbyte(html2db(Request("itemsize")),128,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,sourcearea='" & chrbyte(html2db(Request("sourcearea")),128,"") & "'" + vbCrlf 
    	sqlStr = sqlStr & " ,sourcekind ='" & Request("rdArea") & "'" & vbCrlf 	''������ ��ǰ�� �߰�(2017.01.02 ������)
			sqlStr = sqlStr & " ,makername='" & chrbyte(html2db(Request("makername")),64,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,keywords='" & chrbyte(html2db(Request("keywords")),500,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,usinghtml='" & Request("usinghtml") & "'" + vbCrlf
			sqlStr = sqlStr & " ,ordercomment='" & chrbyte(html2db(Request("ordercomment")),1024,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,designercomment='" & chrbyte(html2db(Request("designercomment")),255,"") & "'" + vbCrlf
			sqlStr = sqlStr & " ,requireMakeDay='" &  Request("requireMakeDay") & "'" + vbCrlf
    		'' ��ǰǰ��,������������ (2012.10.19;������ �߰�)
    		sqlStr = sqlStr & " ,infoDiv='" & Request("infoDiv") & "'" & vbCrlf
    		sqlStr = sqlStr & " ,safetyYn='" & Request("safetyYn") & "'" & vbCrlf
    		'sqlStr = sqlStr & " ,safetyDiv='" & Request("safetyDiv") & "'" & vbCrlf
    		'sqlStr = sqlStr & " ,safetyNum='" & chrbyte(html2db(Request("safetyNum")),25,"") & "'" & vbCrlf 
			'' ISBN ����
			sqlStr = sqlStr & " ,isbn13='" & chrbyte(html2db(Request("isbn13")),13,"") & "'" & vbCrlf
			sqlStr = sqlStr & " ,isbn10='" & chrbyte(html2db(Request("isbn10")),10,"") & "'" & vbCrlf
			sqlStr = sqlStr & " ,isbn_sub='" & chrbyte(html2db(Request("isbn_sub")),5,"") & "'" & vbCrlf
			sqlStr = sqlStr & " where itemid=" & CStr(realitemid) & "" + vbCrlf

	        dbget.execute(sqlStr)

			'// 2016.2.16 �ű��߰� ��ǰ�󼼼��� ������ �߰� - ������
			'// 2016-12-13  iframe ���� ��� - iframe ���� ����
			'// ������ ������ �� ���Խ����� src, width, height�� �̾Ƴ�
			If Trim(request("itemvideo")) <> "" Then
				Dim itemvideo, RetStr, RetSrc, RetWidth, RetHeight, regEx, Matches, Match, VideoTempSrc, VideoTempWidth, VideoTempHeight, videoType, dbsql
				itemvideo = request("itemvideo")
				'// 2016-12-13 �߰� iframe ���� �ּҸ� �Ѿ� �ð��
				If InStr(itemvideo ,"iframe") > 0 Then
				Else 
					'// ���� ��ȯ �� �⺻�� (���������� ��޿�����)
					If InStr(itemvideo , "youtu.be")>0 Then
						itemvideo = "<iframe width=""560"" height=""315"" src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" frameborder=""0"" allowfullscreen></iframe>"
					ElseIf InStr(itemvideo, "vimeo")>0 Then
						itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://vimeo.com/","https://player.vimeo.com/video/")) &""" width=""640"" height=""360"" frameborder=""0"" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"
					End If
				End If 

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

					VideoTempWidth =  Mid(Match.Value, InStrRev(Match.Value,"width='")+7)
					RetWidth = Left(VideoTempWidth, InStr(VideoTempWidth, "'")-1)

					VideoTempHeight =  Mid(Match.Value, InStrRev(Match.Value,"height='")+8)
					RetHeight = Left(VideoTempHeight, InStr(VideoTempHeight, "'")-1)
				Next
				Set regEx = Nothing
				Set Matches = Nothing

				sqlStr = " select idx FROM db_item.dbo.tbl_item_videos WHERE videogubun='video1' And itemid =" + CStr(realitemid)  
				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof Then
					If Not(videoType="etc") Then
						'// �����Ͱ� �ִٸ� ������Ʈ ����.
						dbsql = "update [db_item].[dbo].tbl_item_videos" + vbCrlf
						dbsql = dbsql & " set videourl='" &RetSrc& "'" + vbCrlf
						dbsql = dbsql & " ,videowidth='" & RetWidth & "'" + vbCrlf
						dbsql = dbsql & " ,videoheight='" & RetHeight & "'" + vbCrlf
						dbsql = dbsql & " ,videotype='" & videoType & "'" + vbCrlf
						dbsql = dbsql & " ,videofullurl='" & chrbyte(html2db(itemvideo),255,"") & "'" + vbCrlf
						dbsql = dbsql & " ,modifydate=getdate()" + vbCrlf
						dbsql = dbsql & " where idx='"&rsget("idx")&"' And itemid='" & CStr(realitemid) & "'" + vbCrlf
						dbget.execute(dbsql)
					End If
				Else
					If Not(videoType="etc") Then
						'// �����Ͱ� ������ �μ�Ʈ ����.
						dbsql = " insert into [db_item].[dbo].tbl_item_videos (itemid, videogubun, videotype, videourl, videowidth, videoheight, videofullurl, regdate) values " + vbCrlf
						dbsql = dbsql & " ('"&CStr(realitemid)&"', 'video1', '"&videoType&"', '"&RetSrc&"', '"&RetWidth&"', '"&RetHeight&"','"&chrbyte(html2db(itemvideo),255,"")&"', getdate()) " + vbCrlf
						dbget.execute(dbsql)
					End If
				end if
				rsget.close
			Else
				'// �ƹ����� �ȳѾ�Դµ� db�� ���� ������ ������� �Ǵ�. ������.
				sqlStr = " select idx FROM db_item.dbo.tbl_item_videos WHERE videogubun='video1' And itemid =" + CStr(realitemid)  
				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof Then
					dbsql = " Delete from [db_item].[dbo].tbl_item_videos Where videogubun='video1' And itemid=" + CStr(realitemid)
					dbget.execute(dbsql)
				End If
				rsget.close
			End If

			'// ���� ��ǰ �ֱ� //
			If (Request("relateItems")<>"") Then
				sqlStr = "delete from db_item.dbo.tbl_item_relation Where mainItemid='" & realitemid & "';" & vbCrLf
				sqlStr = sqlStr & " Insert into db_item.dbo.tbl_item_relation (mainItemid, subItemid) "
				sqlStr = sqlStr & " Select '" & realitemid & "', itemid "
				sqlStr = sqlStr & " From db_item.dbo.tbl_item "
				sqlStr = sqlStr & " Where itemid in (" & Request("relateItems") & ")"
				dbget.execute(sqlStr)
				
				vChangeContents = vChangeContents & "- ������ǰ��� : itemid = " & Request("relateItems") & vbCrLf
			end if

			'// ����ī�װ� �ֱ� //
			sqlStr = "delete from db_item.dbo.tbl_display_cate_item Where itemid='" & realitemid & "';" & vbCrLf
			If (Request("catecode").Count>0) Then
				sqlStr = sqlStr & "update db_item.dbo.tbl_item set dispcate1=null Where itemid='" & realitemid & "';" & vbCrLf
				vChangeContentsCa = "- ����ī�װ� : "
				for i=1 to Request("catecode").Count
					sqlStr = sqlStr & "Insert into db_item.dbo.tbl_display_cate_item (catecode, itemid, depth, sortNo, isDefault) values "
					sqlStr = sqlStr & "('" & Request("catecode")(i) & "'"
					sqlStr = sqlStr & ",'" & realitemid & "'"
					sqlStr = sqlStr & ",'" & Request("catedepth")(i) & "',9999"
					sqlStr = sqlStr & ",'" & Request("isDefault")(i) & "');" & vbCrLf
					
					vChangeContentsCa = vChangeContentsCa & Request("catecode")(i) & ","
					if Request("isDefault")(i)="y" then
						sqlStr = sqlStr & "update db_item.dbo.tbl_item set dispcate1='" & left(Request("catecode")(i),3) & "' Where itemid='" & realitemid & "';" & vbCrLf
						vChangeContentsCa = vChangeContentsCa & " �⺻���� : " & left(Request("catecode")(i),3) & ","
					end if
				next
			end if
			dbget.execute(sqlStr)

			'// ��ǰ�Ӽ� �ֱ� //
			If (Request("attribCd").Count>0) Then
				sqlStr = "delete from db_item.dbo.tbl_itemAttrib_item Where itemid='" & realitemid & "';" & vbCrLf
				for i=1 to Request("attribCd").Count
					sqlStr = sqlStr & "Insert into db_item.dbo.tbl_itemAttrib_item (attribCd, itemid) values "
					sqlStr = sqlStr & "('" & Request("attribCd")(i) & "'"
					sqlStr = sqlStr & ",'" & realitemid & "')" & vbCrLf
				next
				dbget.execute(sqlStr)
			end if
			
			
			'' ��ǰ�Ӽ� ���Ӹ� 2015/03/10 �߰�
            sqlStr = "exec db_item.[dbo].[sp_Ten_KSearch_Attribute_Summary] '"&realitemid&"'"& vbCrLf
            dbget.execute(sqlStr)

			'//��ǰ ǰ�������� ����
			if Request("infoDiv")<>"" then
				dim infoCd, infoCont, infoChk, infoType
	
				'�迭�� ó��
				redim infoCd(Request("infoCd").Count)
				redim infoCont(Request("infoCont").Count)
				redim infoChk(Request("infoChk").Count)
				redim infoType(Request("infoType").Count)

				for i=1 to Request("infoCd").Count
					infoCd(i) = Request("infoCd")(i)
					infoCont(i) = Request("infoCont")(i)
					infoChk(i) = Request("infoChk")(i)
					infoType(i) = Request("infoType")(i)
				next
	
				'������ ����
				sqlStr = "Delete From db_item.dbo.tbl_item_infoCont Where itemid='" & CStr(realitemid) & "'"
				dbget.execute(sqlStr)

				dim regEx_infoCont, infoContresult

				'DB�� ó��
				for i=1 to ubound(infoCd)
					'�Է°��� �ִ� ��츸 ����
					if infoChk(i)<>"" or infoCont(i)<>"" then
						if infoType(i)="A" then
							if infoCont(i)="" or isnull(infoCont(i)) then
								dbget.RollBackTrans				'�ѹ�(�����߻���)
								Response.Write "<script type='text/javascript'>alert('A/S å����/��ȭ��ȣ�� �Է��� �ּ���.');self.history.back();</script>"
								dbget.close()	:	response.End
							else
								Set regEx_infoCont = New RegExp

								With regEx_infoCont
									.Pattern = "([0-9]+)-([0-9]+)-([0-9]+)"
									.IgnoreCase = True
									.Global = True
								End With
								infoContresult = regEx_infoCont.Replace(infoCont(i),"$1-***-***")
								Set regEx_infoCont = nothing

								if instr(infoContresult,"010")>0 or instr(infoContresult,"011")>0 or instr(infoContresult,"016")>0 or instr(infoContresult,"017")>0 or instr(infoContresult,"018")>0 or instr(infoContresult,"019")>0 then
									dbget.RollBackTrans				'�ѹ�(�����߻���)
									Response.Write "<script type='text/javascript'>alert('A/S å����/��ȭ��ȣ ���� ��ǰ�󼼿� ǥ�õǴ� ������ �޴�����ȣ�� �Է��� �Ұ� �մϴ�.');self.history.back();</script>"
									dbget.close()	:	response.End
								end if
							end if
						end if

						sqlStr = "Insert into db_item.dbo.tbl_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
						sqlStr = sqlStr & "('" & CStr(realitemid) & "'"
						sqlStr = sqlStr & ",'" & CStr(infoCd(i)) & "'"
						sqlStr = sqlStr & ",'" & CStr(infoChk(i)) & "'"
						sqlStr = sqlStr & ",'" & html2db(infoCont(i)) & "')"
						dbget.execute(sqlStr)
					end if
				next
			end if

			'###########################################################################
			' �������� ���� ����
			'###########################################################################
			Dim vSafetyYN, vSafetyDiv, vSafetyNum, vSafetyIdx, vQuery, vTmpSafetyNum, vSafetyDeleteNum, vSafetyDeleteDiv
			vSafetyYN = requestCheckVar(Request.Form("safetyYn"),1)
			vSafetyDiv = Split(Replace(Request.Form("real_safetydiv")," ",""),",")
			vSafetyNum = Split(Replace(Request.Form("real_safetynum")," ",""),",")
			vSafetyIdx = Replace(Request.Form("real_safetyidx")," ","")
			vSafetyDeleteNum = Split(Replace(Request.Form("real_safetynum_delete")," ",""),",")
			vSafetyDeleteDiv = Split(Replace(Request.Form("real_safetydiv_delete")," ",""),",")

			dim pattern0, pattern1, pattern2, pattern3, pattern4, pattern5, pattern6
			pattern0 = "[^��-�R]"  '�ѱ�üũ
			pattern1 = "[^-0-9 ]"  '����üũ
			pattern2 = "[^-a-zA-Z]"  '����üũ
			pattern3 = "[^-��-�Ra-zA-Z0-9/ ]" '���ڿ� ���� �ѱ۸�
			pattern4 = "<[^>]*>"   '�±�üũ
			pattern5 = "[^-a-zA-Z0-9/ ]"    '���� ���ڸ�
			pattern6 = "[^A-Za-z0-9\-]"	'����, ����, �����¸�

			If vSafetyYN = "Y" Then
				If Request.Form("real_safetynum_delete") <> "" Then
					'### �����Ұ� ������ ���� ����.
					For i = LBound(vSafetyDeleteNum) To UBound(vSafetyDeleteNum)
						vQuery = "Delete from db_item.[dbo].[tbl_safetycert_tenReg] "
						vQuery = vQuery & "where itemid = '" & CStr(realitemid) & "' and certNum = '" & trim(vSafetyDeleteNum(i)) & "'; "
						vQuery = vQuery & "Delete from db_item.[dbo].[tbl_safetycert_info] "
						vQuery = vQuery & "where itemid = '" & CStr(realitemid) & "' and certNum = '" & trim(vSafetyDeleteNum(i)) & "'; "
						vQuery = vQuery & "Delete from db_item.[dbo].[tbl_safetycert_image] "
						vQuery = vQuery & "where itemid = '" & CStr(realitemid) & "' and certNum = '" & trim(vSafetyDeleteNum(i)) & "'; "

						dbget.execute(vQuery)
					Next
				End If

				'### �߰��Ǵ°� ������ �߰�
				For i = LBound(vSafetyDiv) To UBound(vSafetyDiv)
					If InStr(Request.Form("real_safetydiv_delete"), trim(vSafetyDiv(i))) < 1 Then
						if trim(vSafetyNum(i))<>"" then
							if chkWord(trim(vSafetyNum(i)),pattern6) = False then
								dbget.RollBackTrans				'�ѹ�(�����߻���)
								Response.Write "<script language=javascript>alert('���� ������ȣ���� ����,����,�����¸� �Է��ϽǼ� �ֽ��ϴ�.');self.history.back();</script>"
								dbget.close()	:	response.End
							end if
						end if

						' vQuery = "select" & vbcrlf
						' vQuery = vQuery & " itemid" & vbcrlf
						' vQuery = vQuery & " from db_item.[dbo].[tbl_safetycert_tenReg] with (nolock)" & vbcrlf
						' vQuery = vQuery & " where itemid = '" & CStr(realitemid) & "' and safetyDiv = '" & trim(vSafetyDiv(i)) & "'" & vbcrlf

						' 'response.write vQuery & "<br>"
						' rsget.CursorLocation = adUseClient
						' rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly
						' if Not rsget.Eof then
						' 	dbget.RollBackTrans				'�ѹ�(�����߻���)
						' 	Response.Write	"<script type='text/javascript'>" & vbcrlf
						' 	Response.Write	"	alert('�̹� �������������� �����Ǿ� �ֽ��ϴ�.');" & vbcrlf
						' 	Response.Write	"	self.history.back();" & vbcrlf
						' 	Response.Write	"</script>" & vbcrlf
						' 	rsget.Close : dbget.close()	: response.End				
						' end if
						' rsget.Close

						vQuery = "IF NOT EXISTS(select itemid from db_item.[dbo].[tbl_safetycert_tenReg] where itemid = '" & CStr(realitemid) & "' and certNum = '" & trim(vSafetyNum(i)) & "' and safetyDiv = '" & trim(vSafetyDiv(i)) & "') " & vbCrLf
						vQuery = vQuery & "BEGIN " & vbCrLf
						vQuery = vQuery & "INSERT INTO db_item.[dbo].[tbl_safetycert_tenReg](itemid, certNum, safetyDiv) "
						vQuery = vQuery & "VALUES('" & CStr(realitemid) & "', '" & trim(vSafetyNum(i)) & "', '" & trim(vSafetyDiv(i)) & "') " & vbCrLf
						vQuery = vQuery & "END " & vbCrLf
						
						dbget.execute(vQuery)

						vTmpSafetyNum = vTmpSafetyNum & "'" & vSafetyNum(i) & "',"
					End If
				Next

				If vSafetyIdx <> "" Then
					'### db_temp.[dbo].[tbl_safetycert_info] ����
					vQuery = "INSERT INTO db_item.[dbo].[tbl_safetycert_info](itemid,certUid,certOrganName,certNum,certState,certDiv,certDate,certChgDate,certChgReason,"
					vQuery = vQuery & " firstCertNum,productName,brandName,modelName,categoryName,importDiv,makerName,makerCntryName,importerName) " & vbCrLf
					vQuery = vQuery & " 	SELECT '" & CStr(realitemid) & "', sit.certUid,sit.certOrganName,sit.certNum,sit.certState,sit.certDiv,sit.certDate" & vbCrLf
					vQuery = vQuery & " 	,sit.certChgDate,sit.certChgReason,sit.firstCertNum,sit.productName,sit.brandName,sit.modelName,sit.categoryName" & vbCrLf
					vQuery = vQuery & " 	,sit.importDiv,sit.makerName,sit.makerCntryName,sit.importerName "
					vQuery = vQuery & " 	From db_temp.[dbo].[tbl_safetycert_info_temp] sit" & vbCrLf
					vQuery = vQuery & " 	left join db_item.[dbo].[tbl_safetycert_info] si" & vbCrLf
					vQuery = vQuery & " 		on sit.certNum = si.certNum" & vbCrLf
					vQuery = vQuery & " 		and si.itemid = "& realitemid &"" & vbCrLf
					vQuery = vQuery & " 	WHERE si.itemid is null and sit.idx in(" & vSafetyIdx & ")"

					'response.write vQuery & "<Br>"
					dbget.execute(vQuery)

					vQuery = ""

					'### db_temp.[dbo].[tbl_safetycert_image] ����
					If Right(vTmpSafetyNum,1) = "," Then
						vTmpSafetyNum = Left(vTmpSafetyNum, Len(vTmpSafetyNum)-1)
					End If
					
					vQuery = "INSERT INTO db_item.[dbo].[tbl_safetycert_image](itemid,certNum,ImageUrls) " & vbCrLf
					vQuery = vQuery & " 	SELECT '" & CStr(realitemid) & "', sit.certNum, sit.ImageUrls" & vbCrLf
					vQuery = vQuery & " 	From db_temp.[dbo].[tbl_safetycert_image_temp] sit" & vbCrLf
					vQuery = vQuery & " 	left join db_item.[dbo].[tbl_safetycert_image] si" & vbCrLf
					vQuery = vQuery & " 		on sit.certNum = si.certNum" & vbCrLf
					vQuery = vQuery & " 		and si.itemid = "& realitemid &"" & vbCrLf
					vQuery = vQuery & " 	WHERE si.itemid is null and sit.topidx in(" & vSafetyIdx & ") and sit.certNum in(" & vTmpSafetyNum & ")"

					'response.write vQuery & "<Br>"
					dbget.execute(vQuery)

					vQuery = ""

					vQuery = "DELETE From db_temp.[dbo].[tbl_safetycert_info_temp] WHERE idx in(" & vSafetyIdx & "); "
					vQuery = vQuery & "DELETE From db_temp.[dbo].[tbl_safetycert_image_temp] WHERE topidx in(" & vSafetyIdx & ") and certNum in(" & vTmpSafetyNum & ")"
					dbget.execute(vQuery)

					vQuery = ""

				End If
			Else
				'### �������ƴϰų� ���� ǥ��� ���� ������ ����.
				vQuery = "Delete from db_item.[dbo].[tbl_safetycert_tenReg] where itemid = '" & CStr(realitemid) & "'; "
				vQuery = vQuery & "Delete from db_item.[dbo].[tbl_safetycert_info] where itemid = '" & CStr(realitemid) & "'; "
				vQuery = vQuery & "Delete from db_item.[dbo].[tbl_safetycert_image] where itemid = '" & CStr(realitemid) & "'; "
				dbget.execute(vQuery)
				vQuery = ""
			End If

			'//�����ǰ�� ����
			if html2db(Request("itemnameEng"))<>"" then
				sqlstr = "IF NOT EXISTS(select itemid from db_item.dbo.tbl_item_multiLang where itemid='" & realitemid & "' and countryCd='EN') " & vbCrLf
				sqlstr = sqlstr & " BEGIN "
				sqlstr = sqlstr & "INSERT INTO db_item.dbo.tbl_item_multiLang (itemid,countryCd,itemname,itemcopy,itemContent,useyn,regdate,lastupdate) "
				sqlstr = sqlstr & " VALUES(" & realitemid & ", 'EN', N'" & chrbyte(html2db(Request("itemnameEng")),64,"") & "','','','Y',getdate(),getdate()) "
				sqlstr = sqlstr & " END " & vbCrLf
				sqlstr = sqlstr & " ELSE " & vbCrLf
				sqlstr = sqlstr & " BEGIN "
				sqlstr = sqlstr & "Update db_item.dbo.tbl_item_multiLang "
				sqlstr = sqlstr & " Set "
				sqlstr = sqlstr & " itemname = N'" & chrbyte(html2db(Request("itemnameEng")),64,"") & "'"
				sqlstr = sqlstr & " Where itemid=" & CStr(realitemid)
				sqlstr = sqlstr & "		and countryCd='EN'"
				sqlstr = sqlstr & " END " & vbCrLf
				''�ϴ� �������� ������ ���� (�Է�/������ �߰����� �˾������� ���)
				''dbget.execute(sqlStr)
			end if

		Case "ItemPriceInfo"
			'###########################################################################
			'��ǰ �Ǹ�/�������� ����
			'###########################################################################

			'// ���� ���� ����
	        dim sailprice, sailsuplycash, orgprice, orgsuplycash, sellcash, buycash
	        dim orgSellyn, orgsellSTDate
	        
	         sqlStr = " select sellyn, sellSTDate FROM db_item.dbo.tbl_item WHERE itemid =" + CStr(realitemid)  
            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
            	orgSellyn       = rsget("sellyn") 
            	orgsellSTDate   = rsget("sellSTDate") 
            end if
            rsget.close
	
			if Request("sailyn")= "Y" then
				sailprice = Request("sailprice")
				sailsuplycash = Request("sailsuplycash")
				orgprice = Request("sellcash")
				orgsuplycash = Request("buycash")
				sellcash = Request("sailprice")
				buycash = Request("sailsuplycash")
			else
				sailprice = Request("sailprice")
				sailsuplycash = Request("sailsuplycash")
				orgprice = Request("sellcash")
				orgsuplycash = Request("buycash")
				sellcash = Request("sellcash")
				buycash = Request("buycash")
			end if
            
            ''//��ۺ� ��å ** 
            if (request("mwdiv")="U") then
                ''��ü ����� ��� ��ü�� ��ۺ� �ΰ��Ǵ� ��������� �ƴϸ� 2 - ����⺻
                if (deliveryType<>"9") and (deliveryType<>"7") and (deliveryType<>"6") then
                    deliveryType = "2"
                end if
            else
                ''��ü ����� �ƴѰ�� ������ �Ǵ� ��������� �ƴϸ� 1 - �ٹ�⺻
                if (deliveryType<>"4") and (deliveryType<>"6") then
                    deliveryType = "1"
                end if
            end if
        
        
			'// ��ǰ ������ �Է� //
			sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
			sqlStr = sqlStr & " set sellcash=" & Cstr(sellcash) & "" + vbCrlf
			sqlStr = sqlStr & " ,buycash=" & Cstr(buycash) & "" + vbCrlf
			sqlStr = sqlStr & " ,mileage=" & Request("mileage") & "" + vbCrlf
			sqlStr = sqlStr & " ,vatinclude='" & Request("vatinclude") & "'" + vbCrlf
			sqlStr = sqlStr & " ,sellyn='" & Request("sellyn") & "'" + vbCrlf
			sqlStr = sqlStr & " ,isusing='" & Request("isusing") & "'" + vbCrlf

    		IF (reipgodate<>"") then
    		    sqlStr = sqlStr & " ,reipgodate='" & CStr(reipgodate) & "'" + vbCrlf
    		ELSE
    		    sqlStr = sqlStr & " ,reipgodate=NULL"  + vbCrlf
    		END if
    		
			sqlStr = sqlStr & " ,sailyn='" & Request("sailyn") & "'" + vbCrlf
			sqlStr = sqlStr & " ,sailprice=" & Cstr(sailprice) & "" + vbCrlf
			sqlStr = sqlStr & " ,sailsuplycash=" & Cstr(sailsuplycash) & "" + vbCrlf
			sqlStr = sqlStr & " ,orgprice=" & Cstr(orgprice) & "" + vbCrlf
			sqlStr = sqlStr & " ,orgsuplycash=" & Cstr(orgsuplycash) & "" + vbCrlf
			sqlStr = sqlStr & " ,deliverarea='" & Request("deliverarea") & "'" + vbCrlf
			sqlStr = sqlStr & " ,deliverfixday='" & Request("deliverfixday") & "'" + vbCrlf
			if Request("deliverOverseas")="Y" then
				sqlStr = sqlStr & " ,deliverOverseas='Y' " + vbCrlf
			else
				sqlStr = sqlStr & " ,deliverOverseas='N' " + vbCrlf
			end if
			sqlStr = sqlStr & " ,mwdiv='" & Request("mwdiv") & "'" + vbCrlf
			sqlStr = sqlStr & " ,deliverytype='" & deliverytype & "'" + vbCrlf
			If deliveryType <> "1" Then		'### �ٹ谡 �ƴҰ�� ����(pojangok) �� N ���� ����.
				sqlStr = sqlStr & " ,pojangok='N'" + vbCrlf
			End If
			sqlStr = sqlStr & " ,availPayType='" & Request("availPayType") & "'" + vbCrlf
			sqlStr = sqlStr & " ,orderMinNum='" & Request("orderMinNum") & "'" + vbCrlf
			sqlStr = sqlStr & " ,orderMaxNum='" & Request("orderMaxNum") & "'" + vbCrlf
			sqlStr = sqlStr & " ,lastupdate=getdate()"
			  if orgSellyn <>"Y" and Request("sellyn")  ="Y" and isNull(orgsellSTDate) then
	        sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf        
	          end if
			sqlStr = sqlStr & " where itemid='" & realitemid & "'" + vbCrlf
			dbget.execute(sqlStr)
			
			vChangeContents = vChangeContents & "- �Һ��ڰ� : sellcash = " & Cstr(sellcash) & vbCrLf
			vChangeContents = vChangeContents & "- ���ް� : buycash = " & Cstr(buycash) & vbCrLf
			vChangeContents = vChangeContents & "- �Ǹſ��� : sellyn = " & Request("sellyn") & vbCrLf
			vChangeContents = vChangeContents & "- ��뿩�� : isusing = " & Request("isusing") & vbCrLf
			vChangeContents = vChangeContents & "- ���ο��� : sailyn = " & Request("sailyn") & vbCrLf
			vChangeContents = vChangeContents & "- ����Ư������ : mwdiv = " & Request("mwdiv") & vbCrLf
			vChangeContents = vChangeContents & "- ��۱��� : deliverarea = " & deliverytype & vbCrLf
			vChangeContents = vChangeContents & "- �ּ� �Ǹż� : orderMinNum = " & Request("orderMinNum") & vbCrLf
			vChangeContents = vChangeContents & "- �ִ� �Ǹż� : orderMaxNum = " & Request("orderMaxNum") & vbCrLf

			'// �߰� ���� �Է�
			sqlStr = "update [db_item].[dbo].tbl_item_Contents" + vbCrlf
			sqlStr = sqlStr & " set freight_min='" & getNumeric(Request("freight_min")) & "'" + vbCrlf
			sqlStr = sqlStr & " ,freight_max='" & getNumeric(Request("freight_max")) & "'" + vbCrlf
			sqlStr = sqlStr & " where itemid=" & CStr(realitemid) & "" + vbCrlf
	        dbget.execute(sqlStr)
	End Select


	'�귣�� �̸� �ֱ�
'	sqlStr =	"Update [db_item].[dbo].tbl_item Set " &_
'				"	 brandname=[db_user].[dbo].tbl_user_c.socname" &_
'				"		from [db_user].[dbo].tbl_user_c " &_
'				"		where [db_item].[dbo].tbl_item.itemid=" &  CStr(realitemid) &_
'				"			and [db_item].[dbo].tbl_item.makerid=[db_user].[dbo].tbl_user_c.userid"
    ''2012/03/26 ���� frontMakerid����
    sqlStr = " update I " &VbCRLF
    sqlStr = sqlStr&" set brandName=C.socname " &VbCRLF
    sqlStr = sqlStr&" from db_item.dbo.tbl_item I " &VbCRLF
    sqlStr = sqlStr&" 	Join [db_user].[dbo].tbl_user_c C " &VbCRLF
    sqlStr = sqlStr&" 	on C.userid=(CASE WHEN IsNULL(I.frontMakerid,'')='' THEN I.makerid ELSE I.frontMakerid END) " &VbCRLF
    sqlStr = sqlStr&" where I.itemid=" &  CStr(realitemid) 
    
	dbget.execute(sqlStr)

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
        
        sqlStr = " Delete From [db_item].dbo.tbl_Item_category " 
        sqlStr = sqlStr & " Where itemid=" & realitemid
        sqlStr = sqlStr & " and ((Convert(varchar(10),itemid) + code_div + code_large + code_mid + code_small)"
        sqlStr = sqlStr & "      not in (" & CateArr & ") "
        sqlStr = sqlStr & "      )"
        dbget.execute(sqlStr)
 
		for lp=0 to Ubound(NewDiv)
			if Trim(NewDiv(lp))="D" then
				sqlStr = " Update [db_item].[dbo].tbl_item Set "
				sqlStr = sqlStr & "	 cate_large='" & Trim(NewCd1(lp)) & "' " 
				sqlStr = sqlStr & "	 ,cate_mid='" & Trim(NewCd2(lp)) & "' "
				sqlStr = sqlStr & "	 ,cate_small='" & Trim(NewCd3(lp)) & "' " 
				sqlStr = sqlStr & " Where itemid=" & realitemid 
				sqlStr = sqlStr & " and (cate_large<>'" & Trim(NewCd1(lp)) & "' " 
				sqlStr = sqlStr & "     or cate_mid<>'" & Trim(NewCd2(lp)) & "' " 
				sqlStr = sqlStr & "     or cate_small<>'" & Trim(NewCd3(lp)) & "' " 
				sqlStr = sqlStr & " ) "

				dbget.execute(sqlStr)
			end if
			
			''���� ī�װ��� ���°�츸 �Է�
			sqlStr = "Insert into [db_item].dbo.tbl_Item_category "
			sqlStr = sqlStr & " (itemid,code_large,code_mid,code_small,code_div)  " 
			sqlStr = sqlStr & " select i.itemid" 
			sqlStr = sqlStr & " ,'" & Trim(NewCd1(lp)) & "'" 
			sqlStr = sqlStr & " ,'" & Trim(NewCd2(lp)) & "'" 
			sqlStr = sqlStr & " ,'" & Trim(NewCd3(lp)) & "'" 
			sqlStr = sqlStr & " ,'" & Trim(NewDiv(lp)) & "'"
			sqlStr = sqlStr & " from [db_item].dbo.tbl_Item i"
			sqlStr = sqlStr & "     left join [db_item].dbo.tbl_Item_category c"
			sqlStr = sqlStr & "     on i.itemid=c.itemid"
			sqlStr = sqlStr & "     and c.code_large='" & Trim(NewCd1(lp)) & "'" 
			sqlStr = sqlStr & "     and c.code_mid='" & Trim(NewCd2(lp)) & "'" 
			sqlStr = sqlStr & "     and c.code_small='" & Trim(NewCd3(lp)) & "'" 
			sqlStr = sqlStr & "     and c.code_div='" & Trim(NewDiv(lp)) & "'" 
			sqlStr = sqlStr & " where i.itemid=" & realitemid 
			sqlStr = sqlStr & " and c.itemid Is NULL"
			
			dbget.execute(sqlStr)
		next       
        
	end if

	'##### DB ���� ó�� #####
    If Err.Number = 0 Then
    	dbget.CommitTrans				'Ŀ��(����)
    	
    	'### ���� �α� ����(item)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'item', '" & realitemid & "', '" & Request("menupos") & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
    	
    	'### ���� �α� ����(dispcate)
    	If vChangeContentsCa <> "" Then
    		vChangeContentsCa = vChangeContentsCa & vbCrLf
	    	vSCMChangeSQLCa = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	    	vSCMChangeSQLCa = vSCMChangeSQLCa & "VALUES('" & session("ssBctId") & "', 'dispcate', '" & realitemid & "', '" & Request("menupos") & "', "
	    	vSCMChangeSQLCa = vSCMChangeSQLCa & "'" & vChangeContentsCa & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	    	dbget.execute(vSCMChangeSQLCa)
    	End If
    	
    	
    	Response.Write	"<script language=javascript>" &_
    					"	alert('�����͸� �����Ͽ����ϴ�.');" &_
    					"	opener.history.go(0);" &_
    					"	self.close();" &_
    					"</script>"
    Else
        dbget.RollBackTrans				'�ѹ�(�����߻���)
    	Response.Write	"<script language=javascript>" &_
    					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
    					"	self.history.back();" &_
    					"</script>"
    End If

        
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim mode, sqlStr

Dim pjt_code, ePCode, ptDepth, pjtgroup_desc, pjtgroup_sort, pjt_sortType, pjt_topImgUrl, pjtgroup_BGColor, pjtgroup_FontColor
Dim pjt_name, pjt_kind, pjt_gender, pjt_state, pjt_using, pjtgroup_code
Dim cnt
mode			= request("mode")
pjt_code		= request("pjt_code")
ePCode			= request("selPC")
pjtgroup_desc	= request("pjtgroup_desc")
pjtgroup_sort	= request("pjtgroup_sort")
pjtgroup_BGColor= request("pjtgroup_BGColor")
pjtgroup_FontColor = request("pjtgroup_FontColor")

pjt_name		= request("pjt_name")
pjt_kind		= request("pjt_kind")
pjt_gender		= request("pjt_gender")
pjt_state		= request("pjt_state")
pjt_using		= request("pjt_using")
pjtgroup_code	= request("pjtgroup_code")

pjt_sortType	= request("pjt_sortType")
pjt_topImgUrl	= request("ban")

Select Case mode
	Case "I"
'####### pjt_kind #######
'#		A ����			#
'#		B 100��			#
'#		C 1�ֳ�			#
'#		D ��ȥ�����	#
'#		E �߷�Ÿ�ε���	#
'#		F ȭ��Ʈ����	#
'#		G �����ε���	#
'#		H ũ��������	#
'#		I MD PICK		#
'#		J ETC ��ȹ��	#
'########################
		If pjt_kind <> "J" Then				'ETC ��ȹ���� �ƴ϶��
			If pjt_gender = "A" Then		'������ ��ü��� ƨ���..
				Call sbAlertMsg ("ETC ��ȹ���� ������ ��ü�� ��밡�� �մϴ�.",  "project_list.asp?menupos="&menupos, "self")
			Else							'������ ��ü�� �ƴҶ� �ش籸�п� ������ ������ �ִ��� Ȯ��, �ִٸ� ƨ���.
				sqlStr = ""
				sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_outmall.dbo.tbl_between_project "
				sqlStr = sqlStr & " WHERE pjt_kind = '"& pjt_kind &"' "
				sqlStr = sqlStr & " and pjt_gender = '"& pjt_gender &"' "
				rsCTget.Open sqlStr, dbCTget
				IF not (rsCTget.EOF or rsCTget.BOF) THEN
					cnt = rsCTget("cnt")
				END IF
				rsCTget.Close
				
				If cnt > 0 Then
					Call sbAlertMsg ("�̹� ����� �����Ͱ� �ֽ��ϴ�\n\n�����ؼ� ����ϼ���",  "project_list.asp?menupos="&menupos, "self")
				End If
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_between_project (pjt_name, pjt_kind, pjt_gender, pjt_state, pjt_using, adminid, pjt_sortType) VALUES "
		sqlStr = sqlStr & " ('"& pjt_name &"', '"& pjt_kind &"', '"& pjt_gender &"', '"& pjt_state &"', 'Y', '"& session("ssBctId") &"', '"& pjt_sortType &"') "
		dbCTget.execute sqlStr
		Call sbAlertMsg ("����Ǿ����ϴ�.",  "project_list.asp?menupos="&menupos, "self")
	Case "U"
		If pjt_kind <> "J" Then			'ETC ��ȹ���� �ƴ϶��
			If pjt_gender = "A" Then		'������ ��ü��� ƨ���..
				Call sbAlertMsg ("ETC ��ȹ���� ������ ��ü�� ��밡�� �մϴ�.",  "project_list.asp?menupos="&menupos, "self")
			Else							'������ ��ü�� �ƴҶ� �ش籸�п� ������ ������ �ִ��� Ȯ��, �ִٸ� ƨ���.
				sqlStr = ""
				sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_outmall.dbo.tbl_between_project "
				sqlStr = sqlStr & " WHERE pjt_kind = '"& pjt_kind &"' "
				sqlStr = sqlStr & " and pjt_gender = '"& pjt_gender &"' "
				rsCTget.Open sqlStr, dbCTget
				IF not (rsCTget.EOF or rsCTget.BOF) THEN
					cnt = rsCTget("cnt")
				END IF
				rsCTget.Close
				
				If cnt > 1 Then
					Call sbAlertMsg ("�̹� ����� �����Ͱ� �ֽ��ϴ�\n\n�����ؼ� ����ϼ���",  "project_list.asp?menupos="&menupos, "self")
				End If
			End If
		End If

		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_project SET "
		sqlStr = sqlStr & " pjt_name = '"&pjt_name&"' "
		sqlStr = sqlStr & " ,pjt_kind = '"&pjt_kind&"' "
		sqlStr = sqlStr & " ,pjt_topImgUrl = '"&pjt_topImgUrl&"' "
		sqlStr = sqlStr & " ,pjt_gender = '"&pjt_gender&"' "
		sqlStr = sqlStr & " ,pjt_state = '"&pjt_state&"' "
		sqlStr = sqlStr & " ,pjt_using = '"&pjt_using&"' "
		sqlStr = sqlStr & " ,pjt_sortType = '"&pjt_sortType&"' "
		sqlStr = sqlStr & " ,pjt_lastupdate = getdate() "
		sqlStr = sqlStr & " WHERE pjt_code = "&pjt_code 
		dbCTget.execute sqlStr
		Call sbAlertMsg ("����Ǿ����ϴ�.",  "project_list.asp?menupos="&menupos, "self")
	Case "GI"
		sqlStr = ""
		IF ePCode = "0" THEN
			sqlStr = sqlStr & " SELECT isnull(max(pjtgroup_depth),0) + 100 FROM db_outmall.dbo.tbl_between_project_group where pjt_code = "&pjt_code
		ELSE
			sqlStr = sqlStr & " SELECT isnull(max(pjtgroup_depth),0)+1 FROM db_outmall.dbo.tbl_between_project_group WHERE pjt_code = "&pjt_code&" and (pjtgroup_code = "& ePCode&" OR pjtgroup_pcode ="&ePCode&")"
		END IF
		rsCTget.Open sqlStr, dbCTget
		IF not (rsCTget.EOF or rsCTget.BOF) THEN
			ptDepth = rsCTget(0)
		END IF
		rsCTget.Close

		sqlStr = ""
		sqlStr = sqlStr & "INSERT INTO db_outmall.dbo.tbl_between_project_group (pjt_code, pjtgroup_desc, pjtgroup_sort, pjtgroup_pcode,pjtgroup_depth, regdate, pjtgroup_BGColor, pjtgroup_FontColor) "	&_
			" VALUES ("&pjt_code&", '"&pjtgroup_desc& "', '"&pjtgroup_sort&"', "&ePCode&", "&ptDepth&", getdate(), '"&pjtgroup_BGColor&"', '"&pjtgroup_FontColor&"')"
		dbCTget.execute sqlStr
		response.write "<script language='javascript'>alert('���� �Ǿ����ϴ�.'); opener.location.reload();window.close();</script>"

	CASE "GU"
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_project_group SET pjtgroup_desc ='"&pjtgroup_desc&"', pjtgroup_sort='"&pjtgroup_sort&"',"
		sqlStr = sqlStr & " pjtgroup_pcode = " & ePCode & ",pjtgroup_BGColor='"&pjtgroup_BGColor&"',pjtgroup_FontColor='"&pjtgroup_FontColor&"'"
		sqlStr = sqlStr & " WHERE pjtgroup_code = "&pjtgroup_code
		dbCTget.execute sqlStr
		response.write "<script language='javascript'>alert('���� �Ǿ����ϴ�.'); opener.location.reload();window.close();</script>"

	CASE "GD"	'�׷����
		Dim pGCode
		pGCode= Request("pGC")
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_project_group SET pjtgroup_using ='N' "
		sqlStr = sqlStr & "	WHERE pjtgroup_code = "&pGCode&" OR pjtgroup_pcode ="&pGCode
		dbCTget.execute sqlStr
		Call sbAlertMsg ("�����Ǿ����ϴ�.", "iframe_projectitem_group.asp?pjt_code="&pjt_code&"&menupos="&menupos, "self")
		dbCTget.close()	:	response.End
End Select

%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
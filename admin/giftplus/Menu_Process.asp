<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  ����Ʈ�÷���
' History : 2010.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%  
dim LCode , MCode , SCode , CodeNm, OrderNo , isUsing,guideoffimg ,Depth,mode
dim LCodeImgOn,LCodeImgOFF,MCodeTopImg ,GuideListImg,GuideTopImg , strSQL ,cnt , msg ,listtype
dim minvalue, maxvalue, i	
	Depth = request("Depth")
	mode = request("mode")
	LCode = request("LCode")
	MCode = request("MCode")
	SCode = request("SCode")
	CodeNm= html2db(request("CodeNm"))
	OrderNo= html2db(request("OrderNo"))	
	isUsing = request("isUsing")
	guideoffimg = request("guideoffimg")
	LCodeImgOn = request("LCodeImgOn")
	LCodeImgOFF = request("LCodeImgOFF")
	MCodeTopImg = request("MCodeTopImg")
	GuideListImg = request("GuideListImg")
	GuideTopImg = request("GuideTopImg")
	listtype = request("listtype")

	if guideoffimg<>"" then
		GuideListImg = guideoffimg
	end if

'// �޴� ����	
if mode = "edit" then

	if Depth = "L" then
		strSQL =" UPDATE [db_giftplus].[dbo].[tbl_giftplus_LMenu] " &_
				" SET LCodeNm='" & CodeNm &"'" &_
				" , OrderNo = '" & OrderNo & "'" &_
				" , isUsing='" & isUsing &"'" &_
				" WHERE LCode='" & LCode &"'"
	elseif Depth = "M" then
		strSQL =" UPDATE [db_giftplus].[dbo].[tbl_giftplus_MMenu] " &_
				" SET MCodeNm='" & CodeNm &"'" &_
				" , OrderNo = '" & OrderNo & "'" &_
				" , isUsing='" & isUsing &"'" &_
				" WHERE LCode='" & LCode &"'" &_
				" and MCode ='" & MCode &"'"
	elseif Depth = "S" then
		strSQL =" UPDATE [db_giftplus].[dbo].[tbl_giftplus_SMenu] " &_
				" SET SCodeNm='" & CodeNm &"'" &_
				" , OrderNo = '" & OrderNo & "'" &_
				" , isUsing='" & isUsing &"'" &_
				" WHERE LCode='" & LCode &"'" &_
				" and MCode ='" & MCode &"'" &_
				" and SCode ='" & SCode &"'"
	end if
	
	'// ī�װ� �ڵ庰 ����
	strSQL = strSQL &_
			" UPDATE [db_giftplus].[dbo].[tbl_giftplus_ViewMenu] " &_
			" SET OrderNo = '" & OrderNo & "' " &_
			" , isUsing='" & isUsing &"'" &_		
			" WHERE LCode ='" & LCode & "' " 
			
			IF MCode<>"" THEN
				strSQL=strSQL & " and MCode ='" & MCode & "' " 
			ELSE
				strSQL=strSQL & " and MCode is null " 
			END IF
			IF SCode<>"" THEN
				strSQL=strSQL & " and SCode ='" & SCode & "' "
			ELSE
				strSQL=strSQL & " and SCode is null " 
			END IF
				
	'// ���� �ڵ� ��ü ����
	strSQL = strSQL &_		
			" UPDATE [db_giftplus].[dbo].[tbl_giftplus_ViewMenu] " &_
			" SET " 
			
	if Depth = "L" then
		strSQL = strSQL & " LCodeNm ='" & CodeNm & "' " 	
	elseif Depth = "M" then
		strSQL = strSQL & " MCodeNm ='" & CodeNm & "' " 	
	elseif Depth = "S" then
		strSQL = strSQL & " SCodeNm ='" & CodeNm & "' " 	
	end if  

	if listtype<>"" then
		strSQL = strSQL & "	,listtype = '" & listtype & "' "
	end if						
	if LCodeImgOn<>"" then
		strSQL = strSQL & "	,LCodeImgOn = '" & LCodeImgOn & "' "
	end if
	if LcodeImgOff<>"" then
		strSQL = strSQL & "	,LcodeImgOff = '" & LcodeImgOff & "' "
	end if
	if MCodeTopImg<>"" then
		strSQL = strSQL & "	,MCodeTopImg = '" & MCodeTopImg & "' "
	end if
	if GuideListImg<>"" then
		strSQL = strSQL & "	,GuideListImg = '" & GuideListImg & "' "
	end if
	if GuideTopImg<>"" then
		strSQL = strSQL & "	,GuideTopImg = '" & GuideTopImg & "' "
	end if	
	IF isUsing<>"" Then
		strSQL = strSQL & "	,isUsing = '" & isUsing & "' "
	End if
	
	strSQL = strSQL & _
			" WHERE LCode ='" & LCode & "' " 
	
	IF MCode<>"" THEN
		strSQL=strSQL & " and MCode ='" & MCode & "' " 
	END IF
	
	IF SCode<>"" THEN
		strSQL=strSQL & " and SCode ='" & SCode & "' "
	END IF
						
	msg = "���� �Ǿ����ϴ�"

	msg = "�Է� �Ǿ����ϴ�"
	
	dbget.BeginTrans
	
	'response.write strSQL &"<br>"
	dbget.execute(strSQL)
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
	
		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "'); opener.location.reload(); self.close();"
		response.write	"</script>"
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
	
		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	self.close();" &_
					"</script>"	
	End If
	
'// �޴�����
elseif mode="del" then
	strSQL =" SELECT count(*) as count FROM [db_giftplus].[dbo].[tbl_giftplus_item] " &_
			" WHERE LCode='" & LCode &"'"

		IF MCode<>"" THEN
			strSQL= strSQL & " and MCode ='" & MCode &"'"
		END IF
		IF SCode<>"" THEN
			strSQL= strSQL & " and SCode ='" & SCode &"'"
		END IF

	rsget.open strSQL ,dbget,1

	if not rsget.eof then
		cnt = rsget("count")
	end if

	rsget.close

	if cnt >0 then
		response.write	"<script language='javascript'>"
		response.write	" alert('��ǰ�� �����ִ� ī�װ��� ���� �Ҽ� �����ϴ�.\nȮ���� �ٽ� �Է����ּ���.'); self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

			
	if Depth = "L" then
		strSQL = strSQL & _
				" UPDATE [db_giftplus].[dbo].[tbl_giftplus_LMenu] " &_
				" SET IsUsing='N' " &_
				" WHERE LCode='" & LCode &"'"				
	elseif Depth = "M" then
		strSQL = strSQL & _
		 		" UPDATE [db_giftplus].[dbo].[tbl_giftplus_MMenu] " &_
		 		" SET IsUsing='N' " &_
				" WHERE LCode='" & LCode &"'"
				
				IF MCode<>"" THEN
					strSQL=strSQL & " and MCode ='" & MCode & "' " 
				END IF
	elseif Depth = "S" then
		strSQL = strSQL & _
				" UPDATE [db_giftplus].[dbo].[tbl_giftplus_SMenu] " &_
				" SET IsUsing='N' " &_
				" WHERE LCode='" & LCode &"'" 
				IF MCode<>"" THEN
					strSQL=strSQL & " and MCode ='" & MCode & "' " 
				END IF
				
				IF SCode<>"" THEN
					strSQL=strSQL & " and SCode ='" & SCode & "' "
				END IF
	end if
			
	'// ī�װ� ���� ��ü ����
	strSQL = strSQL &_		
			" UPDATE [db_giftplus].[dbo].[tbl_giftplus_ViewMenu] " &_
			" SET IsUsing='N' " &_
			" WHERE LCode ='" & LCode & "' " 
			
	IF MCode<>"" THEN
		strSQL=strSQL & " and MCode ='" & MCode & "' " 
	END IF
	
	IF SCode<>"" THEN
		strSQL=strSQL & " and SCode ='" & SCode & "' "
	END IF

	msg = "���� �Ǿ����ϴ�"

	msg = "�Է� �Ǿ����ϴ�"
	
	dbget.BeginTrans
	
	'response.write strSQL &"<br>"
	dbget.execute(strSQL)
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
	
		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "'); opener.location.reload(); self.close();"
		response.write	"</script>"
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
	
		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	self.close();" &_
					"</script>"	
	End If
		
''// �˻� ���� ������Ʈ
elseif  mode="cashedit" then

	minvalue = split(request("minvalue"),",")
	maxvalue = split(request("maxvalue"),",")
	
	dim minCnt : minCnt = ubound(minvalue)
	dim maxCnt : maxCnt = ubound(maxvalue)
	
	if minCnt <> maxCnt then
		response.write	"<script language='javascript'>"
		response.write	" alert('ó���� ������ �߻��߽��ϴ�.'); history.go(-1);"
		response.write	"</script>"
		dbget.close()	:	response.End
	else
	
		strSQL =" DELETE [db_giftplus].[dbo].[tbl_giftplus_CashMenu]" &_
				" WHERE LCode ='" & LCode & "' " 
					
					IF MCode<>"" THEN
						strSQL=strSQL & " and MCode ='" & MCode & "' " 
					ELSE
						strSQL=strSQL & " and MCode is null " 
					END IF
					IF SCode<>"" THEN
						strSQL=strSQL & " and SCode ='" & SCode & "' "
					ELSE
						strSQL=strSQL & " and SCode is null " 
					END IF
		
			for i = 0 to minCnt  			
				strSQL = strSQL &_
						" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
						" VALUES ("
						IF LCode<>"" then
							strSQL = strSQL & "'" & LCode & "'"
						else
							strSQL = strSQL & " NULL"
						end if
						
						IF MCode<>"" then
							strSQL = strSQL & ", '" & MCode & "'"
						else
							strSQL = strSQL & ", NULL"
						end if
						
						IF SCode<>"" then
							strSQL = strSQL & ", '" & SCode & "'"
						else
							strSQL = strSQL & ", NULL"
						end if
						
						strSQL = strSQL &_
						" , '" & minvalue(i) & "'" &_
						" , '" & maxvalue(i) & "'" &_
						")"
			next
	end if
	
	msg = "�Է� �Ǿ����ϴ�"

	msg = "�Է� �Ǿ����ϴ�"
	
	dbget.BeginTrans
	
	'response.write strSQL &"<br>"
	dbget.execute(strSQL)
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
	
		response.write	"<script language='javascript'>"
		response.write	"	document.domain = '10x10.co.kr';"
		response.write	"	alert('" & msg & "'); parent.location.reload(); self.close();"
		response.write	"</script>"
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
	
		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	self.close();" &_
					"</script>"	
	End If					

else
'// �޴� �߰�

	if Depth = "L" then
		strSQL =	" SELECT count(*) as count FROM [db_giftplus].[dbo].[tbl_giftplus_LMenu] " &_
		" WHERE LCode='" & LCode &"'"	
	elseif Depth = "M" then	
		strSQL =	" SELECT count(*) as count FROM [db_giftplus].[dbo].[tbl_giftplus_MMenu] " &_
					" WHERE LCode='" & LCode &"'" &_
					" and MCode ='" & MCode &"'"
	elseif Depth = "S" then
		strSQL =	" SELECT count(*) as count FROM [db_giftplus].[dbo].[tbl_giftplus_SMenu] " &_
					" WHERE LCode='" & LCode &"'" &_
					" and MCode ='" & MCode &"'" &_
					" and SCode ='" & SCode &"'"						
	end if

	rsget.open strSQL ,dbget,1
	if not rsget.eof then
		cnt = rsget("count")
	end if

	rsget.close

	if cnt >0 then
		response.write	"<script language='javascript'>"
		response.write	" alert('�ߺ��� �޴��Դϴ�.\nī�װ� �ڵ带 Ȯ���� �ٽ� �Է����ּ���.'); self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Depth = "L" then
		strSQL ="INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_LMenu] (LCode,LCodeNm,OrderNo,isusing)" &_
				" VALUES ('" & LCode &"','" & CodeNm &"','" & OrderNo &"','Y') " &_
				
				"INSERT INTO  [db_giftplus].[dbo].[tbl_giftplus_ViewMenu] (LCode,LCodeNm,LCodeImgOn,LCodeImgOFF,OrderNo,isusing,ListType)" &_
				"VALUES ('" & LCode & "','" & CodeNm & "','" & LCodeImgOn & "','" & LCodeImgOFF & "','" & OrderNo & "','Y','" & listtype & "')" &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"',NULL,NULL, '',30000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"',NULL,NULL,30000 ,60000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"',NULL,NULL,60000 ,90000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"',NULL,NULL,90000 ,'') " 
	elseif Depth = "M" then	
		strSQL =" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_MMenu] (LCode,MCode,MCodeNm,OrderNo,isusing)" &_
				" VALUES ('" & LCode &"','" & MCode &"','" & CodeNm &"','" & OrderNo &"','Y') " &_
				
				" INSERT INTO  [db_giftplus].[dbo].[tbl_giftplus_ViewMenu] (LCode,MCode,LCodeNm,MCodeNm,MCodeTopImg,OrderNo,isusing)" &_
				" SELECT top 1 LCode,'" & MCode &"',LCodeNm,'" & CodeNm & "','" & MCodeTopImg & "','" & OrderNo & "','Y'" &_
				" FROM db_giftplus.dbo.tbl_giftplus_ViewMenu  " &_
				" WHERE Lcode='" & LCode & "' "&_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"',NULL, '',30000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"',NULL,30000 ,60000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"',NULL,60000 ,90000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"',NULL,90000 ,'') " 		
	elseif Depth = "S" then
		strSQL ="INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_SMenu] (LCode,MCode,SCode,SCodeNm,OrderNo,isusing)" &_
				" VALUES ('" & LCode &"','" & MCode &"','" & SCode & "','" & CodeNm &"','" & OrderNo &"','Y') " &_
				
				" INSERT INTO  [db_giftplus].[dbo].[tbl_giftplus_ViewMenu] (LCode,MCode,SCode,LCodeNm,MCodeNm,SCodeNm,GuideListImg,GuideTopImg,OrderNo,isusing)" &_
				" SELECT top 1 LCode,MCode,'" & SCode & "',LCodeNm,MCodeNm,'" & CodeNm & "','" & GuideListImg & "','" & GuideTopImg &"','" & OrderNo & "','Y'" &_
				" FROM db_giftplus.dbo.tbl_giftplus_ViewMenu  " &_
				" WHERE Lcode='" & LCode & "' " &_
				" and MCode='" & MCode & "'" &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"','" & SCode & "','' ,30000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"','" & SCode & "',30000 ,60000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"','" & SCode & "',60000 ,90000) " &_
				
				" INSERT INTO [db_giftplus].[dbo].[tbl_giftplus_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
				" VALUES ('" & LCode &"','" & MCode &"','" & SCode & "',90000 ,'') " 											
	end if

	msg = "�Է� �Ǿ����ϴ�"
	
	dbget.BeginTrans
	
	'response.write strSQL &"<br>"
	dbget.execute(strSQL)
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
	
		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "'); opener.location.reload(); self.close();"
		response.write	"</script>"
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
	
		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	self.close();" &_
					"</script>"	
	End If					
	
end if


%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%

dim Depth,mode

Depth = request("Depth")

mode = request("mode")


dim Mn_idx,MnM_code,MnS_code
dim Mn_codeNm,Mn_viewIDX,Mn_codeL_img

dim LCode , MCode , SCode , CodeNm, OrderNo , SortMethod ,ListType,isUsing

	LCode = request("LCode")
	MCode = request("MCode")
	SCode = request("SCode")

	CodeNm= html2db(request("CodeNm"))
	OrderNo= html2db(request("OrderNo"))
	SortMethod = request("SortMethod")
	ListType = request("ListType")
	isUsing = request("isUsing")
	'code = request("code")
	'viewidx = request("viewidx")
	'name = html2db(request("name"))
	'strType = request("strType")
	
	

dim LCodeImgOn,LCodeImgOFF,MCodeTopImg ,GuideListImg,GuideTopImg
LCodeImgOn = request("LCodeImgOn")
LCodeImgOFF = request("LCodeImgOFF")
MCodeTopImg = request("MCodeTopImg")
GuideListImg = request("GuideListImg")
GuideTopImg = request("GuideTopImg")


'response.write Depth & "," & LCode & "," & MCode & "," & SCode & "," & CodeNm & "," & OrderNo & "," & SortMethod & "," & ListType

'response.write LCodeImgOn & "," & LCodeImgOFF & "," & MCodeTopImg & "," & GuideListImg & "," & GuideTopImg
'dbget.close()	:	response.End
dim strSQL ,cnt , msg

if mode = "edit" then
'// �޴� ����

		SELECT CASE Depth
			CASE "L"
				strSQL =" UPDATE [db_giftManager].[dbo].[tbl_gift_LMenu] " &_
						" SET LCodeNm='" & CodeNm &"'" &_
						" , OrderNo = '" & OrderNo & "'" &_
						" , isUsing='" & isUsing &"'" &_
						" WHERE LCode='" & LCode &"'"
			CASE "M"
				strSQL =" UPDATE [db_giftManager].[dbo].[tbl_gift_MMenu] " &_
						" SET MCodeNm='" & CodeNm &"'" &_
						" , OrderNo = '" & OrderNo & "'" &_
						" , isUsing='" & isUsing &"'" &_
						" WHERE LCode='" & LCode &"'" &_
						" and MCode ='" & MCode &"'"
			CASE "S"
				strSQL =" UPDATE [db_giftManager].[dbo].[tbl_gift_SMenu] " &_
						" SET SCodeNm='" & CodeNm &"'" &_
						" , OrderNo = '" & OrderNo & "'" &_
						" , isUsing='" & isUsing &"'" &_
						" WHERE LCode='" & LCode &"'" &_
						" and MCode ='" & MCode &"'" &_
						" and SCode ='" & SCode &"'"
		END SELECT
				'// ī�װ� �ڵ庰 ����
				strSQL = strSQL &_
						" UPDATE [db_giftManager].[dbo].[tbl_gift_ViewMenu] " &_
						" SET SortMethod = '" & SortMethod & "' " &_
						" , OrderNo = '" & OrderNo & "' " &_
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
						" UPDATE [db_giftManager].[dbo].[tbl_gift_ViewMenu] " &_
						" SET " 
						
						SELECT CASE Depth
							CASE "L"
								strSQL = strSQL & " LCodeNm ='" & CodeNm & "' " 	
							CASE "M"
								strSQL = strSQL & " MCodeNm ='" & CodeNm & "' " 	
							CASE "S"
								strSQL = strSQL & " SCodeNm ='" & CodeNm & "' " 	
						END SELECT
						
						
						if ListType<>"" then
							strSQL = strSQL & "	,ListType = '" & ListType & "' "
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
						
						strSQL = strSQL & _
								" WHERE LCode ='" & LCode & "' " 
						
						IF MCode<>"" THEN
							strSQL=strSQL & " and MCode ='" & MCode & "' " 
						END IF
						
						IF SCode<>"" THEN
							strSQL=strSQL & " and SCode ='" & SCode & "' "
						END IF
						
		msg = "���� �Ǿ����ϴ�"

elseif mode="del" then

'// �޴�����

		strSQL =" SELECT count(*) as count FROM [db_giftManager].[dbo].[tbl_Gift_item] " &_
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

				
				
				strSQL = strSQL & _
						" DELETE [db_giftManager].[dbo].[tbl_gift_LMenu] " &_
						" WHERE LCode='" & LCode &"'"
				
				strSQL = strSQL & _
				 		" DELETE [db_giftManager].[dbo].[tbl_gift_MMenu] " &_
						" WHERE LCode='" & LCode &"'"
						
						IF MCode<>"" THEN
							strSQL=strSQL & " and MCode ='" & MCode & "' " 
						END IF
				
				strSQL = strSQL & _
						" DELETE [db_giftManager].[dbo].[tbl_gift_SMenu] " &_
						" WHERE LCode='" & LCode &"'" 
						IF MCode<>"" THEN
							strSQL=strSQL & " and MCode ='" & MCode & "' " 
						END IF
						
						IF SCode<>"" THEN
							strSQL=strSQL & " and SCode ='" & SCode & "' "
						END IF
				
				
				'// ī�װ� ���� ��ü ����
				strSQL = strSQL &_		
						" DELETE [db_giftManager].[dbo].[tbl_gift_ViewMenu] " &_
						" WHERE LCode ='" & LCode & "' " 
						
						IF MCode<>"" THEN
							strSQL=strSQL & " and MCode ='" & MCode & "' " 
						END IF
						
						IF SCode<>"" THEN
							strSQL=strSQL & " and SCode ='" & SCode & "' "
						END IF


		msg = "���� �Ǿ����ϴ�"
		
''// �˻� ���� ������Ʈ
elseif  mode="cashedit" then
	
		dim minvalue,maxvalue,i
	
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
		
			strSQL =" DELETE [db_giftManager].[dbo].[tbl_gift_CashMenu]" &_
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
							" INSERT INTO [db_giftManager].[dbo].[tbl_gift_CashMenu](LCode ,MCode ,SCode,MinCash ,MaxCash) " &_
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
else
'// �޴� �߰�

		SELECT CASE Depth
			CASE "L"
				strSQL =	" SELECT count(*) as count FROM [db_giftManager].[dbo].[tbl_gift_LMenu] " &_
							" WHERE LCode='" & LCode &"'"
			CASE "M"
				strSQL =	" SELECT count(*) as count FROM [db_giftManager].[dbo].[tbl_gift_MMenu] " &_
							" WHERE LCode='" & LCode &"'" &_
							" and MCode ='" & MCode &"'"
			CASE "S"
				strSQL =	" SELECT count(*) as count FROM [db_giftManager].[dbo].[tbl_gift_SMenu] " &_
							" WHERE LCode='" & LCode &"'" &_
							" and MCode ='" & MCode &"'" &_
							" and SCode ='" & SCode &"'"
		END SELECT

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

		SELECT CASE Depth

			CASE "L"
				strSQL ="INSERT INTO [db_giftManager].[dbo].[tbl_gift_LMenu] (LCode,LCodeNm,OrderNo)" &_
						" VALUES ('" & LCode &"','" & CodeNm &"','" & OrderNo &"') " &_
						
						"INSERT INTO  [db_giftManager].[dbo].[tbl_gift_ViewMenu] (LCode,LCodeNm,LCodeImgOn,LCodeImgOFF,ListType,SortMethod,OrderNo)" &_
						"VALUES ('" & LCode & "','" & CodeNm & "','" & LCodeImgOn & "','" & LCodeImgOFF & "','" & ListType & "','" & SortMethod & "','" & OrderNo & "')"
			CASE "M"
				strSQL =" INSERT INTO [db_giftManager].[dbo].[tbl_gift_MMenu] (LCode,MCode,MCodeNm,OrderNo)" &_
						" VALUES ('" & LCode &"','" & MCode &"','" & CodeNm &"','" & OrderNo &"') " &_
						
						" INSERT INTO  [db_giftManager].[dbo].[tbl_gift_ViewMenu] (LCode,MCode,LCodeNm,MCodeNm,MCodeTopImg,ListType,SortMethod,OrderNo)" &_
						" SELECT top 1 LCode,'" & MCode &"',LCodeNm,'" & CodeNm & "','" & MCodeTopImg & "',ListType,'" & SortMethod & "','" & OrderNo & "'" &_
						" FROM db_giftmanager.dbo.tbl_gift_ViewMenu  " &_
						" WHERE Lcode='60' "
			CASE "S"
				strSQL ="INSERT INTO [db_giftManager].[dbo].[tbl_gift_SMenu] (LCode,MCode,SCode,SCodeNm,OrderNo)" &_
						" VALUES ('" & LCode &"','" & MCode &"','" & SCode & "','" & CodeNm &"','" & OrderNo &"') " &_
						
						" INSERT INTO  [db_giftManager].[dbo].[tbl_gift_ViewMenu] (LCode,MCode,SCode,LCodeNm,MCodeNm,SCodeNm,GuideListImg,GuideTopImg,ListType,SortMethod,OrderNo)" &_
						" SELECT top 1 LCode,MCode,'" & SCode & "',LCodeNm,MCodeNm,'" & CodeNm & "','" & GuideListImg & "','" & GuideTopImg &"',ListType,'" & SortMethod & "','" & OrderNo & "'" &_
						" FROM db_giftmanager.dbo.tbl_gift_ViewMenu  " &_
						" WHERE Lcode='" & LCode & "' " &_
						" and MCode='" & MCode & "'"
		END SELECT
		msg = "�Է� �Ǿ����ϴ�"
		

end if
	
	'response.write strSQL
	
	'dbget.close()	:	response.End
	dbget.BeginTrans

	dbget.execute(strSQL)

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "'); opener.location.reload();self.close();"
		response.write	"</script>"
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	self.close();" &_
					"</script>"

	End If

%>

<script>
opener.refresh();
self.close();
</script>


<!-- #include virtual="/lib/db/dbclose.asp" -->
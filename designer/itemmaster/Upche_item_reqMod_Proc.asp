<%@ language=vbscript %>
<% option explicit %>
<%
'########################################################### 
' Description : ��ü��� ��ǰ�� ������û ó��
' History : 2014.03.19 ������ ��� 
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itemidarr, olditemname, itemname, etcstr
dim i,itemid
dim objCmd, returnValue, errID,ItemCount,chkReturnCount
Dim mode, menupos
Dim strResultMsg
Dim oldsellcash, oldbuycash, sellcash, buycash
Dim makerid, edtype
 
mode=requestCheckvar(request("hidM"),1)
menupos=requestCheckvar(request("menupos"),10)
itemidarr = ReplaceRequestSpecialChar(request("itemidarr")) 
olditemname= ReplaceRequestSpecialChar(request("olditemname"))
itemname= ReplaceRequestSpecialChar(request("itemname"))
oldsellcash= ReplaceRequestSpecialChar(request("oldsellcash"))
oldbuycash= ReplaceRequestSpecialChar(request("oldbuycash"))
sellcash= ReplaceRequestSpecialChar(request("sellcash"))
buycash= ReplaceRequestSpecialChar(request("buycash"))
etcstr=  requestCheckvar(request("etcStr"),64)
ItemCount=requestCheckvar(request("itemcount"),10)
makerid = session("ssBctID")
 
SELECT  CASE mode
CASE "N" '--��ǰ�� ������û
	itemid = split(itemidarr,"|,|")
	olditemname = split(olditemname,"|,|")
	itemname = split(itemname,"|,|")
	chkReturnCount = 0
	For i=0 To UBound(itemid)
		itemid(i) = Left(trim(itemid(i)),16)
		olditemname(i) = Left(trim(olditemname(i)),64)
		itemname(i) = Left(trim(itemname(i)),64)
		 
		if itemname(i) = "" then 
			Call Alert_return ("��ǰ���� ��ϵ��� �ʾҽ��ϴ�.")
	 		response.end
		end if
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_upche_itemedit_itemanmeInsert]("&trim(itemid(i))&", '"&trim(olditemname(i))&"' ,'"&trim(itemname(i))&"','"&etcstr&"','"&makerid&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		IF returnValue <> "1" THEN 
			if errID = "" THEN
				errID = itemid(i)
			else
			errID = errID+","+itemid(i)
			end if 
		ELSE
			chkReturnCount = chkReturnCount  + 1
		END IF	
	
	Next


 strResultMsg = "" 
 	IF errID <> "" THEN
 		strResultMsg = strResultMsg & "��ǰ�ڵ� ["&errID &"] �� ó�� �����߽��ϴ�.\n"
	END IF
	strResultMsg = strResultMsg & "�����Ͻ� [" & ItemCount &"]���� ��ǰ �� "& "["&chkReturnCount&"]���� ��ǰ�� ������û�Ǿ����ϴ�.\n ��ǰ���� ��û����� ��ü��ǰ����>>��ǰ����ó��������� Ȯ���ϼ���"
	Call Alert_move(strResultMsg, "/designer/itemmaster/upche_item_reqMod_itemname.asp?menupos="&menupos)
CASE "P" '--���� ������û
	itemid = split(itemidarr,",")
	oldsellcash = split(oldsellcash,",")
	sellcash = split(sellcash,",")
	oldbuycash = split(oldbuycash,",")
	buycash = split(buycash,",")
	chkReturnCount = 0
	For i=0 To UBound(itemid)
		itemid(i) = Left(trim(itemid(i)),16) 
		
	 '������� Ȯ��
	 if trim(sellcash(i)) = ""  then
	 	 Call Alert_return ("�ǸŰ��� ��ϵ��� �ʾҽ��ϴ�.")
	 response.end
	 end if
	 
	  if (Clng(trim(sellcash(i))) <=100 or   Clng(trim(buycash(i))) <=100 ) then
	 	 Call Alert_return ("�ǸŰ��� ���ް��� 100�� �̻� �����մϴ�.")
	 response.end
	 end if
	 
	  if Clng(trim(sellcash(i))) < Clng(trim(buycash(i))) then
	 	 Call Alert_return ("�ǸŰ��� ���ް����� ū ���ݸ� �����մϴ�." & sellcash(i)&"-" &buycash(i))
	 response.end
	 end if
	
	    ''2015/03/10 �߰�
		if (UBOUND(itemid)<>UBOUND(oldsellcash)) or (UBOUND(itemid)<>UBOUND(sellcash)) or (UBOUND(itemid)<>UBOUND(oldbuycash)) or (UBOUND(itemid)<>UBOUND(buycash)) then
		    Call Alert_return ("���� �Ķ���� ����-�����ڹ��� ���")
	        response.end
		end if
		
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_upche_itemedit_sellcashInsert]("&trim(itemid(i))&", '"&trim(oldsellcash(i))&"' ,'"&trim(oldbuycash(i))&"', '"&trim(sellcash(i))&"' ,'"&trim(buycash(i))&"','"&etcstr&"','"&makerid&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		IF returnValue <> "1" THEN 
			if errID = "" THEN
				errID = itemid(i)
			else
			errID = errID+","+itemid(i)
			end if
		ELSE
			chkReturnCount = chkReturnCount  + 1
		END IF	
	Next
 
 strResultMsg = "" 
 	IF errID <> "" THEN
 		strResultMsg = strResultMsg & "��ǰ�ڵ� ["&errID &"] �� ó�� �����߽��ϴ�.\n"
	END IF
	strResultMsg = strResultMsg & "�����Ͻ� [" & ItemCount &"]���� ��ǰ �� "& "["&chkReturnCount&"]���� ��ǰ�� ������û�Ǿ����ϴ�.\n ��ǰ���� ��û����� ��ü��ǰ����>>��ǰ����ó��������� Ȯ���ϼ���"
	Call Alert_move(strResultMsg, "/designer/itemmaster/upche_item_reqMod_sellprice.asp?menupos="&menupos)
CASE "C" '������û ��� 
itemidarr = left(itemidarr,16)
olditemname = left(olditemname,64)
 
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call db_temp.[dbo].[sp_Ten_upche_itemedit_cancel]("&itemidarr&",'"&makerid&"','"&olditemname&"','"&oldsellcash&"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
		
		IF returnValue <> 1 THEN
				Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.-error: case 'c' returnValue")
 		ELSE
 				Call Alert_move("��ǰ������û�� ��ҵǾ����ϴ�.", "/designer/itemmaster/upche_item_reqMod_result.asp?menupos="&menupos)
		END IF
CASE ELSE
		Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.-error: case else")
END SELECT	
%>
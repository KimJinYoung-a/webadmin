<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode
dim itemid, cksel, sellyn, isusing, mwdiv, deliveryType,dsellreservedate, danjongyn
dim ArrCnt, ChkCnt, preParam
dim limityn, orgLimityn
mode = request("mode")
dsellreservedate = request("dSR")
preParam = request("preparam")
dim sqlStr, i, vChangeContents

dim refer
''refer = request.ServerVariables("HTTP_REFERER")
refer = "itemviewset.asp?" & preParam

'response.write "������" & "<br><br><br>"
'dbget.close()	:	response.End

''�Ǹ�, ��뿩�� �ϰ� ����
if (mode="ModiSellArr") then
    ChkCnt = request("cksel").Count
    for i=1 to ChkCnt
        ''cksel.value is ItemID
        itemid = Trim(request("cksel")(i))
        sellyn = Trim(request("sellyn_" + CStr(itemid)))
        isusing = Trim(request("usingyn_" + CStr(itemid)))
				if isusing ="N" then sellyn="N"		'��뿩�ΰ� N�϶� �Ǹſ��ε� Nó�� 2016.08.09

        mwdiv  = Trim(request("mwdiv_" + CStr(itemid)))
        limityn = Trim(request("limityn_" + CStr(itemid)))
        orgLimityn = Trim(request("orgLimityn_" + CStr(itemid)))
        deliveryType = Trim(request("deliveryTypePolicy_" + CStr(itemid)))
        danjongyn = Trim(request("danjongyn_" + CStr(itemid)))
 'response.write ":" & deliveryType

        if (mwdiv="U") then
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

   '���¿��� ó��
			dim objCmd, returnValue
			IF dsellreservedate<> "" THEN '���¿��� �������� + ���¿��� ���� ��������(�ٹ�-�������, ��ü���)
					 Set objCmd = Server.CreateObject("ADODB.COMMAND")
								With objCmd
									.ActiveConnection = dbget
									.CommandType = adCmdText
									.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_sellreserve_Insert]("&itemid&",'"&dsellreservedate&"','"&session("ssBctId")&"')}"
									.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
									.Execute, , adExecuteNoRecords
									End With
								    returnValue = objCmd(0).Value
							Set objCmd = nothing
							IF returnValue <>  1  THEN
					        Call Alert_return("ó���� ������ �߻��߽��ϴ�. errcode : ���¿���" )
					        response.end
							END IF
			END IF
         dim orgSellyn, orgsellSTDate
        'if (Len(itemid)>0) and (Len(sellyn)>0) and (Len(isusing)>0) and (Len(mwdiv)>0) then '08/07/10 ������ ���� --mwdiv���� �Ѿ���� ����,
		if (Len(itemid)>0) and (Len(sellyn)>0) and (Len(isusing)>0) then
		    sqlStr = " select sellyn, sellSTDate FROM db_item.dbo.tbl_item WHERE itemid =" + CStr(itemid)
            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
            	orgSellyn       = rsget("sellyn")
            	orgsellSTDate   = rsget("sellSTDate")
            end if
            rsget.close

            sqlStr = "update [db_item].[dbo].tbl_item" & VbCrlf
            sqlStr = sqlStr + " set sellyn='" & sellyn & "'" & VbCrlf
            sqlStr = sqlStr + " , isusing='" & isusing & "'" & VbCrlf
            sqlStr = sqlStr + " , mwdiv='" & mwdiv & "'" & VbCrlf
            sqlStr = sqlStr + " , deliveryType='" & deliveryType & "'" & VbCrlf
            sqlStr = sqlStr + " , danjongyn='" & danjongyn & "'" & VbCrlf                '' �߰�
            sqlStr = sqlStr + " , lastupdate=getdate()" & VbCrlf
              if orgSellyn <>"Y" and sellyn  ="Y" and isNull(orgsellSTDate) then
	        sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf
	          end if
            sqlStr = sqlStr + " where itemid=" & CStr(itemid)
            dbget.Execute sqlStr

            '�������� ����(����->�������� ����)
            if limityn="N" and orgLimityn="Y" then
                sqlStr = "update [db_item].[dbo].[tbl_item]" & VbCrlf
                sqlStr = sqlStr + " set limityn='N'" & VbCrlf
                sqlStr = sqlStr + " where itemid=" & CStr(itemid) & " and limityn='Y';" & vbCrLf
                sqlStr = sqlStr + " update [db_item].[dbo].[tbl_item_option]" & VbCrlf
                sqlStr = sqlStr + " set optlimityn='N'" & VbCrlf
                sqlStr = sqlStr + " where itemid=" & CStr(itemid) & " and optlimityn='Y';"
                dbget.Execute sqlStr
            end if

			vChangeContents = ""
			vChangeContents = vChangeContents & "- refer : refer = " & refer & vbCrLf
			vChangeContents = vChangeContents & "- �Ǹſ��� : sellyn = " & sellyn & vbCrLf
			vChangeContents = vChangeContents & "- ��뿩�� : isusing = " & isusing & vbCrLf
			vChangeContents = vChangeContents & "- ���Ա��� : mwdiv = " & mwdiv & vbCrLf
			vChangeContents = vChangeContents & "- ��۱��� : deliveryType = " & deliveryType & vbCrLf
			vChangeContents = vChangeContents & "- �������� : danjongyn = " & danjongyn & vbCrLf
            vChangeContents = vChangeContents & "- �����Ǹſ��� : limityn = N" & vbCrLf

    		'### ���� �α� ����(item)
    		sqlStr = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    		sqlStr = sqlStr & "VALUES('" & session("ssBctId") & "', 'item', '" & itemid & "', '" & Request("menupos") & "', "
    		sqlStr = sqlStr & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    		dbget.execute(sqlStr)

        end if
    next

    response.write "<script type='text/javascript'>alert('�����Ǿ����ϴ�.');</script>"
    response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
    dbget.close()	:	response.End

end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

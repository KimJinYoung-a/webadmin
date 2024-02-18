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

'response.write "수정중" & "<br><br><br>"
'dbget.close()	:	response.End

''판매, 사용여부 일괄 수정
if (mode="ModiSellArr") then
    ChkCnt = request("cksel").Count
    for i=1 to ChkCnt
        ''cksel.value is ItemID
        itemid = Trim(request("cksel")(i))
        sellyn = Trim(request("sellyn_" + CStr(itemid)))
        isusing = Trim(request("usingyn_" + CStr(itemid)))
				if isusing ="N" then sellyn="N"		'사용여부가 N일때 판매여부도 N처리 2016.08.09

        mwdiv  = Trim(request("mwdiv_" + CStr(itemid)))
        limityn = Trim(request("limityn_" + CStr(itemid)))
        orgLimityn = Trim(request("orgLimityn_" + CStr(itemid)))
        deliveryType = Trim(request("deliveryTypePolicy_" + CStr(itemid)))
        danjongyn = Trim(request("danjongyn_" + CStr(itemid)))
 'response.write ":" & deliveryType

        if (mwdiv="U") then
            ''업체 배송인 경우 업체별 배송비 부과가 아니면 2 - 업배기본
            if (deliveryType<>"9") and (deliveryType<>"7") then
                deliveryType = "2"
            end if
        else
            ''업체 배송이 아닌경우 무료배송이 아니면 1 - 텐배기본
            if (deliveryType<>"4") then
                deliveryType = "1"
            end if
        end if

   '오픈예약 처리
			dim objCmd, returnValue
			IF dsellreservedate<> "" THEN '오픈예약 설정여부 + 오픈예약 조건 충족여부(텐배-재고있음, 업체배송)
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
					        Call Alert_return("처리중 에러가 발생했습니다. errcode : 오픈예약" )
					        response.end
							END IF
			END IF
         dim orgSellyn, orgsellSTDate
        'if (Len(itemid)>0) and (Len(sellyn)>0) and (Len(isusing)>0) and (Len(mwdiv)>0) then '08/07/10 김정인 수정 --mwdiv값이 넘어오지 않음,
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
            sqlStr = sqlStr + " , danjongyn='" & danjongyn & "'" & VbCrlf                '' 추가
            sqlStr = sqlStr + " , lastupdate=getdate()" & VbCrlf
              if orgSellyn <>"Y" and sellyn  ="Y" and isNull(orgsellSTDate) then
	        sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf
	          end if
            sqlStr = sqlStr + " where itemid=" & CStr(itemid)
            dbget.Execute sqlStr

            '한정여부 변경(한정->비한정만 진행)
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
			vChangeContents = vChangeContents & "- 판매여부 : sellyn = " & sellyn & vbCrLf
			vChangeContents = vChangeContents & "- 사용여부 : isusing = " & isusing & vbCrLf
			vChangeContents = vChangeContents & "- 매입구분 : mwdiv = " & mwdiv & vbCrLf
			vChangeContents = vChangeContents & "- 배송구분 : deliveryType = " & deliveryType & vbCrLf
			vChangeContents = vChangeContents & "- 단종구분 : danjongyn = " & danjongyn & vbCrLf
            vChangeContents = vChangeContents & "- 한정판매여부 : limityn = N" & vbCrLf

    		'### 수정 로그 저장(item)
    		sqlStr = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    		sqlStr = sqlStr & "VALUES('" & session("ssBctId") & "', 'item', '" & itemid & "', '" & Request("menupos") & "', "
    		sqlStr = sqlStr & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    		dbget.execute(sqlStr)

        end if
    next

    response.write "<script type='text/javascript'>alert('수정되었습니다.');</script>"
    response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
    dbget.close()	:	response.End

end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

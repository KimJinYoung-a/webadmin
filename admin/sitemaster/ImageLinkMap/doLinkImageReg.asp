<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 이미지링크관리
' History : 2019.08.06 원승현 : 신규작성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim idx, title, Link_Image, Isusing, mode, Link_Admin_Image, reguserfrontname, sqlStr
	idx = requestCheckVar(request("idx"),10)
    title = RequestCheckVar(request("title"),256)
    Link_Image = requestCheckVar(request("Link_Image"),128)
    Isusing = requestCheckVar(request("Isusing"),1)
    Link_Admin_Image = requestCheckVar(request("Link_Admin_Image"),128)
    reguserfrontname = requestCheckVar(request("reguserfrontname"),32)


	if idx="" then idx=0
	If idx=0 Then
	mode = "add"
	Else
	mode = "edit"
	End If

if (mode = "add") then
    if Title <> "" and not(isnull(Title)) then
        Title = ReplaceBracket(Title)
    end If
    if RegUserFrontName <> "" and not(isnull(RegUserFrontName)) then
        RegUserFrontName = ReplaceBracket(RegUserFrontName)
    end If

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_ImageLink_Master]" + VbCrlf
    sqlStr = sqlStr + " (Title, Image, RegUser, ModifyUser, Isusing, RegDate, LastUpDate, RegUserImage, RegUserFrontName)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + Title + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Link_Image + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Session("ssBctId") + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Session("ssBctId") + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,getdate()" + VbCrlf
    sqlStr = sqlStr + " ,getdate()" + VbCrlf
    sqlStr = sqlStr + " ,'" + Link_Admin_Image + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reguserfrontname + "'" + VbCrlf
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr

elseif mode = "edit" then
    if Title <> "" and not(isnull(Title)) then
        Title = ReplaceBracket(Title)
    end If
    if RegUserFrontName <> "" and not(isnull(RegUserFrontName)) then
        RegUserFrontName = ReplaceBracket(RegUserFrontName)
    end If

   sqlStr = " update [db_sitemaster].[dbo].[tbl_ImageLink_Master]" + VbCrlf
   sqlStr = sqlStr + " set Title='" + Title + "'" + VbCrlf
   sqlStr = sqlStr + " ,Image='" + Link_Image + "'" + VbCrlf
   sqlStr = sqlStr + " ,ModifyUser='" + Session("ssBctId") + "'" + VbCrlf
   sqlStr = sqlStr + " ,Isusing='" + Isusing + "'" + VbCrlf
   sqlStr = sqlStr + " ,LastUpDate=getdate()" + VbCrlf
   sqlStr = sqlStr + " ,RegUserImage='" + Link_Admin_Image + "'" + VbCrlf
   sqlStr = sqlStr + " ,RegUserFrontName='" + reguserfrontname + "'" + VbCrlf
   sqlStr = sqlStr + " where idx=" + cstr(idx)   
   dbget.Execute sqlStr

end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
	response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
'###########################################################
' Description : 상품 설명 수정 (에디터용)
' History : 2018.01.12 허진원 생성
'###########################################################

session.codePage = 65001		'세션코드 UTF-8 강제 설정
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'// 저장 모드 접수
	dim mode, vChangeContents, vSCMChangeSQL
	mode = requestCheckvar(Request("mode"),16)

	Select Case mode
		Case "ItemContentEdit"
			'상품 설명 저장
			dim sqlStr
		    dim itemid, usinghtml, itemcontent, menupos

		    itemid = requestCheckvar(Request("itemid"),10)
		    usinghtml = requestCheckvar(Request("usinghtml"),1)
		    itemcontent = html2db(Request("itemcontent"))
		    menupos = 594		'상품수정 메뉴

			'// 데이터 저장
			sqlStr = "update [db_item].[dbo].tbl_item_Contents" + vbCrlf
			sqlStr = sqlStr & " set itemcontent='" & itemcontent & "'" + vbCrlf
			sqlStr = sqlStr & " ,usinghtml='" & usinghtml & "'" + vbCrlf
			sqlStr = sqlStr & " where itemid=" & itemid & "" + vbCrlf

		    dbget.execute(sqlStr)


			'// 수정 로그 저장(item)
			vChangeContents = "- HTTP_REFERER : " & request.ServerVariables("HTTP_REFERER") & vbCrLf
			vChangeContents = vChangeContents & "- html사용여부 : usinghtml = " & usinghtml & vbCrLf
			vChangeContents = vChangeContents & "- 상품설명 수정 (에디터) " & vbCrLf

			vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
			vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'item', '" & itemid & "', '" & 594 & "', "
			vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
			dbget.execute(vSCMChangeSQL)
		    	
			Response.Write	"<script type=""text/javascript"">" &_
							"	alert('데이터를 저장하였습니다.');" &_
							"	self.close();" &_
							"</script>"
		Case "ImageInsert"
			'이미지 저장 후 처리 (from http://upload.10x10.co.kr/linkweb/items/itemEditorContentUpload.asp); Origin Site Issued
			Dim isComplete, funcNum, fileUrl, message
		    isComplete = requestCheckvar(Request("isComplete"),1)
		    funcNum = requestCheckvar(Request("funcNum"),10)
		    fileUrl = requestCheckvar(Request("fileUrl"),128)
		    message = requestCheckvar(Request("message"),256)

			Response.Write "<script type=""text/javascript"">" & vbCrLf
			if isComplete then
			    Response.Write "window.parent.CKEDITOR.tools.callFunction(" & funcNum & ", '" & fileUrl & "', '" & message & "');" & vbCrLf
			else
			    Response.Write "var ref = window.parent.CKEDITOR.tools.addFunction( function() { alert( '" & message & "');} );" & vbCrLf
			    Response.Write "window.parent.CKEDITOR.tools.callFunction(ref);" & vbCrLf
			    Response.Write "history.go(-1);" & vbCrLf
			end if
			Response.Write "</script>"

	End Select

	session.codePage = 949		'세션코드 EUC-KR 원복
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
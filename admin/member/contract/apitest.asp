<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 2009.04.07 서동석 생성
'			 	 2010.05.26 한용민 수정
' 			2017.06.23 정윤정 전자계약 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<% 
'dim mailfrom, mailto, mailtitle, mailcontent, innerContents ,CurrState,NextState, sendOpenMail 
'dim mngEmail, mngHp
'dim ocontract, oMdInfoList
'        mailfrom = "@10x10.co.kr"
'	'	 mngEmail ="soso120222@ithinkso.co.kr"
''        if (mailfrom="") or (InStr(mailfrom,"@")<0) or (Len(mailfrom)<8) then
''            response.write "<script>alert('발송자 Email  주소가 유효하지 않습니다.마이 정보에서 Email 수정 후 사용하시기 바랍니다.(등록된 이메일주소:"&mailfrom&")');</script>"
''            response.write "<script>location.replace('" & refer & "');</script>"
''            dbget.close()	:	response.End
''        end if
''    
' 
'    
'
'    
'        '' SMS 발송
'        mngHp ="010-6249-2706"
'
'        ''call SendNormalSMS(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
'       ' call SendNormalSMS_LINK(mngHp,"1644-6030","[텐바이텐] 신규 계약서가 발송되었습니다. email 또는 SCM 업체계약관리 메뉴 참조")
'    
'
'    
'        '' 이메일 발송
'        set ocontract = new CPartnerContract
'        ocontract.FPageSize=50
'        ocontract.FCurrPage = 1
'        ocontract.FRectContractState = 1 ''오픈
'        ocontract.FRectGroupID = "G06657"
'        ocontract.FRectCtrKeyArr = "35181"
'        ocontract.GetNewContractList
'
'        set oMdInfoList = new CPartnerContract
'        oMdInfoList.FRectGroupID = "G06657"
'        oMdInfoList.FRectContractState = 1 ''오픈
'        oMdInfoList.FRectCtrKeyArr = "35181"
'        oMdInfoList.getContractEmailMdList(FALSE)
'
'        mailtitle       = "[텐바이텐] 신규 계약서가 발송 되었습니다."
'       
'        	 mailcontent   = makeEcCtrMailContents(ocontract,oMdInfoList,False,manageUrl)
'       
'
'        Call SendMail(mailfrom, mngEmail, mailtitle, mailcontent)
'
'        set ocontract=nothing
'        set oMdInfoList=nothing



    
%>
	<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script> 
<script >
	function viewCont(){
		var data_cont ={ 
			"seq":"14957",
			"venderno":"2158764039",
			"users":encodeURIComponent("admin")
			};
		$.ajax({
			url : "https://edocu.uplus.co.kr/api/viewCont?access_token=7fcac127-7998-4fbf-ac2b-0a31d44a527e",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				if(data.status=="fail"){
					alert("err_msg:"+data.info);
				}else{
					alert("result \n title"+data.title+" B company:"+data.memb_list[1].company);
				}
			}
		});
	}
</script>
 <a href="javascript:viewCont()">viewCont</a>
 
 <% dim ecstate
 ecstate = fnViewEcCont1("14936" ,"1081430367","김진성","7fcac127-7998-4fbf-ac2b-0a31d44a527e")
 
 response.write "ecstate="&ecstate
 
 
 public Function fnViewEcCont1(ecCtrseq, bcompno, ecBuser,access_token)			 
 	dim APIpath,strParam
 		dim objXML, iRbody,jsResult
 		dim con_status,con_info,con_error,ecCtrState
	APIpath =FecURL&"/api/viewCont" 
	strParam = "?seq="&ecCtrseq&"&venderno="&bcompno&"&users=" &server.URLEncode(ecBUser)
	
 	On Error Resume Next 
 Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		  objXML.Open "GET", APIpath&strParam , False
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
			objXML.SetRequestHeader "Authorization", "Bearer " & access_token
			objXML.Send()				
			iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
		
			iRbody= replace(iRbody,"tmpCallBack({","{")
		 	iRbody = replace(iRbody,"})","}") 
  
			If objXML.Status = "200" Then
				Set jsResult = JSON.parse(iRbody)
					con_status	= jsResult.status
					con_info= jsResult.info 
					'con_error = jsResult.error
					response.write con_info
					if con_status ="fail" Then 
							 if (con_info="001") then
							 		FErrMsg= "계약서번호 값 없음"
							 elseif (con_info="002")then
							  	FErrMsg= "사업자번호(venderno) 값 없음,venderno 에 사용자 존재하지 않음"
							 elseif (con_info="003")then
							  	FErrMsg= "해당 정보에 대한 문서가 존재하지않음"  
							 elseif (con_info="004")then
							  	FErrMsg= "해당문서가 삭제되었습니다."   	
							  	ecCtrState = "-1"
							 end if 
					else
						ecCtrState = jsResult.nowstat_vl			  
					end if
					
			Set jsResult = Nothing
			End If
		Set objXML = Nothing		 
	On Error Goto 0 
	fnViewEcCont1 = ecCtrState
	
	End Function
	
	
	'response.write "a="& fnCheckUser("2158764039","7fcac127-7998-4fbf-ac2b-0a31d44a527e")
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
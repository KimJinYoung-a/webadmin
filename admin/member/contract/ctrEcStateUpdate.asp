<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�� ��� ����
' Hieditor : 2009.04.07 ������ ����
'			 	 2010.05.26 �ѿ�� ����
' 			2017.06.23 ������ ���ڰ�� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<%
dim sqlStr
dim oneContract,acctoken,reftoken,ecCtrState
dim  arrList, intLoop

 		sqlStr = " select  m.ctrKey, ecctrseq, g.company_no, ecBUser , m.ctrstate "
 		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_ctr_master as m  "
 		sqlStr = sqlStr & "	inner join db_partner.dbo.tbl_partner_group as g on m.groupid = g.groupid "
 		sqlStr = sqlStr & "	where CtrState > 0 and CtrState not in (7,9) "
 		sqlStr = sqlStr & "	 and ecCtrseq > 0 	"
 	     rsget.Open sqlStr,dbget,1
 	     if not rsget.eof Then
 	     	 arrList = rsget.getrows()
 	    end if
 	    rsget.close
 	    
 	    	if isArray(arrList) Then
 	    		
		'token ��������(db����)
		 set oneContract = new CPartnerContract
				oneContract.fnGetContractToken
				acctoken = oneContract.Facctoken 	
				reftoken = oneContract.Freftoken 
  		set oneContract = nothing
  		
  		'token�� ������ token ����
 				if   not isNull(acctoken) then  
 	 	
 	 				for intLoop = 0 To uBound(arrList,2)
 	 					ecCtrState =  fnViewEcCont(arrList(1,intLoop),replace(arrList(2,intLoop),"-",""),arrList(3,intLoop),acctoken)
 	 				 
 	 					if Fchkerror ="invalid_token" then
				 				call sbGetRefToken(reftoken)
				 				acctoken = Faccess_token
				 				ecCtrState =  fnViewEcCont(arrList(1,intLoop),replace(arrList(2,intLoop),"-",""),arrList(3,intLoop),acctoken)
				 		end if	
				 		
				 		if ecCtrState <> "" then				 	 				 	 
				 		sqlStr = "update db_partner.dbo.tbl_partner_ctr_master set ctrstate = "&GetContractEcState(ecCtrState)&", lastupdate =getdate()"
			 			sqlstr = sqlstr & " where ctrKey="&arrList(0,intLoop)&" and ctrstate <> " &GetContractEcState(ecCtrState)
			 			dbget.Execute  sqlstr, 1		
			 			end if
 	 				next
 	 			end if	
 			end if
 	  
	 
%>		
<script type="text/javascript">
	alert("�Ϸ�Ǿ����ϴ�.");
	history.back();
</script>	
<!-- #include virtual="/lib/db/dbclose.asp" -->				
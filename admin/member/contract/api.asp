<html>
<head>
	<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script> 
<script type="text/javascript" language="javascript">

	function test(){
		var data_cont ={ 
			 "code":"582",
			 "code_nm":encodeURIComponent("�ѱ��׽�Ʈ"),
			 "ptestList[0].id":"A",
			 "ptestList[0].pw":encodeURIComponent("������"),
			 "ptestList[1].id":"B",
			 "ptestList[1].pw":encodeURIComponent("���ٻ�"),
			 "ptestList[2].id":"C",
			 "ptestList[2].pw":encodeURIComponent("������"),
			};
		$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/test?access_token=e17a08dc-3453-4cfd-a332-f056c8c87e0c",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n code:"+data.code+" \n code_nm:"+data.code_nm);
				for(var i=0;i<data.ptestList.length;i++){
					alert("\n id"+data.ptestList[i].id+" \n pw"+data.ptestList[i].pw);
				}		
			}
		});
	}


	//cross domain ������ ���� ũ�ҿ��� �������� ������ �׽�Ʈ �����θ� ����Ͻð� ������� ���ø����̼� ���� �����Ͻñ� �ٶ��ϴ�.
	function getToken(){
		var data_cont ={ 
			 "grant_type":"password",
			 "client_id":"edocuopenapi",
			 "username":"CTENTEN",
			 "password":"tax7882!"};
		var jsontext = $.post( "https://edocu.uplus.co.kr/oauth/token",
		data_cont,
		function(data) { 
		  var json = JSON.parse(data);
		  alert(json);
				 alert("token:"+json.access_token);
				 alert("refresh_token:"+json.refresh_token);
			
		});
	}
	
	function checkuser(){
		var data_cont ={ 
			 "corp_id":"6831800305"
			};
		$.ajax({
			url : "https://edocu.uplus.co.kr/api/checkuser?access_token=7fcac127-7998-4fbf-ac2b-0a31d44a527e",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n corp_id:"+data.corp_id+" \n corp_nm:"+data.corp_nm+" \n accnt_cd:"+data.accnt_cd+" \n status_cd:"+data.status_cd);
			}
		});
	}

	function inputuser(){ 
		var data_cont ={ 
			"accnt_cd":"UB",
			"address":encodeURIComponent("����� ���α� ������"),
			"cell":"",
			"company":encodeURIComponent("C7321"),
			"condition":encodeURIComponent("�׽�Ʈ ����"),
			"email":"",
			"email_noti_yn":"",
			"emp_name":encodeURIComponent("C7321"),
			"hubcompany_id":"C7321",
			"items":encodeURIComponent("�׽�Ʈ ����"),
			"name":encodeURIComponent("��ǥ��"),
			"nlfr_nxt_yn":"",
			"nltx_now_yn":"",
			"password":"a1234567",
			"regno":"1288619698",
			"sms_noti_yn":"",
			"tell":"",
			"userid":"C7321"
			};
		$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/inputuser?access_token=7fcac127-7998-4fbf-ac2b-0a31d44a527e",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n status"+data.status+" \n info"+data.info);
			}
		});
	}

	function createCont(){
		var data_cont ={ 
			 "type_seq":"2270",
			 "cancel_limit":"0",
			 "contract_dt":"2017-07-24",
			 "contract_key":"",
			 "contract_money":"0",
			 "expire_dt":"2017-12-31",
			 "venderno":"2118700620",
			 "search_word":encodeURIComponent("�� ������"),
			 "start_dt":"2017-07-24",
			 "title":"html test",
			 "membList[0].company":encodeURIComponent("(��)�ٹ�����"),
			 "membList[0].gubun":"A",
			 "membList[0].users":encodeURIComponent("CTENTEN"),
			 "membList[0].venderno":"2118700620",
			 "membList[1].company":encodeURIComponent("�� ������"),
			 "membList[1].gubun":"B",
			 "membList[1].users":encodeURIComponent("������"),
			 "membList[1].venderno":"6831800305",
			 "usertagList[0].tag_nm":"JUNGSAN_DATE",
			 "usertagList[0].tag_vl":encodeURIComponent("�Ǹ�(����)���� �Ϳ� ����")
			};
		$.ajax({
			url : "https://edocu.uplus.co.kr/api/createCont?access_token=7fcac127-7998-4fbf-ac2b-0a31d44a527e",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n status"+data.status+" \n info"+data.info);
			}
		});
	}

	function viewCont(){
		var data_cont ={ 
			"seq":"556",
			"venderno":"2118734585",
			"users":encodeURIComponent("ADMIN")
			};
		$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/viewCont?access_token=e17a08dc-3453-4cfd-a332-f056c8c87e0c",
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


	function checkEmail(){
		
		$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/mailCheck?access_token=093eb86e-f2cc-42d2-b643-c822c9334d09",
			dataType : "jsonp",
			type:'get',
			data : {"email":"dhshin@lgdacom.biz"},
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				if(!data.error){
					alert("result:"+data.info);
				}else{
					alert("result:"+data.error+"|"+data.error_description);
				}
			}
		});
	}


function checkCont(){
	var data_cont ={ 
			 "seq":"556",
   			"venderno":"2118734585",
   			"users":encodeURIComponent("mmmg")
			};
			$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/checkCont?access_token=e17a08dc-3453-4cfd-a332-f056c8c87e0c",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n status"+data.status+" \n info"+data.info);
			}
		});
} 



function checkCont2(){
	var data_cont ={ 
			 "seq":"305",
   			"venderno":"211873458",
   			"users":encodeURIComponent("CIUTEL")
			};
			$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/checkCont?access_token=e17a08dc-3453-4cfd-a332-f056c8c87e0c",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n status"+data.status+" \n info"+data.info);
			}
		});
}

 

function approveCont(){
	var data_cont ={ 
			 "seq":"780",
   			"venderno":"2118734585",
   			"users":encodeURIComponent("mmmg")
			};
			$.ajax({
			url : "http://webtax-test.uplus.co.kr/api/approveCont?access_token=e17a08dc-3453-4cfd-a332-f056c8c87e0c",
			dataType : "jsonp",
			type:'get',
			data :  data_cont,
			complete: function(xhr,status){
				if(status == "error"){
					alert("err");
				}
			},
			success : function(data){
				alert("result \n status"+data.status+" \n info"+data.info);
			}
		});
} 

</script>
</head>
<body>
<a href="javascript:getToken()">getToken</a> 
<br/>
<a href="javascript:checkuser()">checkuser</a>
<br/>
<a href="javascript:inputuser()">inputuser</a>
<br/>
<a href="javascript:createCont()">createCont</a>
<br/>
<a href="javascript:viewCont()">viewCont</a>
<br/>
<a href="javascript:checkEmail()">eamilcheck</a>
<br/>
<a href="javascript:checkCont()">checkCont</a>
<br/>
<a href="javascript:checkCont()">approveCont</a>

<form name="frmecView" method="post" action="http://w20-test.webtax21.com/w20/contractView.do" target="_blank">
<input type="hidden" name="remote_id" value="CTENTEN" />  <!-- �ۼ��� LOGIN ID -->
<input type="hidden" name="cont_seq" value="802" />  <!-- ��༭ ��ȣ -->
<input type="hidden" name="corp_id" value="2118700620" /> <!-- ����� ȭ���Ϸ��� ����ڹ�ȣ -->
</form>
 
 <form name="frmecView1" method="post" action="http://w20-test.webtax21.com/w20/contractView.do" target="_blank"> 
	<input type="hidden" name="remote_id" value="CTENTEN" />  <!-- �ۼ��� LOGIN ID -->
	<input type="hidden" name="cont_seq" value="802" />  <!-- ��༭ ��ȣ -->
	<input type="hidden" name="corp_id" value="2118700620" /> <!-- ����� ȭ���Ϸ��� ����ڹ�ȣ -->
	</form>
	
<a href="javascript:document.frm.submit();">viewer</a>
<a href="javascript:document.frmecView.submit();">viewer1</a>

</body>
</html>


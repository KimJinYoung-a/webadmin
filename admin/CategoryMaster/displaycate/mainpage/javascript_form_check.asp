	if($("input[name=multiimg1]").val() == ""){
		alert("multi �̹����� ������ּ���.");
		return;
	}
	if($("input[name=multiimg2]").val() == ""){
		alert("multi �̹����� ������ּ���.");
		return;
	}
	if($("input[name=multiimg3]").val() == ""){
		alert("multi �̹����� ������ּ���.");
		return;
	}
	
	for(var i=1; i<13; i++){
		if($("input[name=itemimg"+i+"]").val() == ""){
			alert("item"+i+" �̹����� �����ϼ���.");
			return;
		}
	}
	
	for(var i=1; i<5; i++){
		if($("input[name=eventimg"+i+"]").val() == ""){
			alert("event"+i+" �̹����� �����ϼ���.");
			return;
		}
	}
	
	if($("input[name=bookimg]").val() == ""){
		alert("bookimg �̹����� ������ּ���.");
		return;
	}
	
	if($("input[name=recipeimg]").val() == ""){
		alert("recipeimg �̹����� ������ּ���.");
		return;
	}
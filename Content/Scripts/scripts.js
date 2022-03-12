var Helpers = window.Helpers = {

	Submit: function(obj){
		obj = $(obj);
		var params = obj.closest('div').find('input[type=text],input[type=password]').serialize();
		Helpers.Send('http://localhost:51001/index/usuarios/create', params);
	},

	Send: function(url, params){
		$.ajax({
			url:url,
			data:params,
			type:'POST',
			dataType:'HTML', 
			success: function(html){
				console.log(html);
			}
		});
	}

}

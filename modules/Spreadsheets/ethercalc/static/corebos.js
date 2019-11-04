document.addEventListener('DOMContentLoaded', function (event) {
	var myURL = new URL(window.location.href);
	var cbUsr = myURL.searchParams.get('usr');
	var cbPwd = myURL.searchParams.get('pwd');
	var cbURL = myURL.searchParams.get('url');
	var cbMtd = myURL.searchParams.get('mtd');
	var cbSid = myURL.searchParams.get('sid');
	setTimeout(
		function () {
			document.getElementById('SocialCalc-button_corebos').onclick=function () {
				var oDiv = document.createElement('div');
				oDiv.className='vex vex-theme-flat-attack';
				var vDiv = document.createElement('div');
				vDiv.className='vex-overlay';
				var iDiv = document.createElement('div');
				iDiv.id = 'cbinfo';
				iDiv.className='vex-content';
				iDiv.innerHTML='<b>Saving information...</b>';
				oDiv.appendChild(vDiv);
				oDiv.appendChild(iDiv);
				var cbdiv=document.getElementsByTagName('body')[0].appendChild(oDiv);
				var cbconn = new cbWSClient(cbURL);
				cbconn.doLogin(cbUsr, cbPwd, false)
				.then(function () {
					cbconn.doInvoke(cbMtd, {'spreadsheetid':cbSid}, 'GET')
					.then(function (response) {
						return response.json();
					})
					.then(function (result) {
						document.getElementById('cbinfo').innerHTML = '<b>'+result+'</b>';
						setTimeout(
							function () {
								cbdiv.remove();
							},
							4000
						);
					});
				});
			};
		},
		3000
	);
});

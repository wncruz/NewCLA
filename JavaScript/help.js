function showAjudaImg(img,div,dt,dl){
				var img = document.getElementById(img);
			
				var top = 0; //img.offsetTop;
				var left = img.offsetLeft;
				var pai = img.parentNode;
				var filho = img;
				
				while (pai != null) {
					if (!isNaN(pai.offsetTop)) {
						top += regraSoma(pai,filho);
					}
					if (!isNaN(pai.offsetLeft)) {
						left += pai.offsetLeft;
					}
					filho = pai;
					pai = pai.parentNode;
				}
				
				var layer = document.getElementById(div);
				layer.style.visibility = "visible";
				layer.style.top = (top + dt) + "px";
				layer.style.left = (left + dl) + "px";
				
				document.getElementById("duvida1").src='../imagens/fechar_b.gif';
			}
			
			function hideAjuda(div){
				var layer = document.getElementById(div);
				layer.style.top = "-1000px"
				layer.style.left = "-1000px"
				layer.style.visibility = "hidden";
				
				document.getElementById("duvida1").src='../imagens/question_b.gif';
			}
			
			function regraSoma(pai,filho) {
				//alert(pai.nodeName + ":" + pai.offsetTop + " - > " + filho.nodeName + ":" + filho.offsetTop);
				
				if ((pai.nodeName == "TR") || (pai.nodeName == "TABLE")){
					return pai.offsetTop;
				}
				
				return 0;
			}
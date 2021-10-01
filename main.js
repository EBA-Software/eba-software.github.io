//Get Header and Footer
$(function(){
  $("#header").load("https://eba-software.github.io/header.html");
  $("#footer").load("https://eba-software.github.io/footer.html");
});

//Create Functions
function dwnld(fileDir, name) {
  var a = document.createElement("a");
  a.href = fileDir;
  a.setAttribute("download",name);
  a.click();
}

//Microsft Clarity
(function(c,l,a,r,i,t,y){
  c[a]=c[a]||function(){(c[a].q=c[a].q||[]).push(arguments)};
  t=l.createElement(r);t.async=1;t.src="https://www.clarity.ms/tag/"+i;
  y=l.getElementsByTageName(r)[0];y.parentNode.insertBefore(t,y);
})(window, document, "clarity", "script", "69wj77p43q");

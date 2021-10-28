//Get Header, Footer, and Status
$(function(){
  $("#header").load("https://eba-software.github.io/header.html");
  $("#footer").load("https://eba-software.github.io/footer.html");
  $("#status").load("https://eba-software.github.io/status.html");
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
  y=l.getElementsByTagName(r)[0];y.parentNode.insertBefore(t,y);
})(window, document, "clarity", "script", "69wj77p43q");

//Dark Mode
function toggleDark() {
  const themeStylesheet = document.getElementById('dark');
  if(themeStylesheet.href.includes('dark')){
    themeStylesheet.href = 'https://eba-software.github.io/styles-dark.css';
    localStorage.setItem('dark', 'https://eba-software.github.io/styles-dark.css');
  } else {
    themeStylesheet.href = 'https://eba-software.github.io/styles.css';
    localStorage.setItem('dark', 'https://eba-software.github.io/styles.css');
  }
}

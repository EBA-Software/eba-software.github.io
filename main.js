//If debug page
var debugMode = false;
if (document.location == 'https://eba-software.github.io/debug.html') {
  debugMode = true;
}
if (document.location == 'https://eba-software.github.io/debug') {
  debugMode = true;
}

//Get Header, Footer, and Status
$(function(){
  $("#header").load("https://eba-software.github.io/header.html");
  $("#footer").load("https://eba-software.github.io/footer.html");
  $("#status").load("https://eba-software.github.io/status.html");
  $("#status-dark").load("https://eba-software.github.io/status-dark.html");
});

//Create Functions
function dwnld(fileDir, name) {
  if (debugMode == true) {alert("Downloading: " + fileDir);}
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

//Cookies
function createCookie(name,value) {
  if (debugMode == true) {alert("Creating Cookie: " + name + "=" + value + "; path=/");}
  document.cookie = name + "=" + value + "; path=/";
}
function getCookie(name) {
  if (debugMode == true) {alert("Reading Cookie: " + name);}
  name = name + "=";
  let decodedCookie = decodeURIComponent(document.cookie);
  let ca = decodedCookie.split(';');
  for(let i = 0; i <ca.length; i++) {
    let c = ca[i];
    while (c.charAt(0) == ' ') {
      c = c.substring(1);
    }
    if (c.indexOf(name) == 0) {
      return c.substring(name.length, c.length);
    }
  }
  return "";
}

//Dark Mode
function toggleDark() {
  const themeStylesheet = document.getElementById('dark');
  if(themeStylesheet.href.includes('dark')){
    themeStylesheet.href = 'https://eba-software.github.io/styles.css';
    createCookie('dark','false');
  } else {
    themeStylesheet.href = 'https://eba-software.github.io/styles-dark.css';
    createCookie('dark','true');
  }
}

//Read Dark Mode Cookie
if (getCookie('dark') == 'true') {
  toggleDark();
}

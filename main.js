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

//Cookies
function createCookie(name,value) {
  document.cookie = name + "=" + value + "; path=/";
}
function getCookie(name) {
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

function getCookie(c_name) {
    if (document.cookie.length > 0) {
        c_start = document.cookie.indexOf(c_name + "=");
        if (c_start != -1) {
            c_start = c_start + c_name.length + 1;
            c_end = document.cookie.indexOf(";", c_start);
            if (c_end == -1) {
                c_end = document.cookie.length;
            }
            return unescape(document.cookie.substring(c_start, c_end));
        }
    }
    return "";
}

//Dark Mode
function toggleDark() {
  const themeStylesheet = document.getElementById('dark');
  if(themeStylesheet.href.includes('dark')){
    themeStylesheet.href = 'https://eba-software.github.io/styles.css';
    localStorage.setItem('dark', 'https://eba-software.github.io/styles.css');
  } else {
    themeStylesheet.href = 'https://eba-software.github.io/styles-dark.css';
    localStorage.setItem('dark', 'https://eba-software.github.io/styles-dark.css');
  }
}

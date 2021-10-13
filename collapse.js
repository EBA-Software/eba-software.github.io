var coll = document.getElementsByClassName("col");
var i;

for (i = 0; i < coll.length; i++) {
  coll[i].addEventListener("click",function() {
    this.classList.toggle("active");
    var cont = this.nextElementSibling;
    if (cont.syle.maxHeight){
      cont.style.maxHeight = null;
    } else {
      cont.style.maxHeight = cont.scrollHeight + "px";
    }
  });
}

function move2() {
    var elem = document.getElementById("myBar");
    var width = getWidth();
    elem.style.width = width + '%';
    elem.innerHTML = width * 1 + '%';
}

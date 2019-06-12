<script type="text/javascript">
function createButtons() {
    var btn = document.createElement("<input type=\"button\" class=\"button\" value=\"fill\" onclick=\"fillRand();\">");
    var parent = document.getElementById("sbtn").parentNode;
    parent.appendChild(btn);
}

function fillRand() {
    for(var i=1;i<11;i++) {
        var id = "rd1" + i;
        var value = Math.random();
        if(value < 0.5)
            id += "1";
        else 
            id += "2";
        
        var el = document.getElementById(id);
        el.checked = "checked";
    }
}
createButtons();
</script>
jbMap = window.jbMap || {};
function jbViaJs(locationId) {
    var _f = undefined;
    var _fconv = 'jbMap[\"' + locationId + '\"]';
    try {
        _f = eval(_fconv);
        if (_f != undefined) {
            _f()
        }
    } catch(e) {}
}
function jbLoader(closetag) {
    var jbTest = null,
    jbTestPos = document.getElementsByTagName("span");
    for (var i = 0; i < jbTestPos.length; i++) {
        if (jbTestPos[i].className == "jbTestPos") {
            jbTest = jbTestPos[i];
            break
        }
    }
    if (jbTest == null) return;
    if (!closetag) {
        document.write("<span id=jbTestPos_" + jbTest.id + " style=display:none>");
        jbViaJs(jbTest.id);
        return
    }
    document.write("</span>");
    var real = document.getElementById("jbTestPos_" + jbTest.id);
    for (var i = 0; i < real.childNodes.length; i++) {
        var node = real.childNodes[i];
        if (node.tagName == "SCRIPT" && /closetag/.test(node.className)) continue;
        jbTest.parentNode.insertBefore(node, jbTest);
        i--
    }
    jbTest.parentNode.removeChild(jbTest);
    real.parentNode.removeChild(real)
}

var adv_nav = '<script type="text/javascript">var cpro_id="u1645651";</script>';
adv_nav += '<scri'+'pt src="http://cpro.baidustatic.com/cpro/ui/c.js" type="text/javascript"></sc'+'ript>';
var adv_side = '<script type="text/javascript">var cpro_id="u1645657";</script>';
adv_side += '<scri'+'pt src="http://cpro.baidustatic.com/cpro/ui/c.js" type="text/javascript"></sc'+'ript>';

var advtop = '<scr'+'ipt async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></scr'+'ipt><!-- tools_1000 --><ins class="adsbygoogle"     style="display:inline-block;width:970px;height:90px"     data-ad-client="ca-pub-6384567588307613"     data-ad-slot="8927361285"></ins><scr'+'ipt>(adsbygoogle = window.adsbygoogle || []).push({});</scr'+'ipt>';
var advbottom = '<scr'+'ipt async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></scr'+'ipt><!-- tools_1000 --><ins class="adsbygoogle"     style="display:inline-block;width:970px;height:90px"     data-ad-client="ca-pub-6384567588307613"     data-ad-slot="8927361285"></ins><scr'+'ipt>(adsbygoogle = window.adsbygoogle || []).push({});</scr'+'ipt>';

jbMap['adv_nav'] = function() {
	document.writeln(adv_nav);
};
jbMap['adv_side'] = function() {
	document.writeln(adv_side);
};
jbMap['advtop'] = function() {
	document.writeln(advtop);
};
jbMap['advbottom'] = function() {
	document.writeln(advbottom);
};


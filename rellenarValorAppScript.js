function rellenarVacios(rango) {
    var last = "";

    for (var i = rango.length; i > 0; i--) {

        if (rango[i] == "") {
            rango[i] = last;
        } else {
            last = rango[i];
        }

    }

    return rango;
}
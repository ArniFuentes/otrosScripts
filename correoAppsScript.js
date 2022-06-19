function enviarCorreo(proveedor, linkData, destinatarios, asunto) {

    try {
        // hoja con la tabla dinámica y la nomenclatura
        var sheet = SpreadsheetApp.getActive().getSheetByName(proveedor);

        // rango de la tabla dinámica
        var rango = sheet.getRange(11, 1, sheet.getLastRow(), 10).getValues();


        // declaración e inicialización de la variable contenedora del body 
        var body = '';

        // salto de linea
        body += "<br>";

        // Encabezados de la tabla
        body += "<table style='border:1px solid #dddddd;border-collapse:collapse;text-align:center' border = 1 cellpadding = 5>";
        body += "<tr style='background-color: #4b5796; color: white'>";
        body += "<th>" + sheet.getRange('A11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('B11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('C11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('D11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('E11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('F11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('G11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('H11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('I11').getValue() + "</th>";
        body += "<th>" + sheet.getRange('J11').getValue() + "</th>";
        body += "</tr>";


        // Insertando los registros de la tabla (rango.length es 140)
        for (var i = 1; i < rango.length - 10; i++) {

            // si es el último registro, que le ponga letra negra remarcada
            if (i == rango.length - 11) {

                body += "<tr style='background-color: #f2f2f2'>";

                if (String(rango[i][0]).match('Total') != null) {
                    body += "<td><b>" + rango[i][0] + "</b></td>";
                } else if (String(rango[i][0]).match('Suma') != null) {
                    body += "<td><b>" + rango[i][0] + "</b></td>";
                } else {
                    body += "<td><b>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</b></td>";
                }

                // body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                body += "<td><b>" + rango[i][1] + "</b></td>";
                body += "<td><b>" + (rango[i][2]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][3]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][4]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][5]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][6]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][7] * 100).toFixed(2) + '%' + "</b></td>";
                body += "<td><b>" + (rango[i][8] * 100).toFixed(2) + '%' + "</b></td>";
                body += "<td><b>" + (rango[i][9] * 100).toFixed(2) + '%' + "</b></td>";

                body += "</tr>";


            } else {

                // Si match devuelve el string 'Total' entonces que la fila tenga un color mas oscuro
                if (String(rango[i][0]).match('Total') != null) {

                    body += "<tr style='background-color: #f2f2f2'>";

                    if (String(rango[i][0]).match('Total') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else if (String(rango[i][0]).match('Suma') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else {
                        body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    }

                    // body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    body += "<td>" + rango[i][1] + "</td>";
                    body += "<td>" + (rango[i][2]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][3]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][4]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][5]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][6]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][7] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][8] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][9] * 100).toFixed(2) + '%' + "</td>";

                    body += "</tr>";


                } else {

                    body += "<tr>";

                    if (String(rango[i][0]).match('Total') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else if (String(rango[i][0]).match('Suma') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else {
                        body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    }

                    // body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    body += "<td>" + rango[i][1] + "</td>";
                    body += "<td>" + (rango[i][2]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][3]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][4]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][5]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][6]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][7] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][8] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][9] * 100).toFixed(2) + '%' + "</td>";

                    body += "</tr>";

                }

            }

        }

        // final del tag table 
        body += "</table>";

        // salto de linea
        body += "<br>";

        // Link a la data
        body += 'Link: ' + linkData;


        GmailApp.sendEmail(
            destinatarios,
            asunto,
            "Requires HTML",
            { htmlBody: body }
        );


    } catch (error) {
        console.log(error);
    }

}
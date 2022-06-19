function getDirection() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mapSheet = ss.getSheetByName("Data");

    for (var i = 2; i < mapSheet.getLastRow() + 1; i++) {

        try {
            var start = mapSheet.getRange(i, 54).getValue();
            var end = mapSheet.getRange(i, 55).getValue();

            var directions = Maps.newDirectionFinder()
                .setOrigin(start)
                .setDestination(end)
                .setMode(Maps.DirectionFinder.Mode.DRIVING)
                .getDirections();

            var route = directions.routes[0];

            var distance = route.legs[0].distance.text;

            mapSheet.getRange(i, 56).setValue(distance);

        } catch (error) {
            mapSheet.getRange(i, 56).setValue(0);
        }

    }

}
'use strict';

var clientContext;
var completedItems;

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', sharePointReady);

function sharePointReady() {
    clientContext = SP.ClientContext.get_current();
}

function purgeCompletedItems() {

    var list = clientContext.get_web().get_lists().getByTitle('New Employees In Seattle');
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        '<View><Query><Where><Eq>' +
          '<FieldRef Name=\'OrientationStage\'/><Value Type=\'Choice\'>Completed</Value>' +
        '</Eq></Where></Query></View>');
    completedItems = list.getItems(camlQuery);
    clientContext.load(completedItems);
    
    clientContext.executeQueryAsync(deleteCompletedItems, onGetCompletedItemsFail);
}

function deleteCompletedItems() {

    var itemArray = new Array();
    var listItemEnumerator = completedItems.getEnumerator();
        
    while (listItemEnumerator.moveNext()) {
        var item = listItemEnumerator.get_current();
        itemArray.push(item);
    }

    var i;
    for (i = 0; i < itemArray.length; i++) {
        itemArray[i].deleteObject();
    }

    clientContext.executeQueryAsync(onDeleteCompletedItemsSuccess, onDeleteCompletedItemsFail);
}

function onDeleteCompletedItemsSuccess(sender, args) {
    alert('Completed orientations deleted.');
    location.reload(true);
}

// Failure callbacks

function onGetCompletedItemsFail(sender, args) {
    alert('Unable to get completed items. Error:' + args.get_message() + '\n' + args.get_stackTrace());
}

function onDeleteCompletedItemsFail(sender, args) {
    alert('Unable to delete completed items. Error:' + args.get_message() + '\n' + args.get_stackTrace());
}
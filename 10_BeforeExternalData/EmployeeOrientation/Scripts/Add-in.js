'use strict';

var clientContext = SP.ClientContext.get_current();
var completedItems;
var notStartedItems;
var calendarList;
var scheduledItems;
var hostWebURL = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
var hostWebContext = new SP.AppContextSite(clientContext, hostWebURL);


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

    clientContext.executeQueryAsync(null, onDeleteCompletedItemsFail);
}

function ensureOrientationScheduling() {

    var employeeList = clientContext.get_web().get_lists().getByTitle('New Employees In Seattle');
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        '<View><Query><Where><Eq>' +
            '<FieldRef Name=\'OrientationStage\'/><Value Type=\'Choice\'>Not started</Value>' +
        '</Eq></Where></Query></View>');
    notStartedItems = employeeList.getItems(camlQuery);
    clientContext.load(notStartedItems);

    clientContext.executeQueryAsync(getScheduledOrientations, onGetNotStartedItemsFail);
}

function getScheduledOrientations() {

    calendarList = hostWebContext.get_web().get_lists().getByTitle('Employee Orientation Schedule');
    var camlQuery = new SP.CamlQuery();
    scheduledItems = calendarList.getItems(camlQuery);
    clientContext.load(scheduledItems);

    clientContext.executeQueryAsync(scheduleAsNeeded, onGetScheduledItemsFail);
}

function scheduleAsNeeded() {
    var unscheduledItems = false;
    var dayOfMonth = '10';

    var listItemEnumerator = notStartedItems.getEnumerator();

    while (listItemEnumerator.moveNext()) {
        var alreadyScheduled = false;
        var notStartedItem = listItemEnumerator.get_current();

        var calendarEventEnumerator = scheduledItems.getEnumerator();
        while (calendarEventEnumerator.moveNext()) {
            var scheduledEvent = calendarEventEnumerator.get_current();
            if (scheduledEvent.get_item('Title').indexOf(notStartedItem.get_item('Title')) > -1) {
                alreadyScheduled = true;
                break;
            }
        }
        if (alreadyScheduled === false) {
            var calendarItem = new SP.ListItemCreationInformation();
            var itemToCreate = calendarList.addItem(calendarItem);
            itemToCreate.set_item('Title', 'Orient ' + notStartedItem.get_item('Title'));
            itemToCreate.set_item('EventDate', '2015-06-' + dayOfMonth + 'T21:00:00Z');
            itemToCreate.set_item('EndDate', '2015-06-' + dayOfMonth + 'T23:00:00Z');
            dayOfMonth++;
            itemToCreate.update();
            unscheduledItems = true;
        }
    }
    if (unscheduledItems) {
        calendarList.update();
        clientContext.executeQueryAsync(onScheduleItemsSuccess, onScheduleItemsFail);
    }
}

function onScheduleItemsSuccess(sender, args) {
    alert('There was at least one unscheduled orientation and it has been added to the '
              + 'Employee Orientation Schedule calendar.');
}

// Failure callbacks

function onGetCompletedItemsFail(sender, args) {
    alert('Unable to get completed items. Error:' + args.get_message() + '\n' + args.get_stackTrace());
}

function onDeleteCompletedItemsFail(sender, args) {
    alert('Unable to delete completed items. Error:' + args.get_message() + '\n' + args.get_stackTrace());
}

function onGetNotStartedItemsFail(sender, args) {
    alert('Unable to get the not-started items. Error:'
        + args.get_message() + '\n' + args.get_stackTrace());
}

function onGetScheduledItemsFail(sender, args) {
    alert('Unable to get scheduled items from host web. Error:'
        + args.get_message() + '\n' + args.get_stackTrace());
}

function onScheduleItemsFail(sender, args) {
    alert('Unable to schedule items on host web calendar. Error:'
        + args.get_message() + '\n' + args.get_stackTrace());
}

// Utility functions

function getQueryStringParameter(paramToRetrieve) {
    var params = document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve) {
            return singleParam[1];
        }
    }
}


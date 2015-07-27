// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file. 

'use strict';

var clientContext = SP.ClientContext.get_current();
var employeeList = clientContext.get_web().get_lists().getByTitle('New Employees In Seattle');;
var completedItems;

function purgeCompletedItems() {

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
        '<View><Query><Where><Eq>' +
          '<FieldRef Name=\'OrientationStage\'/><Value Type=\'Choice\'>Completed</Value>' +
        '</Eq></Where></Query></View>');
    completedItems = employeeList.getItems(camlQuery);
    clientContext.load(completedItems);

    clientContext.executeQueryAsync(deleteCompletedItems, onGetCompletedItemsFail);
    return false;
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
        employeeList.getItemById(itemArray[i].get_id()).deleteObject();
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


/*
SharePoint-hosted SharePoint Add-in Tutorials, https://github.com/OfficeDev/SharePoint_SP-hosted_Add-ins_Tutorials
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
*/
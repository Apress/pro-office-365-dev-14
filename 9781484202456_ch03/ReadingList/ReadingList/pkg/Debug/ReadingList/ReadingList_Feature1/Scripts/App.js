$(document).ready(function () {
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', retrieveBooks);
    SP.SOD.executeOrDelayUntilScriptLoaded(ModifyRibbon, 'sp.ribbon.js');
});

function retrieveBooks() {
    var context = new SP.ClientContext.get_current();
    var oList = context.get_web().get_lists().getByTitle('Books');

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><OrderBy><FieldRef Name=\'Order1\' ' +
        'Ascending=\'TRUE\' /></OrderBy></Query><ViewFields><FieldRef Name=\'Id\' ' +
        '/><FieldRef Name=\'Title\' /><FieldRef Name=\'URL\' ' +
        '/><FieldRef Name=\'Background\' /><FieldRef Name=\'Order1\' ' +
        '/></ViewFields></View>');
    this.collListItem = oList.getItems(camlQuery);

    context.load(collListItem, 'Include(Id, Title, URL, Background, Order1)');

    context.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded),
                              Function.createDelegate(this, this.onQueryFailed));
}

function onQueryFailed(sender, args) {
    SP.UI.Notify.addNotification('Request failed. ' + args.get_message() + '\n' +
                                                      args.get_stackTrace(), true);
}

function onQuerySucceeded(sender, args) {

    var listEnumerator = collListItem.getEnumerator();
    var listInfo = "";

    while (listEnumerator.moveNext()) {
        var oListItem = listEnumerator.get_current();
        listInfo +=
            "<div id='" + oListItem.get_id() + "' class='Book " +
                          oListItem.get_item('Background') + "'>" +
                "<a href='" + oListItem.get_item('URL').get_url() +
                     "' target='_blank'>" +
                    "<div class='BookDescription'>" + oListItem.get_item('Title') +
                    "</div>" +
                "</a>" +
                "<div class='EditIcon'>" +
                    "<a href='#' onclick='ShowDialog(" + oListItem.get_id() +
                        ")'><img src='../Images/EditIcon.png' /></a>" +
                "</div>" +
            "</div>";
    }

    $("#results").html(listInfo);
}

function ShowDialog(ID) {

    var options = {
        url: "../Lists/Books/EditForm.aspx?ID=" + ID,
        allowMaximize: true,
        title: "Edit Book",
        dialogReturnValueCallback: scallback
    };
    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
    return false;
}

function scallback(dialogResult, returnValue) {
    if (dialogResult == SP.UI.DialogResult.OK) {
        SP.UI.ModalDialog.RefreshPage(SP.UI.DialogResult.OK);
    }
}

// Methods for the ribbon
function ModifyRibbon() {

    var pm = SP.Ribbon.PageManager.get_instance();

    pm.add_ribbonInited(function () {
        AddBookTab();
    });

    var ribbon = null;
    try {
        ribbon = pm.get_ribbon();
    }
    catch (e) { }

    if (!ribbon) {
        if (typeof (_ribbonStartInit) == "function")
            _ribbonStartInit(_ribbon.initialTabId, false, null);
    }
    else {
        AddBookTab();
    }
}

function AddBookTab() {
    var sTitleHtml = "";
    var sManageHtml = "";

    sTitleHtml += "<a href='../Lists/Books/Title%20List.aspx' >' ";
    sTitleHtml += "<img src='../images/ViewIcon.png' /></a><br/>Title List";
    sManageHtml += "<a href='../Lists/Books/AllItems.aspx' >";
    sManageHtml += "<img src='../images/ViewIcon.png' /></a><br/>Manage Books";

    var ribbon = SP.Ribbon.PageManager.get_instance().get_ribbon();
    if (ribbon !== null) {
        var tab = new CUI.Tab(ribbon, 'Books.Tab', 'Books',
            'Use this tab to view and modify the book list',
            'Books.Tab.Command', false, '', null);
        ribbon.addChildAtIndex(tab, 1);
        var group = new CUI.Group(ribbon, 'Books.Tab.Group', 'Views',
            'Use this group to view a list of titles',
            'Books.Group.Command', null);
        tab.addChild(group);
        var group = new CUI.Group(ribbon, 'Books.Tab.Group', 'Actions',
            'Use this group to add/update/delete books',
            'Books.Group.Command', null);
        tab.addChild(group);
    }
    SelectRibbonTab('Books.Tab', true);
    $("span:contains('Views')").prev("span").html(sTitleHtml);
    $("span:contains('Actions')").prev("span").html(sManageHtml);
    SelectRibbonTab('Ribbon.Read', true);
}


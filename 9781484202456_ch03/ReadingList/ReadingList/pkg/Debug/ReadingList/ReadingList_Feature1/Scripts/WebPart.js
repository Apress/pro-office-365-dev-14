
$(document).ready(function () {
    var spHostUrl = decodeURIComponent(getQueryStringParameter('SPHostUrl'));
    var layoutsRoot = spHostUrl + '/_layouts/15/';

    $.getScript(layoutsRoot + "SP.Runtime.js", function () {
        $.getScript(layoutsRoot + "SP.js", formatWebPart);
    }
    );
});

function formatWebPart() {
    context = new SP.ClientContext.get_current();
    web = context.get_web();

    retrieveBooks();
}

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

    context.load(this.collListItem, 'Include(Id, Title, URL, Background, Order1)');

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

    if (listEnumerator.moveNext()) {
        var oListItem = listEnumerator.get_current();
        listInfo +=
            "<div id='" + oListItem.get_id() + "' class='Book " +
                          oListItem.get_item('Background') + "'>" +
                "<a href='" + oListItem.get_item('URL').get_url() +
                   "' target='_blank'>" +
                    "<div class='BookDescription'>" + oListItem.get_item('Title') +
                    "</div>" +
                "</a>" +
            "</div>";
    }

    $("#bookList").html(listInfo);
}

function getQueryStringParameter(urlParameterKey) {
    var params = document.URL.split('?')[1].split('&');
    var strParams = '';
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split('=');
        if (singleParam[0] == urlParameterKey)
            return decodeURIComponent(singleParam[1]);
    }
}




'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage() {
    var linkItems;
    var links = [];
    var lists;
    function getLinks() {
        var clientContext = new SP.ClientContext.get_current();
        var hostWebURL = decodeURIComponent("https://myefdomain.sharepoint.com/sites/replylabtask4");

        var hostWebContext = new SP.AppContextSite(clientContext, hostWebURL);
        var groupLists = hostWebContext.get_web().get_lists().getByTitle("list of links");
        /*lists = clientContext.get_web().get_lists();

        clientContext.load(lists);
        clientContext.executeQueryAsync(onListsQuerySucceeded, onGetItemsFail);*/

        var camlQuery = new SP.CamlQuery();
        linkItems = groupLists.getItems(camlQuery);

        clientContext.load(linkItems);
        clientContext.executeQueryAsync(getLinkItem, onGetItemsFail);
    }

    function onListsQuerySucceeded() {
        debugger;
        var listEnumerator = lists.getEnumerator();
        while (listEnumerator.moveNext()) {
            var oList = listEnumerator.get_current();
            //getLinkItem(oList);            
            console.log('Title: ' + oList.get_title() + oList.get_item());
        }
        //renderLinks();
    }

    function getLinkItem() {
        debugger;
        var listItemEnumerator = linkItems.getEnumerator();
        while (listItemEnumerator.moveNext()) {
            var link = listItemEnumerator.get_current();
            var title = link.get_item('Title');
            var description = link.get_item('Description');
            var URL = link.get_item('URL');
            var group = link.get_item('Group');
            links.push({
                description,
                title,
                URL: URL.$2_1,
                name: URL.$1_1,
                group: group
            });
        }
        if (links.length !== 0) {
            var linksContainer = document.getElementById('linksList');
            links.forEach(link => {
                var linkContainer = document.createElement('li');
                linkContainer.innerHTML = `
                    <h2>Group: ${link.group}</h2>
                    <a href="${link.URL}" title="${link.title}" target="_blank">${link.name}</a>
                    <p>Description: ${link.description}</p>
                    `;
                linksContainer.appendChild(linkContainer);
            });
        }
    }

    function onGetItemsFail(sender, args) {
        alert(args.get_message());
    }

    $(document).ready(function () {
        getLinks();
    });
}
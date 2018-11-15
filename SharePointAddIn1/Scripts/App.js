'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage() {
    var linkItems;
    var links = [];
    var lists;
    var result = [];


    function getData() {
        var clientContext = new SP.ClientContext.get_current();
        var hostWebURL = decodeURIComponent("https://myefdomain.sharepoint.com/sites/replylabtask4");

        var hostWebContext = new SP.AppContextSite(clientContext, hostWebURL);
        var listOfLinks = hostWebContext.get_web().get_lists().getByTitle("list of links");
       /* var ctx = new SP.ClientContext.get_current();
        lists = ctx.get_web().get_lists();
        ctx.load(lists, "Include(Id,Title)");
        ctx.executeQueryAsync(
            function () {
                debugger;
                lists.get_data().forEach(function (list) {
                    var items = list.getItems(SP.CamlQuery.createAllItemsQuery());
                    ctx.load(items, "Include(Id,FileRef)");
                    var listEntry = {
                        id: list.get_id().toString(),
                        title: list.get_title()
                    }
                    result.push({
                        list: listEntry,
                        items: items
                    });
                });
                ctx.executeQueryAsync(
                    function () {
                        //transform listitem properties
                        result.forEach(function (item) {

                            item.items = item.items.get_data().map(function (listItem) {
                                return listItem.get_fieldValues();
                            });
                        });

                        console.log(JSON.stringify(result));
                    }, logError);


            }, logError);
            */
        var camlQuery = new SP.CamlQuery();
        linkItems = listOfLinks.getItems(camlQuery);

        clientContext.load(linkItems);
        clientContext.executeQueryAsync(getLinkItem, onGetItemsFail);
    }

    function logError(sender, args) {
        console.log(args.get_message());
    } 

    function getLinkItem() {
        debugger;
        var listItemEnumerator = linkItems.getEnumerator();
        while (listItemEnumerator.moveNext()) {
            var link = listItemEnumerator.get_current();            
            links.push(getLinkData(link));
        }
        renderLinks();
    }

    function getLinkData(link) {        
        var title = link.get_item('Title');
        var description = link.get_item('Description');
        var URL = link.get_item('URL');
        var group = link.get_item('Group');
        return {
            description,
            title,
            URL: URL.$2_1,
            name: URL.$1_1,
            group: group
        };
    }

    function renderLinks() {
        if (links.length !== 0) {
            var linksContainer = document.getElementById('linksList');
            links.forEach(link => {
                linksContainer.appendChild(createLink(link));
            });
        }
    }

    function createLink(link) {
        var linkContainer = document.createElement('li');
        linkContainer.innerHTML = `
                    <h2>Group: ${link.group !== undefined ? link.group : ""}</h2>
                    <a
                        href="${link.URL !== undefined ? link.URL : ""}" 
                        title="${link.title !== undefined ? link.title : ""}" 
                        target="_blank">${link.name !== undefined ? link.name : ""}
                    </a>
                    <p>Description: ${link.description !== undefined ? link.description : ""}</p>
                    `;
        return linkContainer;
    }

    function onGetItemsFail(sender, args) {
        alert(args.get_message());
    }

    $(document).ready(function () {
        getData();
    });
}